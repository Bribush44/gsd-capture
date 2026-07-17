"use strict";

import { createHash, timingSafeEqual } from "node:crypto";

import {
  createVoiceReviewId,
  getVoiceReviewIndexKey,
  getVoiceReviewItemKey,
} from "../lib/voice-user-store.js";

const CONNECTION_KEY = "gsd:siri:microsoft:connection"; const USER_CONNECTION_PREFIX = "gsd:siri:user:";
const REQUEST_PREFIX = "gsd:siri:request:";
const RATE_PREFIX = "gsd:siri:rate:";
const REVIEW_ITEM_TTL_SECONDS = 15552000;
const MICROSOFT_AUTHORITY = "https://login.microsoftonline.com/common";
const MICROSOFT_REVIEW_LIST_NAME = "GSD Review";
const MICROSOFT_SCOPES = [
  "openid",
  "profile",
  "offline_access",
  "User.Read",
  "Tasks.ReadWrite",
  "Calendars.ReadWrite",
  "MailboxSettings.Read",
];

export default async function handler(request, response) {
  response.setHeader("Cache-Control", "no-store");

  if (request.method !== "POST") {
    response.setHeader("Allow", "POST");
    return response.status(405).json({
      error: "Only POST requests are allowed.",
message: "Capture is unavailable from this request. Please try again.",
    });
  }

  try {
    requireEnvironment([
      "SIRI_CAPTURE_SECRET",
      "MICROSOFT_CLIENT_ID",
      "MICROSOFT_CLIENT_SECRET",
      "KV_REST_API_URL",
      "KV_REST_API_TOKEN",
      "OPENAI_API_KEY",
    ]);

    const suppliedSecret =
      readSuppliedSecret(request);

    const captureIdentity =
      await resolveCaptureIdentity(
        suppliedSecret
      );

    if (!captureIdentity) {
      return response.status(401).json({
        error:
          "The Siri capture key is invalid.",
        message:
          "Capture needs to be set up again. Please open GSD Capture.",
      });
    }

    const body = parseRequestBody(request.body);
    const text = typeof body.text === "string" ? body.text.trim() : "";
    const requestId = sanitizeRequestId(body.requestId);
    const currentDate = sanitizeCurrentDate(body.currentDate);

    if (!text) {
      return response.status(400).json({
        error: "Task text is required.",
message: "I did not hear a task. Please try Capture again.",
      });
    }

    if (text.length > 1000) {
      return response.status(400).json({
        error: "Task text must be 1,000 characters or fewer.",
message: "That task was too long. Please try a shorter capture.",
      });
    }

    await enforceRateLimit(request, captureIdentity.id);

    if (requestId) {
      const existingResult = await redisCommand([
        "GET",
        REQUEST_PREFIX + captureIdentity.id + ":" + requestId,
      ]);

      if (existingResult) {
        const parsedResult = JSON.parse(existingResult);
        parsedResult.duplicateRequest = true;
        return response.status(200).json(parsedResult);
      }
    }

    const connection = await readConnection(captureIdentity);

    if (!connection || !connection.refreshToken) {
      return response.status(409).json({
        error: "Siri background capture is not connected to Microsoft To Do.",
message: "Microsoft To Do needs to be reconnected. Open GSD Capture and reconnect Siri capture.",
        connectUrl: "https://gsd-capture.vercel.app/api/microsoft-connect",
      });
    }

    const tokenResult = await refreshMicrosoftAccessToken(
      connection.refreshToken
    );

    if (tokenResult.refresh_token) {
      connection.refreshToken = tokenResult.refresh_token;
      connection.refreshedAt = new Date().toISOString();
      await saveConnection(connection, captureIdentity);
    }

    const classification = await classifyTask(text, currentDate);
    const accessToken = tokenResult.access_token;
    const reviewId = createVoiceReviewId();

    let microsoftTimeZone = "";
    let schedulingConnectionError = "";

    if (hasSchedulingIntent(classification)) {
      try {
        microsoftTimeZone = await readMicrosoftTimeZone(accessToken);
      } catch (timeZoneError) {
        schedulingConnectionError =
          timeZoneError && timeZoneError.message
            ? timeZoneError.message
            : "Calendar and reminder access needs to be reconnected.";
      }
    }

    const reminderStatus = getInitialReminderStatus(
      classification,
      microsoftTimeZone
    );
    const destinationList = await resolveDestinationList(
      classification,
      accessToken
    );
    const taskPayload = buildMicrosoftTaskPayload(
      text,
      classification,
      requestId,
      reviewId,
      microsoftTimeZone,
      reminderStatus
    );
    const microsoftTask = await createMicrosoftTask(
      destinationList.id,
      taskPayload,
      accessToken
    );

    let calendarStatus = getInitialCalendarStatus(
      classification,
      microsoftTimeZone
    );
    let calendarEvent = null;
    let calendarError = "";

    if (calendarStatus === "ready") {
      try {
        calendarEvent = await createMicrosoftCalendarEvent(
          classification,
          microsoftTimeZone,
          reviewId,
          text,
          accessToken
        );
        calendarStatus = "created";
      } catch (calendarCreateError) {
        calendarStatus = "failed";
        calendarError =
          calendarCreateError && calendarCreateError.message
            ? calendarCreateError.message
            : "The calendar event could not be created.";
      }
    }

    const createdAt = new Date().toISOString();

    const reviewItem = {
      id: reviewId,
      source: "voice",
      status: "pending",
      originalText: text,
      title: classification.summary,
      suggestedCategory: classification.category,
      category: classification.category,
      priority: classification.priority,
      dueDate: classification.dueDate,
      context: classification.context,
      project: classification.project,
      estimatedMinutes: classification.estimatedMinutes,
      confidence: classification.confidence,
      explanation: classification.explanation,
      sortingMethod: classification.sortingMethod,
      schedulingExplanation: classification.schedulingExplanation,
      calendarIntent: classification.calendarIntent,
      calendarTitle: classification.calendarTitle,
      calendarStartDateTime: classification.calendarStartDateTime,
      calendarEndDateTime: classification.calendarEndDateTime,
      calendarLocation: classification.calendarLocation,
      calendarStatus,
      calendarEventId: calendarEvent && calendarEvent.id
        ? calendarEvent.id
        : "",
      calendarWebLink: calendarEvent && calendarEvent.webLink
        ? calendarEvent.webLink
        : "",
      calendarError,
      reminderIntent: classification.reminderIntent,
      reminderDateTime: classification.reminderDateTime,
      reminderStatus,
      reminderError:
        reminderStatus === "needs-reconnect"
          ? schedulingConnectionError
          : "",
      microsoftTimeZone,
      schedulingConnectionError,
      needsReview: true,
      requestId: requestId || "",
      createdAt,
      microsoftTaskId: microsoftTask.id,
      microsoftReviewListId: destinationList.id,
      microsoftReviewListName:
        destinationList.displayName || MICROSOFT_REVIEW_LIST_NAME,
    };

    await saveVoiceReviewItem(
      captureIdentity,
      reviewItem
    );

    const messageParts = [
      "Captured and added to GSD Review for review."
    ];

    if (calendarStatus === "created") {
      messageParts.push("Calendar event created.");
    } else if (calendarStatus === "needs-details") {
      messageParts.push("Calendar details need review in GSD Capture.");
    } else if (calendarStatus === "needs-reconnect") {
      messageParts.push("Reconnect Voice Capture to enable calendar access.");
    } else if (calendarStatus === "failed") {
      messageParts.push("The calendar event needs attention in GSD Capture.");
    }

    if (reminderStatus === "set") {
      messageParts.push("Reminder set.");
    } else if (reminderStatus === "needs-details") {
      messageParts.push("Reminder details need review in GSD Capture.");
    } else if (reminderStatus === "needs-reconnect") {
      messageParts.push("Reconnect Voice Capture to enable reminders.");
    }

    const result = {
      ok: true,
      reviewId,
      summary: classification.summary,
      category: classification.category,
      suggestedCategory: classification.category,
      confidence: classification.confidence,
      explanation: classification.explanation,
      priority: classification.priority,
      dueDate: classification.dueDate,
      needsReview: true,
      calendarIntent: classification.calendarIntent,
      calendarStatus,
      calendarEventId: reviewItem.calendarEventId,
      reminderIntent: classification.reminderIntent,
      reminderStatus,
      listName:
        destinationList.displayName || MICROSOFT_REVIEW_LIST_NAME,
      microsoftTaskId: microsoftTask.id,
      message: messageParts.join(" "),
    };

    if (requestId) {
      await redisCommand([
        "SET",
        REQUEST_PREFIX + captureIdentity.id + ":" + requestId,
        JSON.stringify(result),
        "EX",
        604800,
      ]);
    }

    return response.status(200).json(result);
  } catch (error) {
    console.error("Siri capture failed:", error);

    const technicalMessage =
      error && error.message ? error.message : "Capture failed.";

    const lowerTechnicalMessage = technicalMessage.toLowerCase();

    const reconnectRequired =
      lowerTechnicalMessage.includes("reconnect") ||
      lowerTechnicalMessage.includes("refresh token") ||
      lowerTechnicalMessage.includes("not connected") ||
      lowerTechnicalMessage.includes("interaction_required") ||
      lowerTechnicalMessage.includes("consent") ||
      lowerTechnicalMessage.includes("insufficient privileges");

    return response.status(reconnectRequired ? 409 : 500).json({
      error: technicalMessage,
      message: getFriendlyCaptureErrorMessage(technicalMessage),
      reconnectRequired: reconnectRequired,
      connectUrl: reconnectRequired
        ? "https://gsd-capture.vercel.app/api/microsoft-connect"
        : undefined,
    });
  }
}

function getFriendlyCaptureErrorMessage(message) {
  const lowerMessage = String(message || "").toLowerCase();

  if (
    lowerMessage.includes("reconnect") ||
    lowerMessage.includes("refresh token") ||
    lowerMessage.includes("not connected")
  ) {
    return "Microsoft To Do needs to be reconnected. Open GSD Capture and reconnect Siri capture.";
  }

  if (lowerMessage.includes("too many siri capture requests")) {
    return "There have been too many capture requests. Please wait a few minutes and try again.";
  }

  if (lowerMessage.includes("request id")) {
    return "I could not create a valid capture request. Please try again.";
  }

  if (
    lowerMessage.includes("microsoft") ||
    lowerMessage.includes("graph")
  ) {
    return "Microsoft To Do could not save the task. Please try again.";
  }

  if (
    lowerMessage.includes("secure storage") ||
    lowerMessage.includes("environment variable")
  ) {
    return "Capture is temporarily unavailable. Please try again in a few minutes.";
  }

  return "I could not save that task. Please try again.";
}

async function resolveCaptureIdentity(
  suppliedSecret
) {
  if (
    safeSecretMatch(
      suppliedSecret,
      process.env.SIRI_CAPTURE_SECRET
    )
  ) {
    return {
      id: "legacy",
      voiceKeyHash: "legacy",
      connectionKey: CONNECTION_KEY,
      connection: null,
    };
  }

  const voiceKey =
    String(suppliedSecret || "").trim();

  if (
    !/^gsvc_[A-Za-z0-9_-]{20,}$/.test(
      voiceKey
    )
  ) {
    return null;
  }

  const voiceKeyHash =
    createHash("sha256")
      .update(voiceKey)
      .digest("hex");

  const connectionKey =
    USER_CONNECTION_PREFIX +
    voiceKeyHash +
    ":connection";

  const storedConnection =
    await redisCommand([
      "GET",
      connectionKey,
    ]);

  if (!storedConnection) {
    return null;
  }

  try {
    return {
      id: voiceKeyHash.slice(0, 24),
      voiceKeyHash,
      connectionKey,
      connection:
        JSON.parse(storedConnection),
    };
  } catch (error) {
    console.error(
      "Stored Voice Capture connection was invalid:",
      error
    );

    return null;
  }
}

function parseRequestBody(value) {
  if (!value) {
    return {};
  }

  if (typeof value === "string") {
    return JSON.parse(value);
  }

  return value;
}

function readSuppliedSecret(request) {
  const headerValue =
    request.headers["x-gsd-capture-key"] ||
    request.headers["X-GSD-Capture-Key"] ||
    "";

  if (headerValue) {
    return String(headerValue);
  }

  const authorization = String(request.headers.authorization || "");
  return authorization.toLowerCase().startsWith("bearer ")
    ? authorization.slice(7).trim()
    : "";
}

function safeSecretMatch(supplied, expected) {
  const suppliedBuffer = Buffer.from(String(supplied || ""));
  const expectedBuffer = Buffer.from(String(expected || ""));

  if (
    suppliedBuffer.length === 0 ||
    suppliedBuffer.length !== expectedBuffer.length
  ) {
    return false;
  }

  return timingSafeEqual(suppliedBuffer, expectedBuffer);
}

function sanitizeRequestId(value) {
  const requestId = typeof value === "string" ? value.trim() : "";

  if (!requestId) {
    return "";
  }

  if (!/^[A-Za-z0-9._:-]{1,160}$/.test(requestId)) {
    throw new Error("The request ID is invalid.");
  }

  return requestId;
}

function sanitizeCurrentDate(value) {
  const currentDate = typeof value === "string" ? value.trim() : "";
  return /^\d{4}-\d{2}-\d{2}$/.test(currentDate)
    ? currentDate
    : new Date().toISOString().slice(0, 10);
}

async function enforceRateLimit(request, identityId) {
  const forwardedFor = String(request.headers["x-forwarded-for"] || "unknown");
  const ip = forwardedFor.split(",")[0].trim().replace(/[^A-Za-z0-9:._-]/g, "");
  const hour = new Date().toISOString().slice(0, 13);
  const key = RATE_PREFIX + identityId + ":" + ip + ":" + hour;
  const count = Number(await redisCommand(["INCR", key]));

  if (count === 1) {
    await redisCommand(["EXPIRE", key, 3700]);
  }

  if (count > 120) {
    throw new Error("Too many Siri capture requests. Try again later.");
  }
}

async function readConnection(
  captureIdentity
) {
  if (captureIdentity.connection) {
    return captureIdentity.connection;
  }

  const storedConnection =
    await redisCommand([
      "GET",
      captureIdentity.connectionKey,
    ]);

  return storedConnection
    ? JSON.parse(storedConnection)
    : null;
}

async function saveConnection(
  connection,
  captureIdentity
) {
  captureIdentity.connection = connection;

  await redisCommand([
    "SET",
    captureIdentity.connectionKey,
    JSON.stringify(connection),
  ]);
}

async function refreshMicrosoftAccessToken(refreshToken) {
  const tokenResponse = await fetch(
    MICROSOFT_AUTHORITY + "/oauth2/v2.0/token",
    {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: new URLSearchParams({
        client_id: process.env.MICROSOFT_CLIENT_ID,
        client_secret: process.env.MICROSOFT_CLIENT_SECRET,
        grant_type: "refresh_token",
        refresh_token: refreshToken,
        scope: MICROSOFT_SCOPES.join(" "),
      }).toString(),
    }
  );

  const tokenData = await tokenResponse.json().catch(function () {
    return null;
  });

  if (!tokenResponse.ok || !tokenData || !tokenData.access_token) {
    const message =
      tokenData && tokenData.error_description
        ? tokenData.error_description
        : "The Microsoft refresh token is no longer valid. Please reconnect Siri capture.";
    throw new Error(message);
  }

  return tokenData;
}

async function classifyTask(text, currentDate) {
  try {
    const openAIResponse = await fetch("https://api.openai.com/v1/responses", {
      method: "POST",
      headers: {
        Authorization: "Bearer " + process.env.OPENAI_API_KEY,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model: "gpt-5.6-luna",
        store: false,
        reasoning: { effort: "none" },
        max_output_tokens: 550,
        input: [
          {
            role: "system",
            content: [
              {
                type: "input_text",
                text:
                  "You organize captured tasks using GTD principles. " +
                  "Return a practical classification without inventing details. " +
                  "Today's date is " +
                  currentDate +
                  ". Convert clear relative dates such as today or tomorrow into YYYY-MM-DD. " +
                  "Use null when a date, project, or duration cannot reasonably be determined. " +
                  "Also return a confidence score from 0 to 100 for the suggested category and a short explanation a user can review. " +
                  "Detect calendar and reminder intent carefully. Use calendarIntent explicit only when the user directly asks to add, create, book, schedule, or put something on a calendar. Use suggested when a calendar event would clearly help but the user did not directly ask for one. Use reminderIntent explicit only when the user directly asks to be reminded, alerted, or notified. Use suggested when a reminder would clearly help but was not directly requested. Use none otherwise. Never invent a date or time. Return local date-times as YYYY-MM-DDTHH:mm:ss without an offset. When a calendar start is known but no end is stated, use a reasonable 30-minute end time. Use null for missing scheduling details.",
              },
            ],
          },
          {
            role: "user",
            content: [{ type: "input_text", text: text }],
          },
        ],
        text: {
          format: {
            type: "json_schema",
            name: "gsd_siri_task_classification",
            strict: true,
            schema: {
              type: "object",
              properties: {
                summary: { type: "string" },
                category: {
                  type: "string",
                  enum: [
                    "Purchasing",
                    "Operations",
                    "Leadership",
                    "Personal",
                    "Ideas",
                    "General",
                  ],
                },
                priority: {
                  type: "string",
                  enum: ["High", "Normal", "Low"],
                },
                dueDate: { type: ["string", "null"] },
                context: {
                  type: "string",
                  enum: [
                    "Calls",
                    "Computer",
                    "Errands",
                    "Work",
                    "Home",
                    "Anywhere",
                  ],
                },
                project: { type: ["string", "null"] },
                estimatedMinutes: {
                  type: ["integer", "null"],
                  minimum: 1,
                  maximum: 480,
                },
                needsReview: { type: "boolean" },
                confidence: {
                  type: "integer",
                  minimum: 0,
                  maximum: 100,
                },
                explanation: { type: "string" },
                calendarIntent: {
                  type: "string",
                  enum: ["none", "explicit", "suggested"],
                },
                calendarTitle: { type: ["string", "null"] },
                calendarStartDateTime: { type: ["string", "null"] },
                calendarEndDateTime: { type: ["string", "null"] },
                calendarLocation: { type: ["string", "null"] },
                reminderIntent: {
                  type: "string",
                  enum: ["none", "explicit", "suggested"],
                },
                reminderDateTime: { type: ["string", "null"] },
                schedulingExplanation: { type: "string" },
              },
              required: [
                "summary",
                "category",
                "priority",
                "dueDate",
                "context",
                "project",
                "estimatedMinutes",
                "needsReview",
                "confidence",
                "explanation",
                "calendarIntent",
                "calendarTitle",
                "calendarStartDateTime",
                "calendarEndDateTime",
                "calendarLocation",
                "reminderIntent",
                "reminderDateTime",
                "schedulingExplanation",
              ],
              additionalProperties: false,
            },
          },
        },
      }),
    });

    const openAIData = await openAIResponse.json().catch(function () {
      return null;
    });

    if (!openAIResponse.ok || !openAIData) {
      throw new Error("AI sorting failed.");
    }

    const outputText =
      openAIData.output_text ||
      openAIData.output
        ?.flatMap(function (item) {
          return Array.isArray(item.content) ? item.content : [];
        })
        .find(function (content) {
          return content.type === "output_text";
        })?.text;

    if (!outputText) {
      throw new Error("AI sorting returned no result.");
    }

    return normalizeClassification(JSON.parse(outputText), text);
  } catch (error) {
    console.warn("Siri AI sorting failed; using safe local fallback:", error);
    return localFallbackClassification(text, currentDate);
  }
}

function normalizeClassification(value, originalText) {
  const allowedCategories = [
    "Purchasing",
    "Operations",
    "Leadership",
    "Personal",
    "Ideas",
    "General",
  ];
  const allowedPriorities = ["High", "Normal", "Low"];
  const allowedContexts = [
    "Calls",
    "Computer",
    "Errands",
    "Work",
    "Home",
    "Anywhere",
  ];

  return {
    summary:
      typeof value.summary === "string" && value.summary.trim()
        ? value.summary.trim().slice(0, 300)
        : originalText.slice(0, 300),
    category: allowedCategories.includes(value.category)
      ? value.category
      : "General",
    priority: allowedPriorities.includes(value.priority)
      ? value.priority
      : "Normal",
    dueDate:
      typeof value.dueDate === "string" &&
      /^\d{4}-\d{2}-\d{2}$/.test(value.dueDate)
        ? value.dueDate
        : null,
    context: allowedContexts.includes(value.context)
      ? value.context
      : "Anywhere",
    project:
      typeof value.project === "string" && value.project.trim()
        ? value.project.trim().slice(0, 200)
        : null,
    estimatedMinutes:
      Number.isInteger(value.estimatedMinutes) &&
      value.estimatedMinutes >= 1 &&
      value.estimatedMinutes <= 480
        ? value.estimatedMinutes
        : null,
    needsReview: true,
    confidence:
      Number.isInteger(value.confidence) &&
      value.confidence >= 0 &&
      value.confidence <= 100
        ? value.confidence
        : 50,
    explanation:
      typeof value.explanation === "string" &&
      value.explanation.trim()
        ? value.explanation.trim().slice(0, 300)
        : "AI suggested this folder from the words in the capture.",
    calendarIntent: normalizeSchedulingIntent(value.calendarIntent),
    calendarTitle: normalizeOptionalText(value.calendarTitle, 300),
    calendarStartDateTime: normalizeLocalDateTime(
      value.calendarStartDateTime
    ),
    calendarEndDateTime: normalizeLocalDateTime(
      value.calendarEndDateTime
    ),
    calendarLocation: normalizeOptionalText(
      value.calendarLocation,
      300
    ),
    reminderIntent: normalizeSchedulingIntent(value.reminderIntent),
    reminderDateTime: normalizeLocalDateTime(value.reminderDateTime),
    schedulingExplanation:
      typeof value.schedulingExplanation === "string" &&
      value.schedulingExplanation.trim()
        ? value.schedulingExplanation.trim().slice(0, 400)
        : "No additional scheduling suggestion was made.",
    sortingMethod: "ai",
  };
}

function localFallbackClassification(text, currentDate) {
  const lowerText = text.toLowerCase();
  let category = "General";
  let priority = "Normal";
  let context = "Work";
  let dueDate = null;

  if (/supplier|vendor|quote|purchase/.test(lowerText)) {
    category = "Purchasing";
  } else if (/warehouse|inventory|shipment|receiving|production/.test(lowerText)) {
    category = "Operations";
  } else if (/team|employee|meeting|coach|leader/.test(lowerText)) {
    category = "Leadership";
  } else if (/home|family|personal/.test(lowerText)) {
    category = "Personal";
  } else if (/idea|maybe|someday/.test(lowerText)) {
    category = "Ideas";
  }

  if (/important|urgent|asap|high priority/.test(lowerText)) {
    priority = "High";
  }

  if (/call|phone/.test(lowerText)) {
    context = "Calls";
  } else if (/computer|email/.test(lowerText)) {
    context = "Computer";
  } else if (/pick up|store|errand/.test(lowerText)) {
    context = "Errands";
  } else if (category === "Personal") {
    context = "Home";
  }

  if (lowerText.includes("today")) {
    dueDate = currentDate;
  } else if (lowerText.includes("tomorrow")) {
    const date = new Date(currentDate + "T12:00:00Z");
    date.setUTCDate(date.getUTCDate() + 1);
    dueDate = date.toISOString().slice(0, 10);
  }

  return {
    summary: text.slice(0, 300),
    category: category,
    priority: priority,
    dueDate: dueDate,
    context: context,
    project: null,
    estimatedMinutes: null,
    needsReview: true,
    confidence: 35,
    explanation:
      "Local fallback suggested this folder because AI sorting was unavailable.",
    calendarIntent: detectLocalCalendarIntent(lowerText),
    calendarTitle: null,
    calendarStartDateTime: null,
    calendarEndDateTime: null,
    calendarLocation: null,
    reminderIntent: detectLocalReminderIntent(lowerText),
    reminderDateTime: null,
    schedulingExplanation:
      "Scheduling details need review because AI scheduling was unavailable.",
    sortingMethod: "local",
  };
}

function normalizeSchedulingIntent(value) {
  return ["none", "explicit", "suggested"].includes(value)
    ? value
    : "none";
}

function normalizeOptionalText(value, maximumLength) {
  return typeof value === "string" && value.trim()
    ? value.trim().slice(0, maximumLength)
    : null;
}

function normalizeLocalDateTime(value) {
  if (typeof value !== "string") {
    return null;
  }

  const trimmed = value.trim();
  const match = trimmed.match(
    /^(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2})(?::(\d{2}))?$/
  );

  if (!match) {
    return null;
  }

  const normalized =
    match[1] + "T" + match[2] + ":" + (match[3] || "00");
  const testDate = new Date(normalized + "Z");

  return Number.isNaN(testDate.getTime())
    ? null
    : normalized;
}

function detectLocalCalendarIntent(lowerText) {
  return /(add|put|create|book|schedule).{0,40}(calendar|appointment|meeting)|(calendar).{0,40}(add|put|create|book|schedule)/.test(
    lowerText
  )
    ? "explicit"
    : "none";
}

function detectLocalReminderIntent(lowerText) {
  return /\bremind me\b|\bset (a |an )?reminder\b|\balert me\b|\bnotify me\b/.test(
    lowerText
  )
    ? "explicit"
    : "none";
}

function hasSchedulingIntent(classification) {
  return (
    classification.calendarIntent !== "none" ||
    classification.reminderIntent !== "none"
  );
}

function getInitialCalendarStatus(classification, microsoftTimeZone) {
  if (classification.calendarIntent === "none") {
    return "none";
  }

  if (classification.calendarIntent === "suggested") {
    return "suggested";
  }

  if (
    !classification.calendarStartDateTime ||
    !classification.calendarEndDateTime
  ) {
    return "needs-details";
  }

  return microsoftTimeZone
    ? "ready"
    : "needs-reconnect";
}

function getInitialReminderStatus(classification, microsoftTimeZone) {
  if (classification.reminderIntent === "none") {
    return "none";
  }

  if (classification.reminderIntent === "suggested") {
    return "suggested";
  }

  if (!classification.reminderDateTime) {
    return "needs-details";
  }

  return microsoftTimeZone
    ? "set"
    : "needs-reconnect";
}

async function readMicrosoftTimeZone(accessToken) {
  const graphResponse = await fetch(
    "https://graph.microsoft.com/v1.0/me/mailboxSettings?$select=timeZone",
    {
      headers: {
        Authorization: "Bearer " + accessToken,
        Accept: "application/json",
      },
    }
  );

  const data = await graphResponse.json().catch(function () {
    return null;
  });

  if (!graphResponse.ok) {
    throw new Error(readGraphError(data, graphResponse.status));
  }

  const timeZone =
    data && typeof data.timeZone === "string"
      ? data.timeZone.trim()
      : "";

  if (!timeZone) {
    throw new Error(
      "Microsoft did not return a calendar time zone. Open Outlook calendar settings and choose a time zone."
    );
  }

  return timeZone.slice(0, 100);
}

async function resolveDestinationList(classification, accessToken) {
  const lists = await listMicrosoftTodoLists(accessToken);
  const desiredName = MICROSOFT_REVIEW_LIST_NAME;
  let matchingList = findListByName(lists, desiredName);

  if (!matchingList) {
    matchingList = await createMicrosoftTodoList(
      desiredName,
      accessToken
    );
  }

  if (!matchingList) {
    throw new Error(
      "The Microsoft To Do GSD Review list could not be found."
    );
  }

  return matchingList;
}

async function listMicrosoftTodoLists(accessToken) {
  const graphResponse = await fetch(
    "https://graph.microsoft.com/v1.0/me/todo/lists",
    {
      headers: {
        Authorization: "Bearer " + accessToken,
        Accept: "application/json",
      },
    }
  );

  const data = await graphResponse.json().catch(function () {
    return null;
  });

  if (!graphResponse.ok) {
    throw new Error(readGraphError(data, graphResponse.status));
  }

  return data && Array.isArray(data.value) ? data.value : [];
}

function findListByName(lists, name) {
  const normalizedName = String(name).trim().toLocaleLowerCase();
  return lists.find(function (list) {
    return (
      String(list.displayName || "").trim().toLocaleLowerCase() ===
      normalizedName
    );
  });
}

async function createMicrosoftTodoList(displayName, accessToken) {
  const graphResponse = await fetch(
    "https://graph.microsoft.com/v1.0/me/todo/lists",
    {
      method: "POST",
      headers: {
        Authorization: "Bearer " + accessToken,
        Accept: "application/json",
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ displayName: displayName.slice(0, 80) }),
    }
  );

  const data = await graphResponse.json().catch(function () {
    return null;
  });

  if (!graphResponse.ok || !data || !data.id) {
    throw new Error(readGraphError(data, graphResponse.status));
  }

  return data;
}

function buildMicrosoftTaskPayload(
  originalText,
  classification,
  requestId,
  reviewId,
  microsoftTimeZone,
  reminderStatus
) {
  const notes = [];

  if (originalText !== classification.summary) {
    notes.push("Original capture: " + originalText);
  }

  notes.push("Suggested category: " + classification.category);
  notes.push("AI confidence: " + classification.confidence + "%");
  notes.push("Suggestion reason: " + classification.explanation);
  notes.push("Priority: " + classification.priority);
  notes.push("Review status: Awaiting approval in GSD Capture");

  if (classification.context) {
    notes.push("Context: " + classification.context);
  }

  if (classification.project) {
    notes.push("Project: " + classification.project);
  }

  if (classification.estimatedMinutes) {
    notes.push(
      "Estimated duration: " + classification.estimatedMinutes + " minutes"
    );
  }

  notes.push(
    "Sorted by: " +
      (classification.sortingMethod === "ai"
        ? "GSD Capture AI"
        : "GSD Capture local fallback")
  );
  notes.push("Captured through: Siri Shortcut");

  if (classification.calendarIntent !== "none") {
    notes.push(
      "Calendar: " +
        classification.calendarIntent +
        (classification.calendarStartDateTime
          ? " at " + classification.calendarStartDateTime
          : " — details needed")
    );
  }

  if (classification.reminderIntent !== "none") {
    notes.push(
      "Reminder: " +
        classification.reminderIntent +
        (classification.reminderDateTime
          ? " at " + classification.reminderDateTime
          : " — details needed")
    );
  }

  if (classification.schedulingExplanation) {
    notes.push(
      "Scheduling note: " + classification.schedulingExplanation
    );
  }

  if (reviewId) {
    notes.push("GSD Voice Review ID: " + reviewId);
  }

  if (requestId) {
    notes.push("GSD Siri Request ID: " + requestId);
  }

  const payload = {
    title: classification.summary,
    importance: getMicrosoftImportance(classification.priority),
    body: {
      contentType: "text",
      content: notes.join("\n"),
    },
  };

  if (classification.dueDate) {
    payload.dueDateTime = {
      dateTime: classification.dueDate + "T12:00:00",
      timeZone: microsoftTimeZone || "UTC",
    };
  }

  if (
    reminderStatus === "set" &&
    classification.reminderDateTime &&
    microsoftTimeZone
  ) {
    payload.isReminderOn = true;
    payload.reminderDateTime = {
      dateTime: classification.reminderDateTime,
      timeZone: microsoftTimeZone,
    };
  }

  return payload;
}

async function createMicrosoftCalendarEvent(
  classification,
  microsoftTimeZone,
  reviewId,
  originalText,
  accessToken
) {
  const payload = {
    subject:
      classification.calendarTitle ||
      classification.summary ||
      originalText,
    body: {
      contentType: "text",
      content:
        "Created by GSD Capture from voice input.\n\n" +
        "Original capture: " +
        originalText +
        "\nGSD Voice Review ID: " +
        reviewId,
    },
    start: {
      dateTime: classification.calendarStartDateTime,
      timeZone: microsoftTimeZone,
    },
    end: {
      dateTime: classification.calendarEndDateTime,
      timeZone: microsoftTimeZone,
    },
    transactionId: reviewId,
  };

  if (classification.calendarLocation) {
    payload.location = {
      displayName: classification.calendarLocation,
    };
  }

  const graphResponse = await fetch(
    "https://graph.microsoft.com/v1.0/me/events",
    {
      method: "POST",
      headers: {
        Authorization: "Bearer " + accessToken,
        Accept: "application/json",
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    }
  );

  const data = await graphResponse.json().catch(function () {
    return null;
  });

  if (!graphResponse.ok || !data || !data.id) {
    throw new Error(readGraphError(data, graphResponse.status));
  }

  return data;
}

function getMicrosoftImportance(priority) {
  if (priority === "High") {
    return "high";
  }

  if (priority === "Low") {
    return "low";
  }

  return "normal";
}

async function saveVoiceReviewItem(
  captureIdentity,
  reviewItem
) {
  const voiceKeyHash =
    captureIdentity.voiceKeyHash ||
    captureIdentity.id;
  const itemKey = getVoiceReviewItemKey(
    voiceKeyHash,
    reviewItem.id
  );
  const indexKey = getVoiceReviewIndexKey(
    voiceKeyHash
  );
  const createdScore = Date.parse(
    reviewItem.createdAt
  );

  await redisCommand([
    "SET",
    itemKey,
    JSON.stringify(reviewItem),
    "EX",
    REVIEW_ITEM_TTL_SECONDS,
  ]);

  await redisCommand([
    "ZADD",
    indexKey,
    String(
      Number.isFinite(createdScore)
        ? createdScore
        : Date.now()
    ),
    reviewItem.id,
  ]);

  await redisCommand([
    "EXPIRE",
    indexKey,
    REVIEW_ITEM_TTL_SECONDS,
  ]);
}

async function createMicrosoftTask(listId, payload, accessToken) {
  const graphResponse = await fetch(
    "https://graph.microsoft.com/v1.0/me/todo/lists/" +
      encodeURIComponent(listId) +
      "/tasks",
    {
      method: "POST",
      headers: {
        Authorization: "Bearer " + accessToken,
        Accept: "application/json",
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    }
  );

  const data = await graphResponse.json().catch(function () {
    return null;
  });

  if (!graphResponse.ok || !data || !data.id) {
    throw new Error(readGraphError(data, graphResponse.status));
  }

  return data;
}

function readGraphError(data, status) {
  return data && data.error && data.error.message
    ? data.error.message
    : "Microsoft Graph returned " + status + ".";
}

function requireEnvironment(names) {
  const missing = names.filter(function (name) {
    return !process.env[name];
  });

  if (missing.length > 0) {
    throw new Error("Missing environment variables: " + missing.join(", "));
  }
}

async function redisCommand(command) {
  const redisResponse = await fetch(process.env.KV_REST_API_URL, {
    method: "POST",
    headers: {
      Authorization: "Bearer " + process.env.KV_REST_API_TOKEN,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(command),
  });

  const data = await redisResponse.json().catch(function () {
    return null;
  });

  if (!redisResponse.ok || !data || data.error) {
    throw new Error(
      data && data.error ? data.error : "Secure storage request failed."
    );
  }

  return data.result;
}
