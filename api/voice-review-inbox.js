"use strict";

import {
  getUserConnectionKey,
  getVoiceKeyHash,
  getVoiceReviewIndexKey,
  getVoiceReviewItemKey,
  redisCommand,
  requireEnvironment,
  sanitizeVoiceReviewId,
} from "../lib/voice-user-store.js";

const REVIEW_ITEM_TTL_SECONDS = 15552000;
const MAX_REVIEW_ITEMS = 200;
const MICROSOFT_AUTHORITY =
  "https://login.microsoftonline.com/common";
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

  if (request.method !== "GET" && request.method !== "PATCH") {
    response.setHeader("Allow", "GET, PATCH");
    return response.status(405).json({
      error: "Only GET and PATCH requests are allowed.",
    });
  }

  try {
    requireEnvironment([
      "KV_REST_API_URL",
      "KV_REST_API_TOKEN",
    ]);

    const voiceKey = readVoiceKey(request);
    const voiceKeyHash = getVoiceKeyHash(voiceKey);

    if (!voiceKeyHash) {
      return response.status(401).json({
        error: "A valid Voice Capture key is required.",
      });
    }

    const connectionKey = getUserConnectionKey(voiceKeyHash);
    const storedConnection = await redisCommand([
      "GET",
      connectionKey,
    ]);

    if (!storedConnection) {
      return response.status(401).json({
        error: "Voice Capture needs to be set up again.",
      });
    }

    let connection = null;

    try {
      connection = JSON.parse(storedConnection);
    } catch (error) {
      return response.status(401).json({
        error: "Voice Capture needs to be set up again.",
      });
    }

    if (request.method === "GET") {
      return await listReviewItems(
        response,
        voiceKeyHash
      );
    }

    return await updateReviewItem(
      request,
      response,
      voiceKeyHash,
      connectionKey,
      connection
    );
  } catch (error) {
    console.error("Voice Review Inbox failed:", error);

    return response.status(500).json({
      error:
        error && error.message
          ? error.message
          : "The Voice Review Inbox could not be loaded.",
    });
  }
}

async function listReviewItems(
  response,
  voiceKeyHash
) {
  const indexKey = getVoiceReviewIndexKey(
    voiceKeyHash
  );

  const reviewIds = await redisCommand([
    "ZREVRANGE",
    indexKey,
    "0",
    String(MAX_REVIEW_ITEMS - 1),
  ]);

  if (!Array.isArray(reviewIds) || reviewIds.length === 0) {
    return response.status(200).json({
      ok: true,
      items: [],
    });
  }

  const itemKeys = reviewIds.map(function (reviewId) {
    return getVoiceReviewItemKey(
      voiceKeyHash,
      reviewId
    );
  });

  const storedItems = await redisCommand([
    "MGET",
    ...itemKeys,
  ]);

  const items = [];

  if (Array.isArray(storedItems)) {
    storedItems.forEach(function (storedItem) {
      if (!storedItem) {
        return;
      }

      try {
        const parsed = JSON.parse(storedItem);

        if (
          parsed &&
          parsed.status === "pending"
        ) {
          items.push(parsed);
        }
      } catch (error) {
        console.warn(
          "A stored Voice Review item was invalid:",
          error
        );
      }
    });
  }

  return response.status(200).json({
    ok: true,
    items,
  });
}

async function updateReviewItem(
  request,
  response,
  voiceKeyHash,
  connectionKey,
  connection
) {
  const body = parseRequestBody(request.body);
  const reviewId = sanitizeVoiceReviewId(
    body.reviewId
  );

  if (!reviewId) {
    return response.status(400).json({
      error: "A valid review ID is required.",
    });
  }

  const itemKey = getVoiceReviewItemKey(
    voiceKeyHash,
    reviewId
  );
  const storedItem = await redisCommand([
    "GET",
    itemKey,
  ]);

  if (!storedItem) {
    return response.status(404).json({
      error: "The Voice Review item was not found.",
    });
  }

  const item = JSON.parse(storedItem);
  const action = normalizeAction(body.action);

  if (action) {
    return await performSchedulingAction(
      response,
      body,
      action,
      item,
      itemKey,
      connectionKey,
      connection
    );
  }

  const allowedStatuses = [
    "approved",
    "kept",
    "rejected",
    "completed",
  ];
  const status = allowedStatuses.includes(
    body.status
  )
    ? body.status
    : "";

  if (!status) {
    return response.status(400).json({
      error:
        "A valid review status or scheduling action is required.",
    });
  }

  item.status = status;
  item.resolvedAt = new Date().toISOString();

  if (
    typeof body.finalCategory === "string" &&
    body.finalCategory.trim()
  ) {
    item.finalCategory = body.finalCategory
      .trim()
      .slice(0, 80);
  }

  if (
    typeof body.finalListName === "string" &&
    body.finalListName.trim()
  ) {
    item.finalListName = body.finalListName
      .trim()
      .slice(0, 80);
  }

  if (body.dismissPendingScheduling !== false) {
    dismissPendingScheduling(item);
  }

  await saveReviewItem(itemKey, item);

  await redisCommand([
    "ZREM",
    getVoiceReviewIndexKey(voiceKeyHash),
    reviewId,
  ]);

  return response.status(200).json({
    ok: true,
    item,
  });
}

async function performSchedulingAction(
  response,
  body,
  action,
  item,
  itemKey,
  connectionKey,
  connection
) {
  if (action === "dismiss-calendar") {
    item.calendarStatus = "dismissed";
    item.calendarError = "";
    item.calendarResolvedAt = new Date().toISOString();
    await saveReviewItem(itemKey, item);
    return response.status(200).json({ ok: true, item });
  }

  if (action === "dismiss-reminder") {
    item.reminderStatus = "dismissed";
    item.reminderError = "";
    item.reminderResolvedAt = new Date().toISOString();
    await saveReviewItem(itemKey, item);
    return response.status(200).json({ ok: true, item });
  }

  requireEnvironment([
    "MICROSOFT_CLIENT_ID",
    "MICROSOFT_CLIENT_SECRET",
  ]);

  if (!connection || !connection.refreshToken) {
    return response.status(409).json({
      error:
        "Voice Capture needs to be set up again for calendar and reminder access.",
      reconnectRequired: true,
    });
  }

  try {
    const tokenResult = await refreshMicrosoftAccessToken(
      connection.refreshToken
    );

    if (tokenResult.refresh_token) {
      connection.refreshToken = tokenResult.refresh_token;
      connection.refreshedAt = new Date().toISOString();
      await redisCommand([
        "SET",
        connectionKey,
        JSON.stringify(connection),
      ]);
    }

    const accessToken = tokenResult.access_token;
    const microsoftTimeZone =
      item.microsoftTimeZone ||
      (await readMicrosoftTimeZone(accessToken));

    item.microsoftTimeZone = microsoftTimeZone;
    item.schedulingConnectionError = "";

    if (action === "create-calendar") {
      const calendarStartDateTime =
        normalizeLocalDateTime(body.calendarStartDateTime) ||
        normalizeLocalDateTime(item.calendarStartDateTime);
      let calendarEndDateTime =
        normalizeLocalDateTime(body.calendarEndDateTime) ||
        normalizeLocalDateTime(item.calendarEndDateTime);

      if (!calendarStartDateTime) {
        return response.status(400).json({
          error:
            "Choose a valid calendar start date and time.",
        });
      }

      if (
        !calendarEndDateTime ||
        calendarEndDateTime <= calendarStartDateTime
      ) {
        calendarEndDateTime = addMinutes(
          calendarStartDateTime,
          30
        );
      }

      const calendarTitle =
        normalizeText(body.calendarTitle, 300) ||
        normalizeText(item.calendarTitle, 300) ||
        normalizeText(item.title, 300) ||
        "GSD Capture event";
      const calendarLocation =
        normalizeText(body.calendarLocation, 300) ||
        normalizeText(item.calendarLocation, 300) ||
        "";

      const event = await createMicrosoftCalendarEvent(
        {
          title: calendarTitle,
          startDateTime: calendarStartDateTime,
          endDateTime: calendarEndDateTime,
          location: calendarLocation,
          originalText: item.originalText || item.title || "",
          reviewId: item.id,
        },
        microsoftTimeZone,
        accessToken
      );

      item.calendarTitle = calendarTitle;
      item.calendarStartDateTime = calendarStartDateTime;
      item.calendarEndDateTime = calendarEndDateTime;
      item.calendarLocation = calendarLocation;
      item.calendarStatus = "created";
      item.calendarEventId = event.id || "";
      item.calendarWebLink = event.webLink || "";
      item.calendarError = "";
      item.calendarResolvedAt = new Date().toISOString();
      await saveReviewItem(itemKey, item);

      return response.status(200).json({
        ok: true,
        item,
        message: "Calendar event created.",
      });
    }

    if (action === "set-reminder") {
      const reminderDateTime =
        normalizeLocalDateTime(body.reminderDateTime) ||
        normalizeLocalDateTime(item.reminderDateTime);

      if (!reminderDateTime) {
        return response.status(400).json({
          error: "Choose a valid reminder date and time.",
        });
      }

      if (
        !item.microsoftReviewListId ||
        !item.microsoftTaskId
      ) {
        return response.status(409).json({
          error:
            "The Microsoft To Do task could not be found for this reminder.",
        });
      }

      await updateMicrosoftTaskReminder(
        item.microsoftReviewListId,
        item.microsoftTaskId,
        reminderDateTime,
        microsoftTimeZone,
        accessToken
      );

      item.reminderDateTime = reminderDateTime;
      item.reminderStatus = "set";
      item.reminderError = "";
      item.reminderResolvedAt = new Date().toISOString();
      await saveReviewItem(itemKey, item);

      return response.status(200).json({
        ok: true,
        item,
        message: "Reminder set in Microsoft To Do.",
      });
    }

    return response.status(400).json({
      error: "The scheduling action was not recognized.",
    });
  } catch (error) {
    const message =
      error && error.message
        ? error.message
        : "The scheduling action could not be completed.";
    const reconnectRequired = isReconnectError(message);

    if (action === "create-calendar") {
      item.calendarStatus = reconnectRequired
        ? "needs-reconnect"
        : "failed";
      item.calendarError = message;
    }

    if (action === "set-reminder") {
      item.reminderStatus = reconnectRequired
        ? "needs-reconnect"
        : "failed";
      item.reminderError = message;
    }

    item.schedulingConnectionError = reconnectRequired
      ? message
      : item.schedulingConnectionError || "";

    await saveReviewItem(itemKey, item);

    return response.status(reconnectRequired ? 409 : 502).json({
      error: message,
      reconnectRequired,
      item,
    });
  }
}

function normalizeAction(value) {
  const action =
    typeof value === "string"
      ? value.trim()
      : "";
  const allowed = [
    "create-calendar",
    "dismiss-calendar",
    "set-reminder",
    "dismiss-reminder",
  ];

  return allowed.includes(action) ? action : "";
}

function dismissPendingScheduling(item) {
  if (
    [
      "suggested",
      "needs-details",
      "needs-reconnect",
      "failed",
    ].includes(item.calendarStatus)
  ) {
    item.calendarStatus = "dismissed";
    item.calendarResolvedAt = new Date().toISOString();
  }

  if (
    [
      "suggested",
      "needs-details",
      "needs-reconnect",
      "failed",
    ].includes(item.reminderStatus)
  ) {
    item.reminderStatus = "dismissed";
    item.reminderResolvedAt = new Date().toISOString();
  }
}

async function saveReviewItem(itemKey, item) {
  item.updatedAt = new Date().toISOString();

  await redisCommand([
    "SET",
    itemKey,
    JSON.stringify(item),
    "EX",
    REVIEW_ITEM_TTL_SECONDS,
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
    throw new Error(
      tokenData && tokenData.error_description
        ? tokenData.error_description
        : "Microsoft access needs to be reconnected."
    );
  }

  return tokenData;
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

async function createMicrosoftCalendarEvent(
  eventDetails,
  microsoftTimeZone,
  accessToken
) {
  const payload = {
    subject: eventDetails.title,
    body: {
      contentType: "text",
      content:
        "Created by GSD Capture.\n\nOriginal capture: " +
        eventDetails.originalText +
        "\nGSD Voice Review ID: " +
        eventDetails.reviewId,
    },
    start: {
      dateTime: eventDetails.startDateTime,
      timeZone: microsoftTimeZone,
    },
    end: {
      dateTime: eventDetails.endDateTime,
      timeZone: microsoftTimeZone,
    },
    transactionId: eventDetails.reviewId,
  };

  if (eventDetails.location) {
    payload.location = {
      displayName: eventDetails.location,
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

async function updateMicrosoftTaskReminder(
  listId,
  taskId,
  reminderDateTime,
  microsoftTimeZone,
  accessToken
) {
  const graphResponse = await fetch(
    "https://graph.microsoft.com/v1.0/me/todo/lists/" +
      encodeURIComponent(listId) +
      "/tasks/" +
      encodeURIComponent(taskId),
    {
      method: "PATCH",
      headers: {
        Authorization: "Bearer " + accessToken,
        Accept: "application/json",
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        isReminderOn: true,
        reminderDateTime: {
          dateTime: reminderDateTime,
          timeZone: microsoftTimeZone,
        },
      }),
    }
  );

  const data = await graphResponse.json().catch(function () {
    return null;
  });

  if (!graphResponse.ok) {
    throw new Error(readGraphError(data, graphResponse.status));
  }

  return data || {};
}

function normalizeLocalDateTime(value) {
  if (typeof value !== "string") {
    return "";
  }

  const trimmed = value.trim().replace(" ", "T");
  const match = trimmed.match(
    /^(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2})(?::(\d{2}))?$/
  );

  if (!match) {
    return "";
  }

  const normalized =
    match[1] + "T" + match[2] + ":" + (match[3] || "00");
  const testDate = new Date(normalized + "Z");

  return Number.isNaN(testDate.getTime())
    ? ""
    : normalized;
}

function addMinutes(dateTime, minutes) {
  const date = new Date(dateTime + "Z");
  date.setUTCMinutes(date.getUTCMinutes() + minutes);
  return date.toISOString().slice(0, 19);
}

function normalizeText(value, maximumLength) {
  return typeof value === "string" && value.trim()
    ? value.trim().slice(0, maximumLength)
    : "";
}

function isReconnectError(message) {
  const lower = String(message || "").toLowerCase();
  return (
    lower.includes("interaction_required") ||
    lower.includes("consent") ||
    lower.includes("insufficient privileges") ||
    lower.includes("invalidauthenticationtoken") ||
    lower.includes("refresh token") ||
    lower.includes("reconnect") ||
    lower.includes("aadsts")
  );
}

function readGraphError(data, status) {
  return data && data.error && data.error.message
    ? data.error.message
    : "Microsoft Graph returned " + status + ".";
}

function readVoiceKey(request) {
  const headerValue =
    request.headers["x-gsd-capture-key"] ||
    request.headers["X-GSD-Capture-Key"] ||
    "";

  if (headerValue) {
    return String(headerValue);
  }

  const authorization = String(
    request.headers.authorization || ""
  );

  return authorization
    .toLowerCase()
    .startsWith("bearer ")
    ? authorization.slice(7).trim()
    : "";
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
