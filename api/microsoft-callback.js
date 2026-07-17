"use strict";

import {
  LEGACY_CONNECTION_KEY,
  SETUP_PREFIX,
  SETUP_RESULT_PREFIX,
  STATE_PREFIX,
  getUserConnectionKey,
  redisCommand,
  requireEnvironment,
  sanitizeSetupId,
} from "../lib/voice-user-store.js";

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

export default async function handler(
  request,
  response
) {
  response.setHeader(
    "Cache-Control",
    "no-store"
  );

  if (request.method !== "GET") {
    response.setHeader("Allow", "GET");

    return response
      .status(405)
      .send("Only GET requests are allowed.");
  }

  try {
    requireEnvironment([
      "MICROSOFT_CLIENT_ID",
      "MICROSOFT_CLIENT_SECRET",
      "MICROSOFT_REDIRECT_URI",
      "KV_REST_API_URL",
      "KV_REST_API_TOKEN",
    ]);

    const query = request.query || {};

    if (query.error) {
      const description = String(
        query.error_description ||
          query.error ||
          "Microsoft sign-in was cancelled."
      );

      return sendHtml(
        response,
        false,
        description,
        "/"
      );
    }

    const code =
      typeof query.code === "string"
        ? query.code
        : "";

    const state =
      typeof query.state === "string"
        ? query.state
        : "";

    if (!code || !state) {
      return sendHtml(
        response,
        false,
        "Microsoft did not return the required authorization information.",
        "/"
      );
    }

    const storedState = await redisCommand([
      "GET",
      STATE_PREFIX + state,
    ]);

    if (!storedState) {
      return sendHtml(
        response,
        false,
        "This Microsoft connection link expired or was already used. Start the connection again.",
        "/"
      );
    }

    await redisCommand([
      "DEL",
      STATE_PREFIX + state,
    ]);

    const stateRecord =
      parseStateRecord(storedState);

    if (!stateRecord) {
      return sendHtml(
        response,
        false,
        "The Microsoft connection information was invalid. Start the connection again.",
        "/"
      );
    }

    const tokenData =
      await exchangeAuthorizationCode(code);

    if (
      !tokenData.access_token ||
      !tokenData.refresh_token
    ) {
      throw new Error(
        "Microsoft did not return the required access and refresh tokens."
      );
    }

    const profile =
      await readMicrosoftProfile(
        tokenData.access_token
      );

    if (!profile.id) {
      throw new Error(
        "The Microsoft profile did not include a user ID."
      );
    }

    const connection = {
      refreshToken:
        tokenData.refresh_token,
      displayName:
        profile.displayName ||
        "Microsoft user",
      email:
        profile.mail ||
        profile.userPrincipalName ||
        "",
      microsoftUserId:
        profile.id,
      connectedAt:
        new Date().toISOString(),
      scope:
        tokenData.scope ||
        MICROSOFT_SCOPES.join(" "),
    };

    if (stateRecord.mode === "per-user") {
      return completePerUserSetup(
        response,
        stateRecord,
        connection
      );
    }

    await redisCommand([
      "SET",
      LEGACY_CONNECTION_KEY,
      JSON.stringify(connection),
    ]);

    return sendHtml(
      response,
      true,
      "Siri background capture is connected to " +
        connection.displayName +
        ".",
      "/"
    );
  } catch (error) {
    console.error(
      "Microsoft callback failed:",
      error
    );

    return sendHtml(
      response,
      false,
      "The Microsoft connection could not be completed. " +
        (
          error && error.message
            ? error.message
            : "Please try again."
        ),
      "/"
    );
  }
}

async function completePerUserSetup(
  response,
  stateRecord,
  connection
) {
  const setupId = sanitizeSetupId(
    stateRecord.setupId
  );

  if (!setupId) {
    return sendHtml(
      response,
      false,
      "The Voice Capture setup ID was invalid.",
      "/"
    );
  }

  const storedSetup = await redisCommand([
    "GET",
    SETUP_PREFIX + setupId,
  ]);

  if (!storedSetup) {
    return sendHtml(
      response,
      false,
      "This Voice Capture setup expired. Return to GSD Capture and start again.",
      "/"
    );
  }

  const setupRecord =
    JSON.parse(storedSetup);

  if (
    connection.microsoftUserId !==
    setupRecord.microsoftUserId
  ) {
    return sendHtml(
      response,
      false,
      "The Microsoft account did not match the account that started Voice Capture setup.",
      "/"
    );
  }

  const connectionKey =
    getUserConnectionKey(
      setupRecord.voiceKeyHash
    );

  await redisCommand([
    "SET",
    connectionKey,
    JSON.stringify(connection),
  ]);

  const setupResult = {
    ok: true,
    complete: true,
    setupId,
    displayName:
      connection.displayName,
    email:
      connection.email,
    connectedAt:
      connection.connectedAt,
  };

  await redisCommand([
    "SET",
    SETUP_RESULT_PREFIX + setupId,
    JSON.stringify(setupResult),
    "EX",
    3600,
  ]);

  await redisCommand([
    "DEL",
    SETUP_PREFIX + setupId,
  ]);

  return sendHtml(
    response,
    true,
    "Voice Capture is connected to " +
      connection.displayName +
      ". Return to GSD Capture to install the shortcut.",
    "/?voiceSetup=complete&setupId=" +
      encodeURIComponent(setupId)
  );
}

function parseStateRecord(value) {
  try {
    const parsed = JSON.parse(value);

    if (
      !parsed ||
      typeof parsed !== "object"
    ) {
      return null;
    }

    if (
      parsed.mode !== "legacy" &&
      parsed.mode !== "per-user"
    ) {
      return null;
    }

    return parsed;
  } catch (error) {
    return null;
  }
}

async function exchangeAuthorizationCode(
  code
) {
  const tokenResponse = await fetch(
    MICROSOFT_AUTHORITY +
      "/oauth2/v2.0/token",
    {
      method: "POST",
      headers: {
        "Content-Type":
          "application/x-www-form-urlencoded",
      },
      body: new URLSearchParams({
        client_id:
          process.env.MICROSOFT_CLIENT_ID,
        client_secret:
          process.env.MICROSOFT_CLIENT_SECRET,
        grant_type:
          "authorization_code",
        code,
        redirect_uri:
          process.env.MICROSOFT_REDIRECT_URI,
        scope:
          MICROSOFT_SCOPES.join(" "),
      }).toString(),
    }
  );

  const tokenData = await tokenResponse
    .json()
    .catch(function () {
      return null;
    });

  if (!tokenResponse.ok) {
    throw new Error(
      tokenData &&
        tokenData.error_description
        ? tokenData.error_description
        : "Microsoft token exchange failed."
    );
  }

  return tokenData || {};
}

async function readMicrosoftProfile(
  accessToken
) {
  const profileResponse = await fetch(
    "https://graph.microsoft.com/v1.0/me?$select=id,displayName,mail,userPrincipalName",
    {
      headers: {
        Authorization:
          "Bearer " + accessToken,
        Accept: "application/json",
      },
    }
  );

  const profile = await profileResponse
    .json()
    .catch(function () {
      return null;
    });

  if (!profileResponse.ok) {
    throw new Error(
      profile &&
        profile.error &&
        profile.error.message
        ? profile.error.message
        : "The Microsoft profile could not be read."
    );
  }

  return profile || {};
}

function sendHtml(
  response,
  success,
  message,
  returnPath
) {
  const heading = success
    ? "Voice Capture Connected"
    : "Voice Capture Connection Failed";

  const accent = success
    ? "#166534"
    : "#991b1b";

  const background = success
    ? "#f0fdf4"
    : "#fef2f2";

  response.setHeader(
    "Content-Type",
    "text/html; charset=utf-8"
  );

  return response
    .status(success ? 200 : 400)
    .send(
      "<!doctype html>" +
        '<html lang="en">' +
        "<head>" +
        '<meta charset="utf-8">' +
        '<meta name="viewport" content="width=device-width,initial-scale=1">' +
        "<title>" +
        escapeHtml(heading) +
        "</title>" +
        "</head>" +
        '<body style="font-family:-apple-system,BlinkMacSystemFont,Segoe UI,sans-serif;' +
        "margin:0;padding:32px;background:" +
        background +
        ';color:#172033">' +
        '<main style="max-width:620px;margin:0 auto;background:white;padding:28px;' +
        'border-radius:18px;box-shadow:0 12px 36px rgba(0,0,0,.08)">' +
        '<h1 style="margin-top:0;color:' +
        accent +
        '">' +
        escapeHtml(heading) +
        "</h1>" +
        "<p>" +
        escapeHtml(message) +
        "</p>" +
        '<p><a href="' +
        escapeHtml(returnPath || "/") +
        '">Return to GSD Capture</a></p>' +
        "</main>" +
        "</body>" +
        "</html>"
    );
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
