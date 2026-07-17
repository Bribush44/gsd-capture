"use strict";

const STATE_PREFIX = "gsd:siri:microsoft:oauth-state:";
const CONNECTION_KEY = "gsd:siri:microsoft:connection";
const MICROSOFT_AUTHORITY = "https://login.microsoftonline.com/common";
const MICROSOFT_SCOPES = [
  "openid",
  "profile",
  "offline_access",
  "User.Read",
  "Tasks.ReadWrite",
];

export default async function handler(request, response) {
  if (request.method !== "GET") {
    response.setHeader("Allow", "GET");
    return response.status(405).send("Only GET requests are allowed.");
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
        query.error_description || query.error || "Microsoft sign-in was cancelled."
      );
      return sendHtml(response, false, description);
    }

    const code = typeof query.code === "string" ? query.code : "";
    const state = typeof query.state === "string" ? query.state : "";

    if (!code || !state) {
      return sendHtml(
        response,
        false,
        "Microsoft did not return the required authorization information."
      );
    }

    const storedState = await redisCommand(["GETDEL", STATE_PREFIX + state]);

    if (storedState !== "pending") {
      return sendHtml(
        response,
        false,
        "This Microsoft connection link expired or was already used. Start the connection again."
      );
    }

    const tokenData = await exchangeAuthorizationCode(code);

    if (!tokenData.refresh_token || !tokenData.access_token) {
      throw new Error("Microsoft did not return the required tokens.");
    }

    const profile = await readMicrosoftProfile(tokenData.access_token);
    const connection = {
      refreshToken: tokenData.refresh_token,
      displayName: profile.displayName || "Microsoft user",
      email: profile.mail || profile.userPrincipalName || "",
      microsoftUserId: profile.id || "",
      connectedAt: new Date().toISOString(),
      scope: tokenData.scope || MICROSOFT_SCOPES.join(" "),
    };

    await redisCommand(["SET", CONNECTION_KEY, JSON.stringify(connection)]);

    return sendHtml(
      response,
      true,
      "Siri background capture is connected to " + connection.displayName + "."
    );
  } catch (error) {
    console.error("Microsoft callback failed:", error);
    return sendHtml(
      response,
      false,
      "The Microsoft connection could not be completed. " +
        (error && error.message ? error.message : "")
    );
  }
}

async function exchangeAuthorizationCode(code) {
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
        grant_type: "authorization_code",
        code: code,
        redirect_uri: process.env.MICROSOFT_REDIRECT_URI,
        scope: MICROSOFT_SCOPES.join(" "),
      }).toString(),
    }
  );

  const tokenData = await tokenResponse.json().catch(function () {
    return null;
  });

  if (!tokenResponse.ok) {
    const message =
      tokenData && tokenData.error_description
        ? tokenData.error_description
        : "Microsoft token exchange failed.";
    throw new Error(message);
  }

  return tokenData || {};
}

async function readMicrosoftProfile(accessToken) {
  const profileResponse = await fetch("https://graph.microsoft.com/v1.0/me", {
    headers: {
      Authorization: "Bearer " + accessToken,
      Accept: "application/json",
    },
  });

  const profile = await profileResponse.json().catch(function () {
    return null;
  });

  if (!profileResponse.ok) {
    throw new Error(
      profile && profile.error && profile.error.message
        ? profile.error.message
        : "The Microsoft profile could not be read."
    );
  }

  return profile || {};
}

function sendHtml(response, success, message) {
  const heading = success
    ? "Siri Capture Connected"
    : "Siri Capture Connection Failed";
  const safeMessage = escapeHtml(message || "");
  const accent = success ? "#166534" : "#991b1b";
  const background = success ? "#f0fdf4" : "#fef2f2";

  response.setHeader("Cache-Control", "no-store");
  response.setHeader("Content-Type", "text/html; charset=utf-8");

  return response.status(success ? 200 : 400).send(
    "<!doctype html>" +
      '<html lang="en"><head><meta charset="utf-8">' +
      '<meta name="viewport" content="width=device-width, initial-scale=1">' +
      "<title>" +
      heading +
      "</title>" +
      "<style>body{font-family:-apple-system,BlinkMacSystemFont,Segoe UI,sans-serif;" +
      "margin:0;padding:32px;background:#f8fafc;color:#0f172a}" +
      ".card{max-width:640px;margin:10vh auto;background:white;border-radius:20px;" +
      "padding:28px;box-shadow:0 16px 45px rgba(15,23,42,.12)}" +
      "h1{color:" +
      accent +
      ";margin-top:0}.message{background:" +
      background +
      ";padding:16px;border-radius:12px;line-height:1.5}" +
      "a{display:inline-block;margin-top:20px;color:#1d4ed8;font-weight:700}</style>" +
      "</head><body><main class=\"card\"><h1>" +
      heading +
      "</h1><div class=\"message\">" +
      safeMessage +
      "</div><a href=\"/\">Return to GSD Capture</a></main></body></html>"
  );
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
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
