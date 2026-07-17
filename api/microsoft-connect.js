"use strict";

import { randomBytes } from "node:crypto";

const STATE_PREFIX = "gsd:siri:microsoft:oauth-state:";
const STATE_TTL_SECONDS = 600;
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
    return response.status(405).json({ error: "Only GET requests are allowed." });
  }

  try {
    requireEnvironment([
      "MICROSOFT_CLIENT_ID",
      "MICROSOFT_REDIRECT_URI",
      "KV_REST_API_URL",
      "KV_REST_API_TOKEN",
    ]);

    const state = randomBytes(32).toString("hex");
    await redisCommand([
      "SET",
      STATE_PREFIX + state,
      "pending",
      "EX",
      STATE_TTL_SECONDS,
      "NX",
    ]);

    const authorizationUrl = new URL(
      MICROSOFT_AUTHORITY + "/oauth2/v2.0/authorize"
    );

    authorizationUrl.searchParams.set(
      "client_id",
      process.env.MICROSOFT_CLIENT_ID
    );
    authorizationUrl.searchParams.set("response_type", "code");
    authorizationUrl.searchParams.set(
      "redirect_uri",
      process.env.MICROSOFT_REDIRECT_URI
    );
    authorizationUrl.searchParams.set("response_mode", "query");
    authorizationUrl.searchParams.set("scope", MICROSOFT_SCOPES.join(" "));
    authorizationUrl.searchParams.set("state", state);
    authorizationUrl.searchParams.set("prompt", "consent");

    response.setHeader("Cache-Control", "no-store");
    response.setHeader("Location", authorizationUrl.toString());
    return response.status(302).end();
  } catch (error) {
    console.error("Microsoft connection start failed:", error);
    return response.status(500).json({
      error: "The Microsoft connection could not be started.",
    });
  }
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
