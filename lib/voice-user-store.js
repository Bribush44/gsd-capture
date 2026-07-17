"use strict";

import { createHash, randomBytes } from "node:crypto";

export const SETUP_PREFIX = "gsd:siri:setup:";
export const SETUP_RESULT_PREFIX =
  "gsd:siri:setup-result:";
export const STATE_PREFIX =
  "gsd:siri:microsoft:oauth-state:";

export const LEGACY_CONNECTION_KEY =
  "gsd:siri:microsoft:connection";

export const USER_CONNECTION_PREFIX =
  "gsd:siri:user:";

export function createSetupId() {
  return randomBytes(24).toString("hex");
}

export function createVoiceKey() {
  return (
    "gsvc_" +
    randomBytes(32).toString("base64url")
  );
}

export function hashVoiceKey(voiceKey) {
  return createHash("sha256")
    .update(String(voiceKey || ""))
    .digest("hex");
}

export function getUserConnectionKey(
  voiceKeyHash
) {
  return (
    USER_CONNECTION_PREFIX +
    voiceKeyHash +
    ":connection"
  );
}

export function sanitizeSetupId(value) {
  const setupId =
    typeof value === "string"
      ? value.trim()
      : "";

  return /^[a-f0-9]{48}$/.test(setupId)
    ? setupId
    : "";
}

export function requireEnvironment(
  variableNames
) {
  const missingVariables =
    variableNames.filter(function (
      variableName
    ) {
      return !process.env[variableName];
    });

  if (missingVariables.length > 0) {
    throw new Error(
      "Missing environment variables: " +
        missingVariables.join(", ")
    );
  }
}

export async function redisCommand(command) {
  requireEnvironment([
    "KV_REST_API_URL",
    "KV_REST_API_TOKEN",
  ]);

  const redisResponse = await fetch(
    process.env.KV_REST_API_URL,
    {
      method: "POST",
      headers: {
        Authorization:
          "Bearer " +
          process.env.KV_REST_API_TOKEN,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(command),
    }
  );

  const data = await redisResponse
    .json()
    .catch(function () {
      return null;
    });

  if (
    !redisResponse.ok ||
    !data ||
    data.error
  ) {
    throw new Error(
      data && data.error
        ? data.error
        : "Secure storage request failed."
    );
  }

  return data.result;
}
