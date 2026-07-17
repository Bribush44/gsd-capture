"use strict";

import { randomBytes } from "node:crypto";

import {
  SETUP_PREFIX,
  STATE_PREFIX,
  redisCommand,
  requireEnvironment,
  sanitizeSetupId,
} from "../lib/voice-user-store.js";

const STATE_TTL_SECONDS = 600;

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

    return response.status(405).json({
      error: "Only GET requests are allowed.",
    });
  }

  try {
    requireEnvironment([
      "MICROSOFT_CLIENT_ID",
      "MICROSOFT_REDIRECT_URI",
      "KV_REST_API_URL",
      "KV_REST_API_TOKEN",
    ]);

    const rawSetupId = Array.isArray(
      request.query?.setup
    )
      ? request.query.setup[0]
      : request.query?.setup || "";

    const setupId =
      sanitizeSetupId(rawSetupId);

    if (rawSetupId && !setupId) {
      return response.status(400).json({
        error:
          "The Voice Capture setup link is invalid.",
      });
    }

    if (setupId) {
      const pendingSetup =
        await redisCommand([
          "GET",
          SETUP_PREFIX + setupId,
        ]);

      if (!pendingSetup) {
        return response.status(410).json({
          error:
            "This Voice Capture setup link expired. Start setup again in GSD Capture.",
        });
      }
    }

    const state =
      randomBytes(32).toString("hex");

    const stateRecord = {
      mode: setupId
        ? "per-user"
        : "legacy",
      setupId: setupId || "",
      createdAt:
        new Date().toISOString(),
    };

    await redisCommand([
      "SET",
      STATE_PREFIX + state,
      JSON.stringify(stateRecord),
      "EX",
      STATE_TTL_SECONDS,
      "NX",
    ]);

    const authorizationUrl = new URL(
      MICROSOFT_AUTHORITY +
        "/oauth2/v2.0/authorize"
    );

    authorizationUrl.searchParams.set(
      "client_id",
      process.env.MICROSOFT_CLIENT_ID
    );

    authorizationUrl.searchParams.set(
      "response_type",
      "code"
    );

    authorizationUrl.searchParams.set(
      "redirect_uri",
      process.env.MICROSOFT_REDIRECT_URI
    );

    authorizationUrl.searchParams.set(
      "response_mode",
      "query"
    );

    authorizationUrl.searchParams.set(
      "scope",
      MICROSOFT_SCOPES.join(" ")
    );

    authorizationUrl.searchParams.set(
      "state",
      state
    );

    authorizationUrl.searchParams.set(
      "prompt",
      "consent"
    );

    response.setHeader(
      "Location",
      authorizationUrl.toString()
    );

    return response.status(302).end();
  } catch (error) {
    console.error(
      "Microsoft connection start failed:",
      error
    );

    return response.status(500).json({
      error:
        "The Microsoft connection could not be started.",
      message:
        error && error.message
          ? error.message
          : "Please try again.",
    });
  }
}
