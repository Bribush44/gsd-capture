"use strict";

import {
  SETUP_PREFIX,
  SETUP_RESULT_PREFIX,
  redisCommand,
  requireEnvironment,
  sanitizeSetupId,
} from "../lib/voice-user-store.js";

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
      "KV_REST_API_URL",
      "KV_REST_API_TOKEN",
    ]);

    const setupId = sanitizeSetupId(
      request.query &&
        typeof request.query.setup === "string"
        ? request.query.setup
        : ""
    );

    if (!setupId) {
      return response.status(400).json({
        error: "A valid setup ID is required.",
      });
    }

    const completedSetup =
      await redisCommand([
        "GET",
        SETUP_RESULT_PREFIX + setupId,
      ]);

    if (completedSetup) {
      return response.status(200).json(
        JSON.parse(completedSetup)
      );
    }

    const pendingSetup =
      await redisCommand([
        "GET",
        SETUP_PREFIX + setupId,
      ]);

    return response.status(200).json({
      ok: true,
      complete: false,
      pending: Boolean(pendingSetup),
      expired: !pendingSetup,
    });
  } catch (error) {
    console.error(
      "Voice setup status failed:",
      error
    );

    return response.status(500).json({
      error:
        "Voice Capture setup status could not be checked.",
    });
  }
}
