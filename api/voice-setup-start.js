"use strict";

import {
  SETUP_PREFIX,
  createSetupId,
  createVoiceKey,
  hashVoiceKey,
  redisCommand,
  requireEnvironment,
} from "../lib/voice-user-store.js";

const SETUP_TTL_SECONDS = 900;

export default async function handler(
  request,
  response
) {
  response.setHeader(
    "Cache-Control",
    "no-store"
  );

  if (request.method !== "POST") {
    response.setHeader("Allow", "POST");

    return response.status(405).json({
      error: "Only POST requests are allowed.",
    });
  }

  try {
    requireEnvironment([
      "KV_REST_API_URL",
      "KV_REST_API_TOKEN",
    ]);

    const accessToken =
      readBearerToken(request);

    if (!accessToken) {
      return response.status(401).json({
        error: "Microsoft sign-in is required.",
      });
    }

    const profile =
      await readMicrosoftProfile(accessToken);

    if (!profile.id) {
      throw new Error(
        "The Microsoft profile did not include a user ID."
      );
    }

    const setupId = createSetupId();
    const voiceKey = createVoiceKey();
    const voiceKeyHash =
      hashVoiceKey(voiceKey);

    const setupRecord = {
      setupId,
      voiceKeyHash,
      microsoftUserId: profile.id,
      displayName:
        profile.displayName ||
        "Microsoft user",
      email:
        profile.mail ||
        profile.userPrincipalName ||
        "",
      createdAt: new Date().toISOString(),
    };

    await redisCommand([
      "SET",
      SETUP_PREFIX + setupId,
      JSON.stringify(setupRecord),
      "EX",
      SETUP_TTL_SECONDS,
      "NX",
    ]);

    return response.status(200).json({
      ok: true,
      setupId,
      voiceKey,
      displayName:
        setupRecord.displayName,
      email: setupRecord.email,
      expiresInSeconds:
        SETUP_TTL_SECONDS,
      connectUrl:
        "/api/microsoft-connect?setup=" +
        encodeURIComponent(setupId),
    });
  } catch (error) {
    console.error(
      "Per-user voice setup failed to start:",
      error
    );

    return response.status(500).json({
      error:
        "Voice Capture setup could not be started.",
      message:
        error && error.message
          ? error.message
          : "Please try again.",
    });
  }
}

function readBearerToken(request) {
  const authorization = String(
    request.headers.authorization || ""
  );

  return authorization
    .toLowerCase()
    .startsWith("bearer ")
    ? authorization.slice(7).trim()
    : "";
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
        : "The Microsoft profile could not be verified."
    );
  }

  return profile || {};
}
