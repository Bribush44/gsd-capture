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

    const connection = await redisCommand([
      "GET",
      getUserConnectionKey(voiceKeyHash),
    ]);

    if (!connection) {
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
      voiceKeyHash
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
  voiceKeyHash
) {
  const body = parseRequestBody(request.body);
  const reviewId = sanitizeVoiceReviewId(
    body.reviewId
  );
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

  if (!reviewId || !status) {
    return response.status(400).json({
      error:
        "A valid review ID and status are required.",
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

  await redisCommand([
    "SET",
    itemKey,
    JSON.stringify(item),
    "EX",
    REVIEW_ITEM_TTL_SECONDS,
  ]);

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
