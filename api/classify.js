"use strict";

export default async function handler(request, response) {
  if (request.method !== "POST") {
    response.setHeader("Allow", "POST");
    return response.status(405).json({
      error: "Only POST requests are allowed."
    });
  }

  if (!process.env.OPENAI_API_KEY) {
    return response.status(500).json({
      error: "The OpenAI API key has not been configured."
    });
  }

  try {
    const body =
      typeof request.body === "string"
        ? JSON.parse(request.body)
        : request.body || {};

    const taskText =
      typeof body.text === "string"
        ? body.text.trim()
        : "";

    if (!taskText) {
      return response.status(400).json({
        error: "Task text is required."
      });
    }

    if (taskText.length > 1000) {
      return response.status(400).json({
        error: "Task text must be 1,000 characters or fewer."
      });
    }

    const currentDate =
      typeof body.currentDate === "string" && body.currentDate
        ? body.currentDate
        : new Date().toISOString().slice(0, 10);

    const openAIResponse = await fetch(
      "https://api.openai.com/v1/responses",
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${process.env.OPENAI_API_KEY}`,
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          model: "gpt-5.6-luna",
          store: false,
          reasoning: {
            effort: "none"
          },
          max_output_tokens: 400,
          input: [
            {
              role: "system",
              content: [
                {
                  type: "input_text",
                  text:
                    "You organize captured tasks using GTD principles. " +
                    "Return a practical classification without inventing details. " +
                    `Today's date is ${currentDate}. ` +
                    "Convert clear relative dates such as today or tomorrow into YYYY-MM-DD. " +
                    "Use null when a date, project, or duration cannot reasonably be determined."
                }
              ]
            },
            {
              role: "user",
              content: [
                {
                  type: "input_text",
                  text: taskText
                }
              ]
            }
          ],
          text: {
            format: {
              type: "json_schema",
              name: "gsd_task_classification",
              strict: true,
              schema: {
                type: "object",
                properties: {
                  summary: {
                    type: "string"
                  },
                  category: {
                    type: "string",
                    enum: [
                      "Purchasing",
                      "Operations",
                      "Leadership",
                      "Personal",
                      "Ideas",
                      "General"
                    ]
                  },
                  priority: {
                    type: "string",
                    enum: ["High", "Normal", "Low"]
                  },
                  dueDate: {
                    type: ["string", "null"]
                  },
                  context: {
                    type: "string",
                    enum: [
                      "Calls",
                      "Computer",
                      "Errands",
                      "Work",
                      "Home",
                      "Anywhere"
                    ]
                  },
                  project: {
                    type: ["string", "null"]
                  },
                  estimatedMinutes: {
                    type: ["integer", "null"],
                    minimum: 1,
                    maximum: 480
                  },
                  needsReview: {
                    type: "boolean"
                  }
                },
                required: [
                  "summary",
                  "category",
                  "priority",
                  "dueDate",
                  "context",
                  "project",
                  "estimatedMinutes",
                  "needsReview"
                ],
                additionalProperties: false
              }
            }
          }
        })
      }
    );

    const openAIData = await openAIResponse.json();

    if (!openAIResponse.ok) {
      console.error("OpenAI API error:", openAIData);

      return response.status(502).json({
        error: "AI sorting failed. Local sorting can still be used."
      });
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
      throw new Error("The AI response did not contain classification text.");
    }

    const classification = JSON.parse(outputText);

    return response.status(200).json({
      classification: classification
    });
  } catch (error) {
    console.error("Classification endpoint error:", error);

    return response.status(500).json({
      error: "The task could not be classified."
    });
  }
}