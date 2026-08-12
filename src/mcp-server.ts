#!/usr/bin/env node

import {
  acceptedContent,
  inputRequired,
  inputResponse,
  McpServer,
} from "@modelcontextprotocol/server";
import { serveStdio } from "@modelcontextprotocol/server/stdio";
import { z } from "zod";
import { formsActions } from "./actions.js";
import { MicrosoftFormsClient } from "./forms-client.js";
import { packageName, packageVersion } from "./package-metadata.js";
import { serverInstructions } from "./server-instructions.js";

const confirmationSchema = z.object({ confirm: z.boolean() });

function jsonObject(value: unknown): Record<string, unknown> {
  return {
    result: JSON.parse(JSON.stringify(value)) as unknown,
  };
}

function createClient(): MicrosoftFormsClient {
  const browserChannel =
    process.env.MSFORMS_BROWSER_CHANNEL === "chrome" ||
    process.env.MSFORMS_BROWSER_CHANNEL === "msedge"
      ? process.env.MSFORMS_BROWSER_CHANNEL
      : undefined;
  return new MicrosoftFormsClient({
    browserProfileDirectory: process.env.MSFORMS_PROFILE_DIRECTORY,
    browserChannel,
    headless: process.env.MSFORMS_HEADLESS !== "false",
  });
}

export function createMicrosoftFormsMcpServer(): McpServer {
  const server = new McpServer(
    { name: packageName, version: packageVersion },
    { instructions: serverInstructions },
  );
  const client = createClient();

  server.registerTool(
    "forms_authenticate",
    {
      title: "Authenticate Microsoft Forms",
      description:
        "Open a visible persistent browser for Microsoft Forms login. Use when another tool reports missing or expired authentication.",
      inputSchema: z.object({
        form: z
          .string()
          .optional()
          .describe("Optional form URL or ID to verify after login."),
      }),
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: true,
      },
    },
    async ({ form }) => {
      const result = await MicrosoftFormsClient.authenticate({
        browserProfileDirectory: process.env.MSFORMS_PROFILE_DIRECTORY,
        browserChannel:
          process.env.MSFORMS_BROWSER_CHANNEL === "chrome" ||
          process.env.MSFORMS_BROWSER_CHANNEL === "msedge"
            ? process.env.MSFORMS_BROWSER_CHANNEL
            : undefined,
        form,
      });
      const structuredContent = jsonObject(result);
      return {
        content: [
          { type: "text", text: JSON.stringify(structuredContent, null, 2) },
        ],
        structuredContent,
      };
    },
  );

  for (const action of formsActions) {
    server.registerTool(
      `forms_${action.name.replaceAll("-", "_")}`,
      {
        title: action.title,
        description: action.description,
        inputSchema: action.inputSchema,
        annotations: action.annotations,
      },
      async (parameters, context) => {
        if (action.requiresConfirmation) {
          const confirmationResponse = inputResponse(
            context.mcpReq.inputResponses,
            "submission_confirmation",
          );
          if (
            confirmationResponse.kind === "elicit" &&
            confirmationResponse.action !== "accept"
          ) {
            return {
              content: [
                {
                  type: "text",
                  text: "Submission cancelled because confirmation was declined.",
                },
              ],
              isError: true,
            };
          }
          const confirmation = acceptedContent(
            context.mcpReq.inputResponses,
            "submission_confirmation",
            confirmationSchema,
          );
          if (!confirmation) {
            return inputRequired({
              inputRequests: {
                submission_confirmation: inputRequired.elicit({
                  message:
                    "Submit this response to Microsoft Forms? Submission creates a permanent response and cannot be undone by this tool.",
                  requestedSchema: confirmationSchema,
                }),
              },
            });
          }
          if (!confirmation.confirm) {
            return {
              content: [
                {
                  type: "text",
                  text: "Submission cancelled because confirmation was declined.",
                },
              ],
              isError: true,
            };
          }
        }

        const result = await action.execute(
          client,
          parameters as Record<string, unknown>,
        );
        const structuredContent = jsonObject(result);
        return {
          content: [
            { type: "text", text: JSON.stringify(structuredContent, null, 2) },
          ],
          structuredContent,
        };
      },
    );
  }
  return server;
}

if (require.main === module) {
  serveStdio(createMicrosoftFormsMcpServer, {
    onerror(error) {
      console.error(`MCP server error: ${error.message}`);
    },
  });
}
