import { z } from "zod";
import type { MicrosoftFormsClient } from "./forms-client.js";
import type { FormAnswers } from "./types.js";

export interface ActionParameter {
  description: string;
  name: string;
  required: boolean;
  type: "boolean" | "json" | "string";
}

export interface FormsAction {
  annotations: {
    destructiveHint: boolean;
    idempotentHint: boolean;
    openWorldHint: boolean;
    readOnlyHint: boolean;
  };
  description: string;
  execute(
    client: MicrosoftFormsClient,
    input: Record<string, unknown>,
  ): Promise<unknown>;
  inputSchema: z.ZodObject<z.ZodRawShape>;
  name: string;
  parameters: ActionParameter[];
  requiresConfirmation?: boolean;
  title: string;
}

const formSchema = z
  .string()
  .min(1)
  .describe("Microsoft Forms response URL or form ID.");
const answersSchema = z
  .record(z.string(), z.unknown())
  .describe(
    "Answers keyed by exact question ID. Use option labels or IDs for choices, ordered label/ID arrays for ranking, YYYY-MM-DD for dates, and upload results for files.",
  );

export const formsActions: FormsAction[] = [
  {
    name: "inspect",
    title: "Inspect Microsoft Form",
    description:
      "Get complete normalized form metadata, raw API data, questions, options, validation, upload rules, and branch targets.",
    inputSchema: z.object({ form: formSchema }),
    parameters: [
      {
        name: "form",
        type: "string",
        required: true,
        description: "Microsoft Forms response URL or form ID.",
      },
    ],
    annotations: {
      readOnlyHint: true,
      destructiveHint: false,
      idempotentHint: true,
      openWorldHint: true,
    },
    async execute(client, input) {
      const parameters = this.inputSchema.parse(input);
      return client.getForm(String(parameters.form));
    },
  },
  {
    name: "validate",
    title: "Validate Microsoft Forms Answers",
    description:
      "Resolve branching, validate answers, and preview the exact serialized response without submitting.",
    inputSchema: z.object({ form: formSchema, answers: answersSchema }),
    parameters: [
      {
        name: "form",
        type: "string",
        required: true,
        description: "Microsoft Forms response URL or form ID.",
      },
      {
        name: "answers",
        type: "json",
        required: true,
        description: "JSON object, or @file path in CLI, keyed by question ID.",
      },
    ],
    annotations: {
      readOnlyHint: true,
      destructiveHint: false,
      idempotentHint: true,
      openWorldHint: true,
    },
    async execute(client, input) {
      const parameters = this.inputSchema.parse(input);
      return client.validateAnswers(
        String(parameters.form),
        parameters.answers as FormAnswers,
      );
    },
  },
  {
    name: "upload-file",
    title: "Upload File to Microsoft Form",
    description:
      "Upload one local file for a visible file-upload question and return metadata to place in that question's answer.",
    inputSchema: z.object({
      form: formSchema,
      questionId: z.string().min(1).describe("File-upload question ID."),
      filePath: z.string().min(1).describe("Local path of the file to upload."),
    }),
    parameters: [
      {
        name: "form",
        type: "string",
        required: true,
        description: "Microsoft Forms response URL or form ID.",
      },
      {
        name: "questionId",
        type: "string",
        required: true,
        description: "File-upload question ID.",
      },
      {
        name: "filePath",
        type: "string",
        required: true,
        description: "Local path of the file to upload.",
      },
    ],
    annotations: {
      readOnlyHint: false,
      destructiveHint: false,
      idempotentHint: false,
      openWorldHint: true,
    },
    async execute(client, input) {
      const parameters = this.inputSchema.parse(input);
      return client.uploadFile(
        String(parameters.form),
        String(parameters.questionId),
        String(parameters.filePath),
      );
    },
  },
  {
    name: "submit",
    title: "Submit Microsoft Forms Response",
    description:
      "Validate and submit a response. This creates a permanent form response and always requires explicit user confirmation.",
    inputSchema: z.object({
      form: formSchema,
      answers: answersSchema,
      saveResponse: z
        .boolean()
        .optional()
        .describe("Save the submitted response link to the signed-in account."),
      submitLanguage: z
        .string()
        .optional()
        .describe("Optional Forms locale identifier."),
      emailReceiptConsent: z
        .boolean()
        .optional()
        .describe("Whether the responder consents to an email receipt."),
    }),
    parameters: [
      {
        name: "form",
        type: "string",
        required: true,
        description: "Microsoft Forms response URL or form ID.",
      },
      {
        name: "answers",
        type: "json",
        required: true,
        description: "JSON object, or @file path in CLI, keyed by question ID.",
      },
      {
        name: "saveResponse",
        type: "boolean",
        required: false,
        description: "Save response link to the signed-in account.",
      },
      {
        name: "submitLanguage",
        type: "string",
        required: false,
        description: "Optional Forms locale identifier.",
      },
      {
        name: "emailReceiptConsent",
        type: "boolean",
        required: false,
        description: "Consent to an email receipt.",
      },
    ],
    requiresConfirmation: true,
    annotations: {
      readOnlyHint: false,
      destructiveHint: true,
      idempotentHint: false,
      openWorldHint: true,
    },
    async execute(client, input) {
      const parameters = this.inputSchema.parse(input);
      return client.submitResponse(
        String(parameters.form),
        parameters.answers as FormAnswers,
        {
          saveResponse: parameters.saveResponse as boolean | undefined,
          submitLanguage: parameters.submitLanguage as string | undefined,
          emailReceiptConsent: parameters.emailReceiptConsent as
            boolean | undefined,
        },
      );
    },
  },
  {
    name: "save-response",
    title: "Save Microsoft Forms Response Link",
    description:
      "Save an existing submitted response link to the signed-in Microsoft Forms account.",
    inputSchema: z.object({
      form: formSchema,
      responseId: z.string().min(1).describe("Submitted response ID."),
      submitDate: z
        .string()
        .datetime()
        .describe("Submission time as an ISO-8601 timestamp."),
    }),
    parameters: [
      {
        name: "form",
        type: "string",
        required: true,
        description: "Microsoft Forms response URL or form ID.",
      },
      {
        name: "responseId",
        type: "string",
        required: true,
        description: "Submitted response ID.",
      },
      {
        name: "submitDate",
        type: "string",
        required: true,
        description: "Submission time as an ISO-8601 timestamp.",
      },
    ],
    annotations: {
      readOnlyHint: false,
      destructiveHint: false,
      idempotentHint: true,
      openWorldHint: true,
    },
    async execute(client, input) {
      const parameters = this.inputSchema.parse(input);
      return client.saveResponse(
        String(parameters.form),
        String(parameters.responseId),
        String(parameters.submitDate),
      );
    },
  },
  {
    name: "list-saved-responses",
    title: "List Saved Microsoft Forms Responses",
    description:
      "List response links saved to the signed-in account. A form reference supplies the authenticated tenant context.",
    inputSchema: z.object({ form: formSchema }),
    parameters: [
      {
        name: "form",
        type: "string",
        required: true,
        description: "Microsoft Forms response URL or form ID.",
      },
    ],
    annotations: {
      readOnlyHint: true,
      destructiveHint: false,
      idempotentHint: true,
      openWorldHint: true,
    },
    async execute(client, input) {
      const parameters = this.inputSchema.parse(input);
      return client.listSavedResponses(String(parameters.form));
    },
  },
];
