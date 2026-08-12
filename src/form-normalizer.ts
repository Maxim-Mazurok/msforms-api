import { MicrosoftFormsValidationError } from "./errors.js";
import type {
  FileUploadRules,
  FormDefinition,
  FormOption,
  FormQuestion,
  FormValidation,
  RawChoice,
  RawForm,
  RawQuestion,
} from "./types.js";

function parseObject(
  value: string | Record<string, unknown> | undefined,
  fieldName: string,
): Record<string, unknown> {
  if (!value) {
    return {};
  }
  if (typeof value === "object") {
    return value;
  }
  try {
    const parsedValue: unknown = JSON.parse(value);
    return parsedValue && typeof parsedValue === "object"
      ? (parsedValue as Record<string, unknown>)
      : {};
  } catch (error) {
    throw new MicrosoftFormsValidationError(
      `Microsoft Forms returned invalid JSON for ${fieldName}: ${
        error instanceof Error ? error.message : String(error)
      }`,
    );
  }
}

function toOptionalNumber(value: unknown): number | undefined {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  if (typeof value === "string" && value.trim()) {
    const parsedValue = Number(value);
    return Number.isFinite(parsedValue) ? parsedValue : undefined;
  }
  return undefined;
}

function toOptionalStringOrNumber(value: unknown): string | number | undefined {
  return typeof value === "string" || typeof value === "number"
    ? value
    : undefined;
}

function normalizeOptions(
  rawQuestion: RawQuestion,
  configuration: Record<string, unknown>,
): FormOption[] {
  const configuredChoices = Array.isArray(configuration.Choices)
    ? (configuration.Choices as RawChoice[])
    : [];
  const rawChoices = rawQuestion.choices ?? [];
  const choices = configuredChoices.length > 0 ? configuredChoices : rawChoices;

  return choices
    .filter((choice): choice is RawChoice => Boolean(choice))
    .map((choice, index) => {
      const raw = choice as Record<string, unknown>;
      const branchInformation =
        choice.BranchInfo ??
        (raw.BranchInfo as { TargetQuestionId?: string } | undefined);
      return {
        id: choice.id ?? toOptionalStringOrNumber(raw.Id),
        key: choice.key ?? toOptionalStringOrNumber(raw.Key),
        order: toOptionalNumber(choice.order ?? raw.Order) ?? index,
        label: String(
          choice.description ??
            raw.Description ??
            choice.displayText ??
            raw.DisplayText ??
            raw.Label ??
            raw.Text ??
            "",
        ),
        branchTargetId: branchInformation?.TargetQuestionId,
        raw,
      };
    })
    .sort(
      (firstOption, secondOption) => firstOption.order - secondOption.order,
    );
}

function normalizeValidation(
  configuration: Record<string, unknown>,
): FormValidation | undefined {
  const rawValidation = configuration.Validation;
  if (!rawValidation || typeof rawValidation !== "object") {
    return undefined;
  }
  const validation = rawValidation as Record<string, unknown>;
  return {
    rule: validation.rule as number | string | undefined,
    minimum: toOptionalNumber(
      validation.numberMinBoundary ?? validation.minimum,
    ),
    maximum: toOptionalNumber(
      validation.numberMaxBoundary ?? validation.maximum,
    ),
    text:
      typeof validation.textInput === "string"
        ? validation.textInput
        : undefined,
    raw: validation,
  };
}

function normalizeFileUploadRules(
  configuration: Record<string, unknown>,
): FileUploadRules | undefined {
  const isFileUpload = Object.keys(configuration).some((key) =>
    /file|upload|extension|size/i.test(key),
  );
  if (!isFileUpload) {
    return undefined;
  }

  const rawExtensions =
    configuration.FileType ??
    configuration.FileTypes ??
    configuration.AllowedFileTypes ??
    configuration.AllowedExtensions;
  let allowedExtensions = Array.isArray(rawExtensions)
    ? rawExtensions.map(String)
    : typeof rawExtensions === "string"
      ? rawExtensions
          .split(/[;,]/)
          .map((extension) => extension.trim())
          .filter(Boolean)
      : [];
  if (
    rawExtensions &&
    typeof rawExtensions === "object" &&
    !Array.isArray(rawExtensions)
  ) {
    const fileTypeExtensions: Record<string, string[]> = {
      Audio: ["aac", "mp3", "wav"],
      Excel: ["xls", "xlsx"],
      Image: ["gif", "jpeg", "jpg", "png", "svg", "tif", "tiff"],
      PDF: ["pdf"],
      PowerPoint: ["ppt", "pptx"],
      Video: ["avi", "mov", "mp4", "wmv"],
      Word: ["doc", "docx"],
    };
    allowedExtensions = Object.entries(rawExtensions)
      .filter(([, enabled]) => enabled === true)
      .flatMap(([fileType]) => fileTypeExtensions[fileType] ?? []);
  }

  return {
    allowedExtensions,
    maximumFileCount: toOptionalNumber(
      configuration.FileLimit ??
        configuration.MaxFileCount ??
        configuration.MaximumFileCount,
    ),
    maximumFileSizeMegabytes: toOptionalNumber(
      configuration.FileSize ??
        configuration.MaxFileSize ??
        configuration.MaximumFileSize,
    ),
    raw: configuration,
  };
}

function normalizeQuestion(rawQuestion: RawQuestion): FormQuestion {
  const deserializedQuestionInformation =
    rawQuestion.deserializedQuestionInfo &&
    typeof rawQuestion.deserializedQuestionInfo === "object"
      ? (rawQuestion.deserializedQuestionInfo as Record<string, unknown>)
      : undefined;
  const configuration = parseObject(
    rawQuestion.questionInfo ?? deserializedQuestionInformation,
    `question ${rawQuestion.id}`,
  );
  return {
    id: rawQuestion.id,
    groupId: rawQuestion.groupId,
    order: rawQuestion.order ?? 0,
    title: rawQuestion.title ?? "",
    subtitle: rawQuestion.subtitle,
    type: rawQuestion.type,
    required: Boolean(
      rawQuestion.isRequired ??
      rawQuestion.required ??
      configuration.Required ??
      configuration.IsRequired,
    ),
    options: normalizeOptions(rawQuestion, configuration),
    rows: [],
    validation: normalizeValidation(configuration),
    fileUpload:
      rawQuestion.type === "Question.FileUpload"
        ? (normalizeFileUploadRules(configuration) ?? {
            allowedExtensions: [],
            raw: configuration,
          })
        : undefined,
    configuration,
    raw: rawQuestion,
  };
}

export function normalizeForm(
  value: unknown,
  responsePageUrl: string,
): FormDefinition {
  if (!value || typeof value !== "object") {
    throw new MicrosoftFormsValidationError(
      "Microsoft Forms returned an invalid form payload.",
    );
  }

  const response = value as Record<string, unknown>;
  const rawForm = (response.form ??
    (Array.isArray(response.value) ? response.value[0] : undefined) ??
    value) as RawForm;
  if (!rawForm.id) {
    throw new MicrosoftFormsValidationError(
      "Microsoft Forms payload does not contain a form ID.",
    );
  }

  const allQuestions = (rawForm.questions ?? [])
    .map(normalizeQuestion)
    .sort(
      (firstQuestion, secondQuestion) =>
        firstQuestion.order - secondQuestion.order,
    );
  const questionsById = new Map(
    allQuestions.map((question) => [question.id, question]),
  );
  const topLevelQuestions = allQuestions.filter(
    (question) => !question.groupId,
  );

  for (const question of allQuestions) {
    if (!question.groupId) {
      continue;
    }
    const parentQuestion = questionsById.get(question.groupId);
    if (parentQuestion) {
      question.options = parentQuestion.options;
      question.required = parentQuestion.required || question.required;
      parentQuestion.rows.push(question);
    }
  }

  let ordinal = 0;
  for (const question of topLevelQuestions) {
    if (question.type !== "Question.MatrixChoice") {
      ordinal += 1;
      question.ordinal = ordinal;
    }
  }

  return {
    id: rawForm.id,
    title: rawForm.title ?? "",
    description: rawForm.description,
    ownerId: rawForm.ownerId ?? "",
    tenantId: rawForm.ownerTenantId ?? "",
    responsePageUrl,
    settings: parseObject(
      rawForm.settings ??
        (rawForm.deserializedSettings as Record<string, unknown> | undefined),
      "form settings",
    ),
    questions: topLevelQuestions,
    raw: rawForm,
  };
}
