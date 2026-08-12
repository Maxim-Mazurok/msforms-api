import { MicrosoftFormsValidationError } from "./errors.js";
import type {
  AnswerValidationIssue,
  AnswerValidationResult,
  FormAnswer,
  FormAnswers,
  FormDefinition,
  FormOption,
  FormQuestion,
  MatrixSelection,
  SerializedAnswer,
  UploadedFile,
} from "./types.js";

const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const urlPattern = /^https?:\/\/\S+$/i;
const datePattern = /^\d{4}-\d{2}-\d{2}$/;

function allQuestions(form: FormDefinition): FormQuestion[] {
  return form.questions.flatMap((question) => [question, ...question.rows]);
}

function isEmptyAnswer(answer: FormAnswer | undefined): boolean {
  return (
    answer === undefined ||
    answer === null ||
    answer === "" ||
    (Array.isArray(answer) && answer.length === 0)
  );
}

function answerStrings(answer: FormAnswer | undefined): string[] {
  if (typeof answer === "string" || typeof answer === "number") {
    return [String(answer)];
  }
  if (Array.isArray(answer)) {
    return answer.filter((entry): entry is string => typeof entry === "string");
  }
  return [];
}

function selectedBranchTarget(
  question: FormQuestion,
  answer: FormAnswer | undefined,
): string | undefined {
  const selectedValues = answerStrings(answer);
  for (const selectedValue of selectedValues) {
    const option = question.options.find(
      (candidate) =>
        candidate.id === selectedValue ||
        candidate.label === selectedValue ||
        String(candidate.key) === selectedValue,
    );
    if (option?.branchTargetId) {
      return option.branchTargetId;
    }
  }
  return undefined;
}

export function resolveVisibleQuestionIds(
  form: FormDefinition,
  answers: FormAnswers,
): string[] {
  const questions = form.questions;
  const questionIndexes = new Map(
    questions.map((question, index) => [question.id, index]),
  );
  const visibleQuestionIds: string[] = [];
  const visitedQuestionIds = new Set<string>();

  let index = 0;
  while (index < questions.length) {
    const question = questions[index];
    if (!question || visitedQuestionIds.has(question.id)) {
      break;
    }
    visitedQuestionIds.add(question.id);
    visibleQuestionIds.push(question.id);
    visibleQuestionIds.push(...question.rows.map((row) => row.id));

    const branchTargetId = selectedBranchTarget(question, answers[question.id]);
    if (!branchTargetId) {
      index += 1;
      continue;
    }
    if (
      branchTargetId === "EndOfForm" ||
      branchTargetId === "Submit" ||
      branchTargetId === "-1"
    ) {
      break;
    }
    const branchTargetIndex = questionIndexes.get(branchTargetId);
    index = branchTargetIndex ?? index + 1;
  }

  return visibleQuestionIds;
}

function findOption(
  question: FormQuestion,
  value: string | number,
): FormOption | undefined {
  const stringValue = String(value);
  return question.options.find(
    (option) =>
      String(option.id) === stringValue ||
      option.label === stringValue ||
      String(option.key) === stringValue,
  );
}

function normalizeChoiceAnswer(
  question: FormQuestion,
  answer: FormAnswer,
): string | string[] {
  const values = Array.isArray(answer) ? answer : [answer];
  const normalizedValues = values.map((value) => {
    if (typeof value !== "string" && typeof value !== "number") {
      throw new MicrosoftFormsValidationError(
        `Question "${question.title}" requires option labels or IDs.`,
      );
    }
    const option = findOption(question, value);
    const allowsOtherAnswer = Boolean(question.configuration.AllowOtherAnswer);
    if (!option && !allowsOtherAnswer) {
      throw new MicrosoftFormsValidationError(
        `Unknown option "${String(value)}" for question "${question.title}".`,
      );
    }
    return option?.label ?? String(value);
  });

  const choiceType = Number(question.configuration.ChoiceType);
  return choiceType === 2 ? normalizedValues : (normalizedValues[0] ?? "");
}

function normalizeMatrixSelection(
  question: FormQuestion,
  answer: FormAnswer,
): { id: string | number; key: string | number } {
  const selection =
    typeof answer === "string" || typeof answer === "number"
      ? { label: String(answer) }
      : (answer as MatrixSelection);
  const option = question.options.find(
    (candidate) =>
      String(candidate.id) === String(selection.id) ||
      String(candidate.key) === String(selection.key) ||
      candidate.label === selection.label,
  );
  if (option?.id === undefined || option.key === undefined) {
    throw new MicrosoftFormsValidationError(
      `Unknown matrix option for row "${question.title}".`,
    );
  }
  return { id: option.id, key: option.key };
}

function normalizeRankingAnswer(
  question: FormQuestion,
  answer: FormAnswer,
): Array<{ id: string | number; order: number; answerOrder: number }> {
  if (
    !Array.isArray(answer) ||
    !answer.every((value): value is string => typeof value === "string")
  ) {
    throw new MicrosoftFormsValidationError(
      `Question "${question.title}" requires an ordered array of option labels or IDs.`,
    );
  }
  if (answer.length !== question.options.length) {
    throw new MicrosoftFormsValidationError(
      `Question "${question.title}" requires all ${question.options.length} options.`,
    );
  }

  const usedOptionIds = new Set<string>();
  return answer.map((value, answerOrder) => {
    const option = findOption(question, value);
    const optionId = option?.id;
    if (
      !option ||
      optionId === undefined ||
      usedOptionIds.has(String(optionId))
    ) {
      throw new MicrosoftFormsValidationError(
        `Ranking answer for "${question.title}" contains an unknown or duplicate option.`,
      );
    }
    usedOptionIds.add(String(optionId));
    return { id: optionId, order: option.order, answerOrder };
  });
}

function isUploadedFileArray(answer: FormAnswer): answer is UploadedFile[] {
  return (
    Array.isArray(answer) &&
    answer.every(
      (entry) =>
        typeof entry === "object" &&
        entry !== null &&
        "uploadSessionUrl" in entry &&
        "id" in entry,
    )
  );
}

function serializeQuestionAnswer(
  question: FormQuestion,
  answer: FormAnswer,
): unknown {
  switch (question.type) {
    case "Question.TextField":
      return String(answer);
    case "Question.Rating":
    case "Question.NPS": {
      const numericAnswer = Number(answer);
      if (!Number.isFinite(numericAnswer)) {
        throw new MicrosoftFormsValidationError(
          `Question "${question.title}" requires a number.`,
        );
      }
      const minimum =
        question.type === "Question.NPS"
          ? 0
          : Number(question.configuration.MinRating ?? 1);
      const maximum =
        question.type === "Question.NPS"
          ? 10
          : Number(question.configuration.Length ?? 5);
      if (
        !Number.isInteger(numericAnswer) ||
        numericAnswer < minimum ||
        numericAnswer > maximum
      ) {
        throw new MicrosoftFormsValidationError(
          `Question "${question.title}" requires a whole number from ${minimum} to ${maximum}.`,
        );
      }
      return numericAnswer;
    }
    case "Question.Choice": {
      const normalizedAnswer = normalizeChoiceAnswer(question, answer);
      return Array.isArray(normalizedAnswer)
        ? JSON.stringify(normalizedAnswer)
        : normalizedAnswer;
    }
    case "Question.Ranking":
      return normalizeRankingAnswer(question, answer);
    case "Question.FileUpload":
      if (!isUploadedFileArray(answer)) {
        throw new MicrosoftFormsValidationError(
          `Question "${question.title}" requires uploaded-file objects returned by uploadFile().`,
        );
      }
      if (
        question.fileUpload?.maximumFileCount !== undefined &&
        answer.length > question.fileUpload.maximumFileCount
      ) {
        throw new MicrosoftFormsValidationError(
          `Question "${question.title}" accepts at most ${question.fileUpload.maximumFileCount} files.`,
        );
      }
      return JSON.stringify(answer);
    case "Question.DateTime":
      if (
        typeof answer !== "string" ||
        !datePattern.test(answer) ||
        Number.isNaN(Date.parse(`${answer}T00:00:00Z`))
      ) {
        throw new MicrosoftFormsValidationError(
          `Question "${question.title}" requires a date in YYYY-MM-DD format.`,
        );
      }
      return answer;
    case "Question.MatrixChoice":
      return normalizeMatrixSelection(question, answer);
    default:
      return answer;
  }
}

function validateTextRule(
  question: FormQuestion,
  answer: string,
): string | undefined {
  if (answer.length > 4_000) {
    return "Maximum length is 4000.";
  }
  const validation = question.validation;
  if (!validation) {
    return undefined;
  }

  const rule = validation.rule;
  const numericAnswer = Number(answer);
  switch (rule) {
    case 9:
    case "MaxLength":
      return validation.maximum !== undefined &&
        answer.length > validation.maximum
        ? `Maximum length is ${validation.maximum}.`
        : undefined;
    case 10:
    case "MinLength":
      return validation.minimum !== undefined &&
        answer.length < validation.minimum
        ? `Minimum length is ${validation.minimum}.`
        : undefined;
    case 11:
    case "Email":
      return emailPattern.test(answer.trim())
        ? undefined
        : "Answer must be an email address.";
    case 14:
    case "Url":
      return urlPattern.test(answer.trim())
        ? undefined
        : "Answer must be an HTTP or HTTPS URL.";
    case 12:
    case "Contains":
      return validation.text && !answer.includes(validation.text)
        ? `Answer must contain "${validation.text}".`
        : undefined;
    case 13:
    case "NotContains":
      return validation.text && answer.includes(validation.text)
        ? `Answer must not contain "${validation.text}".`
        : undefined;
    case 0:
    case "IsNumber":
      return Number.isFinite(numericAnswer)
        ? undefined
        : "Answer must be a number.";
    case 3:
    case "Less":
      return validation.maximum !== undefined &&
        !(numericAnswer < validation.maximum)
        ? `Answer must be less than ${validation.maximum}.`
        : undefined;
    case 4:
    case "LessOrEqual":
      return validation.maximum !== undefined &&
        !(numericAnswer <= validation.maximum)
        ? `Answer must be at most ${validation.maximum}.`
        : undefined;
    case 1:
    case "Greater":
      return validation.minimum !== undefined &&
        !(numericAnswer > validation.minimum)
        ? `Answer must be greater than ${validation.minimum}.`
        : undefined;
    case 2:
    case "GreaterOrEqual":
      return validation.minimum !== undefined &&
        !(numericAnswer >= validation.minimum)
        ? `Answer must be at least ${validation.minimum}.`
        : undefined;
    case 5:
    case "Equal":
      return validation.minimum !== undefined &&
        numericAnswer !== validation.minimum
        ? `Answer must equal ${validation.minimum}.`
        : undefined;
    case 6:
    case "NotEqual":
      return validation.minimum !== undefined &&
        numericAnswer === validation.minimum
        ? `Answer must not equal ${validation.minimum}.`
        : undefined;
    case 7:
    case "Between":
      return validation.minimum !== undefined &&
        validation.maximum !== undefined &&
        !(
          numericAnswer >= validation.minimum &&
          numericAnswer <= validation.maximum
        )
        ? `Answer must be between ${validation.minimum} and ${validation.maximum}.`
        : undefined;
    case 8:
    case "NotBetween":
      return validation.minimum !== undefined &&
        validation.maximum !== undefined &&
        numericAnswer >= validation.minimum &&
        numericAnswer <= validation.maximum
        ? `Answer must not be between ${validation.minimum} and ${validation.maximum}.`
        : undefined;
    case 15:
    case "WholeNumber":
      return Number.isInteger(numericAnswer) && numericAnswer >= 0
        ? undefined
        : "Answer must be a non-negative whole number.";
    default:
      return undefined;
  }
}

export function validateAndSerializeAnswers(
  form: FormDefinition,
  answers: FormAnswers,
): AnswerValidationResult {
  const questions = allQuestions(form);
  const questionsById = new Map(
    questions.map((question) => [question.id, question]),
  );
  const visibleQuestionIds = resolveVisibleQuestionIds(form, answers);
  const visibleQuestionIdSet = new Set(visibleQuestionIds);
  const issues: AnswerValidationIssue[] = [];
  const serializedAnswers: SerializedAnswer[] = [];

  for (const answerQuestionId of Object.keys(answers)) {
    if (!questionsById.has(answerQuestionId)) {
      issues.push({
        questionId: answerQuestionId,
        questionTitle: "",
        code: "unknown-question",
        message: `Unknown question ID: ${answerQuestionId}`,
      });
    }
  }

  for (const question of questions) {
    if (!visibleQuestionIdSet.has(question.id)) {
      continue;
    }
    if (question.type === "Question.MatrixChoiceGroup") {
      continue;
    }
    const answer = answers[question.id];
    if (question.required && isEmptyAnswer(answer)) {
      issues.push({
        questionId: question.id,
        questionTitle: question.title,
        code: "missing-required",
        message: `Required question "${question.title}" is unanswered.`,
      });
      continue;
    }
    if (isEmptyAnswer(answer)) {
      if (question.type === "Question.MatrixChoice") {
        continue;
      }
      const emptyAnswer =
        question.type === "Question.Ranking"
          ? []
          : question.type === "Question.FileUpload"
            ? JSON.stringify([])
            : null;
      serializedAnswers.push({
        questionId: question.id,
        answer1: emptyAnswer,
      });
      continue;
    }
    if (answer === undefined) {
      continue;
    }

    try {
      if (
        question.type === "Question.TextField" &&
        typeof answer === "string"
      ) {
        const validationMessage = validateTextRule(question, answer);
        if (validationMessage) {
          issues.push({
            questionId: question.id,
            questionTitle: question.title,
            code: "validation-failed",
            message: validationMessage,
          });
          continue;
        }
      }
      serializedAnswers.push({
        questionId: question.id,
        answer1: serializeQuestionAnswer(question, answer),
      });
    } catch (error) {
      issues.push({
        questionId: question.id,
        questionTitle: question.title,
        code: "invalid-answer",
        message: error instanceof Error ? error.message : String(error),
      });
    }
  }

  return {
    valid: issues.length === 0,
    visibleQuestionIds,
    issues,
    serializedAnswers,
  };
}
