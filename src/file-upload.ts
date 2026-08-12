import { basename, extname, resolve } from "node:path";
import { stat } from "node:fs/promises";
import { MicrosoftFormsValidationError } from "./errors.js";
import {
  openAuthenticatedFormPage,
  type OpenFormPage,
} from "./forms-session.js";
import type {
  FormDefinition,
  MicrosoftFormsClientOptions,
  UploadedFile,
} from "./types.js";

interface UploadSessionResponse {
  badger?: string | null;
  value?: string;
}

async function clearPersistedUploadAnswer(
  page: OpenFormPage["page"],
  formId: string,
  questionId: string,
): Promise<void> {
  await page.addInitScript(
    ({ currentFormId, currentQuestionId }) => {
      const storageKey = Object.keys(localStorage).find((key) =>
        key.startsWith(`officeforms.answermap.${currentFormId}.`),
      );
      if (!storageKey) {
        return;
      }
      const storedValue = localStorage.getItem(storageKey);
      if (!storedValue) {
        return;
      }
      const answers = JSON.parse(storedValue) as Array<{
        Answer?: unknown;
        QuestionId?: string;
      }>;
      const answer = answers.find(
        (candidate) => candidate.QuestionId === currentQuestionId,
      );
      if (!answer || answer.Answer === undefined) {
        return;
      }
      delete answer.Answer;
      localStorage.setItem(storageKey, JSON.stringify(answers));
    },
    { currentFormId: formId, currentQuestionId: questionId },
  );
  await page.reload({ waitUntil: "domcontentloaded" });
}

function findQuestion(form: FormDefinition, questionId: string) {
  return form.questions
    .flatMap((question) => [question, ...question.rows])
    .find((question) => question.id === questionId);
}

async function validateUploadFile(
  form: FormDefinition,
  questionId: string,
  filePath: string,
): Promise<string> {
  const question = findQuestion(form, questionId);
  if (!question || question.type !== "Question.FileUpload") {
    throw new MicrosoftFormsValidationError(
      `Question ${questionId} is not a file-upload question.`,
    );
  }

  const absoluteFilePath = resolve(filePath);
  const fileInformation = await stat(absoluteFilePath);
  if (!fileInformation.isFile()) {
    throw new MicrosoftFormsValidationError(
      `Upload path is not a file: ${absoluteFilePath}`,
    );
  }

  const allowedExtensions = question.fileUpload?.allowedExtensions ?? [];
  const fileExtension = extname(absoluteFilePath)
    .replace(/^\./, "")
    .toLowerCase();
  if (
    allowedExtensions.length > 0 &&
    !allowedExtensions.some(
      (extension) =>
        extension.replace(/^\./, "").toLowerCase() === fileExtension,
    )
  ) {
    throw new MicrosoftFormsValidationError(
      `File extension .${fileExtension} is not allowed. Expected: ${allowedExtensions.join(", ")}.`,
    );
  }

  const maximumFileSizeMegabytes =
    question.fileUpload?.maximumFileSizeMegabytes;
  if (
    maximumFileSizeMegabytes !== undefined &&
    fileInformation.size > maximumFileSizeMegabytes * 1_024 * 1_024
  ) {
    throw new MicrosoftFormsValidationError(
      `File exceeds the ${maximumFileSizeMegabytes} MB limit.`,
    );
  }
  return absoluteFilePath;
}

export async function uploadFormFile(
  formReference: string,
  form: FormDefinition,
  questionId: string,
  filePath: string,
  options: MicrosoftFormsClientOptions,
): Promise<UploadedFile> {
  const absoluteFilePath = await validateUploadFile(form, questionId, filePath);
  const question = findQuestion(form, questionId);
  if (!question) {
    throw new MicrosoftFormsValidationError(`Unknown question: ${questionId}`);
  }

  const openFormPage = await openAuthenticatedFormPage(formReference, options);
  try {
    await clearPersistedUploadAnswer(
      openFormPage.page,
      openFormPage.session.formId,
      questionId,
    );
    await openFormPage.page.waitForSelector(
      '[data-automation-id="questionItem"]',
      { timeout: 30_000 },
    );
    const questionItem = openFormPage.page
      .locator('[data-automation-id="questionItem"]')
      .filter({ hasText: question.title })
      .first();
    const uploadButton = questionItem.locator(
      '[data-automation-id="fileUploadButton"]',
    );
    if ((await uploadButton.count()) === 0) {
      const visibleQuestionTitles = await openFormPage.page
        .locator(
          '[data-automation-id="questionItem"] [data-automation-id="questionTitle"]',
        )
        .allInnerTexts();
      throw new MicrosoftFormsValidationError(
        `File question "${question.title}" is not currently visible. Visible questions: ${visibleQuestionTitles.join(
          " | ",
        )}. Supply answers that reveal its branch before uploading.`,
      );
    }
    let uploadSession: UploadSessionResponse | undefined;
    const uploadSessionResponsePromise = openFormPage.page.waitForResponse(
      (response) =>
        response.url().includes("/CreateUploadSession") && response.ok(),
      { timeout: 2 * 60 * 1_000 },
    );
    const sharePointResponsePromise = openFormPage.page.waitForResponse(
      (response) =>
        response.request().method() === "PUT" &&
        response.url().includes("/uploadSession") &&
        response.ok(),
      { timeout: 5 * 60 * 1_000 },
    );
    const fileChooserPromise = openFormPage.page.waitForEvent("filechooser");
    await uploadButton.click();
    const fileChooser = await fileChooserPromise;
    await fileChooser.setFiles(absoluteFilePath);

    const uploadSessionResponse = await uploadSessionResponsePromise;
    uploadSession =
      (await uploadSessionResponse.json()) as UploadSessionResponse;
    await sharePointResponsePromise;
    const uploadSessionPathMatch = new URL(
      uploadSession.value ?? "",
    ).pathname.match(/\/items\/([^/]+)\/uploadSession$/i);
    const sharePointItemId = uploadSessionPathMatch?.[1]
      ? decodeURIComponent(uploadSessionPathMatch[1])
      : undefined;

    if (!uploadSession.value || !sharePointItemId) {
      throw new MicrosoftFormsValidationError(
        "Microsoft Forms upload completed without required file metadata.",
      );
    }

    return {
      name: basename(absoluteFilePath),
      uploadSessionUrl: uploadSession.value,
      link: uploadSession.value,
      badgerToken: uploadSession.badger ?? null,
      time: Date.now(),
      status: 3,
      id: sharePointItemId,
    };
  } finally {
    await openFormPage.browserContext.close();
  }
}
