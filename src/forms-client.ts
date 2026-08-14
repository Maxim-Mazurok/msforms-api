import { validateAndSerializeAnswers } from "./answers.js";
import {
  MicrosoftFormsApiError,
  MicrosoftFormsAuthenticationError,
  MicrosoftFormsValidationError,
} from "./errors.js";
import { buildFormOperationUrl, formsApiRequest } from "./forms-api.js";
import {
  authenticateFormsUser,
  bootstrapFormsSession,
  type FormsApiSession,
} from "./forms-session.js";
import { uploadFormFile } from "./file-upload.js";
import { normalizeForm } from "./form-normalizer.js";
import { parseFormReference } from "./form-reference.js";
import type {
  AnswerValidationResult,
  AuthenticateOptions,
  AuthenticatedUser,
  FormAnswers,
  FormDefinition,
  MicrosoftFormsClientOptions,
  SavedResponseLink,
  SubmissionResult,
  SubmissionSaveResponseError,
  SubmitOptions,
  UploadedFile,
} from "./types.js";

interface CachedSession {
  acquiredAt: number;
  session: FormsApiSession;
}

function createSubmissionSaveResponseError(
  error: unknown,
): SubmissionSaveResponseError {
  return {
    message: error instanceof Error ? error.message : String(error),
    ...(error instanceof MicrosoftFormsApiError
      ? { status: error.status }
      : {}),
  };
}

export class MicrosoftFormsClient {
  readonly #options: MicrosoftFormsClientOptions;
  readonly #sessions = new Map<string, CachedSession>();

  public constructor(options: MicrosoftFormsClientOptions = {}) {
    this.#options = options;
  }

  public static authenticate(
    options: AuthenticateOptions = {},
  ): Promise<AuthenticatedUser> {
    return authenticateFormsUser(options);
  }

  async #getSession(
    formReference: string,
    forceRefresh = false,
  ): Promise<FormsApiSession> {
    const cachedSession = this.#sessions.get(formReference);
    const sessionMaximumAgeMilliseconds =
      this.#options.sessionMaximumAgeMilliseconds ?? 10 * 60 * 1_000;
    if (
      !forceRefresh &&
      cachedSession &&
      Date.now() - cachedSession.acquiredAt < sessionMaximumAgeMilliseconds
    ) {
      return cachedSession.session;
    }

    const session = await bootstrapFormsSession(formReference, this.#options);
    this.#sessions.set(formReference, {
      acquiredAt: Date.now(),
      session,
    });
    return session;
  }

  async #request<T>(
    formReference: string,
    requestUrl: (session: FormsApiSession) => string,
    requestInit?: RequestInit,
  ): Promise<T> {
    let session = await this.#getSession(formReference);
    try {
      return await formsApiRequest<T>(
        session,
        requestUrl(session),
        requestInit,
      );
    } catch (error) {
      if (!(
        error instanceof MicrosoftFormsAuthenticationError ||
        (error instanceof Error && /HTTP 401|HTTP 403/.test(error.message))
      )) {
        throw error;
      }
      session = await this.#getSession(formReference, true);
      return formsApiRequest<T>(session, requestUrl(session), requestInit);
    }
  }

  public async getForm(formReference: string): Promise<FormDefinition> {
    const session = await this.#getSession(formReference);
    const rawForm = await this.#request<unknown>(
      formReference,
      (currentSession) => currentSession.formApiUrl,
    );
    return normalizeForm(rawForm, session.responsePageUrl);
  }

  public async validateAnswers(
    formReference: string,
    answers: FormAnswers,
  ): Promise<AnswerValidationResult> {
    const form = await this.getForm(formReference);
    return validateAndSerializeAnswers(form, answers);
  }

  public async uploadFile(
    formReference: string,
    questionId: string,
    filePath: string,
  ): Promise<UploadedFile> {
    const form = await this.getForm(formReference);
    return uploadFormFile(
      formReference,
      form,
      questionId,
      filePath,
      this.#options,
    );
  }

  public async submitResponse(
    formReference: string,
    answers: FormAnswers,
    options: SubmitOptions = {},
  ): Promise<SubmissionResult> {
    const form = await this.getForm(formReference);
    const validationResult = validateAndSerializeAnswers(form, answers);
    if (!validationResult.valid) {
      throw new MicrosoftFormsValidationError(
        `Cannot submit invalid answers:\n${validationResult.issues
          .map((issue) => `- ${issue.message}`)
          .join("\n")}`,
      );
    }

    const startDate = new Date().toJSON();
    const submitDate = new Date().toJSON();
    const response = await this.#request<Record<string, unknown>>(
      formReference,
      (session) => `${buildFormOperationUrl(session, "forms")}/responses`,
      {
        method: "POST",
        headers: { "Content-Type": "application/json;charset=UTF-8" },
        body: JSON.stringify({
          startDate,
          submitDate,
          answers: JSON.stringify(validationResult.serializedAnswers),
          submitLanguage: options.submitLanguage,
          emailReceiptConsent: options.emailReceiptConsent,
        }),
      },
    );
    const responseId = String(response.id ?? "");
    if (!responseId) {
      throw new MicrosoftFormsValidationError(
        "Microsoft Forms accepted the submission but returned no response ID.",
      );
    }

    if (!options.saveResponse) {
      return {
        responseId,
        submissionStatus: "submitted",
        submitDate,
        response,
        saveResponseStatus: "not-requested",
      };
    }

    try {
      const savedResponse = await this.saveResponse(
        formReference,
        responseId,
        submitDate,
      );
      return {
        responseId,
        submissionStatus: "submitted",
        submitDate,
        response,
        saveResponseStatus: "saved",
        savedResponse,
      };
    } catch (error) {
      return {
        responseId,
        submissionStatus: "submitted",
        submitDate,
        response,
        saveResponseStatus: "failed",
        saveResponseError: createSubmissionSaveResponseError(error),
      };
    }
  }

  public async saveResponse(
    formReference: string,
    responseId: string,
    submitDate: string,
  ): Promise<Record<string, unknown>> {
    return this.#request<Record<string, unknown>>(
      formReference,
      (session) =>
        `${session.formsHost}/formapi/api/${encodeURIComponent(
          session.ownerTenantId,
        )}/users/${encodeURIComponent(session.ownerId)}/saveResponseLink`,
      {
        method: "POST",
        headers: { "Content-Type": "application/json;charset=UTF-8" },
        body: JSON.stringify({
          submitDate,
          formId: parseFormReference(formReference).formId,
          responseId,
        }),
      },
    );
  }

  public async listSavedResponses(
    formReference: string,
  ): Promise<SavedResponseLink[]> {
    const response = await this.#request<
      SavedResponseLink[] | { value?: SavedResponseLink[] }
    >(
      formReference,
      (session) =>
        `${session.formsHost}/formapi/api/${encodeURIComponent(
          session.responderTenantId,
        )}/users/${encodeURIComponent(session.responderId)}/getResponseLinks`,
    );
    return Array.isArray(response) ? response : (response.value ?? []);
  }
}
