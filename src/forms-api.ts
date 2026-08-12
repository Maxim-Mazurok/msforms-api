import {
  MicrosoftFormsApiError,
  MicrosoftFormsValidationError,
} from "./errors.js";
import type { FormsApiSession } from "./forms-session.js";

function formsRequestHeaders(
  session: FormsApiSession,
  additionalHeaders?: HeadersInit,
): Headers {
  const headers = new Headers(additionalHeaders);
  headers.set("Accept", "application/json");
  headers.set("__RequestVerificationToken", session.antiForgeryToken);
  headers.set("Cookie", session.cookieHeader);
  headers.set("X-UserSessionId", session.serverSessionId);
  if (session.multiUserIdentifier) {
    headers.set("X-Ms-Form-Muid", session.multiUserIdentifier);
  }
  return headers;
}

export async function formsApiRequest<T>(
  session: FormsApiSession,
  requestUrl: string,
  requestInit: RequestInit = {},
): Promise<T> {
  const response = await fetch(requestUrl, {
    ...requestInit,
    headers: formsRequestHeaders(session, requestInit.headers),
  });
  const responseText = await response.text();
  if (!response.ok) {
    throw new MicrosoftFormsApiError(response.status, responseText, requestUrl);
  }
  if (!responseText) {
    return {} as T;
  }
  try {
    return JSON.parse(responseText) as T;
  } catch (error) {
    throw new MicrosoftFormsValidationError(
      `Microsoft Forms returned invalid JSON for ${requestUrl}: ${
        error instanceof Error ? error.message : String(error)
      }`,
    );
  }
}

export function buildFormOperationUrl(
  session: FormsApiSession,
  collection: "forms" | "runtimeForms",
): string {
  return `${session.formsHost}/formapi/api/${encodeURIComponent(
    session.tenantId,
  )}/users/${encodeURIComponent(session.ownerId)}/${collection}('${encodeURIComponent(
    session.formId,
  )}')`;
}
