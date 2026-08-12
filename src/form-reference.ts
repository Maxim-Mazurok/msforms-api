import { MicrosoftFormsValidationError } from "./errors.js";

export const defaultFormsResponsePageUrl =
  "https://forms.cloud.microsoft/Pages/ResponsePage.aspx";

export function parseFormReference(reference: string): {
  formId: string;
  responsePageUrl: string;
} {
  const trimmedReference = reference.trim();
  if (!trimmedReference) {
    throw new MicrosoftFormsValidationError("Form reference cannot be empty.");
  }

  if (!/^https?:\/\//i.test(trimmedReference)) {
    const responsePageUrl = new URL(defaultFormsResponsePageUrl);
    responsePageUrl.searchParams.set("id", trimmedReference);
    return {
      formId: trimmedReference,
      responsePageUrl: responsePageUrl.toString(),
    };
  }

  let url: URL;
  try {
    url = new URL(trimmedReference);
  } catch {
    throw new MicrosoftFormsValidationError(
      `Invalid Microsoft Forms URL: ${trimmedReference}`,
    );
  }

  const formId =
    url.searchParams.get("id") ??
    url.searchParams.get("FormId") ??
    url.searchParams.get("formId");
  if (!formId) {
    throw new MicrosoftFormsValidationError(
      "Microsoft Forms URL must contain an id query parameter.",
    );
  }

  const responsePageUrl = new URL(defaultFormsResponsePageUrl);
  responsePageUrl.searchParams.set("id", formId);
  return { formId, responsePageUrl: responsePageUrl.toString() };
}
