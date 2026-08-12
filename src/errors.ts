export class MicrosoftFormsError extends Error {
  public constructor(message: string) {
    super(message);
    this.name = "MicrosoftFormsError";
  }
}

export class MicrosoftFormsAuthenticationError extends MicrosoftFormsError {
  public constructor(message: string) {
    super(message);
    this.name = "MicrosoftFormsAuthenticationError";
  }
}

export class MicrosoftFormsApiError extends MicrosoftFormsError {
  public readonly status: number;
  public readonly responseBody: string;

  public constructor(status: number, responseBody: string, requestUrl: string) {
    super(
      `Microsoft Forms request failed with HTTP ${status}: ${requestUrl}\n${responseBody}`,
    );
    this.name = "MicrosoftFormsApiError";
    this.status = status;
    this.responseBody = responseBody;
  }
}

export class MicrosoftFormsValidationError extends MicrosoftFormsError {
  public constructor(message: string) {
    super(message);
    this.name = "MicrosoftFormsValidationError";
  }
}
