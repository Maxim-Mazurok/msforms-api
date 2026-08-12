export type JsonPrimitive = boolean | number | string | null;
export type JsonValue =
  JsonPrimitive | JsonValue[] | { [key: string]: JsonValue };

export interface RawChoice {
  id?: string | number;
  description?: string;
  displayText?: string;
  order?: number;
  key?: string | number;
  BranchInfo?: {
    TargetQuestionId?: string;
  };
  [key: string]: unknown;
}

export interface RawQuestion {
  id: string;
  title?: string;
  subtitle?: string;
  type: string;
  order?: number;
  groupId?: string;
  isRequired?: boolean;
  required?: boolean;
  questionInfo?: string | Record<string, unknown>;
  choices?: RawChoice[];
  [key: string]: unknown;
}

export interface RawForm {
  id: string;
  title?: string;
  description?: string;
  ownerId?: string;
  ownerTenantId?: string;
  settings?: string | Record<string, unknown>;
  questions?: RawQuestion[];
  [key: string]: unknown;
}

export interface FormOption {
  id?: string | number;
  key?: string | number;
  order: number;
  label: string;
  branchTargetId?: string;
  raw: Record<string, unknown>;
}

export interface FormValidation {
  rule?: number | string;
  minimum?: number;
  maximum?: number;
  text?: string;
  raw: Record<string, unknown>;
}

export interface FileUploadRules {
  allowedExtensions: string[];
  maximumFileCount?: number;
  maximumFileSizeMegabytes?: number;
  raw: Record<string, unknown>;
}

export interface FormQuestion {
  id: string;
  groupId?: string;
  order: number;
  ordinal?: number;
  title: string;
  subtitle?: string;
  type: string;
  required: boolean;
  options: FormOption[];
  rows: FormQuestion[];
  validation?: FormValidation;
  fileUpload?: FileUploadRules;
  configuration: Record<string, unknown>;
  raw: RawQuestion;
}

export interface FormDefinition {
  id: string;
  title: string;
  description?: string;
  ownerId: string;
  tenantId: string;
  responsePageUrl: string;
  settings: Record<string, unknown>;
  questions: FormQuestion[];
  raw: RawForm;
}

export interface UploadedFile {
  name: string;
  uploadSessionUrl: string;
  link: string;
  badgerToken: string | null;
  time: number;
  status: number;
  id: string;
}

export interface MatrixSelection {
  id?: string | number;
  key?: string | number;
  label?: string;
}

export type FormAnswer =
  | JsonPrimitive
  | string[]
  | MatrixSelection
  | UploadedFile[]
  | Record<string, MatrixSelection>;

export type FormAnswers = Record<string, FormAnswer>;

export interface AnswerValidationIssue {
  questionId: string;
  questionTitle: string;
  code:
    | "invalid-answer"
    | "invalid-option"
    | "missing-required"
    | "unknown-question"
    | "validation-failed";
  message: string;
}

export interface AnswerValidationResult {
  valid: boolean;
  visibleQuestionIds: string[];
  issues: AnswerValidationIssue[];
  serializedAnswers: SerializedAnswer[];
}

export interface SerializedAnswer {
  questionId: string;
  answer1: unknown;
}

export interface SubmitOptions {
  emailReceiptConsent?: boolean;
  saveResponse?: boolean;
  submitLanguage?: string;
}

export interface SubmissionResult {
  responseId: string;
  submitDate: string;
  response: Record<string, unknown>;
  savedResponse?: Record<string, unknown>;
}

export interface SavedResponseLink {
  [key: string]: unknown;
}

export interface MicrosoftFormsClientOptions {
  browserProfileDirectory?: string;
  browserChannel?: "chrome" | "msedge";
  headless?: boolean;
  log?: (...messages: unknown[]) => void;
  sessionMaximumAgeMilliseconds?: number;
}

export interface AuthenticateOptions extends Omit<
  MicrosoftFormsClientOptions,
  "headless"
> {
  form?: string;
  timeoutMilliseconds?: number;
}

export interface AuthenticatedUser {
  displayName?: string;
  email?: string;
  tenantId?: string;
  userId?: string;
}
