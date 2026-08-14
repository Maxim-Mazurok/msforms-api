import type { BrowserContext, Page } from "playwright";
import { launchFormsBrowserContext } from "./browser-runtime.js";
import { MicrosoftFormsAuthenticationError } from "./errors.js";
import { parseFormReference } from "./form-reference.js";
import type {
  AuthenticateOptions,
  AuthenticatedUser,
  MicrosoftFormsClientOptions,
} from "./types.js";

interface FormsUserInformation {
  DisplayName?: string;
  Email?: string;
  TenantId?: string;
  UserId?: string;
}

interface OfficeFormServerInformation {
  antiForgeryToken?: string;
  prefetchFormUrl?: string;
  prefetchFormWithResponsesUrl?: string;
  serverSessionId?: string;
  startupStatus?: number;
  userInfo?: FormsUserInformation;
}

export interface FormsApiSession {
  antiForgeryToken: string;
  cookieHeader: string;
  formApiUrl: string;
  formId: string;
  formsHost: string;
  multiUserIdentifier: string;
  ownerId: string;
  ownerTenantId: string;
  responsePageUrl: string;
  responderId: string;
  responderTenantId: string;
  serverSessionId: string;
}

export interface OpenFormPage {
  browserContext: BrowserContext;
  page: Page;
  serverInformation: OfficeFormServerInformation;
  session: FormsApiSession;
}

function responseTimeoutMilliseconds(headless: boolean): number {
  return headless ? 30_000 : 5 * 60 * 1_000;
}

async function waitForFormsServerInformation(
  page: Page,
  headless: boolean,
): Promise<OfficeFormServerInformation> {
  try {
    await page.waitForFunction(
      () => {
        const serverInformation = (
          window as Window & {
            OfficeFormServerInfo?: OfficeFormServerInformation;
          }
        ).OfficeFormServerInfo;
        return Boolean(
          serverInformation?.prefetchFormUrl &&
          serverInformation?.antiForgeryToken &&
          serverInformation?.userInfo?.UserId,
        );
      },
      undefined,
      { timeout: responseTimeoutMilliseconds(headless) },
    );
  } catch (error) {
    throw new MicrosoftFormsAuthenticationError(
      headless
        ? `Microsoft Forms authentication is missing or expired. Run \`msforms-api auth\`. ${
            error instanceof Error ? error.message : String(error)
          }`
        : `Microsoft Forms login did not complete within five minutes. ${
            error instanceof Error ? error.message : String(error)
          }`,
    );
  }

  return page.evaluate(() => {
    const serverInformation = (
      window as Window & {
        OfficeFormServerInfo?: OfficeFormServerInformation;
      }
    ).OfficeFormServerInfo;
    if (!serverInformation) {
      throw new Error("OfficeFormServerInfo is unavailable.");
    }
    return {
      antiForgeryToken: serverInformation.antiForgeryToken,
      prefetchFormUrl: serverInformation.prefetchFormUrl,
      prefetchFormWithResponsesUrl:
        serverInformation.prefetchFormWithResponsesUrl,
      serverSessionId: serverInformation.serverSessionId,
      startupStatus: serverInformation.startupStatus,
      userInfo: serverInformation.userInfo,
    };
  });
}

function parseFormApiUrl(prefetchFormUrl: string): {
  formApiUrl: string;
  ownerId: string;
  ownerTenantId: string;
} {
  const url = new URL(prefetchFormUrl);
  const pathMatch = url.pathname.match(
    /^\/formapi\/api\/([^/]+)\/users\/([^/]+)\/light\/runtimeForms\(/,
  );
  if (!pathMatch?.[1] || !pathMatch[2]) {
    throw new MicrosoftFormsAuthenticationError(
      `Unsupported Microsoft Forms API URL: ${prefetchFormUrl}`,
    );
  }
  return {
    formApiUrl: prefetchFormUrl,
    ownerId: decodeURIComponent(pathMatch[2]),
    ownerTenantId: decodeURIComponent(pathMatch[1]),
  };
}

async function createSession(
  browserContext: BrowserContext,
  serverInformation: OfficeFormServerInformation,
  formId: string,
  responsePageUrl: string,
): Promise<FormsApiSession> {
  const { antiForgeryToken, prefetchFormUrl, serverSessionId, userInfo } =
    serverInformation;
  if (
    !antiForgeryToken ||
    !prefetchFormUrl ||
    !serverSessionId ||
    !userInfo?.UserId ||
    !userInfo.TenantId
  ) {
    throw new MicrosoftFormsAuthenticationError(
      "Microsoft Forms session bootstrap did not return required authentication data.",
    );
  }

  const cookies = await browserContext.cookies("https://forms.cloud.microsoft");
  const cookieHeader = cookies
    .map((cookie) => `${cookie.name}=${cookie.value}`)
    .join("; ");
  const multiUserIdentifier =
    cookies.find((cookie) => cookie.name === "MUID")?.value ?? "";
  const { formApiUrl, ownerId, ownerTenantId } =
    parseFormApiUrl(prefetchFormUrl);

  return {
    antiForgeryToken,
    cookieHeader,
    formApiUrl,
    formId,
    formsHost: new URL(prefetchFormUrl).origin,
    multiUserIdentifier,
    ownerId,
    ownerTenantId,
    responsePageUrl,
    responderId: userInfo.UserId,
    responderTenantId: userInfo.TenantId,
    serverSessionId,
  };
}

export async function openAuthenticatedFormPage(
  formReference: string,
  options: MicrosoftFormsClientOptions,
  headless = options.headless ?? true,
): Promise<OpenFormPage> {
  const { formId, responsePageUrl } = parseFormReference(formReference);
  const browserContext = await launchFormsBrowserContext({
    browserProfileDirectory: options.browserProfileDirectory,
    browserChannel: options.browserChannel,
    headless,
    log: options.log,
  });

  try {
    const existingPages = browserContext.pages();
    const page = existingPages[0] ?? (await browserContext.newPage());
    await page.goto(responsePageUrl, {
      waitUntil: "domcontentloaded",
      timeout: 2 * 60 * 1_000,
    });
    const serverInformation = await waitForFormsServerInformation(
      page,
      headless,
    );
    const session = await createSession(
      browserContext,
      serverInformation,
      formId,
      responsePageUrl,
    );
    return { browserContext, page, serverInformation, session };
  } catch (error) {
    await browserContext.close();
    throw error;
  }
}

export async function bootstrapFormsSession(
  formReference: string,
  options: MicrosoftFormsClientOptions,
): Promise<FormsApiSession> {
  const openFormPage = await openAuthenticatedFormPage(formReference, options);
  await openFormPage.browserContext.close();
  return openFormPage.session;
}

export async function authenticateFormsUser(
  options: AuthenticateOptions = {},
): Promise<AuthenticatedUser> {
  if (options.form) {
    const openFormPage = await openAuthenticatedFormPage(
      options.form,
      options,
      false,
    );
    const userInformation = openFormPage.serverInformation.userInfo;
    await openFormPage.browserContext.close();
    return {
      displayName: userInformation?.DisplayName,
      email: userInformation?.Email,
      tenantId: userInformation?.TenantId,
      userId: userInformation?.UserId,
    };
  }

  const browserContext = await launchFormsBrowserContext({
    browserProfileDirectory: options.browserProfileDirectory,
    browserChannel: options.browserChannel,
    headless: false,
    log: options.log,
  });
  try {
    const page = browserContext.pages()[0] ?? (await browserContext.newPage());
    await page.goto("https://forms.cloud.microsoft/Pages/DesignPageV2.aspx", {
      waitUntil: "domcontentloaded",
      timeout: 2 * 60 * 1_000,
    });
    const timeoutMilliseconds = options.timeoutMilliseconds ?? 5 * 60 * 1_000;
    await page.waitForFunction(
      () =>
        Boolean(
          (
            window as Window & {
              OfficeFormServerInfo?: OfficeFormServerInformation;
            }
          ).OfficeFormServerInfo?.userInfo?.UserId,
        ),
      undefined,
      { timeout: timeoutMilliseconds },
    );
    const userInformation = await page.evaluate(
      () =>
        (
          window as Window & {
            OfficeFormServerInfo?: OfficeFormServerInformation;
          }
        ).OfficeFormServerInfo?.userInfo,
    );
    return {
      displayName: userInformation?.DisplayName,
      email: userInformation?.Email,
      tenantId: userInformation?.TenantId,
      userId: userInformation?.UserId,
    };
  } catch (error) {
    throw new MicrosoftFormsAuthenticationError(
      `Microsoft Forms login failed: ${
        error instanceof Error ? error.message : String(error)
      }`,
    );
  } finally {
    await browserContext.close();
  }
}
