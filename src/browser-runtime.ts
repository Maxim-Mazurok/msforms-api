import { spawnSync } from "node:child_process";
import { mkdirSync } from "node:fs";
import { createRequire } from "node:module";
import { homedir } from "node:os";
import { dirname, join } from "node:path";
import { chromium, type BrowserContext } from "playwright";
import { MicrosoftFormsAuthenticationError } from "./errors.js";

const requireFromHere = createRequire(__filename);

type BrowserChannel = "chrome" | "msedge";
type LogFunction = (...messages: unknown[]) => void;

export function getDefaultBrowserProfileDirectory(): string {
  return join(homedir(), ".msforms-api", "browser-profile");
}

function installedBrowserChannels(): BrowserChannel[] {
  switch (process.platform) {
    case "darwin":
      return ["chrome"];
    case "win32":
      return ["msedge", "chrome"];
    default:
      return ["chrome", "msedge"];
  }
}

function isMissingPlaywrightBrowserError(error: unknown): boolean {
  return (
    error instanceof Error && error.message.includes("Executable doesn't exist")
  );
}

function installBundledChromium(log: LogFunction): void {
  const playwrightPackageJson = requireFromHere.resolve(
    "playwright/package.json",
  );
  const playwrightCommandPath = join(dirname(playwrightPackageJson), "cli.js");
  log("Installing Playwright Chromium for Microsoft Forms...");
  const installResult = spawnSync(
    process.execPath,
    [playwrightCommandPath, "install", "chromium"],
    { encoding: "utf8" },
  );
  if (installResult.status !== 0) {
    const detail =
      installResult.error?.message ||
      installResult.stderr.trim() ||
      installResult.stdout.trim() ||
      "Unknown installation error";
    throw new MicrosoftFormsAuthenticationError(
      `Could not install Playwright Chromium: ${detail}`,
    );
  }
}

export async function launchFormsBrowserContext(options: {
  browserProfileDirectory?: string;
  browserChannel?: BrowserChannel;
  headless: boolean;
  log?: LogFunction;
}): Promise<BrowserContext> {
  const browserProfileDirectory =
    options.browserProfileDirectory ?? getDefaultBrowserProfileDirectory();
  const log = options.log ?? (() => undefined);
  mkdirSync(browserProfileDirectory, { recursive: true });

  const browserChannels = options.browserChannel
    ? [options.browserChannel]
    : installedBrowserChannels();
  const launchErrors: string[] = [];

  for (const browserChannel of browserChannels) {
    try {
      return await chromium.launchPersistentContext(browserProfileDirectory, {
        channel: browserChannel,
        headless: options.headless,
        viewport: { width: 1_440, height: 1_000 },
      });
    } catch (error) {
      launchErrors.push(
        `${browserChannel}: ${error instanceof Error ? error.message : String(error)}`,
      );
    }
  }

  try {
    return await chromium.launchPersistentContext(browserProfileDirectory, {
      headless: options.headless,
      viewport: { width: 1_440, height: 1_000 },
    });
  } catch (error) {
    if (!isMissingPlaywrightBrowserError(error)) {
      throw new MicrosoftFormsAuthenticationError(
        `Could not launch a browser: ${launchErrors.join("; ")}; ${
          error instanceof Error ? error.message : String(error)
        }`,
      );
    }
  }

  installBundledChromium(log);
  return chromium.launchPersistentContext(browserProfileDirectory, {
    headless: options.headless,
    viewport: { width: 1_440, height: 1_000 },
  });
}
