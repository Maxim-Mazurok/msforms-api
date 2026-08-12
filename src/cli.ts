#!/usr/bin/env node

import { readFileSync } from "node:fs";
import { formsActions, type ActionParameter } from "./actions.js";
import { MicrosoftFormsClient } from "./forms-client.js";
import { packageVersion } from "./package-metadata.js";
import { serverInstructions } from "./server-instructions.js";

interface GlobalOptions {
  browser?: "chrome" | "msedge";
  headed?: boolean;
  profile?: string;
}

function camelCaseToKebabCase(value: string): string {
  return value.replace(/([A-Z])/g, "-$1").toLowerCase();
}

function parseJsonInput(value: string): unknown {
  const jsonText = value.startsWith("@")
    ? readFileSync(value.slice(1), "utf8")
    : value;
  try {
    return JSON.parse(jsonText) as unknown;
  } catch (error) {
    throw new Error(
      `Invalid JSON input: ${error instanceof Error ? error.message : String(error)}`,
    );
  }
}

function parseParameter(value: unknown, parameter: ActionParameter): unknown {
  if (value === undefined) {
    return undefined;
  }
  if (parameter.type === "json") {
    return parseJsonInput(String(value));
  }
  return value;
}

function createClient(options: GlobalOptions): MicrosoftFormsClient {
  return new MicrosoftFormsClient({
    browserProfileDirectory: options.profile,
    browserChannel: options.browser,
    headless: !options.headed,
    log: (...messages) => console.error(...messages),
  });
}

function printResult(value: unknown, compact: boolean): void {
  process.stdout.write(`${JSON.stringify(value, null, compact ? 0 : 2)}\n`);
}

async function main(): Promise<void> {
  const { Command } = await import("commander");
  const program = new Command()
    .name("msforms-api")
    .description("Microsoft Forms SDK command-line interface")
    .version(packageVersion)
    .option(
      "--profile <directory>",
      "Persistent browser profile directory (default: ~/.msforms-api/browser-profile)",
    )
    .option("--browser <channel>", "Browser channel: chrome or msedge")
    .option("--headed", "Show browser while performing the command");

  program
    .command("auth")
    .description("Open a browser and authenticate Microsoft Forms")
    .option("--form <reference>", "Optional form URL or ID to verify")
    .option("--compact", "Print compact JSON")
    .action(async (commandOptions: { compact?: boolean; form?: string }) => {
      const globalOptions = program.opts<GlobalOptions>();
      const result = await MicrosoftFormsClient.authenticate({
        browserProfileDirectory: globalOptions.profile,
        browserChannel: globalOptions.browser,
        form: commandOptions.form,
        log: (...messages) => console.error(...messages),
      });
      printResult(result, Boolean(commandOptions.compact));
    });

  program
    .command("guide")
    .description(
      "Print workflow guidance and safety requirements for Microsoft Forms automation",
    )
    .action(() => {
      process.stdout.write(`${serverInstructions}\n`);
    });

  for (const action of formsActions) {
    const command = program
      .command(action.name)
      .description(action.description)
      .option("--compact", "Print compact JSON");

    for (const parameter of action.parameters) {
      const optionName = camelCaseToKebabCase(parameter.name);
      const optionSyntax =
        parameter.type === "boolean"
          ? `--${optionName}`
          : `--${optionName} <value>`;
      if (parameter.required) {
        command.requiredOption(optionSyntax, parameter.description);
      } else {
        command.option(optionSyntax, parameter.description);
      }
    }
    if (action.requiresConfirmation) {
      command.requiredOption(
        "--confirm",
        "Confirm permanent response submission",
      );
    }

    command.action(async (commandOptions: Record<string, unknown>) => {
      const input: Record<string, unknown> = {};
      for (const parameter of action.parameters) {
        input[parameter.name] = parseParameter(
          commandOptions[parameter.name],
          parameter,
        );
      }
      const result = await action.execute(
        createClient(program.opts<GlobalOptions>()),
        input,
      );
      printResult(result, Boolean(commandOptions.compact));
    });
  }

  await program.parseAsync();
}

main().catch((error: unknown) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exitCode = 1;
});
