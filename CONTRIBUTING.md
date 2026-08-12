# Contributing

Development guide for `msforms-api`. User installation and usage belong in
[README.md](README.md).

## Setup

```bash
git clone https://github.com/Maxim-Mazurok/msforms-api.git
cd msforms-api
npm install
```

Run from source:

```bash
npx -y tsx src/cli.ts --help
npm run mcp
```

## Architecture

```text
src/
  forms-client.ts      Public SDK facade
  forms-session.ts     Persistent browser authentication and bootstrap
  forms-api.ts         Authenticated direct HTTP requests
  form-normalizer.ts   Raw API to stable normalized form model
  answers.ts           Branch resolution, validation, and serialization
  file-upload.ts       Browser-assisted upload flow
  actions.ts           Shared CLI and MCP action definitions
  cli.ts               Commander adapter
  mcp-server.ts        MCP v2 stdio adapter
  server-instructions.ts  Shared MCP and CLI workflow guidance
  types.ts             Public and raw API types
```

The SDK owns behavior. CLI and MCP adapters consume the same action registry.
Keep Microsoft-specific wire shapes in the core instead of duplicating them in
adapters.

`SKILL.md` is the installable, CLI-focused agent workflow. Keep its commands and
safety requirements aligned with `server-instructions.ts` and the action
definitions.

## Development checks

```bash
npm run format
npm run spell-check
npm run type-check
npm test
npm run build
```

Tests must not submit live forms. Live API tests belong in explicit manual
verification using a form owned by the tester.

## Code style

- Strict TypeScript.
- Named exports only.
- ESM syntax in TypeScript; CommonJS package output.
- Full descriptive identifiers.
- No compatibility wrappers when replacing APIs.
- Explicit errors rather than silent fallbacks.

## Reverse-engineering changes

When Forms behavior changes:

1. Capture the current browser request and response contract.
2. Add or update a typed raw representation.
3. Normalize it into the stable public model.
4. Add a fixture-based test.
5. Verify read and validation paths before any live submission.

Never commit captured cookies, verification tokens, respondent data, or private
form payloads.
