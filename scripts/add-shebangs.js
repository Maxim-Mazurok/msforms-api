const { readFileSync, writeFileSync } = require("node:fs");
const { join } = require("node:path");

const executableFiles = ["cli.js", "mcp-server.js"];
const shebang = "#!/usr/bin/env node\n";

for (const executableFile of executableFiles) {
  const executablePath = join(__dirname, "..", "dist", executableFile);
  const content = readFileSync(executablePath, "utf8");
  if (!content.startsWith("#!")) {
    writeFileSync(executablePath, `${shebang}${content}`);
  }
}
