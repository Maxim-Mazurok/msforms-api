import { createRequire } from "node:module";

interface PackageMetadata {
  name: string;
  version: string;
}

const requireFromHere = createRequire(__filename);
const packageMetadata = requireFromHere("../package.json") as PackageMetadata;

export const packageName = packageMetadata.name;
export const packageVersion = packageMetadata.version;
