import { describe, expect, it } from "vitest";
import { parseFormReference } from "../../src/form-reference.js";

describe("parseFormReference", () => {
  it("accepts a form ID", () => {
    expect(parseFormReference("form-identifier")).toEqual({
      formId: "form-identifier",
      responsePageUrl:
        "https://forms.cloud.microsoft/Pages/ResponsePage.aspx?id=form-identifier",
    });
  });

  it("extracts an ID from a Forms URL", () => {
    const parsedReference = parseFormReference(
      "https://forms.cloud.microsoft/Pages/ResponsePage.aspx?id=encoded-form",
    );
    expect(parsedReference.formId).toBe("encoded-form");
  });

  it("rejects URLs without form IDs", () => {
    expect(() => parseFormReference("https://forms.cloud.microsoft/")).toThrow(
      "id query parameter",
    );
  });
});
