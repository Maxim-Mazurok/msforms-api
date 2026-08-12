import { describe, expect, it } from "vitest";
import { normalizeForm } from "../../src/form-normalizer.js";
import { rawForm, responsePageUrl } from "./form-fixture.js";

describe("normalizeForm", () => {
  it("normalizes required fields, raw option shapes, and settings", () => {
    const form = normalizeForm(rawForm, responsePageUrl);
    expect(form.title).toBe("Automation test form");
    expect(form.settings).toEqual({ IsAnonymous: false });
    expect(form.questions).toHaveLength(8);
    expect(form.questions[0]?.required).toBe(true);

    const rankingQuestion = form.questions.find(
      (question) => question.id === "ranking-question",
    );
    expect(rankingQuestion?.options).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          id: 1,
          key: "first-feature",
          label: "Branching",
        }),
      ]),
    );
  });

  it("attaches matrix rows and inherits parent options and required state", () => {
    const form = normalizeForm(rawForm, responsePageUrl);
    const matrixQuestion = form.questions.find(
      (question) => question.id === "matrix-question",
    );
    expect(matrixQuestion?.rows).toHaveLength(2);
    expect(matrixQuestion?.rows[0]?.required).toBe(true);
    expect(matrixQuestion?.rows[0]?.options[1]).toEqual(
      expect.objectContaining({
        id: 2,
        key: "agree-key",
        label: "Agree",
      }),
    );
  });

  it("expands file type categories into extensions", () => {
    const form = normalizeForm(rawForm, responsePageUrl);
    const fileQuestion = form.questions.find(
      (question) => question.id === "file-question",
    );
    expect(fileQuestion?.fileUpload).toEqual(
      expect.objectContaining({
        maximumFileCount: 2,
        maximumFileSizeMegabytes: 10,
        allowedExtensions: expect.arrayContaining(["pdf", "jpg", "png"]),
      }),
    );
    expect(fileQuestion?.fileUpload?.allowedExtensions).not.toContain("docx");
  });

  it("preserves branch targets on choice options", () => {
    const form = normalizeForm(rawForm, responsePageUrl);
    const branchQuestion = form.questions.find(
      (question) => question.id === "branch-question",
    );
    expect(branchQuestion?.options).toEqual([
      expect.objectContaining({
        label: "Yes",
        branchTargetId: "advanced-question",
      }),
      expect.objectContaining({
        label: "No",
        branchTargetId: "date-question",
      }),
    ]);
  });
});
