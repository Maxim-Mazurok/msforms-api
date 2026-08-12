import { describe, expect, it } from "vitest";
import {
  resolveVisibleQuestionIds,
  validateAndSerializeAnswers,
} from "../../src/answers.js";
import { normalizeForm } from "../../src/form-normalizer.js";
import type { FormAnswers, UploadedFile } from "../../src/types.js";
import { rawForm, responsePageUrl } from "./form-fixture.js";

const form = normalizeForm(rawForm, responsePageUrl);
const uploadedFile: UploadedFile = {
  name: "evidence.pdf",
  uploadSessionUrl: "https://sharepoint.example/upload",
  link: "https://sharepoint.example/upload",
  badgerToken: null,
  time: 1_800_000_000_000,
  status: 3,
  id: "sharepoint-item",
};

function validAnswers(): FormAnswers {
  return {
    "email-question": "person@example.com",
    "branch-question": "No",
    "date-question": "2026-08-12",
    "ranking-question": ["Uploads", "Branching"],
    "first-matrix-row": { label: "Agree" },
    "second-matrix-row": { key: "disagree-key" },
    "file-question": [uploadedFile],
    "multiple-choice-question": ["Forms", "Excel"],
  };
}

describe("resolveVisibleQuestionIds", () => {
  it("skips questions bypassed by a selected branch", () => {
    const visibleQuestionIds = resolveVisibleQuestionIds(form, {
      "branch-question": "No",
    });
    expect(visibleQuestionIds).not.toContain("advanced-question");
    expect(visibleQuestionIds).toContain("date-question");
  });
});

describe("validateAndSerializeAnswers", () => {
  it("serializes every supported representative answer type", () => {
    const result = validateAndSerializeAnswers(form, validAnswers());
    expect(result.valid).toBe(true);
    expect(result.issues).toEqual([]);
    expect(result.serializedAnswers).toEqual(
      expect.arrayContaining([
        { questionId: "branch-question", answer1: "No" },
        { questionId: "date-question", answer1: "2026-08-12" },
        {
          questionId: "ranking-question",
          answer1: [
            { id: 2, order: 2, answerOrder: 0 },
            { id: 1, order: 1, answerOrder: 1 },
          ],
        },
        {
          questionId: "first-matrix-row",
          answer1: { id: 2, key: "agree-key" },
        },
        {
          questionId: "file-question",
          answer1: JSON.stringify([uploadedFile]),
        },
        {
          questionId: "multiple-choice-question",
          answer1: JSON.stringify(["Forms", "Excel"]),
        },
      ]),
    );
    expect(
      result.serializedAnswers.some(
        (answer) => answer.questionId === "advanced-question",
      ),
    ).toBe(false);
    expect(
      result.serializedAnswers.some(
        (answer) => answer.questionId === "matrix-question",
      ),
    ).toBe(false);
  });

  it("reports required, format, option, and unknown-question failures", () => {
    const result = validateAndSerializeAnswers(form, {
      "email-question": "not-an-email",
      "branch-question": "Unknown choice",
      "unknown-question": "value",
    });
    expect(result.valid).toBe(false);
    expect(result.issues.map((issue) => issue.code)).toEqual(
      expect.arrayContaining([
        "invalid-answer",
        "missing-required",
        "unknown-question",
        "validation-failed",
      ]),
    );
  });
});
