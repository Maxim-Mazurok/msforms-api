import type { RawForm } from "../../src/types.js";

export const responsePageUrl =
  "https://forms.cloud.microsoft/Pages/ResponsePage.aspx?id=test-form";

export const rawForm: RawForm = {
  id: "test-form",
  title: "Automation test form",
  description: "Representative Microsoft Forms controls",
  ownerId: "owner-identifier",
  ownerTenantId: "tenant-identifier",
  settings: JSON.stringify({ IsAnonymous: false }),
  questions: [
    {
      id: "email-question",
      title: "Email address",
      type: "Question.TextField",
      order: 1,
      required: true,
      questionInfo: JSON.stringify({
        Validation: { rule: 11 },
      }),
    },
    {
      id: "branch-question",
      title: "Show advanced question?",
      type: "Question.Choice",
      order: 2,
      required: true,
      questionInfo: JSON.stringify({
        ChoiceType: 1,
        Choices: [
          {
            Description: "Yes",
            BranchInfo: { TargetQuestionId: "advanced-question" },
          },
          {
            Description: "No",
            BranchInfo: { TargetQuestionId: "date-question" },
          },
        ],
      }),
    },
    {
      id: "advanced-question",
      title: "Advanced details",
      type: "Question.TextField",
      order: 3,
      required: true,
      questionInfo: JSON.stringify({ Multiline: true }),
    },
    {
      id: "date-question",
      title: "Event date",
      type: "Question.DateTime",
      order: 4,
      required: true,
    },
    {
      id: "ranking-question",
      title: "Rank features",
      type: "Question.Ranking",
      order: 5,
      required: true,
      choices: [
        {
          id: 1,
          key: "first-feature",
          displayText: "Branching",
          order: 1,
        },
        {
          id: 2,
          key: "second-feature",
          displayText: "Uploads",
          order: 2,
        },
      ],
    },
    {
      id: "matrix-question",
      title: "Rate statements",
      type: "Question.MatrixChoiceGroup",
      order: 6,
      required: true,
      choices: [
        {
          id: 1,
          key: "disagree-key",
          displayText: "Disagree",
          order: 1,
        },
        {
          id: 2,
          key: "agree-key",
          displayText: "Agree",
          order: 2,
        },
      ],
    },
    {
      id: "first-matrix-row",
      groupId: "matrix-question",
      title: "Forms is useful",
      type: "Question.MatrixChoice",
      order: 7,
      required: false,
    },
    {
      id: "second-matrix-row",
      groupId: "matrix-question",
      title: "Forms is fast",
      type: "Question.MatrixChoice",
      order: 8,
      required: false,
    },
    {
      id: "file-question",
      title: "Upload evidence",
      type: "Question.FileUpload",
      order: 9,
      required: true,
      questionInfo: JSON.stringify({
        HasSpecificFileType: true,
        FileTypes: { PDF: true, Image: true, Word: false },
        MaxFileCount: 2,
        MaxFileSize: 10,
      }),
    },
    {
      id: "multiple-choice-question",
      title: "Select tools",
      type: "Question.Choice",
      order: 10,
      required: true,
      questionInfo: JSON.stringify({
        ChoiceType: 2,
        Choices: [{ Description: "Forms" }, { Description: "Excel" }],
      }),
    },
  ],
};
