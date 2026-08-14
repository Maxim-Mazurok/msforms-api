import {
  beforeEach,
  describe,
  expect,
  it as testCase,
  vi as vitest,
} from "vitest";
import { MicrosoftFormsClient } from "../../src/forms-client.js";
import type { FormsApiSession } from "../../src/forms-session.js";
import type { FormDefinition } from "../../src/types.js";

const bootstrapFormsSession = vitest.hoisted(() => vitest.fn());

vitest.mock("../../src/forms-session.js", async (importOriginal) => {
  const formsSessionModule =
    await importOriginal<typeof import("../../src/forms-session.js")>();
  return { ...formsSessionModule, bootstrapFormsSession };
});

const formsApiSession: FormsApiSession = {
  antiForgeryToken: "anti-forgery-token",
  cookieHeader: "session=cookie",
  formApiUrl: "https://forms.example/formapi/runtime-form",
  formId: "form-identifier",
  formsHost: "https://forms.example",
  multiUserIdentifier: "multi-user-identifier",
  ownerId: "owner/user",
  ownerTenantId: "owner tenant",
  responsePageUrl: "https://forms.example/response-page",
  responderId: "responder/user",
  responderTenantId: "responder tenant",
  serverSessionId: "server-session-identifier",
};

const formDefinition: FormDefinition = {
  id: "form-identifier",
  ownerId: "owner/user",
  tenantId: "owner tenant",
  title: "Test form",
  responsePageUrl: "https://forms.example/response-page",
  settings: {},
  questions: [],
  raw: { id: "form-identifier" },
};

describe("MicrosoftFormsClient", () => {
  const fetchMock = vitest.fn<typeof fetch>();

  beforeEach(() => {
    bootstrapFormsSession.mockReset();
    bootstrapFormsSession.mockResolvedValue(formsApiSession);
    fetchMock.mockReset();
    vitest.stubGlobal("fetch", fetchMock);
  });

  testCase("uses owner identity when saving a response", async () => {
    fetchMock.mockResolvedValue(new Response("{}", { status: 200 }));
    const client = new MicrosoftFormsClient();

    await client.saveResponse(
      "form-identifier",
      "response-identifier",
      "2026-01-02T03:04:05.000Z",
    );

    expect(fetchMock.mock.calls[0]?.[0]).toBe(
      "https://forms.example/formapi/api/owner%20tenant/users/owner%2Fuser/saveResponseLink",
    );
  });

  testCase("uses responder identity when listing saved responses", async () => {
    fetchMock.mockResolvedValue(new Response("[]", { status: 200 }));
    const client = new MicrosoftFormsClient();

    await client.listSavedResponses("form-identifier");

    expect(fetchMock.mock.calls[0]?.[0]).toBe(
      "https://forms.example/formapi/api/responder%20tenant/users/responder%2Fuser/getResponseLinks",
    );
  });

  testCase(
    "reports save failure without hiding successful submission",
    async () => {
      fetchMock
        .mockResolvedValueOnce(
          new Response(JSON.stringify({ id: 42 }), { status: 200 }),
        )
        .mockResolvedValueOnce(
          new Response(JSON.stringify({ error: "save failed" }), {
            status: 404,
          }),
        );
      const client = new MicrosoftFormsClient();
      vitest.spyOn(client, "getForm").mockResolvedValue(formDefinition);

      const result = await client.submitResponse(
        "form-identifier",
        {},
        { saveResponse: true },
      );

      expect(result).toEqual(
        expect.objectContaining({
          responseId: "42",
          submissionStatus: "submitted",
          saveResponseStatus: "failed",
          saveResponseError: expect.objectContaining({ status: 404 }),
        }),
      );
      expect(result.savedResponse).toBeUndefined();
      expect(fetchMock).toHaveBeenCalledTimes(2);
    },
  );
});
