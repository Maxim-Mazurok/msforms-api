import { Client } from "@modelcontextprotocol/client";
import { InMemoryTransport } from "@modelcontextprotocol/server";
import { afterEach, describe, expect, it } from "vitest";
import { createMicrosoftFormsMcpServer } from "../../src/mcp-server.js";

const connectedClients: Client[] = [];

afterEach(async () => {
  await Promise.all(connectedClients.splice(0).map((client) => client.close()));
});

async function connectClient(): Promise<Client> {
  const server = createMicrosoftFormsMcpServer();
  const client = new Client(
    { name: "msforms-api-test", version: "1.0.0" },
    { capabilities: { elicitation: { form: {} } } },
  );
  const [clientTransport, serverTransport] =
    InMemoryTransport.createLinkedPair();
  await server.connect(serverTransport);
  await client.connect(clientTransport);
  connectedClients.push(client);
  return client;
}

describe("Microsoft Forms MCP server", () => {
  it("advertises shared actions with safety annotations", async () => {
    const client = await connectClient();
    const { tools } = await client.listTools();
    expect(tools.map((tool) => tool.name)).toEqual(
      expect.arrayContaining([
        "forms_authenticate",
        "forms_inspect",
        "forms_validate",
        "forms_upload_file",
        "forms_submit",
        "forms_save_response",
        "forms_list_saved_responses",
      ]),
    );
    const submitTool = tools.find((tool) => tool.name === "forms_submit");
    expect(submitTool?.annotations).toEqual(
      expect.objectContaining({
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: false,
        openWorldHint: true,
      }),
    );
  });

  it("does not execute submission when confirmation is false", async () => {
    const client = await connectClient();
    client.setRequestHandler("elicitation/create", async () => ({
      action: "accept",
      content: { confirm: false },
    }));
    const result = await client.callTool({
      name: "forms_submit",
      arguments: { form: "must-not-be-requested", answers: {} },
    });
    expect(result.isError).toBe(true);
    expect(result.content).toEqual([
      expect.objectContaining({
        type: "text",
        text: expect.stringContaining("cancelled"),
      }),
    ]);
  });

  it("does not repeat confirmation after elicitation is declined", async () => {
    const client = await connectClient();
    let confirmationRequestCount = 0;
    client.setRequestHandler("elicitation/create", async () => {
      confirmationRequestCount += 1;
      return { action: "decline" };
    });
    const result = await client.callTool({
      name: "forms_submit",
      arguments: { form: "must-not-be-requested", answers: {} },
    });
    expect(result.isError).toBe(true);
    expect(confirmationRequestCount).toBe(1);
  });
});
