import test from "node:test";
import assert from "node:assert/strict";
import { DataClient, SharePointRestAdapter } from "../src/index.js";
import { ScriptedTransport } from "./helpers.js";

type Todo = { id: number; title: string };

const adapter = (transport: ScriptedTransport) => new SharePointRestAdapter<Todo, { title: string }, { title: string }>(new DataClient(transport, { retry: { maxRetries: 0 } }), {
  siteUrl: "https://tenant.test/sites/demo/",
  listTitle: "Todo's",
  map: { id: "Id", title: "Title" },
  createMap: (input) => ({ Title: input.title }),
  updateMap: (input) => ({ Title: input.title }),
});

test("SharePoint adapter builds queries, follows pages, caps items, and returns ETags", async () => {
  const transport = new ScriptedTransport([
    { status: 200, body: JSON.stringify({ value: [{ Id: 1, Title: "one", "@odata.etag": "\"1\"" }, { Id: 2, Title: "two", "@odata.etag": "\"2\"" }], "@odata.nextLink": "/sites/demo/_api/web/lists/getbytitle('Todo%27%27s')/items?$skiptoken=next" }) },
    { status: 200, body: JSON.stringify({ value: [{ Id: 3, Title: "three", "@odata.etag": "\"3\"" }, { Id: 4, Title: "four", "@odata.etag": "\"4\"" }] }) },
  ]);
  const result = await adapter(transport).list({ select: ["Id", "Title"], expand: ["Owner"], filter: "Title eq 'one'", orderBy: [["Title", false]], top: 3, maxPages: 2, pageSize: 2 });
  const first = new URL(transport.calls[0].url);
  assert.equal(first.searchParams.get("$select"), "Id,Title");
  assert.equal(first.searchParams.get("$expand"), "Owner");
  assert.equal(first.searchParams.get("$filter"), "Title eq 'one'");
  assert.equal(first.searchParams.get("$orderby"), "Title desc");
  assert.equal(first.searchParams.get("$top"), "2");
  assert.equal(transport.calls.length, 2);
  assert.deepEqual(result, { data: [{ id: 1, title: "one" }, { id: 2, title: "two" }, { id: 3, title: "three" }], etags: { "1": "\"1\"", "2": "\"2\"", "3": "\"3\"" } });
});

test("SharePoint adapter sends ETags and re-reads after update", async () => {
  const transport = new ScriptedTransport([
    { status: 204 },
    { status: 200, body: JSON.stringify({ d: { Id: 7, Title: "new", ETag: "\"2\"" } }) },
  ]);
  const result = await adapter(transport).update(7, { title: "new" }, { etag: "\"1\"" });
  assert.equal(transport.calls[0].options.method, "POST");
  assert.equal(transport.calls[0].options.headers?.["IF-MATCH"], "\"1\"");
  assert.deepEqual(result, { data: { id: 7, title: "new" }, etag: "\"2\"" });
});

test("SharePoint adapter parses successful and failed batch children", async () => {
  const boundary = "batch_test";
  const body = [
    `--${boundary}`,
    "Content-Type: application/http",
    "Content-Transfer-Encoding: binary",
    "",
    "HTTP/1.1 200 OK",
    "Content-Type: application/json",
    "",
    '{"Id":1}',
    `--${boundary}`,
    "Content-Type: application/http",
    "Content-Transfer-Encoding: binary",
    "",
    "HTTP/1.1 404 Not Found",
    "Content-Type: application/json",
    "",
    '{"error":{"code":"missing"}}',
    `--${boundary}--`,
    "",
  ].join("\r\n");
  const transport = new ScriptedTransport([{ status: 200, headers: { "Content-Type": `multipart/mixed; boundary=${boundary}` }, body }]);
  const result = await adapter(transport).batch([{ id: "one", method: "GET", url: "/items(1)" }, { id: "two", method: "GET", url: "/items(2)" }], boundary);
  assert.deepEqual(result.responses.map(({ id, status, ok, body: child }) => ({ id, status, ok, body: child })), [
    { id: "one", status: 200, ok: true, body: { Id: 1 } },
    { id: "two", status: 404, ok: false, body: { error: { code: "missing" } } },
  ]);
  assert.equal(result.failures.length, 1);
});
