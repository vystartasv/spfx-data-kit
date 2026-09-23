import test from "node:test";
import assert from "node:assert/strict";
import { DataClient, DataError, GraphAdapter } from "../src/index.js";
import type { RequestTransport, TransportResponse } from "../src/index.js";
import { ScriptedTransport } from "./helpers.js";

const request = (id: string) => ({ id, method: "GET", url: `/users/${id}` });

test("Graph adapter rejects batches larger than 20 before transport", async () => {
  const transport = new ScriptedTransport([]);
  const graph = new GraphAdapter(new DataClient(transport, { retry: { maxRetries: 0 } }));
  await assert.rejects(graph.batch(Array.from({ length: 21 }, (_, index) => request(String(index)))), (error: unknown) => error instanceof DataError && error.kind === "validation");
  assert.equal(transport.calls.length, 0);
});

test("Graph adapter returns child failures without hiding successful children", async () => {
  const transport = new ScriptedTransport([{ status: 200, body: JSON.stringify({ responses: [
    { id: "ok", status: 200, headers: { ETag: "\"1\"" }, body: { id: "1" } },
    { id: "bad", status: 403, headers: {}, body: { error: { code: "Forbidden" } } },
  ] }) }]);
  const graph = new GraphAdapter(new DataClient(transport, { retry: { maxRetries: 0 } }));
  const result = await graph.batch([request("ok"), request("bad")]);
  assert.deepEqual(result.responses.map(({ id, status, ok }) => ({ id, status, ok })), [
    { id: "ok", status: 200, ok: true },
    { id: "bad", status: 403, ok: false },
  ]);
  assert.deepEqual(result.failures.map(({ id, body }) => ({ id, body })), [{ id: "bad", body: { error: { code: "Forbidden" } } }]);
  assert.equal(transport.calls[0].options.method, "POST");
});

test("Graph GET preserves response headers through cache and deduplication", async () => {
  let calls = 0;
  let release!: () => void;
  const gate = new Promise<void>((resolve) => { release = resolve; });
  const transport: RequestTransport = {
    async request(): Promise<TransportResponse> {
      calls++;
      await gate;
      return { status: 200, headers: new Headers([["ETag", '"header"']]), text: async () => JSON.stringify({ id: "1" }) };
    },
  };
  const client = new DataClient(transport, { cache: { maxEntries: 2, ttlMs: 1_000 }, retry: { maxRetries: 0 } });
  const graph = new GraphAdapter(client);
  const first = graph.request<{ id: string }>("/me");
  const second = graph.request<{ id: string }>("/me");
  release();
  assert.deepEqual(await Promise.all([first, second]), [
    { data: { id: "1" }, etag: '"header"' },
    { data: { id: "1" }, etag: '"header"' },
  ]);
  assert.equal(calls, 1);
  assert.equal(client.diagnostics().deduplicated, 1);
  assert.deepEqual(await graph.request<{ id: string }>("/me"), { data: { id: "1" }, etag: '"header"' });
  assert.equal(client.diagnostics().cacheHits, 1);
});

test("Graph batch validates relative URLs and methods", async () => {
  const transport = new ScriptedTransport([]);
  const graph = new GraphAdapter(new DataClient(transport, { retry: { maxRetries: 0 } }));
  for (const invalid of [
    { id: " ", method: "GET", url: "/me" },
    { id: "one", method: "", url: "/me" },
    { id: "one", method: "GET", url: "//other.example/me" },
    { id: "one", method: "GET", url: "" },
  ]) await assert.rejects(graph.batch([invalid]), (error: unknown) => error instanceof DataError && error.kind === "validation");
  assert.equal(transport.calls.length, 0);
});

test("Graph batch rejects a response that omits a child", async () => {
  const transport = new ScriptedTransport([{ status: 200, body: JSON.stringify({ responses: [{ id: "ok", status: 200, headers: {}, body: {} }] }) }]);
  const graph = new GraphAdapter(new DataClient(transport, { retry: { maxRetries: 0 } }));
  await assert.rejects(graph.batch([request("ok"), request("missing")]), (error: unknown) => error instanceof DataError && error.kind === "unknown");
});

test("Graph paging stops before requesting after maxPages or maxItems", async () => {
  const transport = new ScriptedTransport([
    { status: 200, body: JSON.stringify({ value: [{ id: 1 }, { id: 2 }], "@odata.nextLink": "/users?page=2" }) },
  ]);
  const graph = new GraphAdapter(new DataClient(transport, { retry: { maxRetries: 0 } }));
  const pages = [];
  for await (const page of graph.pages<{ id: number }>("/users", { maxPages: 1 })) pages.push(page);
  assert.deepEqual(pages, [{ value: [{ id: 1 }, { id: 2 }], nextLink: "/users?page=2" }]);
  assert.equal(transport.calls.length, 1);

  const limited = new ScriptedTransport([{ status: 200, body: JSON.stringify({ value: [{ id: 1 }, { id: 2 }], "@odata.nextLink": "/users?page=2" }) }]);
  const limitedGraph = new GraphAdapter(new DataClient(limited, { retry: { maxRetries: 0 } }));
  const items: { id: number }[] = [];
  for await (const item of limitedGraph.iterate<{ id: number }>("/users", { maxItems: 1 })) items.push(item);
  assert.deepEqual(items, [{ id: 1 }]);
  assert.equal(limited.calls.length, 1);
});

test("Graph delta collects pages and preserves the delta link", async () => {
  const transport = new ScriptedTransport([
    { status: 200, body: JSON.stringify({ value: [{ id: 1 }], "@odata.nextLink": "/users/delta?page=2" }) },
    { status: 200, body: JSON.stringify({ value: [{ id: 2 }], "@odata.deltaLink": "https://graph.microsoft.com/v1.0/users/delta?$deltatoken=next" }) },
  ]);
  const graph = new GraphAdapter(new DataClient(transport, { retry: { maxRetries: 0 } }));
  assert.deepEqual(await graph.delta<{ id: number }>("/users/delta", { maxPages: 2 }), {
    value: [{ id: 1 }, { id: 2 }],
    deltaLink: "https://graph.microsoft.com/v1.0/users/delta?$deltatoken=next",
  });
  assert.equal(transport.calls.length, 2);
});
