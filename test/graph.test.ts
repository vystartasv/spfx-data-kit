import test from "node:test";
import assert from "node:assert/strict";
import { DataClient, DataError, GraphAdapter } from "../src/index.js";
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
