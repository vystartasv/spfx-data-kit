import test from "node:test";
import assert from "node:assert/strict";
import { DataClient, DataError } from "../src/index.js";
import type { RequestTransport, TransportResponse } from "../src/index.js";

test("DataClient deduplicates in-flight GETs, expires cache, and invalidates by URL", async () => {
  let requests = 0;
  let now = 0;
  let release!: () => void;
  const gate = new Promise<void>((resolve) => { release = resolve; });
  const transport: RequestTransport = {
    async request(): Promise<TransportResponse> {
      requests++;
      if (requests === 1) await gate;
      return { status: 200, text: async () => JSON.stringify({ value: requests }) };
    },
  };
  const client = new DataClient(transport, { cache: { maxEntries: 4, ttlMs: 100 }, retry: { now: () => now } });
  const first = client.get<{ value: number }>("https://example.test/items");
  const second = client.get<{ value: number }>("https://example.test/items");
  release();
  assert.deepEqual(await Promise.all([first, second]), [{ value: 1 }, { value: 1 }]);
  assert.equal(client.diagnostics().deduplicated, 1);
  assert.deepEqual(await client.get("https://example.test/items"), { value: 1 });
  assert.equal(client.diagnostics().cacheHits, 1);
  now = 100;
  assert.deepEqual(await client.get("https://example.test/items"), { value: 2 });
  assert.equal(requests, 2);
  assert.equal(client.invalidateUrl("https://example.test/items"), 1);
  await client.get("https://example.test/items");
  assert.equal(requests, 3);
});

test("DataClient retries throttled and transient GET responses deterministically", async () => {
  const sleeps: number[] = [];
  let calls = 0;
  const transport: RequestTransport = {
    request: async () => {
      calls++;
      const status = calls === 1 ? 429 : calls === 2 ? 503 : 200;
      return { status, headers: status === 429 ? { "Retry-After": "0.25" } : {}, text: async () => status === 200 ? "{\"ok\":true}" : "" };
    },
  };
  const client = new DataClient(transport, { retry: { maxRetries: 2, sleep: async (ms) => { sleeps.push(ms); } } });
  assert.deepEqual(await client.get<{ ok: boolean }>("https://example.test/retry"), { ok: true });
  assert.deepEqual(sleeps, [250, 2000]);
  assert.deepEqual(client.diagnostics(), { requests: 3, cacheHits: 0, deduplicated: 0, retries: 2, failures: 0 });
});

test("DataClient exposes throttled failures after retry budget is exhausted", async () => {
  const client = new DataClient({ request: async () => ({ status: 429, headers: {}, text: async () => "" }) }, { retry: { maxRetries: 0 } });
  await assert.rejects(client.get("https://example.test/limited"), (error: unknown) => error instanceof DataError && error.kind === "throttled" && error.status === 429);
});

test("DataClient invalidation prevents an in-flight GET from repopulating cache", async () => {
  let calls = 0;
  let release!: () => void;
  const gate = new Promise<void>((resolve) => { release = resolve; });
  const transport: RequestTransport = {
    request: async () => {
      calls++;
      if (calls === 1) await gate;
      return { status: 200, text: async () => JSON.stringify({ calls }) };
    },
  };
  const client = new DataClient(transport, { cache: { maxEntries: 2, ttlMs: 1_000 } });
  const first = client.get<{ calls: number }>("https://example.test/items");
  assert.equal(client.invalidateUrl("https://example.test/items"), 0);
  release();
  await first;
  assert.deepEqual(await client.get("https://example.test/items"), { calls: 2 });
  assert.equal(calls, 2);
});

test("DataClient forwards signal and timeout and does not retry an aborted request", async () => {
  const controller = new AbortController();
  const calls: { signal?: AbortSignal; timeoutMs?: number }[] = [];
  const client = new DataClient({ request: async (_url, options) => {
    calls.push(options);
    throw controller.signal.reason ?? new DOMException("aborted", "AbortError");
  } }, { retry: { maxRetries: 2 }, signal: controller.signal, timeoutMs: 250 });
  controller.abort(new Error("cancelled"));
  await assert.rejects(client.get("https://example.test/abort"), /cancelled/);
  assert.deepEqual(calls, []);

  const active = new AbortController();
  const transport = { request: async (_url: string, options: { signal?: AbortSignal; timeoutMs?: number }) => {
    calls.push(options);
    return { status: 200, text: async () => "{}" };
  } };
  await new DataClient(transport, { signal: active.signal, timeoutMs: 250 }).get("https://example.test/active");
  assert.equal(calls.at(-1)?.signal, active.signal);
  assert.equal(calls.at(-1)?.timeoutMs, 250);
});

test("DataClient retains structured HTTP error details", async () => {
  const client = new DataClient({ request: async () => ({
    status: 400,
    text: async () => JSON.stringify({ error: { code: "InvalidRequest", message: "bad input" }, requestId: "req-1" }),
  }) }, { retry: { maxRetries: 0 } });
  await assert.rejects(client.get("https://example.test/error"), (error: unknown) => error instanceof DataError
    && error.kind === "validation"
    && error.code === "InvalidRequest"
    && error.details?.requestId === "req-1");
});
