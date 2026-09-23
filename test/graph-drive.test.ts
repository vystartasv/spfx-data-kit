import test from "node:test";
import assert from "node:assert/strict";
import { DataClient, DataError, GraphDriveAdapter, graphDriveUploadUrl } from "../src/index.js";
import { ScriptedTransport } from "./helpers.js";

const drive = (transport: ScriptedTransport) => new GraphDriveAdapter(new DataClient(transport, { retry: { maxRetries: 0 } }), { baseUrl: "https://graph.test/beta" });

test("Graph drive metadata uses relative paths resolved against baseUrl and preserves ETags", async () => {
  const transport = new ScriptedTransport([
    { status: 200, headers: { ETag: '"drive"' }, body: JSON.stringify({ id: "d1", driveType: "documentLibrary" }) },
    { status: 200, body: JSON.stringify({ id: "root", name: "Documents" }) },
    { status: 200, body: JSON.stringify({ id: "i1", name: "file.txt", eTag: '"item"' }) },
  ]);
  const adapter = drive(transport);
  assert.deepEqual(await adapter.getDrive("d1"), { data: { id: "d1", driveType: "documentLibrary" }, etag: '"drive"' });
  assert.deepEqual(await adapter.getRoot("d1"), { data: { id: "root", name: "Documents" }, etag: undefined });
  assert.deepEqual(await adapter.getItem("d1", "i1"), { data: { id: "i1", name: "file.txt", eTag: '"item"' }, etag: '"item"' });
  assert.deepEqual(transport.calls.map((call) => call.url), [
    "https://graph.test/beta/drives/d1",
    "https://graph.test/beta/drives/d1/root",
    "https://graph.test/beta/drives/d1/items/i1",
  ]);
});

test("Graph drive children reuse bounded Graph paging and make no extra requests", async () => {
  const transport = new ScriptedTransport([
    { status: 200, body: JSON.stringify({ value: [{ id: "i1" }, { id: "i2" }], "@odata.nextLink": "/drives/d1/root/children?page=2" }) },
    { status: 200, body: JSON.stringify({ value: [{ id: "i3" }], "@odata.nextLink": "/drives/d1/root/children?page=3" }) },
  ]);
  const items = [];
  for await (const item of drive(transport).iterateChildren("d1", { maxPages: 2 })) items.push(item.id);
  assert.deepEqual(items, ["i1", "i2", "i3"]);
  assert.deepEqual(transport.calls.map((call) => call.url), [
    "https://graph.test/beta/drives/d1/root/children",
    "https://graph.test/beta/drives/d1/root/children?page=2",
  ]);

  const limitedTransport = new ScriptedTransport([{ status: 200, body: JSON.stringify({ value: [{ id: "i1" }, { id: "i2" }], "@odata.nextLink": "/drives/d1/root/children?page=2" }) }]);
  const limited = [];
  for await (const item of drive(limitedTransport).iterateChildren("d1", { maxItems: 1 })) limited.push(item.id);
  assert.deepEqual(limited, ["i1"]);
  assert.equal(limitedTransport.calls.length, 1);
});

test("Graph drive download and upload use Uint8Array bodies with controls and ETags", async () => {
  const controller = new AbortController();
  const transport = new ScriptedTransport([
    { status: 200, bytes: new Uint8Array([0, 255, 3]) },
    { status: 201, headers: { ETag: '"2"' }, body: JSON.stringify({ id: "i2", name: "new.txt" }) },
  ]);
  const adapter = drive(transport);
  assert.deepEqual([...await adapter.downloadFile("d1", "i1", { signal: controller.signal, timeoutMs: 125 })], [0, 255, 3]);
  const result = await adapter.uploadFile("d1", "folder/new.txt", new Uint8Array([1, 2]), { etag: '"1"', signal: controller.signal, timeoutMs: 125 });
  assert.deepEqual(result, { data: { id: "i2", name: "new.txt" }, etag: '"2"' });
  assert.equal(transport.calls.length, 2);
  assert.equal(transport.calls[0].options.responseType, "binary");
  assert.equal(transport.calls[0].options.signal, controller.signal);
  assert.equal(transport.calls[0].options.timeoutMs, 125);
  assert.equal(transport.calls[1].options.method, "PUT");
  assert.deepEqual(transport.calls[1].options.body, new Uint8Array([1, 2]));
  assert.equal(transport.calls[1].options.headers?.["Content-Type"], "application/octet-stream");
  assert.equal(transport.calls[1].options.headers?.["If-Match"], '"1"');
  assert.equal(transport.calls[1].options.signal, controller.signal);
  assert.equal(transport.calls[1].options.timeoutMs, 125);
});

test("Graph drive update and delete forward ETags without follow-up reads", async () => {
  const transport = new ScriptedTransport([
    { status: 200, body: JSON.stringify({ id: "i1", name: "renamed.txt", eTag: '"2"' }) },
    { status: 204 },
  ]);
  const adapter = drive(transport);
  assert.deepEqual(await adapter.updateItem("d1", "i1", { name: "renamed.txt" }, { etag: '"1"' }), { data: { id: "i1", name: "renamed.txt", eTag: '"2"' }, etag: '"2"' });
  assert.deepEqual(await adapter.deleteItem("d1", "i1", { etag: '"2"' }), { data: undefined, etag: undefined });
  assert.equal(transport.calls.length, 2);
  assert.equal(transport.calls[0].options.method, "PATCH");
  assert.equal(transport.calls[0].options.headers?.["If-Match"], '"1"');
  assert.equal(transport.calls[0].options.body, JSON.stringify({ name: "renamed.txt" }));
  assert.equal(transport.calls[1].options.method, "DELETE");
  assert.equal(transport.calls[1].options.headers?.["If-Match"], '"2"');
});

test("Graph drive validates ids and upload paths before transport", async () => {
  const transport = new ScriptedTransport([]);
  const adapter = drive(transport);
  for (const operation of [
    () => adapter.getDrive(""),
    () => adapter.getItem("d/1", "i1"),
    () => adapter.getItem("d1", "../i1"),
    () => adapter.uploadFile("d1", "/folder/file.txt", new Uint8Array([1])),
    () => adapter.uploadFile("d1", "folder//file.txt", new Uint8Array([1])),
    () => adapter.uploadFile("d1", "large.bin", new Uint8Array(1_500_001)),
  ]) await assert.rejects(async () => operation(), (error: unknown) => error instanceof DataError && error.kind === "validation");
  assert.equal(graphDriveUploadUrl("d1", "folder/file.txt"), "/drives/d1/root:/folder/file.txt:/content");
  assert.equal(transport.calls.length, 0);
});
