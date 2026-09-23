import test from "node:test";
import assert from "node:assert/strict";
import { DataClient, DataError, SharePointFilesAdapter, SharePointRestAdapter, sharePointAttachmentDownloadUrl, sharePointFileDownloadUrl, sharePointFileUploadUrl, sharePointFolderChildrenUrl, sharePointListItemsUrl } from "../src/index.js";
import { ScriptedTransport } from "./helpers.js";

type Todo = { id: number; title: string };

const adapter = (transport: ScriptedTransport) => new SharePointRestAdapter<Todo, { title: string }, { title: string }>(new DataClient(transport, { retry: { maxRetries: 0 } }), {
  siteUrl: "https://tenant.test/sites/demo/",
  listTitle: "Todo's",
  map: { id: "Id", title: "Title" },
  createMap: (input) => ({ Title: input.title }),
  updateMap: (input) => ({ Title: input.title }),
});

test("SharePoint list URLs escape OData quotes and URL characters", () => {
  assert.equal(sharePointListItemsUrl("https://tenant.test/sites/demo", "O'Brien / 100%"), "https://tenant.test/sites/demo/_api/web/lists/getbytitle('O''Brien%20%2F%20100%25')/items");
});

test("SharePoint file URLs encode paths and stay inside the configured site", () => {
  assert.equal(sharePointFileDownloadUrl("https://tenant.test/sites/demo/", "/sites/demo/Shared Documents/O'Brien.txt"), "https://tenant.test/sites/demo/_api/web/GetFileByServerRelativeUrl('/sites/demo/Shared%20Documents/O''Brien.txt')/$value");
  assert.equal(sharePointFolderChildrenUrl("https://tenant.test/sites/demo", "Shared Documents"), "https://tenant.test/sites/demo/_api/web/GetFolderByServerRelativeUrl('/sites/demo/Shared%20Documents')?$expand=Folders,Files");
  assert.equal(sharePointFileUploadUrl("https://tenant.test/sites/demo", "/sites/demo/Shared Documents", "a b.txt", true), "https://tenant.test/sites/demo/_api/web/GetFolderByServerRelativeUrl('/sites/demo/Shared%20Documents')/Files/add(url='a%20b.txt',overwrite=true)");
  assert.throws(() => sharePointFileDownloadUrl("https://tenant.test/sites/demo", "/sites/other/a.txt"), (error: unknown) => error instanceof DataError && error.kind === "validation");
});

test("SharePoint files adapter downloads, uploads, forwards controls, and does not re-fetch", async () => {
  const controller = new AbortController();
  const transport = new ScriptedTransport([
    { status: 200, bytes: new Uint8Array([0, 255, 2]) },
    { status: 201, body: JSON.stringify({ d: { Name: "new.txt", ServerRelativeUrl: "/sites/demo/Shared Documents/new.txt", ETag: '"2"' } }) },
  ]);
  const files = new SharePointFilesAdapter(new DataClient(transport, { retry: { maxRetries: 0 } }), { siteUrl: "https://tenant.test/sites/demo", listTitle: "Todo's" });
  assert.deepEqual([...await files.downloadFile("/sites/demo/Shared Documents/a.bin", { signal: controller.signal, timeoutMs: 125 })], [0, 255, 2]);
  const result = await files.uploadFile("/sites/demo/Shared Documents", "new.txt", new Uint8Array([1, 2]), { overwrite: true, etag: '"1"' });
  assert.equal(result.etag, '"2"');
  assert.equal(transport.calls.length, 2);
  assert.equal(transport.calls[0].options.responseType, "binary");
  assert.equal(transport.calls[0].options.signal, controller.signal);
  assert.equal(transport.calls[0].options.timeoutMs, 125);
  assert.deepEqual(transport.calls[1].options.body, new Uint8Array([1, 2]));
  assert.equal(transport.calls[1].options.headers?.["IF-MATCH"], '"1"');
});

test("SharePoint files adapter reads expanded folder children and attachments with ETags", async () => {
  const transport = new ScriptedTransport([
    { status: 200, body: JSON.stringify({ d: { Files: { results: [{ Name: "a.txt", ServerRelativeUrl: "/sites/demo/Docs/a.txt" }] }, Folders: { results: [{ Name: "sub", ServerRelativeUrl: "/sites/demo/Docs/sub" }] }, ETag: '"folder"' } }) },
    { status: 200, body: JSON.stringify({ value: [{ FileName: "a.txt", ServerRelativeUrl: "/sites/demo/Lists/Todo/Attachments/1/a.txt", ETag: '"attachment"' }] }) },
  ]);
  const files = new SharePointFilesAdapter(new DataClient(transport, { retry: { maxRetries: 0 } }), { siteUrl: "https://tenant.test/sites/demo", listTitle: "Todo" });
  assert.deepEqual(await files.folderChildren("/sites/demo/Docs"), {
    data: { files: [{ Name: "a.txt", ServerRelativeUrl: "/sites/demo/Docs/a.txt" }], folders: [{ Name: "sub", ServerRelativeUrl: "/sites/demo/Docs/sub" }] },
    etag: '"folder"',
  });
  assert.deepEqual(await files.listAttachments(1), {
    data: [{ FileName: "a.txt", ServerRelativeUrl: "/sites/demo/Lists/Todo/Attachments/1/a.txt", ETag: '"attachment"' }],
    etags: { "a.txt": '"attachment"' },
  });
  assert.match(transport.calls[0].url, /\$expand=Folders,Files/);
  assert.match(transport.calls[1].url, /getbytitle\('Todo'\)\/items\(1\)\/AttachmentFiles$/);
  assert.equal(sharePointAttachmentDownloadUrl("https://tenant.test/sites/demo", "Todo", 1, "a.txt"), "https://tenant.test/sites/demo/_api/web/lists/getbytitle('Todo')/items(1)/AttachmentFiles('a.txt')/$value");
});

test("SharePoint attachment delete forwards ETag and binary controls", async () => {
  const controller = new AbortController();
  const transport = new ScriptedTransport([{ status: 204 }]);
  const files = new SharePointFilesAdapter(new DataClient(transport, { retry: { maxRetries: 0 } }), { siteUrl: "https://tenant.test/sites/demo", listTitle: "Todo" });
  await files.deleteAttachment(4, "a.txt", { etag: '"9"', signal: controller.signal, timeoutMs: 300 });
  assert.equal(transport.calls[0].options.headers?.["IF-MATCH"], '"9"');
  assert.equal(transport.calls[0].options.signal, controller.signal);
  assert.equal(transport.calls[0].options.timeoutMs, 300);
});

test("SharePoint files adapter rejects malformed metadata and oversized uploads without network", async () => {
  const transport = new ScriptedTransport([{ status: 200, body: JSON.stringify({ d: { Name: "missing-url" } }) }]);
  const files = new SharePointFilesAdapter(new DataClient(transport, { retry: { maxRetries: 0 } }), { siteUrl: "https://tenant.test/sites/demo", listTitle: "Todo" });
  await assert.rejects(files.getFile("/sites/demo/Docs/a.txt"), (error: unknown) => error instanceof DataError && error.kind === "unknown" && /malformed file/.test(error.message));
  await assert.rejects(files.uploadFile("/sites/demo/Docs", "large.bin", new Uint8Array(1_500_001)), (error: unknown) => error instanceof DataError && error.kind === "validation");
  assert.equal(transport.calls.length, 1);
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

test("SharePoint item reads accept ETags from response headers", async () => {
  const transport = new ScriptedTransport([{ status: 200, headers: new Headers([["ETag", '"header"']]), body: JSON.stringify({ d: { Id: 7, Title: "read" } }) }]);
  assert.deepEqual(await adapter(transport).get(7), { data: { id: 7, title: "read" }, etag: '"header"' });
});

test("SharePoint paging stops when a next link repeats", async () => {
  const link = "/sites/demo/_api/web/lists/getbytitle('Todo''s')/items?skiptoken=repeat";
  const transport = new ScriptedTransport([
    { status: 200, body: JSON.stringify({ value: [{ Id: 1, Title: "one" }], "@odata.nextLink": link }) },
    { status: 200, body: JSON.stringify({ value: [{ Id: 2, Title: "two" }], "@odata.nextLink": link }) },
  ]);
  assert.deepEqual((await adapter(transport).list()).data, [{ id: 1, title: "one" }, { id: 2, title: "two" }]);
  assert.equal(transport.calls.length, 2);
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
  const result = await adapter(transport).batch([{ id: "one", method: "GET", url: "/items(1)", headers: { Prefer: "return=minimal" } }, { id: "two", method: "GET", url: "/items(2)" }], boundary);
  assert.deepEqual(result.responses.map(({ id, status, ok, body: child }) => ({ id, status, ok, body: child })), [
    { id: "one", status: 200, ok: true, body: { Id: 1 } },
    { id: "two", status: 404, ok: false, body: { error: { code: "missing" } } },
  ]);
  assert.equal(result.failures.length, 1);
  assert.match(String(transport.calls[0].options.body), /Prefer: return=minimal/);
});
