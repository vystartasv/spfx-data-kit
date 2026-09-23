import test from "node:test";
import assert from "node:assert/strict";
import {
  DataClient,
  DataError,
  GraphSearchAdapter,
  SharePointSearchAdapter,
} from "../src/index.js";
import { ScriptedTransport } from "./helpers.js";

const client = (transport: ScriptedTransport) => new DataClient(transport, { retry: { maxRetries: 0 } });

test("SharePoint search builds typed POST bodies, maps rows, and follows bounded paging", async () => {
  const controller = new AbortController();
  const transport = new ScriptedTransport([
    { status: 200, body: JSON.stringify({ d: { query: { PrimaryQueryResult: { RelevantResults: { Table: { Rows: { results: [{ Title: "one" }] } }, TotalRows: 3, PagingInfo: "Paged=TRUE&p_ID=1" } } } } }) },
    { status: 200, body: JSON.stringify({ d: { query: { PrimaryQueryResult: { RelevantResults: { Table: { Rows: { results: [{ Title: "two" }] } }, TotalRows: 3 } } } } }) },
  ]);
  const result = await new SharePointSearchAdapter(client(transport), { siteUrl: "https://tenant.test/sites/demo" }).search({
    queryText: "A&B 'quoted'",
    selectProperties: ["Title", "Path"],
    refiners: ["author", "filetype"],
    sort: [{ property: "Created", direction: "descending" }],
    rowLimit: 2,
    startRow: 0,
    maxPages: 2,
    signal: controller.signal,
    timeoutMs: 125,
  });
  const first = JSON.parse(String(transport.calls[0].options.body)) as Record<string, any>;
  assert.deepEqual(first, {
    __metadata: { type: "Microsoft.Office.Server.Search.REST.SearchRequest" },
    Querytext: "A&B 'quoted'",
    SelectProperties: { results: ["Title", "Path"] },
    Refiners: "author,filetype",
    SortList: { results: [{ Property: "Created", Direction: "1" }] },
    RowLimit: 2,
    StartRow: 0,
  });
  assert.equal(transport.calls[0].url, "https://tenant.test/sites/demo/_api/search/query");
  assert.equal(transport.calls[0].options.signal, controller.signal);
  assert.equal(transport.calls[0].options.timeoutMs, 125);
  assert.equal(transport.calls[1].options.method, "POST");
  assert.equal(JSON.parse(String(transport.calls[1].options.body)).StartRow, 1);
  assert.deepEqual(result, { data: [{ Title: "one" }, { Title: "two" }], totalRows: 3, startRow: 0, rowLimit: 2 });
});

test("SharePoint search maps next links and does not over-fetch past maxPages", async () => {
  const transport = new ScriptedTransport([{ status: 200, body: JSON.stringify({ d: { query: { PrimaryQueryResult: { RelevantResults: { Table: { Rows: { results: [{ Title: "one" }] } }, TotalRows: 2, "@odata.nextLink": "/sites/demo/_api/search/query?page=2" } } } } }) }]);
  const result = await new SharePointSearchAdapter(client(transport), { siteUrl: "https://tenant.test/sites/demo" }).search({ queryText: "one", maxPages: 1 });
  assert.equal(result.nextPage, "/sites/demo/_api/search/query?page=2");
  assert.equal(transport.calls.length, 1);
});

test("SharePoint search validates request shapes before transport", async () => {
  const transport = new ScriptedTransport([]);
  const adapter = new SharePointSearchAdapter(client(transport), { siteUrl: "https://tenant.test/sites/demo" });
  for (const request of [
    { queryText: "" },
    { queryText: "x", rowLimit: 0 },
    { queryText: "x", startRow: 50_001 },
    { queryText: "x", sort: [{ property: "", direction: "ascending" as const }] },
  ]) await assert.rejects(adapter.search(request), (error: unknown) => error instanceof DataError && error.kind === "validation");
  assert.equal(transport.calls.length, 0);
});

test("Graph search builds typed JSON, maps hits, pages, and caps the request", async () => {
  const transport = new ScriptedTransport([
    { status: 200, body: JSON.stringify({ value: [{ searchTerms: ["contoso"], hitsContainers: [{ total: 3, moreResultsAvailable: true, hits: [{ hitId: "1", summary: "one", resource: { id: "a" } }] }] }] }) },
    { status: 200, body: JSON.stringify({ value: [{ hitsContainers: [{ total: 3, moreResultsAvailable: false, hits: [{ hitId: "2", resource: { id: "b" } }] }] }] }) },
  ]);
  const adapter = new GraphSearchAdapter(client(transport), { baseUrl: "https://graph.test/v1.0" });
  const result = await adapter.search({
    entityTypes: ["driveItem", "listItem"],
    queryText: "contoso & docs",
    fields: ["title", "webUrl"],
    size: 2,
    sortProperties: [{ name: "lastModifiedDateTime", isDescending: true }],
    maxPages: 2,
  });
  const first = JSON.parse(String(transport.calls[0].options.body)) as Record<string, any>;
  assert.deepEqual(first, { requests: [{ entityTypes: ["driveItem", "listItem"], query: { queryString: "contoso & docs" }, fields: ["title", "webUrl"], from: 0, size: 2, sortProperties: [{ name: "lastModifiedDateTime", isDescending: true }] }] });
  assert.deepEqual(result, {
    data: [{ hitId: "1", summary: "one", resource: { id: "a" } }, { hitId: "2", resource: { id: "b" } }],
    total: 3,
    moreResultsAvailable: false,
    from: 0,
    size: 2,
  });
  assert.equal(JSON.parse(String(transport.calls[1].options.body)).requests[0].from, 1);

  const limitedTransport = new ScriptedTransport([{ status: 200, body: JSON.stringify({ value: [{ hitsContainers: [{ total: 4, moreResultsAvailable: true, hits: [{ hitId: "1" }, { hitId: "2" }] }] }] }) }]);
  const limited = await new GraphSearchAdapter(client(limitedTransport), { baseUrl: "https://graph.test/v1.0" }).search({ entityTypes: ["driveItem"], queryText: "x", size: 25, maxItems: 1 });
  assert.equal(JSON.parse(String(limitedTransport.calls[0].options.body)).requests[0].size, 1);
  assert.deepEqual(limited.data, [{ hitId: "1" }]);
  assert.equal(limitedTransport.calls.length, 1);

  const pageBoundTransport = new ScriptedTransport([{ status: 200, body: JSON.stringify({ value: [{ hitsContainers: [{ total: 4, moreResultsAvailable: true, hits: [{ hitId: "1" }, { hitId: "2" }] }] }] }) }]);
  const pageBound = await new GraphSearchAdapter(client(pageBoundTransport), { baseUrl: "https://graph.test/v1.0" }).search({ entityTypes: ["driveItem"], queryText: "x", maxPages: 1 });
  assert.deepEqual(pageBound.data, [{ hitId: "1" }, { hitId: "2" }]);
  assert.equal(pageBoundTransport.calls.length, 1);
});

test("Graph search validates bounds and preserves structured service errors", async () => {
  const invalidTransport = new ScriptedTransport([]);
  const adapter = new GraphSearchAdapter(client(invalidTransport), { baseUrl: "https://graph.test/v1.0" });
  for (const request of [
    { entityTypes: [], queryText: "x" },
    { entityTypes: ["driveItem"], queryText: "" },
    { entityTypes: ["driveItem"], queryText: "x", from: -1 },
    { entityTypes: ["driveItem"], queryText: "x", size: 1_001 },
  ]) await assert.rejects(adapter.search(request), (error: unknown) => error instanceof DataError && error.kind === "validation");
  assert.equal(invalidTransport.calls.length, 0);

  const errorTransport = new ScriptedTransport([{ status: 400, body: JSON.stringify({ error: { code: "BadRequest", innerError: { reason: "invalid query" } } }) }]);
  await assert.rejects(new GraphSearchAdapter(client(errorTransport), { baseUrl: "https://graph.test/v1.0" }).search({ entityTypes: ["driveItem"], queryText: "x" }), (error: unknown) => error instanceof DataError && error.kind === "validation" && error.code === "BadRequest" && error.details?.error !== undefined);
});
