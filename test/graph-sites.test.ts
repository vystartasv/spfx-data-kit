import test from "node:test";
import assert from "node:assert/strict";
import {
  DataClient,
  DataError,
  GraphSitesAdapter,
  graphListItemFieldsUrl,
  graphSiteByPathUrl,
} from "../src/index.js";
import { ScriptedTransport } from "./helpers.js";

const sites = (transport: ScriptedTransport) => new GraphSitesAdapter(new DataClient(transport, { retry: { maxRetries: 0 } }), { baseUrl: "https://graph.test/v1.0" });

test("Graph sites build relative site/list URLs and escape OData query values", async () => {
  const transport = new ScriptedTransport([
    { status: 200, body: JSON.stringify({ id: "s1", displayName: "HR" }) },
    { status: 200, body: JSON.stringify({ value: [{ id: "l1" }] }) },
    { status: 200, body: JSON.stringify({ value: [{ id: "i1" }] }) },
  ]);
  const adapter = sites(transport);
  assert.equal(graphSiteByPathUrl("contoso.sharepoint.com", "/sites/hr"), "/sites/contoso.sharepoint.com:/sites/hr");
  assert.equal(graphSiteByPathUrl("contoso.sharepoint.com", "/"), "/sites/contoso.sharepoint.com:/");
  await adapter.getSiteByPath("contoso.sharepoint.com", "/sites/hr");
  await adapter.listLists("s1", { select: ["id", "displayName"], top: 2 });
  await adapter.listItems("s1", "l1", {
    select: ["fields"],
    expand: ["fields"],
    filter: "fields/Title eq 'A&B'",
    orderBy: [["fields/Title", false]],
    top: 2,
  });
  assert.deepEqual(transport.calls.map((call) => call.url), [
    "https://graph.test/v1.0/sites/contoso.sharepoint.com:/sites/hr",
    "https://graph.test/v1.0/sites/s1/lists?$select=id%2CdisplayName&$top=2",
    "https://graph.test/v1.0/sites/s1/lists/l1/items?$select=fields&$expand=fields&$filter=fields%2FTitle%20eq%20%27A%26B%27&$orderby=fields%2FTitle%20desc&$top=2",
  ]);
});

test("Graph site and list paging honors next links and bounds without over-fetch", async () => {
  const transport = new ScriptedTransport([
    { status: 200, body: JSON.stringify({ value: [{ id: "i1" }, { id: "i2" }], "@odata.nextLink": "/sites/s1/lists/l1/items?page=2" }) },
    { status: 200, body: JSON.stringify({ value: [{ id: "i3" }], "@odata.nextLink": "/sites/s1/lists/l1/items?page=3" }) },
  ]);
  const items: string[] = [];
  for await (const item of sites(transport).iterateItems("s1", "l1", { maxPages: 2 })) items.push(item.id);
  assert.deepEqual(items, ["i1", "i2", "i3"]);
  assert.deepEqual(transport.calls.map((call) => call.url), [
    "https://graph.test/v1.0/sites/s1/lists/l1/items",
    "https://graph.test/v1.0/sites/s1/lists/l1/items?page=2",
  ]);

  const limitedTransport = new ScriptedTransport([
    { status: 200, body: JSON.stringify({ value: [{ id: "i1" }, { id: "i2" }], "@odata.nextLink": "/sites/s1/lists/l1/items?page=2" }) },
  ]);
  const limited: string[] = [];
  for await (const item of sites(limitedTransport).iterateItems("s1", "l1", { maxItems: 1 })) limited.push(item.id);
  assert.deepEqual(limited, ["i1"]);
  assert.equal(limitedTransport.calls.length, 1);
});

test("Graph list CRUD sends typed bodies, ETags, and transport controls without follow-up reads", async () => {
  const controller = new AbortController();
  const transport = new ScriptedTransport([
    { status: 200, headers: { ETag: '"1"' }, body: JSON.stringify({ id: "i1", fields: { Title: "old" } }) },
    { status: 201, headers: { ETag: '"2"' }, body: JSON.stringify({ id: "i2", fields: { Title: "new" } }) },
    { status: 200, body: JSON.stringify({ Title: "updated" }) },
    { status: 204 },
  ]);
  const adapter = sites(transport);
  assert.deepEqual(await adapter.getItem("s1", "l1", "i1", { signal: controller.signal, timeoutMs: 125 }), {
    data: { id: "i1", fields: { Title: "old" } },
    etag: '"1"',
  });
  assert.deepEqual(await adapter.createItem("s1", "l1", { fields: { Title: "new" } }, { signal: controller.signal, timeoutMs: 125 }), {
    data: { id: "i2", fields: { Title: "new" } },
    etag: '"2"',
  });
  await adapter.updateItem("s1", "l1", "i2", { Title: "updated" }, { etag: '"2"', signal: controller.signal, timeoutMs: 125 });
  await adapter.deleteItem("s1", "l1", "i2", { etag: '"3"', signal: controller.signal, timeoutMs: 125 });
  assert.deepEqual(transport.calls.map((call) => [call.options.method, call.url]), [
    ["GET", "https://graph.test/v1.0/sites/s1/lists/l1/items/i1"],
    ["POST", "https://graph.test/v1.0/sites/s1/lists/l1/items"],
    ["PATCH", "https://graph.test/v1.0/sites/s1/lists/l1/items/i2/fields"],
    ["DELETE", "https://graph.test/v1.0/sites/s1/lists/l1/items/i2"],
  ]);
  assert.equal(transport.calls[1].options.body, JSON.stringify({ fields: { Title: "new" } }));
  assert.equal(transport.calls[2].options.body, JSON.stringify({ Title: "updated" }));
  assert.equal(transport.calls[2].options.headers?.["If-Match"], '"2"');
  assert.equal(transport.calls[3].options.headers?.["If-Match"], '"3"');
  assert.equal(transport.calls[0].options.signal, controller.signal);
  assert.equal(transport.calls[0].options.timeoutMs, 125);
  assert.equal(transport.calls.length, 4);
});

test("Graph sites validate relative identifiers, paths, and bounds before transport", async () => {
  const transport = new ScriptedTransport([]);
  const adapter = sites(transport);
  assert.throws(() => graphSiteByPathUrl("contoso.sharepoint.com", "https://evil.test/sites/hr"), (error: unknown) => error instanceof DataError && error.kind === "validation");
  assert.throws(() => graphListItemFieldsUrl("s/1", "l1", "i1"), (error: unknown) => error instanceof DataError && error.kind === "validation");
  assert.throws(() => adapter.getSite("s/1"), (error: unknown) => error instanceof DataError && error.kind === "validation");
  assert.throws(() => adapter.listItems("s1", "l1", { top: -1 }), (error: unknown) => error instanceof DataError && error.kind === "validation");
  assert.throws(() => adapter.listItems("s1", "l1", { maxItems: -1 }), (error: unknown) => error instanceof DataError && error.kind === "validation");
  assert.equal(transport.calls.length, 0);
});
