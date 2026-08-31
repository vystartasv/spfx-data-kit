import test from "node:test";
import assert from "node:assert/strict";
import type { SPFI } from "@pnp/sp";
import { DataError, PnPjsListResource } from "../src/index.js";
type FakeItem = (() => Promise<Record<string, unknown>>) & { update(values: Record<string, unknown>, etag?: string): Promise<void>; delete(etag?: string): Promise<void> };
test("PnP adapter uses typed field mapping, paging, ETags, and update re-read", async () => {
  const calls: string[] = []; let raw: Record<string, unknown> = { Id: 4, Title: "old", Done: false, "@odata.etag": '"1"' };
  const item = (): FakeItem => Object.assign(async () => { calls.push("get"); return raw; }, { async update(values: Record<string, unknown>, etag?: string) { calls.push(`update:${etag}`); if (etag !== raw["@odata.etag"]) throw { status: 412 }; raw = { ...raw, ...values, "@odata.etag": `"${Number(String(raw["@odata.etag"]).replaceAll('"', "")) + 1}"` }; }, async delete(etag?: string) { calls.push(`delete:${etag}`); if (etag !== raw["@odata.etag"]) throw { status: 412 }; } });
  const pages = [[raw], [{ Id: 5, Title: "second", Done: true }]];
  const items = { select(...fields: string[]) { calls.push(`select:${fields}`); return this; }, expand(...fields: string[]) { calls.push(`expand:${fields}`); return this; }, filter(value: string) { calls.push(`filter:${value}`); return this; }, orderBy(field: string, ascending: boolean) { calls.push(`order:${field}:${ascending}`); return this; }, top(value: number) { calls.push(`top:${value}`); return { async *[Symbol.asyncIterator]() { for (const page of pages) yield page; } }; }, getById(_id: number) { return item(); }, async add(values: Record<string, unknown>) { calls.push(`add:${values.Title}`); return { data: { Id: 4 } }; } };
  const sp = { web: { lists: { getByTitle(title: string) { calls.push(`title:${title}`); return { items }; } } } } as unknown as SPFI;
  const r = new PnPjsListResource<{ id: number; title: string; done: boolean }, { title: string }, { title: string }>(sp, { title: "Todos", map: { id: "Id", title: "Title", done: "Done" }, createMap: x => ({ Title: x.title }), updateMap: x => ({ Title: x.title }) });
  const listed = await r.list({ select: ["Id"], pageSize: 2, maxPages: 2 }); assert.deepEqual(listed.data, [{ id: 4, title: "old", done: false }, { id: 5, title: "second", done: true }]); assert.deepEqual(listed.etags, { "4": '"1"' }); assert.ok(calls.includes("top:2"));
  assert.equal((await r.get(4)).etag, '"1"'); const updated = await r.update(4, { title: "new" }, { etag: "\"1\"" }); assert.equal(updated.data.title, "new"); assert.equal(updated.etag, '"2"'); assert.ok(calls.includes("update:\"1\""));
  await assert.rejects(() => r.update(4, { title: "stale" }, { etag: '"1"' }), (e: DataError) => e.kind === "conflict"); await assert.rejects(() => r.remove(4, { etag: '"1"' }), (e: DataError) => e.kind === "conflict"); await r.create({ title: "new" }); await r.remove(4, { etag: '"2"' }); assert.ok(calls.includes("add:new")); assert.ok(calls.includes("delete:\"2\""));
});
test("maxPages makes exactly two page next calls and no page three", async () => {
  let nextCalls = 0; const iterator = { [Symbol.asyncIterator]() { let index = 0; return { async next() { nextCalls++; return index++ < 3 ? { value: [{ Id: index, Title: String(index) }], done: false } : { value: undefined, done: true }; } }; } };
  const items = { top(_size: number) { return iterator; } }; const sp = { web: { lists: { getByTitle: () => ({ items }) } } } as unknown as SPFI; const r = new PnPjsListResource(sp, { title: "x", map: { id: "Id", title: "Title" } });
  assert.equal((await r.list({ pageSize: 2, maxPages: 2 })).data.length, 2); assert.equal(nextCalls, 2);
});
test("top is an overall cap and uses min(top, pageSize) for the first page", async () => {
  const sizes: number[] = []; const items = { top(size: number) { sizes.push(size); return { async *[Symbol.asyncIterator]() { yield [{ Id: 1, Title: "one" }, { Id: 2, Title: "two" }]; yield [{ Id: 3, Title: "three" }]; } }; } }; const sp = { web: { lists: { getByTitle: () => ({ items }) } } } as unknown as SPFI; const r = new PnPjsListResource(sp, { title: "x", map: { id: "Id", title: "Title" } });
  assert.equal((await r.list({ top: 2, pageSize: 5 })).data.length, 2); assert.deepEqual(sizes, [2]);
});
test("zero bounds still validate pageSize", async () => {
  const items = { top() { throw new Error("must not request"); } }; const sp = { web: { lists: { getByTitle: () => ({ items }) } } } as unknown as SPFI; const r = new PnPjsListResource(sp, { title: "x", map: { id: "Id" } });
  await assert.rejects(() => r.list({ top: 0, pageSize: 0 }), (e: DataError) => e.kind === "validation");
  await assert.rejects(() => r.list({ maxPages: 0, pageSize: 0 }), (e: DataError) => e.kind === "validation");
});
