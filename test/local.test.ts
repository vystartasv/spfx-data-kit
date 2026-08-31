import test from "node:test";
import assert from "node:assert/strict";
import { DataError, LocalCrudResource } from "../src/index.js";

type Todo = { id: number; title: string; done: boolean };
type NewTodo = { title: string; done?: boolean };
const make = () => new LocalCrudResource<Todo, NewTodo, Partial<NewTodo>>({
  toEntity: (x, id) => ({ ...x, id, done: x.done ?? false }),
  initial: [{ id: 1, title: "one", done: false }],
  validateCreate: x => { if (!x.title) throw new Error("title required"); },
});
test("local CRUD happy path and bounded list", async () => { const r = make(); const c = await r.create({ title: "two" }); assert.equal(c.data.id, 2); assert.equal(c.etag, "1"); const listed = await r.list({ pageSize: 1, maxPages: 1 }); assert.deepEqual(listed.data[0], { id: 1, title: "one", done: false }); assert.deepEqual(listed.etags, { "1": "1" }); const u = await r.update(2, { done: true }, { etag: c.etag }); assert.equal(u.data.done, true); assert.equal(u.etag, "2"); await r.remove(2, { etag: u.etag }); await assert.rejects(() => r.get(2), (e: DataError) => e.kind === "not-found"); });
test("local validation and ETag conflicts are typed", async () => { const r = make(); await assert.rejects(() => r.create({ title: "" }), (e: DataError) => e.kind === "validation"); await assert.rejects(() => r.update(1, { title: "x" }, { etag: "0" }), (e: DataError) => e.kind === "conflict"); await assert.rejects(() => r.remove(1, { etag: "0" }), (e: DataError) => e.kind === "conflict"); });
test("local initial data rejects duplicate ids", () => assert.throws(() => new LocalCrudResource<Todo, NewTodo, Partial<NewTodo>>({ initial: [{ id: 1, title: "a", done: false }, { id: 1, title: "b", done: false }] }), (e: DataError) => e.kind === "conflict"));
test("local CRUD clones values at boundaries", async () => { const r = make(); const value = (await r.get(1)).data; value.title = "mutated"; const listed = (await r.list()).data[0]; listed.title = "mutated"; assert.equal((await r.get(1)).data.title, "one"); });
