import test from "node:test";
import assert from "node:assert/strict";
import {
  DataClient,
  DataError,
  GraphDirectoryAdapter,
  graphDirectoryObjectUrl,
  graphGroupMembersRefUrl,
  graphGroupUrl,
  graphUserUrl,
} from "../src/index.js";
import { ScriptedTransport } from "./helpers.js";

const directory = (transport: ScriptedTransport) => new GraphDirectoryAdapter(new DataClient(transport, { retry: { maxRetries: 0 } }), { baseUrl: "https://graph.test/v1.0" });

test("Graph directory builds relative URLs and escapes collection query values", async () => {
  const transport = new ScriptedTransport([
    { status: 200, headers: { ETag: '"u1"' }, body: JSON.stringify({ id: "u1" }) },
    { status: 200, headers: { ETag: '"g1"' }, body: JSON.stringify({ id: "g1" }) },
    { status: 200, body: JSON.stringify({ value: [{ id: "u1" }] }) },
  ]);
  const adapter = directory(transport);
  assert.equal(graphUserUrl("user one"), "/users/user%20one");
  assert.equal(graphGroupUrl("g1"), "/groups/g1");
  assert.equal(graphGroupMembersRefUrl("g1", "u1"), "/groups/g1/members/u1/$ref");
  assert.equal(graphDirectoryObjectUrl("u1"), "/directoryObjects/u1");
  assert.deepEqual(await adapter.getUser("user one", { select: ["id", "displayName"] }), { data: { id: "u1" }, etag: '"u1"' });
  assert.deepEqual(await adapter.getGroup("g1", { select: ["id"] }), { data: { id: "g1" }, etag: '"g1"' });
  await adapter.listUsers({
    select: ["id", "displayName"],
    filter: "displayName eq 'A&B'",
    orderBy: [["displayName", false]],
    top: 2,
  });
  assert.deepEqual(transport.calls.map((call) => call.url), [
    "https://graph.test/v1.0/users/user%20one?$select=id%2CdisplayName",
    "https://graph.test/v1.0/groups/g1?$select=id",
    "https://graph.test/v1.0/users?$select=id%2CdisplayName&$filter=displayName%20eq%20%27A%26B%27&$orderby=displayName%20desc&$top=2",
  ]);
});

test("Graph directory collection paging honors bounds without over-fetch", async () => {
  const transport = new ScriptedTransport([
    { status: 200, body: JSON.stringify({ value: [{ id: "u1" }, { id: "u2" }], "@odata.nextLink": "/users?page=2" }) },
    { status: 200, body: JSON.stringify({ value: [{ id: "u3" }], "@odata.nextLink": "/users?page=3" }) },
  ]);
  const users: string[] = [];
  for await (const user of directory(transport).iterateUsers({ maxPages: 2 })) users.push(user.id);
  assert.deepEqual(users, ["u1", "u2", "u3"]);
  assert.deepEqual(transport.calls.map((call) => call.url), [
    "https://graph.test/v1.0/users",
    "https://graph.test/v1.0/users?page=2",
  ]);

  const membersTransport = new ScriptedTransport([
    { status: 200, body: JSON.stringify({ value: [{ id: "u1" }, { id: "u2" }], "@odata.nextLink": "/groups/g1/members?page=2" }) },
  ]);
  const members: string[] = [];
  for await (const member of directory(membersTransport).iterateMembers("g1", { maxItems: 1 })) members.push(member.id);
  assert.deepEqual(members, ["u1"]);
  assert.equal(membersTransport.calls.length, 1);

  const emptyTransport = new ScriptedTransport([]);
  assert.deepEqual(await directory(emptyTransport).listGroups({ maxPages: 0 }), { value: [] });
  assert.equal(emptyTransport.calls.length, 0);
});

test("Graph directory member writes use Graph ref semantics and transport controls", async () => {
  const controller = new AbortController();
  const transport = new ScriptedTransport([
    { status: 204, headers: { ETag: '"1"' } },
    { status: 204, headers: { ETag: '"2"' } },
  ]);
  const adapter = directory(transport);
  assert.deepEqual(await adapter.addMember("g1", "u1", { signal: controller.signal, timeoutMs: 125 }), { data: undefined, etag: '"1"' });
  assert.deepEqual(await adapter.removeMember("g1", "u1", { etag: '"1"', signal: controller.signal, timeoutMs: 125 }), { data: undefined, etag: '"2"' });
  assert.deepEqual(transport.calls.map((call) => [call.options.method, call.url]), [
    ["POST", "https://graph.test/v1.0/groups/g1/members/$ref"],
    ["DELETE", "https://graph.test/v1.0/groups/g1/members/u1/$ref"],
  ]);
  assert.equal(transport.calls[0].options.body, JSON.stringify({ "@odata.id": "https://graph.test/v1.0/directoryObjects/u1" }));
  assert.equal(transport.calls[1].options.headers?.["If-Match"], '"1"');
  assert.equal(transport.calls[0].options.signal, controller.signal);
  assert.equal(transport.calls[0].options.timeoutMs, 125);
  assert.equal(transport.calls.length, 2);
});

test("Graph directory validates identifiers and preserves membership failures", async () => {
  const transport = new ScriptedTransport([{ status: 403, body: JSON.stringify({ error: { code: "Forbidden" } }) }]);
  const adapter = directory(transport);
  for (const invalid of ["", "g/1", "g%2F1", "g?1"]) {
    assert.throws(() => graphGroupUrl(invalid), (error: unknown) => error instanceof DataError && error.kind === "validation");
  }
  assert.throws(() => graphDirectoryObjectUrl("%ZZ"), (error: unknown) => error instanceof DataError && error.kind === "validation");
  assert.throws(() => adapter.listMembers("g/1"), (error: unknown) => error instanceof DataError && error.kind === "validation");
  assert.throws(() => adapter.listUsers({ top: -1 }), (error: unknown) => error instanceof DataError && error.kind === "validation");
  await assert.rejects(adapter.addMember("g1", "u1"), (error: unknown) => error instanceof DataError && error.kind === "permission");
  assert.equal(transport.calls.length, 1);
});
