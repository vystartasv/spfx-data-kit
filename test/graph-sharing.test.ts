import test from "node:test";
import assert from "node:assert/strict";
import {
  DataClient,
  DataError,
  GraphSharingAdapter,
  encodeGraphSharingUrl,
  graphDriveItemCreateLinkUrl,
  graphDriveItemPermissionUrl,
  graphDriveItemPermissionsUrl,
  graphDriveItemInviteUrl,
  graphSharingPermissionGrantUrl,
  graphSitePermissionUrl,
  graphSitePermissionsUrl,
} from "../src/index.js";
import { ScriptedTransport } from "./helpers.js";

const sharing = (transport: ScriptedTransport) => new GraphSharingAdapter(new DataClient(transport, { retry: { maxRetries: 0 } }), { baseUrl: "https://graph.test/v1.0" });

test("Graph sharing builds relative URLs and preserves permission ETags", async () => {
  const transport = new ScriptedTransport([
    { status: 200, headers: { ETag: '"list"' }, body: JSON.stringify({ value: [{ id: "p1", roles: ["read"] }] }) },
    { status: 200, headers: { ETag: '"get"' }, body: JSON.stringify({ id: "p1", roles: ["read"] }) },
    { status: 204, headers: { ETag: '"delete"' } },
  ]);
  const adapter = sharing(transport);
  assert.equal(graphDriveItemPermissionsUrl("d1", "i1"), "/drives/d1/items/i1/permissions");
  assert.equal(graphDriveItemPermissionUrl("d1", "i1", "p1"), "/drives/d1/items/i1/permissions/p1");
  assert.equal(graphSitePermissionsUrl("s1"), "/sites/s1/permissions");
  assert.equal(graphSitePermissionUrl("s1", "p1"), "/sites/s1/permissions/p1");
  assert.deepEqual(await adapter.listDriveItemPermissions("d1", "i1", { select: ["id", "roles"], etag: '"old"' }), {
    data: { value: [{ id: "p1", roles: ["read"] }] }, etag: '"list"',
  });
  assert.deepEqual(await adapter.getSitePermission("s1", "p1"), { data: { id: "p1", roles: ["read"] }, etag: '"get"' });
  assert.deepEqual(await adapter.deleteSitePermission("s1", "p1", { etag: '"get"' }), { data: undefined, etag: '"delete"' });
  assert.deepEqual(transport.calls.map((call) => [call.options.method, call.url]), [
    ["GET", "https://graph.test/v1.0/drives/d1/items/i1/permissions?$select=id%2Croles"],
    ["GET", "https://graph.test/v1.0/sites/s1/permissions/p1"],
    ["DELETE", "https://graph.test/v1.0/sites/s1/permissions/p1"],
  ]);
  assert.equal(transport.calls[0].options.headers?.["If-Match"], '"old"');
  assert.equal(transport.calls[2].options.headers?.["If-Match"], '"get"');
  assert.equal(transport.calls.length, 3);
});

test("Graph sharing constructs link, invite, and grant requests without follow-up reads", async () => {
  const controller = new AbortController();
  const shareUrl = "https://contoso.sharepoint.com/:w:/s/team/Eabc?e=xyz";
  const transport = new ScriptedTransport([
    { status: 201, headers: { ETag: '"link"' }, body: JSON.stringify({ id: "p1", roles: ["read"], link: { type: "view", scope: "organization" } }) },
    { status: 200, body: JSON.stringify({ value: [{ id: "p2", roles: ["write"] }] }) },
    { status: 200, body: JSON.stringify({ value: [{ id: "p3", roles: ["read"] }, { error: { code: "recipientNotFound" } }] }) },
  ]);
  const adapter = sharing(transport);
  assert.equal(graphDriveItemCreateLinkUrl("d1", "i1"), "/drives/d1/items/i1/createLink");
  assert.equal(graphDriveItemInviteUrl("d1", "i1"), "/drives/d1/items/i1/invite");
  assert.equal(graphSharingPermissionGrantUrl(shareUrl), `/shares/${encodeGraphSharingUrl(shareUrl)}/permission/grant`);
  const link = await adapter.createLink("d1", "i1", { type: "view", scope: "organization", retainInheritedPermissions: true }, { signal: controller.signal, timeoutMs: 125 });
  await adapter.invite("d1", "i1", { recipients: [{ email: "reader@contoso.com" }], roles: ["write"], sendInvitation: true, message: "Review this file." }, { signal: controller.signal, timeoutMs: 125 });
  const grant = await adapter.grantAccess(shareUrl, { recipients: [{ email: "reader@contoso.com" }], roles: ["read"] }, { signal: controller.signal, timeoutMs: 125 });
  assert.deepEqual(link, { data: { id: "p1", roles: ["read"], link: { type: "view", scope: "organization" } }, etag: '"link"' });
  assert.deepEqual(grant.data.value[1], { error: { code: "recipientNotFound" } });
  assert.deepEqual(transport.calls.map((call) => [call.options.method, call.url]), [
    ["POST", "https://graph.test/v1.0/drives/d1/items/i1/createLink"],
    ["POST", "https://graph.test/v1.0/drives/d1/items/i1/invite"],
    ["POST", `https://graph.test/v1.0/shares/${encodeGraphSharingUrl(shareUrl)}/permission/grant`],
  ]);
  assert.equal(transport.calls[0].options.body, JSON.stringify({ type: "view", scope: "organization", retainInheritedPermissions: true }));
  assert.equal(transport.calls[1].options.body, JSON.stringify({ recipients: [{ email: "reader@contoso.com" }], roles: ["write"], sendInvitation: true, message: "Review this file." }));
  assert.equal(transport.calls[2].options.body, JSON.stringify({ recipients: [{ email: "reader@contoso.com" }], roles: ["read"] }));
  assert.equal(transport.calls[0].options.signal, controller.signal);
  assert.equal(transport.calls[0].options.timeoutMs, 125);
  assert.equal(transport.calls.length, 3);
});

test("Graph sharing validates ids, links, scopes, roles, recipients, and messages before transport", async () => {
  const transport = new ScriptedTransport([]);
  const adapter = sharing(transport);
  const invalid = [
    () => adapter.listDriveItemPermissions("d/1", "i1"),
    () => adapter.getSitePermission("s1", "p/1"),
    () => adapter.createLink("d1", "i1", { type: "download" as "view" }),
    () => adapter.createLink("d1", "i1", { type: "view", scope: "tenant" as "organization" }),
    () => adapter.invite("d1", "i1", { recipients: [], roles: ["read"] }),
    () => adapter.invite("d1", "i1", { recipients: [{ email: "not an email" }], roles: ["read"] }),
    () => adapter.invite("d1", "i1", { recipients: [{ email: "a@b.test" }], roles: ["owner" as "read"] }),
    () => adapter.grantAccess("not-a-url", { recipients: [{ email: "a@b.test" }], roles: ["read"] }),
    () => adapter.grantAccess("https://contoso.test/share", { recipients: [{ email: "a@b.test" }], roles: ["read", "write"] }),
    () => adapter.invite("d1", "i1", { recipients: [{ email: "a@b.test" }], roles: ["read"], message: "x".repeat(2_001) }),
  ];
  for (const operation of invalid) await assert.rejects(async () => operation(), (error: unknown) => error instanceof DataError && error.kind === "validation");
  assert.throws(() => graphSharingPermissionGrantUrl("file:///tmp/share"), (error: unknown) => error instanceof DataError && error.kind === "validation");
  assert.equal(transport.calls.length, 0);
});

test("Graph sharing retains structured service errors", async () => {
  const transport = new ScriptedTransport([{ status: 403, body: JSON.stringify({ error: { code: "accessDenied", message: "No access" } }) }]);
  await assert.rejects(sharing(transport).invite("d1", "i1", { recipients: [{ email: "a@b.test" }], roles: ["read"] }), (error: unknown) => {
    return error instanceof DataError && error.kind === "permission" && error.code === "accessDenied" && error.details?.error !== undefined;
  });
  assert.equal(transport.calls.length, 1);
});
