import test from "node:test";
import assert from "node:assert/strict";
import { mapSharePointError } from "../src/index.js";
test("SharePoint statuses map to stable kinds and preserve cause", () => { for (const [status, kind] of [[401,"permission"],[403,"permission"],[404,"not-found"],[409,"conflict"],[412,"conflict"],[429,"transient"],[500,"transient"],[503,"transient"]] as const) { const cause = { status, secret: "not exposed by message" }; const error = mapSharePointError(cause); assert.equal(error.kind, kind); assert.equal(error.cause, cause); assert.equal(error.message.includes("secret"), false); } });
