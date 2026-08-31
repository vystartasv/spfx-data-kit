# spfx-data-kit

`spfx-data-kit` is a small typed CRUD contract for SharePoint list data in SPFx. It keeps application code focused on entities while leaving authentication, permissions, and transport policy to PnPjs/SPFx.

```ts
type Todo = { id: number; title: string; done: boolean };
type NewTodo = { title: string; done?: boolean };
type EditTodo = Partial<NewTodo>;
const todos = new PnPjsListResource<Todo, NewTodo, EditTodo>(sp, {
  title: "Todos", map: { id: "Id", title: "Title", done: "Done" },
  createMap: x => ({ Title: x.title, Done: x.done ?? false }),
  updateMap: x => ({ ...(x.title === undefined ? {} : { Title: x.title }), ...(x.done === undefined ? {} : { Done: x.done }) }),
});
const result = await todos.list({ select: ["Id", "Title", "Done"], top: 50, maxPages: 2 });
const first = result.data[0]; // Todo
const firstEtag = result.etags[String(first.id)]; // string | undefined
```

In a web part, initialize after `super.onInit()` and retain the returned `SPFI`:

```ts
protected async onInit() {
  await super.onInit();
  this.sp = createSpfi(this.context);
}
```

`list` uses PnPjs v4's async iterator, whose values are arrays (server pages). `pageSize` is the number requested per iterator page; `maxPages` is the maximum number of pages consumed, not an item count. The adapter appends each page and breaks before the iterator requests another page. `top` is an overall item cap. When `top` is supplied, the first request uses `min(top, pageSize)` (with the default page size 100), and the result contains at most `top` items. No Infinity arithmetic is used; omitted bounds are simply unbounded.

`DataResult<T>` is `{ data: T; etag?: string }` for single-item `get`, `create`, and `update` results. `ListResult<T>` is `{ data: T[]; etags: Readonly<Record<string, string>> }`: entities stay ergonomic, while available per-item ETags are keyed by SharePoint item ID. Initial local entities and local create/update state use ETag `"1"`; each successful update increments the numeric token. Optional write ETags are compared before local mutation and stale values produce `conflict`. Every local value is cloned on input and output. SharePoint writes pass the optional ETag to v4 `.update(properties, eTag)`/`.delete(eTag)`. SharePoint update can return 204/no content, so the adapter always re-reads the item after `.update`.

Writes accept an optional ETag. A stale ETag yields `DataError` with kind `conflict`; `not-found`, `validation`, `permission`, `transient`, and `unknown` are also stable error kinds. Causes are retained for diagnostics, but raw SharePoint fields are never returned—only the configured field map is exposed.

`LocalCrudResource` supplies deterministic in-memory CRUD for tests and examples, including numeric IDs, version ETags, duplicate/validation hooks, and conflict checks. Tests can call `resource.etag(id)` to obtain the current token.

This library intentionally does not abstract auth, permission discovery, Graph, arbitrary URLs, retries, batching, telemetry, caching, or UI frameworks. Tenant-specific validation (content types, required fields, permissions, and business rules) cannot be proven locally; use validation hooks and SharePoint's response as the authority.

PnPjs v4 assumptions and API details: [getting started](https://github.com/pnp/pnpjs/blob/version-4/docs/getting-started.md), [items](https://github.com/pnp/pnpjs/blob/version-4/docs/sp/items.md), [async paging](https://github.com/pnp/pnpjs/blob/version-4/docs/concepts/async-paging.md), and [SPFx setup](https://github.com/sharepoint/sp-dev-docs/blob/main/docs/spfx/web-parts/guidance/use-sp-pnp-js-with-spfx-web-parts.md).
