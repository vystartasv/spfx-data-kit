# spfx-data-kit

§G
Dependency-free TypeScript ESM library for injected request transport, typed SharePoint REST list CRUD, and Microsoft Graph JSON batch results.

§C
- TypeScript; Node 22; ESM
- Runtime dependencies: none
- Host owns authentication, authorization, URL discovery, network access
- Public `DataError.kind` values stable
- Tests use injected transports; no network
- Build clears `dist` before compile
- PnPjs, local storage, UI, and server dependencies forbidden

§I
- Package root: `spfx-data-kit`
- Capabilities: `spfx-data-kit/capabilities`
- Contracts: `RequestTransport`, `DataRequest`, `DataResult`, `ListResult`, `ListQuery`, `CrudResource`
- Errors: `DataError`, `headerValue`, `retryAfterMilliseconds`, `mapHttpError`, `mapSharePointError`, `parseJson`
- Client: `DataClient`, `requestCacheKey`, cache/retry/diagnostic option types
- SharePoint: `SharePointRestAdapter`, URL helpers, CRUD/query/ETag/batch types
- Graph: `GraphAdapter`, request/batch option/result types
- Manifest: `capabilityManifest`, `CAPABILITY_MANIFEST`

§V
V1: Root export includes contracts, errors, client, SharePoint, Graph, and manifest APIs.
V2: `RequestTransport` remains sole host boundary for request execution.
V3: GET requests support in-flight deduplication; configured cache enforces positive `maxEntries` and `ttlMs`; cache invalidation removes matching entries.
V4: Only GET requests retry; retry count is bounded; non-GET requests never retry automatically.
V5: HTTP statuses map to stable `DataError.kind` values; `cause`, `status`, and retry delay remain available where supplied.
V6: SharePoint list reads support select, expand, filter, orderBy, top, pageSize, maxPages, next links, mapped entities, and per-item ETags.
V7: SharePoint writes use POST/DELETE, optional `IF-MATCH`, invalidate list URL cache entries, and re-read created/updated items.
V8: SharePoint batch children accept GET only; results preserve child status, headers, body, id, and `ok`.
V9: Graph batch child URLs are relative, ids are unique and non-empty, and request count ≤ 20; results preserve child failures.
V10: Build output contains only current `src/**/*.ts` compilation; package exports and files contain no PnPjs/local artifacts.

§T
id|status|task|cites
T1|x|publish dependency-free ESM package and capability subpath|V1,V10,I.api
T2|x|implement injected transport client, cache, invalidation, diagnostics, retries, and errors|V2-V5
T3|x|implement SharePoint REST CRUD, paging, queries, ETags, and GET batches|V6-V8
T4|x|implement Graph requests, JSON batches, and capability manifest|V1,V9,I.api
T5|x|restore concise docs and remove stale PnPjs/local package artifacts|V10

§B
id|date|cause|fix
