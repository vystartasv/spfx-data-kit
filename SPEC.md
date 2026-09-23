# spfx-data-kit

§G
Dependency-free TypeScript ESM library for injected request transport, typed SharePoint REST data access, and Microsoft Graph JSON batch results.

§C
- TypeScript; Node 22; ESM
- Runtime dependencies: none
- Host owns authentication, authorization, URL discovery, network access
- Public `DataError.kind` values stable
- Tests use injected transports; no network
- Build clears `dist` before compile
- PnPjs, local storage, UI, and server dependencies forbidden
- Binary payloads → `Uint8Array`; large/resumable upload ⊥ planned

§I
- Package root: `spfx-data-kit`
- Capabilities: `spfx-data-kit/capabilities`
- Contracts: `RequestTransport`, `DataRequest`, `DataResult`, `ListResult`, `ListQuery`, `CrudResource`
- Errors: `DataError`, `headerValue`, `retryAfterMilliseconds`, `mapHttpError`, `mapSharePointError`, `parseJson`
- Client: `DataClient`, `requestCacheKey`, cache/retry/diagnostic option types
- SharePoint: `SharePointRestAdapter`, URL helpers, CRUD/query/ETag/batch types
- SharePoint files: `SharePointFilesAdapter`, file/folder/attachment URL helpers, metadata, binary download, bounded upload types
- Graph: `GraphAdapter`, request/page/delta/batch option/result types
- Manifest: `capabilityManifest`, `CAPABILITY_MANIFEST`

§V
V1: Root export includes contracts, errors, client, SharePoint, Graph, and manifest APIs.
V2: `RequestTransport` remains sole host boundary for request execution.
V3: GET requests support in-flight deduplication; configured cache enforces positive `maxEntries` and `ttlMs`; metadata-aware GETs retain response headers; cache invalidation removes matching entries.
V4: Only GET requests retry; retry count is bounded; non-GET requests never retry automatically.
V5: HTTP statuses map to stable `DataError.kind` values; `cause`, `status`, and retry delay remain available where supplied.
V6: SharePoint list reads support select, expand, filter, orderBy, top, pageSize, maxPages, next links, mapped entities, and per-item ETags.
V7: SharePoint writes use POST/DELETE, optional `IF-MATCH`, invalidate list URL cache entries, and re-read created/updated items.
V8: SharePoint batch children accept GET only; results preserve child status, headers, body, id, and `ok`.
V9: Graph batch child URLs are relative, ids are unique and non-empty, and request count ≤ 20; results preserve child failures.
V10: Build output contains only current `src/**/*.ts` compilation; package exports and files contain no PnPjs/local artifacts.
V11: Transport and client requests forward optional `AbortSignal` and non-negative `timeoutMs`; aborted requests do not retry; timeout enforcement remains host-owned.
V12: `DataError.kind` values remain stable; structured service `code` and object `details` are retained when an HTTP error body provides them.
V13: Graph page/delta iteration validates non-negative bounds, never requests after `maxPages` or `maxItems`, truncates the final page to `maxItems`, and stops repeated links.
V14: Text callers unchanged; binary req → `responseType: "binary"`, `Uint8Array` body, host `arrayBuffer()` for lossless bytes.
V15: SharePoint file/folder paths → encoded OData literals ∈ configured site; attachments ∈ configured list & positive item id.
V16: File/folder/attachment metadata ! required name & server-relative URL; upload content >1_500_000 bytes → reject before transport.
V17: File/folder/attachment req → forward `AbortSignal`/`timeoutMs`; metadata → preserve ETags; meaningful writes → `IF-MATCH`.
V18: Folder children → one expanded GET; upload/download ⊥ automatic follow-up reads.

§T
id|status|task|cites
T1|x|publish dependency-free ESM package and capability subpath|V1,V10,I.api
T2|x|implement injected transport client, cache, invalidation, diagnostics, retries, and errors|V2-V5
T3|x|implement SharePoint REST CRUD, paging, queries, ETags, and GET batches|V6-V8
T4|x|implement Graph requests, JSON batches, and capability manifest|V1,V9,I.api
T5|x|restore concise docs and remove stale PnPjs/local package artifacts|V10
T6|x|add PnP capability matrix and Graph page/delta iteration|V9,V13
T7|x|forward cancellation/deadline options and retain structured HTTP error details|V11,V12
T8|x|add binary transport, typed SharePoint files/folders/attachments, bounded uploads|V14-V18

§B
id|date|cause|fix
