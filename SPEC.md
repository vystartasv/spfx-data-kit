# spfx-data-kit

§G
Small typed SPFx SharePoint CRUD library; hide repetitive PnPjs calls without becoming HTTP/framework layer.

§C
- TypeScript; Node 22; ESM
- @pnp/sp v4 peer/dev dependency
- bounded async paging; no getPaged
- typed errors; preserve safe original cause
- no auth, permissions, Graph, retries, batching, telemetry, UI

§I
- CrudResource<TEntity,TCreate,TUpdate>
- ListQuery
- DataResult/ListResult/DataError
- LocalCrudResource
- PnPjsListResource
- createSpfi(context)
- package exports: dist/index.js and dist/index.d.ts

§V
V1. Public CRUD methods preserve entity/create/update types.
V2. `pageSize` is items per server iterator page; `maxPages` bounds iterator pages; `top` bounds total items; no page is requested after either bound.
V3. Local writes reject stale etags and invalid/duplicate input with typed errors.
V4. PnPjs v4 mapping exposes only configured fields; update passes optional ETag and re-reads after possible 204/no content; delete passes optional ETag.
V5. SharePoint failures map to stable error kinds and retain original cause.
V6. SPFI helper is used after SPFx super.onInit lifecycle.
V7. Local initial/create ETag is `"1"`; successful update increments it; values are cloned at every boundary.
V8. With `top=N`, first server page size is `min(N,pageSize)` and result has at most N items; omitted limits are unbounded without Infinity arithmetic.
V9. `DataResult<T>` is `{data:T;etag?:string}` for single-item results; `ListResult<T>` is `{data:T[];etags:Readonly<Record<string,string>>}`; list data is plain entities with separate per-id ETags; local values remain cloned.
V10. PnPjs maps `@odata.etag`, `odata.etag`, or `ETag` when present; update re-read supplies the fresh ETag.

§T
id|status|task|cites
T1|.|package TypeScript ESM library|V1,I.api
T2|.|implement local CRUD and errors|V1,V3
T3|.|implement PnPjs v4 adapter and paging|V2,V4,V5
T4|.|tests for local, adapter, paging, errors, ETags|V1-V5,V9-V10
T5|.|README, license, lifecycle docs|V6,I.api

§B
id|date|cause|fix
