# spfx-data-kit

`spfx-data-kit` is a dependency-free TypeScript ESM library for request transport, typed SharePoint REST data access, and typed Microsoft Graph data access.

Authentication, authorization, URL discovery, and network access remain with the host through `RequestTransport`.

```ts
import { DataClient, GraphAdapter, GraphDirectoryAdapter, GraphDriveAdapter, GraphSitesAdapter } from "spfx-data-kit";

const client = new DataClient(transport, {
  retry: { maxRetries: 2 },
  cache: { maxEntries: 100, ttlMs: 30_000 },
});

const graph = new GraphAdapter(client);
const result = await graph.request<{ id: string }>("/me");
const drive = new GraphDriveAdapter(client);
const bytes = await drive.downloadFile("drive-id", "item-id");
const sites = new GraphSitesAdapter(client);
const site = await sites.getSite("site-id");
const directory = new GraphDirectoryAdapter(client);
const user = await directory.getCurrentUser({ select: ["id", "displayName"] });
```

The root export includes:

- transport, CRUD, query, result, and error contracts;
- `DataError` and HTTP/JSON error helpers;
- `DataClient` with GET deduplication, bounded in-memory TTL caching, invalidation, diagnostics, and bounded GET retries;
- binary-safe `Uint8Array` transport bodies/responses through `DataClient.requestBytes`;
- `SharePointRestAdapter` for list paging, query construction, ETags, CRUD, and GET-only batch result parsing;
- `SharePointFilesAdapter` for site-scoped file/folder metadata, one-request folder children, binary downloads, bounded small uploads (≤1,500,000 bytes), and list-item attachment metadata/download/upload/delete;
- `SharePointSearchAdapter` for typed SharePoint REST search requests, select properties, refiners, sort, row bounds, result totals, and bounded result paging;
- `GraphAdapter` for Graph requests with cached/deduplicated GETs, response-header ETags, page/delta parsing, bounded async iteration, and JSON batch result parsing;
- `GraphSearchAdapter` for typed Microsoft Graph Search requests across one or more entity types, selected fields, sort properties, totals, and bounded `from`/`size` paging;
- `GraphDriveAdapter` for typed drive/root/item metadata, bounded children paging, binary file download, ETag-aware item update/delete, and bounded simple `Uint8Array` uploads (≤1,500,000 bytes);
- `GraphSitesAdapter` for typed site-by-id/path and list metadata, OData list/item queries, bounded paging, list-item CRUD, response ETags, and optional `If-Match` writes without follow-up reads;
- `GraphDirectoryAdapter` for typed current-user/user/group reads, bounded users/groups/group-members paging, OData select/filter/orderBy/top queries, and Graph `$ref` member add/remove without follow-up reads;
- transport `AbortSignal`/host-enforced timeout forwarding and structured service error details without changing `DataError.kind`;
- `capabilityManifest` and `CAPABILITY_MANIFEST`.

The capability manifest is also available from `spfx-data-kit/capabilities`.

Build and test commands:

```sh
npm test
npm run check
npm run build
npm run pack
```

The package does not provide authentication, authorization, tenant discovery, persistent storage, offline storage, UI components, or network access.

Graph mail, calendar, Teams, Planner, To Do, admin, provisioning, pages, navigation, profiles, taxonomy, resumable uploads, thumbnails, previews, permissions, broader Graph service families, large/resumable SharePoint uploads, and other PnP wrapper families remain planned or excluded; see [docs/pnp-parity.md](docs/pnp-parity.md).
