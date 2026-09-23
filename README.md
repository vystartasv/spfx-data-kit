# spfx-data-kit

`spfx-data-kit` is a dependency-free TypeScript ESM library for request transport, typed SharePoint REST list CRUD, and Microsoft Graph JSON batch results.

Authentication, authorization, URL discovery, and network access remain with the host through `RequestTransport`.

```ts
import { DataClient, GraphAdapter } from "spfx-data-kit";

const client = new DataClient(transport, {
  retry: { maxRetries: 2 },
  cache: { maxEntries: 100, ttlMs: 30_000 },
});

const graph = new GraphAdapter(client);
const result = await graph.request<{ id: string }>("/me");
```

The root export includes:

- transport, CRUD, query, result, and error contracts;
- `DataError` and HTTP/JSON error helpers;
- `DataClient` with GET deduplication, bounded in-memory TTL caching, invalidation, diagnostics, and bounded GET retries;
- `SharePointRestAdapter` for list paging, query construction, ETags, CRUD, and GET-only batch result parsing;
- `GraphAdapter` for Graph requests and JSON batch result parsing;
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
