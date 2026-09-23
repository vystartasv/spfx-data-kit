# Agent guide

`spfx-data-kit` is a dependency-free TypeScript ESM package for request transport, typed SharePoint REST CRUD, and Microsoft Graph JSON batch results.

## Product map

- `src/contracts.ts`: public transport, CRUD, query, result, and error types.
- `src/client.ts`: injected-transport client with GET deduplication, bounded memory cache, invalidation, diagnostics, and bounded retries.
- `src/sharepoint.ts`: SharePoint REST list CRUD, paging, query construction, ETags, and GET-only batch parsing.
- `src/graph.ts`: Graph requests and JSON batch parsing with the service limit of 20 child requests.
- `src/manifest.ts`: capability manifest exported from `spfx-data-kit` and `spfx-data-kit/capabilities`.
- `src/index.ts`: package public entry.

Authentication, authorization, URL discovery, and network access belong to the host through `RequestTransport`. Do not add PnPjs, local-storage, UI, or server dependencies. Deleted PnP/local source must not return.

## Working rules

- Preserve the public ESM exports and stable `DataError.kind` values.
- Keep tests deterministic: use injected transports, fake clocks, and no network.
- Update `README.md` and `SPEC.md` when public behavior changes.
- `SPEC.md` uses the project’s compact section format; keep invariants testable.
- Build cleans `dist` before compiling so deleted source cannot enter a package.

## Checks

Run from the repository root:

```sh
npm ci
npm test
npm run check
npm run build
npm pack --dry-run
git diff --check
```

Do not commit or push unless explicitly asked.
