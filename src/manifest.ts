export const capabilityManifest = {
  schemaVersion: 1,
  package: "spfx-data-kit",
  capabilities: [
    "injected-request-transport",
    "get-in-flight-deduplication",
    "bounded-ttl-memory-cache",
    "cache-invalidation",
    "request-diagnostics",
    "bounded-get-retries",
    "sharepoint-rest-list-adapter",
    "sharepoint-rest-batch-results",
    "graph-json-batch-results",
    "etag-if-match",
  ],
  limits: { graphJsonBatchMaxRequests: 20, cacheRequiresExplicitBounds: true },
  boundaries: ["authentication", "authorization", "tenant-discovery", "offline-storage", "persistent-storage", "performance-benchmarks"],
} as const;

export const CAPABILITY_MANIFEST = capabilityManifest;
