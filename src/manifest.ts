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
    "graph-page-delta-iteration",
    "abort-signal-timeout-forwarding",
    "structured-error-details",
    "etag-if-match",
    "binary-uint8array-transport",
    "sharepoint-files-folders-attachments",
  ],
  status: {
    sharepointFilesFoldersAttachments: "partial",
    sharepointSmallUpload: "present",
    sharepointLargeResumableUpload: "planned",
    sharepointDrivesProvisioningPagesNavigationProfilesSearchTaxonomy: "planned",
    fluentPnPCompatibility: "excluded",
  },
  limits: { graphJsonBatchMaxRequests: 20, sharepointSmallUploadMaxBytes: 1_500_000, cacheRequiresExplicitBounds: true },
  boundaries: ["authentication", "authorization", "tenant-discovery", "offline-storage", "persistent-storage", "performance-benchmarks"],
} as const;

export const CAPABILITY_MANIFEST = capabilityManifest;
