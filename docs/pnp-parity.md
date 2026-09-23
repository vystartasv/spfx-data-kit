# PnPjs capability matrix

Reference families: [@pnp/sp](https://github.com/pnp/pnpjs/tree/version-4/packages/sp) and [@pnp/graph](https://github.com/pnp/pnpjs/tree/version-4/packages/graph), with their [SharePoint docs](https://pnp.github.io/pnpjs/sp/) and [Graph docs](https://pnp.github.io/pnpjs/graph/).

This is a capability-family comparison, not a claim of full parity.

| Capability family | @pnp/sp | @pnp/graph | Kit status | Scope |
| --- | --- | --- | --- | --- |
| Injected request transport | host/runtime integrations | host/runtime integrations | host-owned | Authentication, authorization, discovery, and network remain in `RequestTransport`. |
| Binary request/response transport | host/runtime integrations | host/runtime integrations | present | `Uint8Array` request bodies and `DataClient.requestBytes`; hosts provide optional `TransportResponse.arrayBuffer()`. |
| SharePoint list reads, queries, paging, and ETags | broad fluent surface | — | partial | Typed list-item REST CRUD, query construction, bounded paging, and item ETags. |
| SharePoint writes and batch | broad CRUD and batching | — | partial | Typed create/update/delete plus GET-only REST batch parsing. |
| SharePoint files, folders, and attachments | files/folders/attachments | — | partial | Site-scoped typed metadata, expanded folder children, binary download, list-item attachment operations, and bounded small uploads ≤1,500,000 bytes. |
| SharePoint large/resumable uploads | large-file APIs | — | planned | Not implemented or claimed in this track. |
| Graph requests | — | broad fluent surface | present | Generic Graph JSON requests, ETags, page/delta parsing, and bounded async iteration. |
| Graph JSON batch | — | batch support | partial | JSON batch result parsing with the service limit of 20 children. |
| Drives, sites, groups, admin, provisioning, pages, navigation, profiles, search, taxonomy, and other service families | present across selected modules | present across selected modules | planned | Add only as independently specified typed REST surfaces. |
| Fluent wrappers and PnPjs behaviors | present | present | excluded | No PnPjs dependency or fluent compatibility layer. |
| Authentication, permissions, and tenant URL discovery | host/configuration dependent | host/configuration dependent | host-owned | The host supplies authenticated URLs and transport behavior. |
