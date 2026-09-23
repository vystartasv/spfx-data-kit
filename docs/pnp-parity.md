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
| SharePoint search | search/query | — | partial | Typed POST search requests with query text, select properties, refiners, sort, row bounds, totals, and bounded paging. No ranking claim or fluent emulation. |
| SharePoint large/resumable uploads | large-file APIs | — | planned | Not implemented or claimed in this track. |
| Graph requests | — | broad fluent surface | present | Generic Graph JSON requests, ETags, page/delta parsing, and bounded async iteration. |
| Graph JSON batch | — | batch support | partial | JSON batch result parsing with the service limit of 20 children. |
| Graph drives and file helpers | drives/files/folders | drives/files/folders | partial | Typed drive/root/item metadata, bounded children paging, binary download, ETag-aware item update/delete, and simple `Uint8Array` uploads ≤1,500,000 bytes. |
| Graph sites, lists, and list items | sites/lists/items | sites/lists/items | partial | Typed site-by-id/path and list metadata, OData select/expand/filter/orderBy/top, bounded next-link paging, list-item CRUD, response ETags, and optional `If-Match` writes without follow-up reads. |
| Graph users, groups, and directory membership | users/groups/directory objects | users/groups/directory objects | partial | Typed current-user/user/group reads, bounded users/groups/member paging, select/filter/orderBy/top queries, and Graph `$ref` member add/remove without follow-up reads. |
| Graph permissions and sharing | permissions, sharing links, invitations | permissions, sharing links, invitations | partial | Typed drive-item/site permission list/get/delete, drive-item `createLink`/`invite`, sharing-link `permission/grant`, strict link/scope/recipient validation, ETags, and host-owned authorization. Site permission creation and permission updates remain outside this track. |
| Graph search | — | search/query | partial | Typed requests across one or more entity types with fields, sort properties, totals, and bounded `from`/`size` paging. No semantic search abstraction or ranking claim. |
| Graph resumable uploads, thumbnails, and previews | selected APIs | selected APIs | planned | Not implemented or claimed; simple upload ceiling is explicit. |
| Graph mail, calendar, Teams, Planner, To Do, admin, provisioning, pages, navigation, profiles, taxonomy, and other service families | present across selected modules | present across selected modules | planned | Not implemented or claimed in this track; add only as independently specified typed REST surfaces. |
| Fluent wrappers and PnPjs behaviors | present | present | excluded | No PnPjs dependency or fluent compatibility layer. |
| Authentication, permissions, and tenant URL discovery | host/configuration dependent | host/configuration dependent | host-owned | The host supplies authenticated URLs and transport behavior. |
