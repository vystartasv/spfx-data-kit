import { spfi, type SPFI } from "@pnp/sp";
import { type IItems } from "@pnp/sp/items/types.js";
import { type IWeb } from "@pnp/sp/webs/types.js";
import { type ILists, type IList } from "@pnp/sp/lists/types.js";
import { SPFx } from "@pnp/sp/behaviors/spfx.js";
import "@pnp/sp/webs/index.js"; import "@pnp/sp/lists/index.js"; import "@pnp/sp/lists/web.js"; import "@pnp/sp/items/index.js";
import "@pnp/sp/items/list.js";
import { CrudResource, DataError, DataResult, ListQuery, ListResult, RemoveOptions, UpdateOptions } from "./contracts.js";
import { mapSharePointError } from "./errors.js";
export type SharePointFieldMap<TEntity> = { [K in keyof TEntity]?: string };
export type PnPjsListOptions<TEntity, TCreate, TUpdate> = { title: string; map: SharePointFieldMap<TEntity>; createMap?: (input: TCreate) => Record<string, unknown>; updateMap?: (input: TUpdate) => Record<string, unknown> };
type Row = Record<string, unknown>;
// PnPjs v4 exposes collection/add response payloads weakly; this is the only boundary that narrows those responses.
const row = (value: unknown): Row => (typeof value === "object" && value !== null ? value as Row : {});
export class PnPjsListResource<TEntity, TCreate, TUpdate> implements CrudResource<TEntity, TCreate, TUpdate> {
  constructor(private readonly sp: SPFI, private readonly options: PnPjsListOptions<TEntity, TCreate, TUpdate>) {}
  private items(): IItems { const list = (this.sp as SPFI & { readonly web: IWeb & { readonly lists: ILists } }).web.lists.getByTitle(this.options.title) as IList & { readonly items: IItems }; return list.items; }
  private map(raw: Row): DataResult<TEntity> { const out = {} as TEntity; for (const key of Object.keys(this.options.map) as (keyof TEntity)[]) { const field = this.options.map[key]; if (field !== undefined) (out as Record<keyof TEntity, unknown>)[key] = raw[field]; } const etag = raw["@odata.etag"] ?? raw["odata.etag"] ?? raw.ETag; return typeof etag === "string" ? { data: out, etag } : { data: out }; }
  private itemId(raw: Row): number | undefined { const id = raw.Id ?? raw.ID; return typeof id === "number" && Number.isInteger(id) ? id : undefined; }
  private async one(id: number): Promise<DataResult<TEntity>> { try { return this.map(row(await this.items().getById(id)())); } catch (e) { throw mapSharePointError(e); } }
  async list(query?: ListQuery): Promise<ListResult<TEntity>> { try {
    let request: IItems = this.items(); if (query?.select?.length) request = request.select(...query.select); if (query?.expand?.length) request = request.expand(...query.expand); if (query?.filter) request = request.filter(query.filter);
    if (query?.orderBy) { const orders: readonly [string, boolean][] = typeof query.orderBy === "string" ? [[query.orderBy, true]] : query.orderBy; for (const [field, ascending] of orders) request = request.orderBy(field, ascending); }
    const requestedPageSize = query?.pageSize ?? 100; if (!Number.isInteger(requestedPageSize) || requestedPageSize < 1) throw new DataError("validation", "pageSize must be a positive integer"); if (query?.top !== undefined && (!Number.isInteger(query.top) || query.top < 0)) throw new DataError("validation", "top must be a non-negative integer"); if (query?.maxPages !== undefined && (!Number.isInteger(query.maxPages) || query.maxPages < 0)) throw new DataError("validation", "maxPages must be a non-negative integer");
    if (query?.top === 0 || query?.maxPages === 0) return { data: [], etags: {} };
    const pageSize = query?.top === undefined ? requestedPageSize : Math.min(query.top, requestedPageSize);
    const data: TEntity[] = []; const etags: Record<string, string> = {}; let pages = 0; for await (const page of request.top(pageSize)) { for (const value of page as unknown[]) { if (query?.top !== undefined && data.length >= query.top) break; const raw = row(value); const mapped = this.map(raw); data.push(mapped.data); const id = this.itemId(raw); if (id !== undefined && mapped.etag !== undefined) etags[String(id)] = mapped.etag; } pages++; if ((query?.maxPages !== undefined && pages >= query.maxPages) || (query?.top !== undefined && data.length >= query.top)) break; } return { data, etags };
  } catch (e) { if (e instanceof DataError) throw e; throw mapSharePointError(e); } }
  async get(id: number): Promise<DataResult<TEntity>> { return this.one(id); }
  async create(input: TCreate): Promise<DataResult<TEntity>> { try { const added = await this.items().add(this.options.createMap?.(input) ?? input as unknown as Record<string, unknown>); const result = added as { data?: Row } | undefined; const id = result?.data?.Id ?? result?.data?.ID; if (typeof id !== "number") throw new DataError("unknown", "SharePoint did not return the new item id"); return this.one(id); } catch (e) { if (e instanceof DataError) throw e; throw mapSharePointError(e); } }
  async update(id: number, input: TUpdate, options?: UpdateOptions): Promise<DataResult<TEntity>> { try { await this.items().getById(id).update(this.options.updateMap?.(input) ?? input as unknown as Record<string, unknown>, options?.etag); return this.one(id); } catch (e) { if (e instanceof DataError) throw e; throw mapSharePointError(e); } }
  async remove(id: number, options?: RemoveOptions): Promise<DataResult<void>> { try { await this.items().getById(id).delete(options?.etag); return { data: undefined }; } catch (e) { throw mapSharePointError(e); } }
}
export type MinimalSPFxContext = Parameters<typeof SPFx>[0];
export function createSpfi(context: MinimalSPFxContext): SPFI { return spfi().using(SPFx(context)); }
