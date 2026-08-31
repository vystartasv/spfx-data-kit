export type ErrorKind = "not-found" | "conflict" | "validation" | "permission" | "transient" | "unknown";
export class DataError extends Error { readonly name = "DataError"; constructor(public readonly kind: ErrorKind, message: string, public readonly cause?: unknown) { super(message); } }
/** A value returned by a resource, with the item's current concurrency token when available. */
export type DataResult<T> = { data: T; etag?: string };
/** A collection result keeps entities ergonomic and exposes tokens separately by item id. */
export type ListResult<T> = { data: T[]; etags: Readonly<Record<string, string>> };
export type ListQuery = { select?: readonly string[]; expand?: readonly string[]; filter?: string; orderBy?: string | readonly [string, boolean][]; top?: number; maxPages?: number; pageSize?: number };
export type UpdateOptions = { etag?: string };
export type RemoveOptions = { etag?: string };
export interface CrudResource<TEntity, TCreate, TUpdate> { list(query?: ListQuery): Promise<ListResult<TEntity>>; get(id: number): Promise<DataResult<TEntity>>; create(input: TCreate): Promise<DataResult<TEntity>>; update(id: number, input: TUpdate, options?: UpdateOptions): Promise<DataResult<TEntity>>; remove(id: number, options?: RemoveOptions): Promise<DataResult<void>>; }
