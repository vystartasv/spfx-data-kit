export type ErrorKind = "validation" | "not-found" | "conflict" | "permission" | "transient" | "throttled" | "unknown";

export interface RequestHeaders { readonly [name: string]: string; }
export interface TransportRequestOptions { readonly method: string; readonly headers?: RequestHeaders; readonly body?: string; }
export interface TransportResponse { readonly status: number; readonly headers?: RequestHeaders; text(): Promise<string>; }

/** Host boundary: SPFx, fetch, tests, or another authenticated runtime can implement it. */
export interface RequestTransport { request(url: string, options: TransportRequestOptions): Promise<TransportResponse>; }

export interface DataRequest {
  readonly url: string;
  readonly method?: string;
  readonly headers?: RequestHeaders;
  readonly body?: string;
  /** Header names that participate in GET cache identity. Omitted means all supplied headers. */
  readonly cacheKeyHeaders?: readonly string[];
}

export interface DataResult<T> { data: T; etag?: string; }
export interface ListResult<T> { data: T[]; etags: Readonly<Record<string, string>>; }

export type ListQuery = {
  select?: readonly string[];
  expand?: readonly string[];
  filter?: string;
  orderBy?: string | readonly [string, boolean][];
  top?: number;
  maxPages?: number;
  pageSize?: number;
};

export type UpdateOptions = { etag?: string };
export type RemoveOptions = { etag?: string };

export interface CrudResource<TEntity, TCreate, TUpdate> {
  list(query?: ListQuery): Promise<ListResult<TEntity>>;
  get(id: number): Promise<DataResult<TEntity>>;
  create(input: TCreate): Promise<DataResult<TEntity>>;
  update(id: number, input: TUpdate, options?: UpdateOptions): Promise<DataResult<TEntity>>;
  remove(id: number, options?: RemoveOptions): Promise<DataResult<void>>;
}
