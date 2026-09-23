import type { RequestTransport, TransportRequestOptions, TransportResponse, ResponseHeaders } from "../src/index.js";

export type Reply = { status: number; body?: string; bytes?: Uint8Array; headers?: ResponseHeaders };

export class ScriptedTransport implements RequestTransport {
  readonly calls: { url: string; options: TransportRequestOptions }[] = [];
  constructor(private readonly replies: Reply[]) {}
  request(url: string, options: TransportRequestOptions): Promise<TransportResponse> {
    this.calls.push({ url, options });
    const reply = this.replies.shift();
    if (!reply) throw new Error(`Unexpected request: ${url}`);
    return Promise.resolve({
      status: reply.status,
      headers: reply.headers,
      text: async () => reply.body ?? "",
      ...(reply.bytes === undefined ? {} : { arrayBuffer: async () => reply.bytes!.slice().buffer as ArrayBuffer }),
    });
  }
}
