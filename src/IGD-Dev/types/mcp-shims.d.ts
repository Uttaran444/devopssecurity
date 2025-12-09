declare module "@modelcontextprotocol/sdk/server/mcp.js" {
  export class McpServer {
    constructor(options: any);
    registerTool(name: string, schema: any, handler: (...args: any[]) => any): void;
    connect(transport: any): Promise<void>;
  }
}

declare module "@modelcontextprotocol/sdk/server/streamableHttp.js" {
  export interface StreamableHTTPServerTransportOptions {
    sessionIdGenerator?: (() => string) | undefined;
    enableJsonResponse?: boolean;
    [key: string]: any;
  }
  export class StreamableHTTPServerTransport {
    constructor(options: StreamableHTTPServerTransportOptions);
    handleRequest(req: any, res: any, parsedBody?: unknown): Promise<void>;
    close(): Promise<void>;
    // Shim to satisfy legacy code usage
    setServerContext(ctx: any): void;
  }
}

declare module "@modelcontextprotocol/sdk/server/transport/http.js" {
  export interface StreamableHTTPServerTransportOptions {
    sessionIdGenerator?: (() => string) | undefined;
    enableJsonResponse?: boolean;
    [key: string]: any;
  }
  export class StreamableHTTPServerTransport {
    constructor(options: StreamableHTTPServerTransportOptions);
    handleRequest(req: any, res: any, parsedBody?: unknown): Promise<void>;
    close(): Promise<void>;
    setServerContext(ctx: any): void;
  }
}

declare module "@modelcontextprotocol/sdk/server/index.js" {
  export class Server {
    constructor(info: { name: string; version: string }, options?: any);
    tool(name: string, description: string, schema: any, handler: (...args: any[]) => any): void;
    connect(transport: any): Promise<void>;
  }
}