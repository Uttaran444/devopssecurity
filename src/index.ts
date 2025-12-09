// server.ts
import express from 'express';
import axios, { AxiosError } from 'axios';
import dotenv from 'dotenv';
import { ConfidentialClientApplication } from '@azure/msal-node';
// MCP SDK imports (ESM subpaths must include .js)
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';

dotenv.config();

// ---- Config (use env — do NOT hard-code secrets) ----
const TENANT_ID = process.env.TENANT_ID ?? '';
const MCP_API_CLIENT_ID = process.env.MCP_API_CLIENT_ID ?? '';
const MCP_API_CLIENT_SECRET = process.env.MCP_API_CLIENT_SECRET ?? '';
const ADO_ORG = process.env.ADO_ORG ?? '';
const ADO_API_VERSION = process.env.ADO_API_VERSION ?? '7.2';
const PORT = Number(process.env.PORT ?? 3000);

const LOG_LEVEL = (process.env.LOG_LEVEL ?? 'debug').toLowerCase(); // debug|info|warn|error
const redact = (v: string | undefined) => (v ? `${v.substring(0, 6)}…REDACTED` : 'undefined');

// ---- Basic logger ----
function log(level: 'debug'|'info'|'warn'|'error', message: string, meta: Record<string, any> = {}) {
  const levelsOrder = { debug: 10, info: 20, warn: 30, error: 40 } as const;
  if (levelsOrder[level] < levelsOrder[LOG_LEVEL as keyof typeof levelsOrder]) return;
  const payload = {
    ts: new Date().toISOString(),
    level,
    msg: message,
    ...meta,
  };
  console.log(JSON.stringify(payload));
}

// ---- Correlation ID middleware ----
function correlationIdMiddleware(req: express.Request, _res: express.Response, next: express.NextFunction) {
  const incoming = (req.headers['x-correlation-id'] || req.headers['x-request-id'] || '') as string;
  const cid = incoming && typeof incoming === 'string' ? incoming : `cid_${Math.random().toString(36).slice(2, 10)}`;
  (req as any).__cid = cid;
  next();
}

// ---- Safe header preview (never logs token value) ----
function previewAuthHeader(headers: Record<string, any>) {
  const raw = (headers?.authorization || headers?.Authorization) as string | undefined;
  const hasBearer = !!raw && /^Bearer\s+/i.test(raw);
  const prefix = hasBearer ? raw!.split(/\s+/)[0] : undefined; // logs "Bearer" or undefined
  return { hasAuthorization: !!raw, hasBearer, prefix };
}

// ---- MSAL (OBO) client ----
const cca = new ConfidentialClientApplication({
  auth: {
    clientId: MCP_API_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    clientSecret: MCP_API_CLIENT_SECRET,
  },
});

const ADO_SCOPE = '499b84ac-1321-427f-aa17-267ca6975798/.default';

async function getAdoAccessToken(onBehalfOfToken: string, cid: string): Promise<string> {
  const start = Date.now();
  log('debug', 'OBO start', { cid });
  try {
    const result = await cca.acquireTokenOnBehalfOf({ oboAssertion: onBehalfOfToken, scopes: [ADO_SCOPE] });
    if (!result?.accessToken) {
      log('error', 'OBO returned no access token', { cid });
      throw new Error('OBO failed: no access token');
    }
    const durMs = Date.now() - start;
    log('info', 'OBO success', { cid, ms: durMs, expiresOn: result.expiresOn?.toISOString?.() });
    return result.accessToken;
  } catch (err: any) {
    const durMs = Date.now() - start;
    const msalCode = err?.errorCode || err?.code || 'unknown';
    log('error', 'OBO error', { cid, ms: durMs, errorCode: msalCode, message: err?.message });
    throw err;
  }
}

// ---- Axios interceptors (timing + error body) ----
axios.interceptors.request.use((config) => { (config as any).__start = Date.now(); return config; });
axios.interceptors.response.use(
  (resp) => {
    const start = (resp.config as any).__start ?? Date.now();
    const ms = Date.now() - start;
    log('info', 'HTTP success', { cid: (resp.config.headers as any)?.['x-correlation-id'], method: resp.config.method, url: resp.config.url, status: resp.status, ms });
    return resp;
  },
  (error: AxiosError) => {
    const config = error.config || {};
    const start = (config as any).__start ?? Date.now();
    const ms = Date.now() - start;
    const data = (error.response?.data ?? {}) as any;
    const condensed = (() => {
      try {
        const s = typeof data === 'string' ? data : JSON.stringify(data);
        return s.substring(0, 300);
      } catch {
        return '';
      }
    })();
    log('error', 'HTTP error', { cid: ((config as any).headers)?.['x-correlation-id'], method: (config as any).method, url: (config as any).url, status: error.response?.status, ms, error: error.message, bodyPreview: condensed });
    return Promise.reject(error);
  }
);

// ---- MCP server ----
const mcpServer = new McpServer({ name: 'ado-mcp-server', version: '1.0.0' });

mcpServer.registerTool(
  'ado_list_projects',
  { title: 'List Azure DevOps projects', description: 'Returns projects from dev.azure.com/{organization}' },
  async (...cbArgs: any[]) => {
    // Support both callback signatures:
    // - (extra) when no inputSchema
    // - (args, extra) when inputSchema is provided
    const extra = cbArgs.length === 1 ? cbArgs[0] : cbArgs[1];
    const cid = (extra as any)?.cid || 'no-cid';
    const hasAuthInfo = !!extra?.authInfo;
    log('debug', 'Tool ado_list_projects invoked', { cid, org: ADO_ORG, apiVersion: ADO_API_VERSION, authInfoPresent: hasAuthInfo });

    const userToken = extra?.authInfo?.token ?? '';
    if (!userToken) {
      log('warn', 'Missing bearer token in tool context', { cid });
      return { content: [{ type: 'text', text: 'Missing bearer token (Copilot OAuth not forwarded).' }], isError: true };
    }

    let adoToken: string;
    try { adoToken = await getAdoAccessToken(userToken, cid); }
    catch (e: any) { return { content: [{ type: 'text', text: `Authentication/OBO error: ${e?.message ?? 'unknown'}` }], isError: true }; }

    const url = `https://dev.azure.com/${ADO_ORG}/_apis/projects?api-version=${ADO_API_VERSION}`;
    try {
      const resp = await axios.get(url, { headers: { Authorization: `Bearer ${adoToken}`, 'x-correlation-id': cid } });
      const projects = (resp.data?.value ?? []).map((p: any) => ({ id: p.id, name: p.name }));
      log('info', 'ADO projects retrieved', { cid, count: projects.length });
      return { content: [{ type: 'text', text: JSON.stringify({ projects }) }], structuredContent: { projects } };
    } catch (err: any) {
      log('error', 'ADO projects API call failed', { cid, message: err?.message });
      return { content: [{ type: 'text', text: `Azure DevOps API error: ${err?.message ?? 'unknown'}` }], isError: true };
    }
  }
);

// ---- Express + transport wiring ----
const app = express();
app.use(express.json());
app.use(correlationIdMiddleware);

app.get('/healthz', (req, res) => {
  const cid = (req as any).__cid; log('info', 'Health check', { cid }); res.status(200).send('ok');
});

app.post('/mcp', async (req, res) => {
  const cid = (req as any).__cid;
  const preview = previewAuthHeader(req.headers);
  log('debug', 'Incoming MCP request', { cid, method: req.method, path: req.path, hasAuthHeader: preview.hasAuthorization, hasBearerPrefix: preview.hasBearer });

  const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined, enableJsonResponse: true });

  const authHeader = (req.headers.authorization ?? '') as string;
  const userToken = authHeader.replace(/^Bearer\s+/i, '');

  const reqWithAuth = req as any;
  reqWithAuth.auth = userToken ? { token: userToken, clientId: MCP_API_CLIENT_ID, scopes: [] } : undefined;

  transport.setServerContext({ headers: req.headers, cid });
  res.on('close', () => { log('debug', 'Response closed', { cid }); transport.close(); });

  try {
    await mcpServer.connect(transport);
    (transport as any).__cid = cid;
    await transport.handleRequest(reqWithAuth, res, req.body);
    log('info', 'MCP request handled', { cid });
  } catch (err: any) {
    log('error', 'Transport handleRequest failed', { cid, message: err?.message });
    try { res.status(500).json({ error: 'MCP server transport error', details: err?.message }); } catch {}
  }
});

function validateEnv() {
  const errors: string[] = [];
  if (!TENANT_ID) errors.push('TENANT_ID missing');
  if (!MCP_API_CLIENT_ID) errors.push('MCP_API_CLIENT_ID missing');
  if (!MCP_API_CLIENT_SECRET) errors.push('MCP_API_CLIENT_SECRET missing');
  if (!ADO_ORG) errors.push('ADO_ORG missing');
  if (errors.length) {
    log('warn', 'Startup env validation failed', { errors });
  } else {
    log('info', 'Startup env validation OK', { tenantIdPreview: TENANT_ID.slice(0, 6)+'…', clientIdPreview: MCP_API_CLIENT_ID.slice(0, 6)+'…', adoOrg: ADO_ORG, adoApiVersion: ADO_API_VERSION });
  }
}

validateEnv();

app.listen(PORT, () => { log('info', 'Server listening', { port: PORT, path: '/mcp' }); });


