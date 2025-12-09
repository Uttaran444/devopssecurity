
import express from 'express';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import axios from 'axios';
import { ConfidentialClientApplication } from '@azure/msal-node';

// --- Config ---
const TENANT_ID = process.env.TENANT_ID!;
const MCP_API_CLIENT_ID = process.env.MCP_API_CLIENT_ID!;
const MCP_API_CLIENT_SECRET = process.env.MCP_API_CLIENT_SECRET!;
const ADO_ORG = process.env.ADO_ORG!;              // e.g., "yourorg"
const ADO_API_VERSION = process.env.ADO_API_VERSION ?? '7.2';

// --- MSAL (OBO) ---
const cca = new ConfidentialClientApplication({
  auth: {
    clientId: MCP_API_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    clientSecret: MCP_API_CLIENT_SECRET
  }
});

// Azure DevOps resource GUID scope
const ADO_SCOPE = '499b84ac-1321-427f-aa17-267ca6975798/.default';

async function getAdoAccessToken(onBehalfOfToken: string): Promise<string> {
  const result = await cca.acquireTokenOnBehalfOf({
    oboAssertion: onBehalfOfToken,
    scopes: [ADO_SCOPE]
  });
  if (!result?.accessToken) throw new Error('OBO failed');
  return result.accessToken;
}

// --- MCP Server ---
const mcpServer = new McpServer({
  name: 'ado-mcp-server',
  version: '1.0.0'
});

// Tool: list projects (no input schema; return text JSON)
mcpServer.registerTool(
  'ado_list_projects',
  {
    title: 'List Azure DevOps projects',
    description: 'Returns projects from dev.azure.com/{organization}'
  },
  async (_input, extra) => {
    const userToken = extra.authInfo?.token ?? '';
    if (!userToken) {
      return {
        content: [{ type: 'text', text: 'Missing bearer token' }],
        isError: true
      };
    }

    const adoToken = await getAdoAccessToken(userToken);
    const url = `https://dev.azure.com/${ADO_ORG}/_apis/projects?api-version=${ADO_API_VERSION}`;
    const resp = await axios.get(url, {
      headers: { Authorization: `Bearer ${adoToken}` }
    });

    const projects = (resp.data?.value ?? []).map((p: any) => ({ id: p.id, name: p.name }));
    return {
      content: [{ type: 'text', text: JSON.stringify({ projects }) }],
      structuredContent: { projects }
    };
  }
);

// --- Transport wiring (HTTP, streamable) ---
const app = express();
app.use(express.json());

app.post('/mcp', async (req, res) => {
  const transport = new StreamableHTTPServerTransport({
    sessionIdGenerator: undefined, // stateless mode
    enableJsonResponse: true
  });

  // Extract bearer token and attach as auth info for tool handlers
  const authHeader = (req.headers.authorization ?? '') as string;
  const userToken = authHeader.replace(/^Bearer\s+/i, '');
  const reqWithAuth = req as any;
  reqWithAuth.auth = userToken
    ? { token: userToken, clientId: MCP_API_CLIENT_ID, scopes: [] }
    : undefined;

  res.on('close', () => transport.close());
  await mcpServer.connect(transport);
  await transport.handleRequest(reqWithAuth, res, req.body);
});

const port = Number(process.env.PORT ?? 3000);
app.listen(port, () => {
  console.log(`MCP server listening on http://localhost:${port}/mcp`);
});
``
