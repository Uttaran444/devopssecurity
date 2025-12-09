
// src/server.ts
import express from 'express'
import { z } from 'zod'

// MCP SDK imports (TypeScript SDK)
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js'

// Create MCP server with basic metadata
const mcp = new McpServer({ name: 'mcp-ts-demo', version: '1.0.0' })

// Example tool: add two numbers
mcp.registerTool(
  'add',
  {
    title: 'Add numbers',
    description: 'Adds two numbers',
    inputSchema: { a: z.number(), b: z.number() },
    outputSchema: { result: z.number() },
  },
  async ({ a, b }) => {
    const result = a + b
    return {
      content: [{ type: 'text', text: JSON.stringify({ result }) }],
      structuredContent: { result },
    }
  }
)

// Express app hosting the MCP endpoint with Streamable HTTP
const app = express()
app.use(express.json())

// Enforce auth in a middleware (filled in step 2)
app.use('/mcp', require('./verifyToken').verifyBearer)

app.post('/mcp', async (req, res) => {
  const transport = new StreamableHTTPServerTransport({ enableJsonResponse: true })
  res.on('close', () => transport.close())
  await mcp.connect(transport)
  await transport.handleRequest(req, res, req.body)
})

// Optional: SSE/GET and DELETE handlers if you add session support later

const port = parseInt(process.env.PORT || '3000', 10)
app
