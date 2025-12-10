
// src/server.ts
import express from 'express'
import { z } from 'zod'

// MCP SDK imports (TypeScript SDK)
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js'
import { randomUUID } from 'crypto'

// Create MCP server with basic metadata
const mcp = new McpServer({ name: 'mcp-ts-demo', version: '1.0.0' })

// Tool: magicaddition -> (a + b) * 10
mcp.registerTool(
  'magicaddition',
  {
    title: 'Magic Addition',
    description: 'Adds two numbers then multiplies the sum by 10',
    inputSchema: { a: z.number(), b: z.number() },
    outputSchema: { a: z.number(), b: z.number(), sum: z.number(), result: z.number() },
  },
  async ({ a, b }) => {
    const sum = a + b
    const result = sum * 10
    const payload = { a, b, sum, result }
    return {
      content: [{ type: 'text', text: JSON.stringify(payload) }],
      structuredContent: payload,
    }
  }
)

// Express app hosting the MCP endpoint with Streamable HTTP
const app = express()
app.use(express.json())

// Health check
app.get('/healthz', (_req, res) => res.status(200).send('ok'))

app.post('/mcp', async (req, res) => {
  const transport = new StreamableHTTPServerTransport({ enableJsonResponse: true, sessionIdGenerator: () => randomUUID() })
  res.on('close', () => transport.close())
  await mcp.connect(transport)
  
  await transport.handleRequest(req, res, req.body)
})

const port = parseInt(process.env.PORT || '3000', 10)
app.listen(port, () => {
  console.log(`[mcp-ado-server] listening on port ${port}`)
})
