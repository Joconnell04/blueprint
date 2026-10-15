import { StdioServerTransport } from "@modelcontextprotocol/server/stdio";
import { createServer } from "./server.js";

/** Default: stdio (what Claude Desktop launches). `--http [port]` serves Streamable HTTP for local testing with basic-host. */
async function main() {
  const httpIdx = process.argv.indexOf("--http");
  if (httpIdx === -1) {
    await createServer().connect(new StdioServerTransport());
    return;
  }
  const port = parseInt(process.argv[httpIdx + 1] ?? process.env.PORT ?? "3001", 10);
  const { createMcpExpressApp } = await import("@modelcontextprotocol/express");
  const { NodeStreamableHTTPServerTransport } = await import("@modelcontextprotocol/node");
  const cors = (await import("cors")).default;
  const app = createMcpExpressApp({ host: "0.0.0.0" });
  app.use(cors());
  app.all("/mcp", async (req: any, res: any) => {
    const server = createServer();
    const transport = new NodeStreamableHTTPServerTransport({ sessionIdGenerator: undefined });
    res.on("close", () => { transport.close().catch(() => {}); server.close().catch(() => {}); });
    try { await server.connect(transport); await transport.handleRequest(req, res, req.body); }
    catch (e) { console.error(e); if (!res.headersSent) res.status(500).json({ jsonrpc: "2.0", error: { code: -32603, message: "Internal server error" }, id: null }); }
  });
  app.listen(port, () => console.error(`Blueprint MCP (http) on http://localhost:${port}/mcp`));
}

main().catch((e) => { console.error(e); process.exit(1); });
