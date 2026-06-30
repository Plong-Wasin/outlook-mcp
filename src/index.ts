import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";

import { registerAuthTools } from "./tools/auth-tools.js";
import { registerEmailTools } from "./tools/email-tools.js";
import { registerGroupTools } from "./tools/group-tools.js";
import { registerContactTools } from "./tools/contact-tools.js";
import { registerUserTools } from "./tools/user-tools.js";
import { registerMailTools } from "./tools/mail-tools.js";

const server = new McpServer({
  name: "outlook-mcp",
  version: "1.0.0",
});

registerAuthTools(server);
registerEmailTools(server);
registerGroupTools(server);
registerContactTools(server);
registerUserTools(server);
registerMailTools(server);

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Outlook MCP server running on stdio");
}

main().catch((error) => {
  console.error("Fatal error in main():", error);
  process.exit(1);
});
