import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getGraphClient } from "../graph/client.js";
import { handleGraphError } from "../utils/error-handler.js";

export function registerUserTools(server: McpServer): void {
  server.registerTool(
    "list_users",
    {
      description:
        "Search or list Azure AD users in the organisation. Use 'search' to find users by name or email. Without 'search', returns all users paginated. Useful for finding recipient email addresses.",
      inputSchema: {
        search: z.string().optional().describe("Search string to filter users by display name or email"),
        top: z.number().optional().describe("Maximum number of users to return (default: 50, max: 999)"),
        skip: z.number().optional().describe("Number of users to skip (for pagination)"),
        select: z.string().optional().describe("Comma-separated properties to return (default: id,displayName,mail,userPrincipalName)"),
      },
    },
    async ({ search, top, skip, select }) => {
      try {
        const client = await getGraphClient();

        const selectedFields = select || "id,displayName,mail,userPrincipalName,jobTitle,department";
        const queryParams: string[] = [`$select=${encodeURIComponent(selectedFields)}`];

        if (top !== undefined) queryParams.push(`$top=${top}`);
        else queryParams.push("$top=50");

        if (skip !== undefined) queryParams.push(`$skip=${skip}`);

        let request = client.api(`/users?${queryParams.join("&")}`);

        if (search) {
          request = request
            .header("ConsistencyLevel", "eventual")
            .search(`"displayName:${search} OR mail:${search} OR userPrincipalName:${search}"`);
        }

        const users = await request.get();

        return {
          content: [{ type: "text", text: JSON.stringify(users, null, 2) }],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: handleGraphError(error) }],
        };
      }
    },
  );
}
