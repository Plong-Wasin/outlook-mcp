import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getGraphClient } from "../graph/client.js";
import { handleGraphError } from "../utils/error-handler.js";

export function registerGroupTools(server: McpServer): void {
  server.registerTool(
    "list_group_members",
    {
      description:
        "List all members of a Microsoft 365 Group. Useful for finding email addresses of everyone in a group before sending.",
      inputSchema: {
        groupId: z.string().describe("The Microsoft 365 Group ID"),
        filter: z.string().optional().describe("OData filter expression (e.g. \"displayName eq 'Jane Doe'\")"),
        search: z.string().optional().describe("Search string for displayName (e.g. 'Jane')"),
        select: z.string().optional().describe("Comma-separated properties to return (e.g. 'id,displayName,mail')"),
        top: z.number().optional().describe("Maximum number of members to return (default: 100, max: 999)"),
      },
    },
    async ({ groupId, filter, search, select, top }) => {
      try {
        const client = await getGraphClient();
        let endpoint = `/groups/${groupId}/members`;

        const queryParams: string[] = [];

        if (filter) queryParams.push(`$filter=${encodeURIComponent(filter)}`);
        if (search) queryParams.push(`$search=${encodeURIComponent(`"${search}"`)}`);
        if (select) queryParams.push(`$select=${encodeURIComponent(select)}`);
        if (top !== undefined) queryParams.push(`$top=${top}`);

        if (queryParams.length > 0) endpoint += `?${queryParams.join("&")}`;

        const members = await client.api(endpoint).get();

        return {
          content: [{ type: "text", text: JSON.stringify(members, null, 2) }],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: handleGraphError(error) }],
        };
      }
    },
  );

  server.registerTool(
    "list_distribution_groups",
    {
      description:
        "List mail-enabled Distribution Groups in the organisation (classic distribution lists). Each group has a single email address that delivers to all its members.",
      inputSchema: {
        search: z.string().optional().describe("Search string to filter groups by display name"),
        top: z.number().optional().describe("Maximum number of groups to return (default: 50, max: 999)"),
        skip: z.number().optional().describe("Number of groups to skip (for pagination)"),
      },
    },
    async ({ search, top, skip }) => {
      try {
        const client = await getGraphClient();

        const queryParams: string[] = [
          "$filter=mailEnabled eq true and NOT groupTypes/any(c:c eq 'Unified')",
          "$select=id,displayName,mail,description",
        ];

        if (top !== undefined) queryParams.push(`$top=${top}`);
        else queryParams.push("$top=50");

        if (skip !== undefined) queryParams.push(`$skip=${skip}`);

        let request = client.api(`/groups?${queryParams.join("&")}`).header("ConsistencyLevel", "eventual");

        if (search) {
          request = request.search(`"${search}"`);
        }

        const groups = await request.get();

        return {
          content: [{ type: "text", text: JSON.stringify(groups, null, 2) }],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: handleGraphError(error) }],
        };
      }
    },
  );
}
