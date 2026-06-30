import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getGraphClient } from "../graph/client.js";
import { handleGraphError } from "../utils/error-handler.js";

export function registerContactTools(server: McpServer): void {
  server.registerTool(
    "list_contacts",
    {
      description:
        "List personal Contacts from the authenticated user's Outlook address book. Optionally filter by name or email.",
      inputSchema: {
        search: z.string().optional().describe("Search string to filter contacts by name or email"),
        top: z.number().optional().describe("Maximum number of contacts to return (default: 50, max: 999)"),
        skip: z.number().optional().describe("Number of contacts to skip (for pagination)"),
      },
    },
    async ({ search, top, skip }) => {
      try {
        const client = await getGraphClient();

        const queryParams: string[] = ["$select=id,displayName,emailAddresses,companyName,jobTitle"];

        if (top !== undefined) queryParams.push(`$top=${top}`);
        else queryParams.push("$top=50");

        if (skip !== undefined) queryParams.push(`$skip=${skip}`);

        let endpoint = `/me/contacts?${queryParams.join("&")}`;

        if (search) {
          endpoint += `&$filter=contains(displayName,'${encodeURIComponent(search)}') or emailAddresses/any(e:contains(e/address,'${encodeURIComponent(search)}'))`;
        }

        const contacts = await client.api(endpoint).get();

        return {
          content: [{ type: "text", text: JSON.stringify(contacts, null, 2) }],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: handleGraphError(error) }],
        };
      }
    },
  );

  server.registerTool(
    "list_contact_folders",
    {
      description:
        "List the user's contact folders (contact lists) from Outlook, such as 'My ThailandPostMart - Dev'. To see contacts inside a folder, use list_contacts with the folderId.",
      inputSchema: {
        folderId: z.string().optional().describe("If provided, list contacts inside this folder instead of listing folders"),
        top: z.number().optional().describe("Maximum number of items to return (default: 50)"),
        skip: z.number().optional().describe("Number of items to skip (for pagination)"),
      },
    },
    async ({ folderId, top, skip }) => {
      try {
        const client = await getGraphClient();

        const queryParams: string[] = [];
        if (top !== undefined) queryParams.push(`$top=${top}`);
        else queryParams.push("$top=50");
        if (skip !== undefined) queryParams.push(`$skip=${skip}`);

        let endpoint: string;
        if (folderId) {
          endpoint = `/me/contactFolders/${folderId}/contacts?$select=id,displayName,emailAddresses,companyName,jobTitle&${queryParams.join("&")}`;
        } else {
          endpoint = `/me/contactFolders?$select=id,displayName,parentFolderId&${queryParams.join("&")}`;
        }

        const result = await client.api(endpoint).get();

        return {
          content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: handleGraphError(error) }],
        };
      }
    },
  );
}
