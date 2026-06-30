import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { getGraphClient } from "../graph/client.js";
import { handleGraphError } from "../utils/error-handler.js";

function buildRecipients(addresses: string[] | undefined) {
  if (!addresses || addresses.length === 0) return undefined;
  return addresses.map((addr) => ({ emailAddress: { address: addr } }));
}

export function registerMailTools(server: McpServer): void {
  // ── list_mail_folders ────────────────────────────────────────────────────
  server.registerTool(
    "list_mail_folders",
    {
      description:
        "List all mail folders in the authenticated user's mailbox, including custom folders. Returns folder IDs needed for list_messages with non-standard folders.",
      inputSchema: {
        top: z.number().optional().describe("Maximum number of folders to return (default: 50)"),
        skip: z.number().optional().describe("Number of folders to skip (for pagination)"),
      },
    },
    async ({ top, skip }) => {
      try {
        const client = await getGraphClient();

        const queryParams: string[] = [
          "$select=id,displayName,unreadItemCount,totalItemCount",
        ];
        queryParams.push(`$top=${top ?? 50}`);
        if (skip !== undefined) queryParams.push(`$skip=${skip}`);

        const folders = await client.api(`/me/mailFolders?${queryParams.join("&")}`).get();

        return {
          content: [{ type: "text", text: JSON.stringify(folders, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleGraphError(error) }] };
      }
    },
  );

  // ── list_messages ────────────────────────────────────────────────────────
  server.registerTool(
    "list_messages",
    {
      description:
        "List Messages from a mail folder. Use folder 'inbox', 'drafts', 'sentitems', or 'deleteditems' (well-known names), or pass a folderId from list_mail_folders. Supports full-text search and OData filter — but NOT both at once.",
      inputSchema: {
        folder: z
          .string()
          .optional()
          .describe(
            "Folder name or ID to list from. Well-known names: 'inbox', 'drafts', 'sentitems', 'deleteditems'. Default: 'inbox'",
          ),
        search: z
          .string()
          .optional()
          .describe("Full-text keyword search (KQL). Cannot be combined with filter."),
        filter: z
          .string()
          .optional()
          .describe(
            "OData filter expression (e.g. 'isRead eq false'). Cannot be combined with search.",
          ),
        top: z.number().optional().describe("Maximum number of messages to return (default: 50)"),
        skip: z.number().optional().describe("Number of messages to skip (for pagination)"),
      },
    },
    async ({ folder, search, filter, top, skip }) => {
      try {
        if (search && filter) {
          return {
            content: [
              {
                type: "text",
                text: "❌ Cannot use 'search' and 'filter' together — the Microsoft Graph API does not support combining $search and $filter in a single request. Use one or the other.",
              },
            ],
          };
        }

        const client = await getGraphClient();
        const folderSegment = folder ?? "inbox";

        const queryParams: string[] = [
          "$select=id,subject,from,receivedDateTime,isRead,hasAttachments,bodyPreview",
        ];
        queryParams.push(`$top=${top ?? 50}`);
        if (skip !== undefined) queryParams.push(`$skip=${skip}`);
        if (filter) queryParams.push(`$filter=${encodeURIComponent(filter)}`);

        let request = client.api(
          `/me/mailFolders/${folderSegment}/messages?${queryParams.join("&")}`,
        );

        if (search) {
          request = request.search(`"${search}"`);
        }

        const messages = await request.get();

        return {
          content: [{ type: "text", text: JSON.stringify(messages, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleGraphError(error) }] };
      }
    },
  );

  // ── get_message ──────────────────────────────────────────────────────────
  server.registerTool(
    "get_message",
    {
      description:
        "Read the full content of a single Message by ID. Returns body (plain text or HTML), all recipients, and attachment metadata (name, size, type — not file content).",
      inputSchema: {
        messageId: z.string().describe("The Message ID (from list_messages)"),
        bodyType: z
          .enum(["text", "html"])
          .optional()
          .describe("Body format to return: 'text' (default) or 'html'"),
      },
    },
    async ({ messageId, bodyType }) => {
      try {
        const client = await getGraphClient();

        const selectFields = [
          "id",
          "subject",
          "from",
          "toRecipients",
          "ccRecipients",
          "receivedDateTime",
          "isRead",
          "body",
          "attachments",
        ].join(",");

        let request = client
          .api(`/me/messages/${messageId}?$select=${selectFields}&$expand=attachments($select=id,name,size,contentType)`)

        if ((bodyType ?? "text") === "text") {
          request = request.header("Prefer", 'outlook.body-content-type="text"');
        }

        const message = await request.get();

        return {
          content: [{ type: "text", text: JSON.stringify(message, null, 2) }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleGraphError(error) }] };
      }
    },
  );

  // ── reply_message ────────────────────────────────────────────────────────
  server.registerTool(
    "reply_message",
    {
      description:
        "Reply to a Message. Thread history is preserved automatically. The reply is sent immediately.",
      inputSchema: {
        messageId: z.string().describe("The Message ID to reply to"),
        comment: z.string().describe("The reply body text"),
      },
    },
    async ({ messageId, comment }) => {
      try {
        const client = await getGraphClient();

        await client.api(`/me/messages/${messageId}/reply`).post({ comment });

        return {
          content: [{ type: "text", text: `✅ Reply sent successfully.` }],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleGraphError(error) }] };
      }
    },
  );

  // ── forward_message ──────────────────────────────────────────────────────
  server.registerTool(
    "forward_message",
    {
      description:
        "Forward a Message to one or more new Recipients. Optionally include a comment prepended to the forwarded message.",
      inputSchema: {
        messageId: z.string().describe("The Message ID to forward"),
        to: z.array(z.string()).describe("Recipient email addresses to forward to"),
        comment: z.string().optional().describe("Optional note to prepend to the forwarded message"),
      },
    },
    async ({ messageId, to, comment }) => {
      try {
        const client = await getGraphClient();

        const payload: any = {
          toRecipients: buildRecipients(to),
        };
        if (comment) payload.comment = comment;

        await client.api(`/me/messages/${messageId}/forward`).post(payload);

        return {
          content: [
            {
              type: "text",
              text: `✅ Message forwarded successfully to ${to.join(", ")}`,
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleGraphError(error) }] };
      }
    },
  );

  // ── mark_message ─────────────────────────────────────────────────────────
  server.registerTool(
    "mark_message",
    {
      description: "Mark a Message as read or unread.",
      inputSchema: {
        messageId: z.string().describe("The Message ID to mark"),
        isRead: z.boolean().describe("true to mark as read, false to mark as unread"),
      },
    },
    async ({ messageId, isRead }) => {
      try {
        const client = await getGraphClient();

        await client.api(`/me/messages/${messageId}`).patch({ isRead });

        return {
          content: [
            {
              type: "text",
              text: `✅ Message marked as ${isRead ? "read" : "unread"}.`,
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleGraphError(error) }] };
      }
    },
  );

  // ── delete_message ───────────────────────────────────────────────────────
  server.registerTool(
    "delete_message",
    {
      description:
        "Delete a Message or Draft by moving it to the Deleted Items folder. The message is NOT permanently erased — it can be recovered from Deleted Items.",
      inputSchema: {
        messageId: z.string().describe("The Message ID to delete"),
      },
    },
    async ({ messageId }) => {
      try {
        const client = await getGraphClient();

        await client.api(`/me/messages/${messageId}`).delete();

        return {
          content: [
            {
              type: "text",
              text: `✅ Message moved to Deleted Items (not permanently deleted — recoverable from Deleted Items folder).`,
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleGraphError(error) }] };
      }
    },
  );
}
