import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import fs from "fs";
import path from "path";
import { getGraphClient } from "../graph/client.js";
import { handleGraphError } from "../utils/error-handler.js";
import { findUnsupportedCss } from "../utils/html-validation.js";

const MAX_ATTACHMENT_BYTES = 3 * 1024 * 1024; // 3 MB

function getMimeType(filePath: string): string {
  const ext = path.extname(filePath).toLowerCase();
  const mimeTypes: Record<string, string> = {
    ".pdf": "application/pdf",
    ".doc": "application/msword",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".xls": "application/vnd.ms-excel",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".ppt": "application/vnd.ms-powerpoint",
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".txt": "text/plain",
    ".csv": "text/csv",
    ".zip": "application/zip",
    ".json": "application/json",
  };
  return mimeTypes[ext] || "application/octet-stream";
}

function buildRecipients(addresses: string[] | undefined) {
  if (!addresses || addresses.length === 0) return undefined;
  return addresses.map((addr) => ({ emailAddress: { address: addr } }));
}

function buildAttachments(filePaths: string[] | undefined): Array<{
  "@odata.type": string;
  name: string;
  contentType: string;
  contentBytes: string;
}> {
  if (!filePaths || filePaths.length === 0) return [];

  const attachments = [];
  for (const filePath of filePaths) {
    if (!fs.existsSync(filePath)) {
      throw new Error(`Attachment not found: ${filePath}`);
    }

    const stats = fs.statSync(filePath);
    if (stats.size > MAX_ATTACHMENT_BYTES) {
      const sizeMB = (stats.size / (1024 * 1024)).toFixed(2);
      throw new Error(
        `Attachment "${path.basename(filePath)}" is ${sizeMB} MB, which exceeds the 3 MB limit. Use a smaller file or send a link instead.`,
      );
    }

    const content = fs.readFileSync(filePath);
    attachments.push({
      "@odata.type": "#microsoft.graph.fileAttachment",
      name: path.basename(filePath),
      contentType: getMimeType(filePath),
      contentBytes: content.toString("base64"),
    });
  }

  return attachments;
}

const emailInputSchema = {
  to: z.array(z.string()).describe("Recipient email addresses (required)"),
  cc: z.array(z.string()).optional().describe("CC email addresses"),
  bcc: z.array(z.string()).optional().describe("BCC email addresses"),
  subject: z.string().describe("Email subject (required)"),
  body: z.string().describe("Email body content (required)"),
  contentType: z
    .enum(["text", "html"])
    .optional()
    .describe("Body content type: 'text' for plain text, 'html' for HTML (default: 'text')"),
  attachments: z
    .array(z.string())
    .optional()
    .describe("Local file paths to attach (max 3 MB per file)"),
};

export function registerEmailTools(server: McpServer): void {
  server.registerTool(
    "send_email",
    {
      description:
        "Send an email immediately via Microsoft Outlook. Supports To, CC, BCC, plain text or HTML body, and file attachments (max 3 MB each).",
      inputSchema: emailInputSchema,
    },
    async ({ to, cc, bcc, subject, body, contentType, attachments }) => {
      try {
        if (contentType === "html") {
          const cssError = findUnsupportedCss(body);
          if (cssError) {
            return { content: [{ type: "text", text: cssError }] };
          }
        }

        let attachmentPayload: ReturnType<typeof buildAttachments>;
        try {
          attachmentPayload = buildAttachments(attachments);
        } catch (attachErr) {
          return {
            content: [{ type: "text", text: `❌ ${(attachErr as Error).message}` }],
          };
        }

        const client = await getGraphClient();

        const message: any = {
          subject,
          body: {
            contentType: contentType === "html" ? "HTML" : "Text",
            content: body,
          },
          toRecipients: buildRecipients(to),
        };

        const ccRecipients = buildRecipients(cc);
        if (ccRecipients) message.ccRecipients = ccRecipients;

        const bccRecipients = buildRecipients(bcc);
        if (bccRecipients) message.bccRecipients = bccRecipients;

        if (attachmentPayload.length > 0) {
          message.attachments = attachmentPayload;
        }

        await client.api("/me/sendMail").post({ message });

        return {
          content: [
            {
              type: "text",
              text: `✅ Email sent successfully to ${to.join(", ")}`,
            },
          ],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: handleGraphError(error) }],
        };
      }
    },
  );

  server.registerTool(
    "save_draft",
    {
      description:
        "Save an email as a draft in Microsoft Outlook without sending it. Supports To, CC, BCC, plain text or HTML body, and file attachments (max 3 MB each).",
      inputSchema: emailInputSchema,
    },
    async ({ to, cc, bcc, subject, body, contentType, attachments }) => {
      try {
        if (contentType === "html") {
          const cssError = findUnsupportedCss(body);
          if (cssError) {
            return { content: [{ type: "text", text: cssError }] };
          }
        }

        let attachmentPayload: ReturnType<typeof buildAttachments>;
        try {
          attachmentPayload = buildAttachments(attachments);
        } catch (attachErr) {
          return {
            content: [{ type: "text", text: `❌ ${(attachErr as Error).message}` }],
          };
        }

        const client = await getGraphClient();

        const message: any = {
          subject,
          body: {
            contentType: contentType === "html" ? "HTML" : "Text",
            content: body,
          },
          toRecipients: buildRecipients(to),
        };

        const ccRecipients = buildRecipients(cc);
        if (ccRecipients) message.ccRecipients = ccRecipients;

        const bccRecipients = buildRecipients(bcc);
        if (bccRecipients) message.bccRecipients = bccRecipients;

        if (attachmentPayload.length > 0) {
          message.attachments = attachmentPayload;
        }

        const draft = await client.api("/me/messages").post(message);

        return {
          content: [
            {
              type: "text",
              text: `✅ Draft saved successfully (ID: ${draft.id})`,
            },
          ],
        };
      } catch (error) {
        return {
          content: [{ type: "text", text: handleGraphError(error) }],
        };
      }
    },
  );

  // ── update_draft ─────────────────────────────────────────────────────────
  server.registerTool(
    "update_draft",
    {
      description:
        "Partially update a saved Draft. Only fields you supply are changed — omitted fields are left as-is on the server. Use list_messages(folder: 'drafts') to find Draft IDs.",
      inputSchema: {
        messageId: z.string().describe("The Draft message ID to update"),
        subject: z.string().optional().describe("New subject line"),
        body: z.string().optional().describe("New body content"),
        contentType: z
          .enum(["text", "html"])
          .optional()
          .describe("Body content type (only used when body is supplied, default: 'text')"),
        to: z.array(z.string()).optional().describe("Replace To recipients"),
        cc: z.array(z.string()).optional().describe("Replace CC recipients"),
        bcc: z.array(z.string()).optional().describe("Replace BCC recipients"),
      },
    },
    async ({ messageId, subject, body, contentType, to, cc, bcc }) => {
      try {
        if (body !== undefined && contentType === "html") {
          const cssError = findUnsupportedCss(body);
          if (cssError) {
            return { content: [{ type: "text", text: cssError }] };
          }
        }

        const patch: any = {};

        if (subject !== undefined) patch.subject = subject;

        if (body !== undefined) {
          patch.body = {
            contentType: contentType === "html" ? "HTML" : "Text",
            content: body,
          };
        }

        const toRecipients = buildRecipients(to);
        if (toRecipients) patch.toRecipients = toRecipients;

        const ccRecipients = buildRecipients(cc);
        if (ccRecipients) patch.ccRecipients = ccRecipients;

        const bccRecipients = buildRecipients(bcc);
        if (bccRecipients) patch.bccRecipients = bccRecipients;

        if (Object.keys(patch).length === 0) {
          return {
            content: [
              {
                type: "text",
                text: "❌ No fields to update — supply at least one of: subject, body, to, cc, bcc.",
              },
            ],
          };
        }

        const client = await getGraphClient();
        await client.api(`/me/messages/${messageId}`).patch(patch);

        return {
          content: [
            {
              type: "text",
              text: `✅ Draft updated successfully (ID: ${messageId})`,
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleGraphError(error) }] };
      }
    },
  );

  // ── send_draft ───────────────────────────────────────────────────────────
  server.registerTool(
    "send_draft",
    {
      description:
        "Send a previously saved Draft by its message ID. Use save_draft to create a Draft, update_draft to edit it, then send_draft to transmit it to Recipients.",
      inputSchema: {
        messageId: z.string().describe("The Draft message ID to send"),
      },
    },
    async ({ messageId }) => {
      try {
        const client = await getGraphClient();
        await client.api(`/me/messages/${messageId}/send`).post({});

        return {
          content: [
            {
              type: "text",
              text: `✅ Draft sent successfully (ID: ${messageId})`,
            },
          ],
        };
      } catch (error) {
        return { content: [{ type: "text", text: handleGraphError(error) }] };
      }
    },
  );
}
