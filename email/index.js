/**
 * Email module for Outlook MCP server
 */
const handleListEmails = require('./list');
const handleSearchEmails = require('./search');
const handleReadEmail = require('./read');
const handleSendEmail = require('./send');
const handleListSharedMailboxes = require('./list-shared-mailboxes');
const handleCreateReplyDraft = require('./reply-draft');
const handleMarkRead = require('./mark-read');
const handleGetAttachments = require('./get-attachments');
const handleCreateDraft = require('./create-draft');

// Email tool definitions
const emailTools = [
  {
    name: "list-emails",
    description: "Lists recent emails from your inbox",
    inputSchema: {
      type: "object",
      properties: {
        folder: {
          type: "string",
          description: "Email folder to list (e.g., 'inbox', 'sent', 'drafts', default: 'inbox')"
        },
        count: {
          type: "number",
          description: "Number of emails to retrieve (default: 10, max: 50)"
        },
        sharedMailbox: {
          type: "string",
          description: "Email address of shared mailbox to access (e.g., 'shared@company.com')"
        }
      },
      required: []
    },
    handler: handleListEmails
  },
  {
    name: "search-emails",
    description: "Search for emails using various criteria",
    inputSchema: {
      type: "object",
      properties: {
        query: {
          type: "string",
          description: "Search query text to find in emails"
        },
        folder: {
          type: "string",
          description: "Email folder to search in (default: 'inbox')"
        },
        from: {
          type: "string",
          description: "Filter by sender email address or name"
        },
        to: {
          type: "string",
          description: "Filter by recipient email address or name"
        },
        subject: {
          type: "string",
          description: "Filter by email subject"
        },
        hasAttachments: {
          type: "boolean",
          description: "Filter to only emails with attachments"
        },
        unreadOnly: {
          type: "boolean",
          description: "Filter to only unread emails"
        },
        count: {
          type: "number",
          description: "Number of results to return (default: 10, max: 50)"
        }
      },
      required: []
    },
    handler: handleSearchEmails
  },
  {
    name: "read-email",
    description: "Reads the content of a specific email",
    inputSchema: {
      type: "object",
      properties: {
        id: {
          type: "string",
          description: "ID of the email to read"
        },
        sharedMailbox: {
          type: "string",
          description: "Email address of shared mailbox to read from (e.g., 'shared@company.com')"
        }
      },
      required: ["id"]
    },
    handler: handleReadEmail
  },
  {
    name: "send-email",
    description: "Composes and sends a new email",
    inputSchema: {
      type: "object",
      properties: {
        to: {
          type: "string",
          description: "Comma-separated list of recipient email addresses"
        },
        cc: {
          type: "string",
          description: "Comma-separated list of CC recipient email addresses"
        },
        bcc: {
          type: "string",
          description: "Comma-separated list of BCC recipient email addresses"
        },
        subject: {
          type: "string",
          description: "Email subject"
        },
        body: {
          type: "string",
          description: "Email body content (can be plain text or HTML)"
        },
        importance: {
          type: "string",
          description: "Email importance (normal, high, low)",
          enum: ["normal", "high", "low"]
        },
        saveToSentItems: {
          type: "boolean",
          description: "Whether to save the email to sent items"
        },
        sharedMailbox: {
          type: "string",
          description: "Email address of shared mailbox to send from (e.g., 'shared@company.com')"
        }
      },
      required: ["to", "subject", "body"]
    },
    handler: handleSendEmail
  },
  {
    name: "list-shared-mailboxes",
    description: "Lists available shared mailboxes that you have access to",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleListSharedMailboxes
  },
  {
    name: "create-reply-draft",
    description: "Creates a draft reply to an existing email",
    inputSchema: {
      type: "object",
      properties: {
        emailId: {
          type: "string",
          description: "ID of the email to reply to"
        },
        body: {
          type: "string",
          description: "Content of the reply message"
        },
        replyAll: {
          type: "boolean",
          description: "Whether to reply to all recipients (default: false)"
        },
        sharedMailbox: {
          type: "string",
          description: "Email address of shared mailbox to reply from (e.g., 'shared@company.com')"
        }
      },
      required: ["emailId", "body"]
    },
    handler: handleCreateReplyDraft
  },
  {
    name: "mark-read",
    description: "Marks emails as read or unread",
    inputSchema: {
      type: "object",
      properties: {
        emailIds: {
          type: "string",
          description: "Comma-separated list of email IDs to mark"
        },
        isRead: {
          type: "boolean",
          description: "Whether to mark as read (true) or unread (false). Default: true"
        },
        sharedMailbox: {
          type: "string",
          description: "Email address of shared mailbox (e.g., 'shared@company.com')"
        }
      },
      required: ["emailIds"]
    },
    handler: handleMarkRead
  },
  {
    name: "download-attachments",
    description: "Downloads information about email attachments",
    inputSchema: {
      type: "object",
      properties: {
        emailId: {
          type: "string",
          description: "ID of the email to get attachments from"
        },
        sharedMailbox: {
          type: "string",
          description: "Email address of shared mailbox (e.g., 'shared@company.com')"
        }
      },
      required: ["emailId"]
    },
    handler: handleGetAttachments
  },
  {
    name: "create-draft",
    description: "Creates a new draft email",
    inputSchema: {
      type: "object",
      properties: {
        to: {
          type: "string",
          description: "Comma-separated list of recipient email addresses"
        },
        cc: {
          type: "string",
          description: "Comma-separated list of CC recipient email addresses"
        },
        bcc: {
          type: "string",
          description: "Comma-separated list of BCC recipient email addresses"
        },
        subject: {
          type: "string",
          description: "Email subject"
        },
        body: {
          type: "string",
          description: "Email body content (can be plain text or HTML)"
        },
        importance: {
          type: "string",
          description: "Email importance (normal, high, low)",
          enum: ["normal", "high", "low"]
        },
        sharedMailbox: {
          type: "string",
          description: "Email address of shared mailbox to create draft from (e.g., 'shared@company.com')"
        }
      },
      required: ["to", "subject", "body"]
    },
    handler: handleCreateDraft
  }
];

module.exports = {
  emailTools,
  handleListEmails,
  handleSearchEmails,
  handleReadEmail,
  handleSendEmail,
  handleListSharedMailboxes,
  handleCreateReplyDraft,
  handleMarkRead,
  handleGetAttachments,
  handleCreateDraft
};
