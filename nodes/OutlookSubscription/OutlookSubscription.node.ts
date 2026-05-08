import type {
  IDataObject,
  IExecuteFunctions,
  ILoadOptionsFunctions,
  INodeExecutionData,
  INodePropertyOptions,
  INodeType,
  INodeTypeDescription,
  JsonObject,
  NodeConnectionType,
} from "n8n-workflow";
import { NodeApiError } from "n8n-workflow";

import {
  buildClientState,
  buildExpirationDateTime,
  buildSubscriptionTargets,
  clampLifetimeMinutes,
  createMailFolder,
  createSubscription,
  deleteMessage,
  deleteSubscription,
  findDuplicateSubscriptions,
  getAttachment,
  getMessage,
  listMessageAttachments,
  listMessages,
  listSubscriptions,
  loadFolderOptions,
  moveMessage,
  normalizeChangeTypes,
  renewSubscription,
  replyToMessage,
  sendMail,
  updateMessage,
} from "../shared/graph";

export class OutlookSubscription implements INodeType {
  description: INodeTypeDescription = {
    displayName: "Outlook",
    name: "outlookSubscription",
    icon: "file:outlookSubscription.svg",
    group: ["transform"],
    version: 1,
    subtitle: '={{$parameter["resource"] + ": " + $parameter["operation"]}}',
    description:
      "Manage Microsoft Graph Outlook subscriptions, messages, and attachments",
    defaults: {
      name: "Outlook",
    },
    inputs: ["main" as NodeConnectionType],
    outputs: ["main" as NodeConnectionType],
    credentials: [
      {
        name: "microsoftOutlookOAuth2Api",
        required: false,
      },
    ],
    properties: [
      // ── Resource ──
      {
        displayName: "Resource",
        name: "resource",
        type: "options",
        noDataExpression: true,
        default: "subscription",
        options: [
          { name: "Attachment", value: "attachment" },
          { name: "Message", value: "message" },
          { name: "Subscription", value: "subscription" },
        ],
      },

      // ── Operations ──
      {
        displayName: "Operation",
        name: "operation",
        type: "options",
        noDataExpression: true,
        default: "create",
        displayOptions: { show: { resource: ["subscription"] } },
        options: [
          { name: "Create", value: "create" },
          { name: "Delete", value: "delete" },
          { name: "List", value: "list" },
          { name: "Renew", value: "renew" },
        ],
      },
      {
        displayName: "Operation",
        name: "operation",
        type: "options",
        noDataExpression: true,
        default: "get",
        displayOptions: { show: { resource: ["message"] } },
        options: [
          { name: "Delete", value: "delete" },
          { name: "Get", value: "get" },
          { name: "List", value: "list" },
          { name: "Move", value: "move" },
          { name: "Reply", value: "reply" },
          { name: "Reply All", value: "replyAll" },
          { name: "Send", value: "send" },
          { name: "Update", value: "update" },
        ],
      },
      {
        displayName: "Operation",
        name: "operation",
        type: "options",
        noDataExpression: true,
        default: "list",
        displayOptions: { show: { resource: ["attachment"] } },
        options: [
          { name: "Download", value: "download" },
          { name: "List", value: "list" },
        ],
      },
      // ── Mailbox (subscription create only) ──
      {
        displayName: "Mailbox",
        name: "mailboxMode",
        type: "options",
        default: "current",
        displayOptions: {
          show: { resource: ["subscription"], operation: ["create"] },
        },
        options: [
          { name: "Current Mailbox", value: "current" },
          { name: "Other Mailbox", value: "other" },
        ],
      },
      {
        displayName: "Other Mailbox Email",
        name: "otherMailboxEmail",
        type: "string",
        default: "",
        required: true,
        displayOptions: {
          show: {
            resource: ["subscription"],
            operation: ["create"],
            mailboxMode: ["other"],
          },
        },
        placeholder: "shared@example.com",
      },

      // ── Subscription Create ──
      {
        displayName: "Entity",
        name: "entity",
        type: "options",
        default: "message",
        displayOptions: {
          show: { resource: ["subscription"], operation: ["create"] },
        },
        options: [
          { name: "Message", value: "message" },
          { name: "Folder", value: "folder" },
        ],
      },
      {
        displayName: "Folder",
        name: "folderId",
        type: "options",
        typeOptions: {
          loadOptionsMethod: "getMailFolders",
        },
        default: "",
        displayOptions: {
          show: {
            resource: ["subscription"],
            operation: ["create"],
            mailboxMode: ["current"],
          },
        },
        description:
          "Optional folder selection from the signed-in mailbox. Leave empty to target the whole mailbox root for the chosen entity",
      },
      {
        displayName: "Folder ID",
        name: "otherMailboxFolderId",
        type: "string",
        default: "",
        required: true,
        displayOptions: {
          show: {
            resource: ["subscription"],
            operation: ["create"],
            mailboxMode: ["other"],
          },
        },
        placeholder: "AAMkAG...=",
        description:
          "Folder ID from the other mailbox. This is combined with the other mailbox email to build the Microsoft Graph subscription target.",
      },
      {
        displayName: "Subscribe Subfolders",
        name: "includeSubfolders",
        type: "boolean",
        default: false,
        displayOptions: {
          show: {
            resource: ["subscription"],
            operation: ["create"],
            mailboxMode: ["current"],
          },
        },
        description:
          "Whether to expand the selected folder into one subscription per descendant folder",
      },
      {
        displayName: "Change Types",
        name: "changeTypes",
        type: "multiOptions",
        default: ["created", "updated"],
        displayOptions: {
          show: { resource: ["subscription"], operation: ["create"] },
        },
        options: [
          { name: "Created", value: "created" },
          { name: "Updated", value: "updated" },
          { name: "Deleted", value: "deleted" },
        ],
      },
      {
        displayName: "Notification URL",
        name: "notificationUrl",
        type: "string",
        default: "",
        required: true,
        displayOptions: {
          show: { resource: ["subscription"], operation: ["create"] },
        },
        placeholder: "https://example.com/webhooks/outlook",
        description:
          "Webhook URL that will receive notifications. Copy this from the Outlook Subscription Trigger node.",
      },
      {
        displayName: "Client State",
        name: "clientState",
        type: "string",
        default: "",
        displayOptions: {
          show: { resource: ["subscription"], operation: ["create"] },
        },
        description:
          "Optional shared secret used later to verify notifications",
      },
      {
        displayName: "Lifetime Minutes",
        name: "lifetimeMinutes",
        type: "number",
        default: 4230,
        typeOptions: {
          minValue: 45,
        },
        displayOptions: {
          show: {
            resource: ["subscription"],
            operation: ["create", "renew"],
          },
        },
      },
      {
        displayName: "Auto Renew",
        name: "autoRenew",
        type: "boolean",
        default: false,
        displayOptions: {
          show: { resource: ["subscription"], operation: ["create"] },
        },
        description:
          "Whether to include a lifecycle notification URL so an external receiver can renew the subscription",
      },
      {
        displayName: "Lifecycle Notification URL",
        name: "lifecycleNotificationUrl",
        type: "string",
        default: "",
        displayOptions: {
          show: {
            resource: ["subscription"],
            operation: ["create"],
            autoRenew: [true],
          },
        },
        placeholder: "https://example.com/webhooks/outlook-lifecycle",
      },

      // ── Subscription Delete / Renew ──
      {
        displayName: "Subscription ID",
        name: "subscriptionId",
        type: "string",
        default: "",
        required: true,
        displayOptions: {
          show: {
            resource: ["subscription"],
            operation: ["delete", "renew"],
          },
        },
      },

      // ── Message ID (message / attachment) ──
      // For attachment resource: all operations need the message ID
      {
        displayName: "Message ID",
        name: "messageId",
        type: "string",
        default: "",
        required: true,
        displayOptions: { show: { resource: ["attachment"] } },
        placeholder: "AAMkAG...=",
      },
      // For message resource: only ops that act on a specific message
      {
        displayName: "Message ID",
        name: "messageId",
        type: "string",
        default: "",
        required: true,
        displayOptions: {
          show: {
            resource: ["message"],
            operation: ["delete", "get", "move", "reply", "replyAll", "update"],
          },
        },
        placeholder: "AAMkAG...=",
      },

      // ── Attachment ID (download) ──
      {
        displayName: "Attachment ID",
        name: "attachmentId",
        type: "string",
        default: "",
        required: true,
        displayOptions: {
          show: { resource: ["attachment"], operation: ["download"] },
        },
        placeholder: "AAMkAG...=",
      },

      // ── Message Get: Query Options ──
      {
        displayName: "$expand",
        name: "expand",
        type: "string",
        default: "",
        typeOptions: { rows: 3 },
        displayOptions: {
          show: { resource: ["message"], operation: ["get"] },
        },
        placeholder:
          "singleValueExtendedProperties($filter=id eq 'String {00020329-0000-0000-C000-000000000046} Name contentCheckSum')",
        description: "Expand related entities (e.g. extended properties)",
      },
      {
        displayName: "$select",
        name: "select",
        type: "multiOptions",
        default: [],
        displayOptions: {
          show: { resource: ["message"], operation: ["get"] },
        },
        options: [
          { name: "BccRecipients", value: "bccRecipients" },
          { name: "Body", value: "body" },
          { name: "BodyPreview", value: "bodyPreview" },
          { name: "Categories", value: "categories" },
          { name: "CcRecipients", value: "ccRecipients" },
          { name: "ChangeKey", value: "changeKey" },
          { name: "ConversationId", value: "conversationId" },
          { name: "CreatedDateTime", value: "createdDateTime" },
          { name: "Flag", value: "flag" },
          { name: "From", value: "from" },
          { name: "HasAttachments", value: "hasAttachments" },
          { name: "ID", value: "id" },
          { name: "Importance", value: "importance" },
          { name: "InternetMessageHeaders", value: "internetMessageHeaders" },
          { name: "InternetMessageId", value: "internetMessageId" },
          {
            name: "IsDeliveryReceiptRequested",
            value: "isDeliveryReceiptRequested",
          },
          { name: "IsDraft", value: "isDraft" },
          { name: "IsRead", value: "isRead" },
          { name: "IsReadReceiptRequested", value: "isReadReceiptRequested" },
          { name: "LastModifiedDateTime", value: "lastModifiedDateTime" },
          { name: "ParentFolderId", value: "parentFolderId" },
          { name: "ReceivedDateTime", value: "receivedDateTime" },
          { name: "ReplyTo", value: "replyTo" },
          { name: "Sender", value: "sender" },
          { name: "SentDateTime", value: "sentDateTime" },
          { name: "Subject", value: "subject" },
          { name: "ToRecipients", value: "toRecipients" },
          { name: "WebLink", value: "webLink" },
        ],
        description:
          "Fields to return. Leave empty to return all fields. In expression mode, enter field names separated by commas.",
      },

      // ── Message Update: Fields ──
      {
        displayName: "Flag Status",
        name: "flagStatus",
        type: "options",
        default: "none",
        displayOptions: {
          show: { resource: ["message"], operation: ["update"] },
        },
        options: [
          { name: "Don't Change", value: "none" },
          { name: "Flagged", value: "flagged" },
          { name: "Complete", value: "complete" },
          { name: "Not Flagged", value: "notFlagged" },
        ],
      },
      {
        displayName: "Extended Properties",
        name: "extendedProperties",
        type: "fixedCollection",
        typeOptions: { multipleValues: true },
        default: {},
        displayOptions: {
          show: { resource: ["message"], operation: ["update"] },
        },
        options: [
          {
            displayName: "Property",
            name: "property",
            values: [
              {
                displayName: "Property ID",
                name: "id",
                type: "string",
                default: "",
                placeholder:
                  "String {00020329-0000-0000-C000-000000000046} Name contentCheckSum",
              },
              {
                displayName: "Value",
                name: "value",
                type: "string",
                default: "",
              },
            ],
          },
        ],
        description: "Single-value extended properties to set on the message",
      },
      {
        displayName: "If-Match ETag",
        name: "ifMatchETag",
        type: "string",
        default: "",
        displayOptions: {
          show: { resource: ["message"], operation: ["update"] },
        },
        placeholder: 'W/"CQAAABYAAAB..."',
        description:
          'Optimistic concurrency: only update if the server-side message state matches. Use the @odata.etag value from the Get response, or format the changeKey as W/"<changeKey>"',
      },
      {
        displayName: "Additional Body (JSON)",
        name: "additionalBody",
        type: "json",
        default: "{}",
        displayOptions: {
          show: { resource: ["message"], operation: ["update"] },
        },
        description:
          "Additional fields to include in the PATCH body, merged with the fields above",
      },

      // ── Message List ──
      {
        displayName: "Folder ID",
        name: "listFolderId",
        type: "string",
        default: "",
        displayOptions: {
          show: { resource: ["message"], operation: ["list"] },
        },
        placeholder: "AAMkAG...=",
        description:
          "Restrict results to a specific folder. Leave empty to query all mailbox messages.",
      },
      {
        displayName: "Return All",
        name: "returnAll",
        type: "boolean",
        default: false,
        displayOptions: {
          show: { resource: ["message"], operation: ["list"] },
        },
        description:
          "Whether to return all pages of results by following @odata.nextLink",
      },
      {
        displayName: "Limit",
        name: "limit",
        type: "number",
        default: 25,
        typeOptions: { minValue: 1, maxValue: 1000 },
        displayOptions: {
          show: {
            resource: ["message"],
            operation: ["list"],
            returnAll: [false],
          },
        },
        description: "Maximum number of messages to return",
      },
      {
        displayName: "Filter",
        name: "filter",
        type: "string",
        default: "",
        displayOptions: {
          show: { resource: ["message"], operation: ["list"] },
        },
        placeholder: "isRead eq false",
        description: "OData $filter expression to restrict results",
      },
      {
        displayName: "Order By",
        name: "orderBy",
        type: "options",
        default: "receivedDateTime desc",
        displayOptions: {
          show: { resource: ["message"], operation: ["list"] },
        },
        options: [
          { name: "Received (Newest First)", value: "receivedDateTime desc" },
          { name: "Received (Oldest First)", value: "receivedDateTime asc" },
          { name: "Sent (Newest First)", value: "sentDateTime desc" },
          {
            name: "Last Modified (Newest First)",
            value: "lastModifiedDateTime desc",
          },
        ],
      },
      {
        displayName: "Select Fields",
        name: "listSelect",
        type: "multiOptions",
        default: [],
        displayOptions: {
          show: { resource: ["message"], operation: ["list"] },
        },
        options: [
          { name: "BccRecipients", value: "bccRecipients" },
          { name: "Body", value: "body" },
          { name: "BodyPreview", value: "bodyPreview" },
          { name: "Categories", value: "categories" },
          { name: "CcRecipients", value: "ccRecipients" },
          { name: "ChangeKey", value: "changeKey" },
          { name: "ConversationId", value: "conversationId" },
          { name: "CreatedDateTime", value: "createdDateTime" },
          { name: "Flag", value: "flag" },
          { name: "From", value: "from" },
          { name: "HasAttachments", value: "hasAttachments" },
          { name: "ID", value: "id" },
          { name: "Importance", value: "importance" },
          { name: "InternetMessageHeaders", value: "internetMessageHeaders" },
          { name: "InternetMessageId", value: "internetMessageId" },
          { name: "IsDraft", value: "isDraft" },
          { name: "IsRead", value: "isRead" },
          { name: "LastModifiedDateTime", value: "lastModifiedDateTime" },
          { name: "ParentFolderId", value: "parentFolderId" },
          { name: "ReceivedDateTime", value: "receivedDateTime" },
          { name: "ReplyTo", value: "replyTo" },
          { name: "Sender", value: "sender" },
          { name: "SentDateTime", value: "sentDateTime" },
          { name: "Subject", value: "subject" },
          { name: "ToRecipients", value: "toRecipients" },
          { name: "WebLink", value: "webLink" },
        ],
        description:
          "Fields to return. Leave empty to return all default fields.",
      },

      // ── Message Move ──
      {
        displayName: "Specify Destination Folder As",
        name: "destinationFolderSource",
        type: "options",
        default: "list",
        displayOptions: {
          show: { resource: ["message"], operation: ["move"] },
        },
        options: [
          {
            name: "Select from List",
            value: "list",
            description: "Pick a folder from the dropdown",
          },
          {
            name: "Enter Folder ID",
            value: "manual",
            description: "Type the folder ID directly",
          },
          {
            name: "Create New Folder",
            value: "create",
            description: "Create a new folder and move the message into it",
          },
        ],
      },
      {
        displayName: "Destination Folder",
        name: "destinationFolderId",
        type: "options",
        typeOptions: { loadOptionsMethod: "getMailFolders" },
        default: "",
        displayOptions: {
          show: {
            resource: ["message"],
            operation: ["move"],
            destinationFolderSource: ["list"],
          },
        },
        description: "The folder to move the message into",
      },
      {
        displayName: "Destination Folder ID",
        name: "destinationFolderIdManual",
        type: "string",
        default: "",
        displayOptions: {
          show: {
            resource: ["message"],
            operation: ["move"],
            destinationFolderSource: ["manual"],
          },
        },
        placeholder: "AAMkAG...=",
      },
      {
        displayName: "New Folder Name",
        name: "newFolderName",
        type: "string",
        default: "",
        required: true,
        displayOptions: {
          show: {
            resource: ["message"],
            operation: ["move"],
            destinationFolderSource: ["create"],
          },
        },
        placeholder: "Processed Invoices",
        description: "Name of the new top-level mail folder to create",
      },

      // ── Message Send ──
      {
        displayName: "To",
        name: "sendTo",
        type: "string",
        default: "",
        required: true,
        displayOptions: {
          show: { resource: ["message"], operation: ["send"] },
        },
        placeholder: "alice@example.com, bob@example.com",
        description: "Comma-separated list of recipient email addresses",
      },
      {
        displayName: "Subject",
        name: "sendSubject",
        type: "string",
        default: "",
        required: true,
        displayOptions: {
          show: { resource: ["message"], operation: ["send"] },
        },
      },
      {
        displayName: "Body",
        name: "sendBody",
        type: "string",
        default: "",
        required: true,
        typeOptions: { rows: 5 },
        displayOptions: {
          show: { resource: ["message"], operation: ["send"] },
        },
        description:
          'For HTML inline images, reference the matching Inline Images Content ID as &lt;img src="cid:logo" &gt;.',
      },
      {
        displayName: "Body Type",
        name: "sendBodyType",
        type: "options",
        default: "html",
        displayOptions: {
          show: { resource: ["message"], operation: ["send"] },
        },
        options: [
          { name: "HTML", value: "html" },
          { name: "Plain Text", value: "text" },
        ],
      },
      {
        displayName: "CC",
        name: "sendCc",
        type: "string",
        default: "",
        displayOptions: {
          show: { resource: ["message"], operation: ["send"] },
        },
        placeholder: "cc@example.com",
        description: "Comma-separated CC recipients",
      },
      {
        displayName: "BCC",
        name: "sendBcc",
        type: "string",
        default: "",
        displayOptions: {
          show: { resource: ["message"], operation: ["send"] },
        },
        placeholder: "bcc@example.com",
        description: "Comma-separated BCC recipients",
      },
      {
        displayName: "Save to Sent Items",
        name: "saveToSentItems",
        type: "boolean",
        default: true,
        displayOptions: {
          show: { resource: ["message"], operation: ["send"] },
        },
      },
      {
        displayName: "Add Attachment",
        name: "addAttachment",
        type: "boolean",
        default: false,
        displayOptions: {
          show: { resource: ["message"], operation: ["send"] },
        },
        description:
          "Whether to attach a binary file from the incoming item to the email",
      },
      {
        displayName: "Binary Field Name",
        name: "binaryFieldName",
        type: "string",
        default: "data",
        displayOptions: {
          show: {
            resource: ["message"],
            operation: ["send"],
            addAttachment: [true],
          },
        },
        placeholder: "data",
        description:
          "Name of the binary property on the incoming item that holds the file to attach",
      },
      {
        displayName: "Inline Images",
        name: "inlineImages",
        type: "fixedCollection",
        typeOptions: { multipleValues: true },
        default: {},
        displayOptions: {
          show: {
            resource: ["message"],
            operation: ["send"],
            sendBodyType: ["html"],
          },
        },
        description:
          'Embed images inline in the HTML body. Reference each image in the body as &lt;img src="cid:your-content-id" &gt;.',
        options: [
          {
            displayName: "Image",
            name: "image",
            values: [
              {
                displayName: "Binary Field Name",
                name: "binaryFieldName",
                type: "string",
                default: "data",
                placeholder: "data",
                description:
                  "Name of the binary property on the incoming item that holds the image file",
              },
              {
                displayName: "Content ID",
                name: "contentId",
                type: "string",
                default: "",
                placeholder: "logo",
                description:
                  'Unique identifier for this image. Use it in the HTML body as &lt;img src="cid:logo"&gt;',
              },
            ],
          },
        ],
      },
      {
        displayName: "Comment",
        name: "replyComment",
        type: "string",
        default: "",
        required: true,
        typeOptions: { rows: 5 },
        displayOptions: {
          show: { resource: ["message"], operation: ["reply", "replyAll"] },
        },
        description: "The reply body content",
      },
      {
        displayName: "Body Type",
        name: "replyBodyType",
        type: "options",
        default: "html",
        displayOptions: {
          show: { resource: ["message"], operation: ["reply", "replyAll"] },
        },
        options: [
          { name: "HTML", value: "html" },
          { name: "Plain Text", value: "text" },
        ],
      },
    ],
  };

  methods = {
    loadOptions: {
      async getMailFolders(
        this: ILoadOptionsFunctions,
      ): Promise<INodePropertyOptions[]> {
        return await loadFolderOptions.call(this);
      },
    },
  };

  async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
    const items = this.getInputData();
    const returnData: INodeExecutionData[] = [];

    for (let itemIndex = 0; itemIndex < items.length; itemIndex++) {
      try {
        const resource = this.getNodeParameter("resource", itemIndex) as string;
        const operation = this.getNodeParameter(
          "operation",
          itemIndex,
        ) as string;

        // ── SUBSCRIPTION ──
        if (resource === "subscription") {
          if (operation === "list") {
            const subscriptions = await listSubscriptions.call(this);
            returnData.push(
              ...this.helpers.returnJsonArray(subscriptions as IDataObject[]),
            );
            continue;
          }

          if (operation === "delete") {
            const subscriptionId = this.getNodeParameter(
              "subscriptionId",
              itemIndex,
            ) as string;
            await deleteSubscription.call(this, subscriptionId);
            returnData.push({
              json: {
                success: true,
                subscriptionId,
                operation: "delete",
              },
            });
            continue;
          }

          if (operation === "renew") {
            const subscriptionId = this.getNodeParameter(
              "subscriptionId",
              itemIndex,
            ) as string;
            const lifetimeMinutes = clampLifetimeMinutes(
              this.getNodeParameter("lifetimeMinutes", itemIndex) as number,
            );
            const renewed = await renewSubscription.call(
              this,
              subscriptionId,
              lifetimeMinutes,
            );
            returnData.push({
              json: {
                operation: "renew",
                subscription: renewed,
              },
            });
            continue;
          }

          // operation === "create"
          const mailboxMode = this.getNodeParameter(
            "mailboxMode",
            itemIndex,
          ) as "current" | "other";
          const otherMailboxEmail = this.getNodeParameter(
            "otherMailboxEmail",
            itemIndex,
            "",
          ) as string;
          const entity = this.getNodeParameter("entity", itemIndex) as
            | "message"
            | "folder";
          const folderId = this.getNodeParameter(
            "folderId",
            itemIndex,
            "",
          ) as string;
          const otherMailboxFolderId = this.getNodeParameter(
            "otherMailboxFolderId",
            itemIndex,
            "",
          ) as string;
          const includeSubfolders =
            mailboxMode === "current"
              ? ((this.getNodeParameter(
                  "includeSubfolders",
                  itemIndex,
                ) as boolean) ?? false)
              : false;
          const notificationUrl = this.getNodeParameter(
            "notificationUrl",
            itemIndex,
          ) as string;
          const changeTypes = this.getNodeParameter(
            "changeTypes",
            itemIndex,
          ) as string[];
          const changeType = normalizeChangeTypes(changeTypes);
          const lifetimeMinutes = clampLifetimeMinutes(
            this.getNodeParameter("lifetimeMinutes", itemIndex) as number,
          );
          const autoRenew = this.getNodeParameter(
            "autoRenew",
            itemIndex,
          ) as boolean;
          const lifecycleNotificationUrl = this.getNodeParameter(
            "lifecycleNotificationUrl",
            itemIndex,
            "",
          ) as string;
          const clientState = buildClientState(
            this.getNodeParameter("clientState", itemIndex, "") as string,
          );

          const targets = await buildSubscriptionTargets.call(this, {
            mailboxMode,
            otherMailboxEmail,
            entity,
            folderId: mailboxMode === "other" ? otherMailboxFolderId : folderId,
            includeSubfolders,
          });

          const existingSubscriptions = await listSubscriptions.call(this);

          for (const target of targets) {
            const duplicates = findDuplicateSubscriptions(
              existingSubscriptions,
              target,
              changeType,
              notificationUrl,
            );

            if (duplicates.length > 0) {
              returnData.push({
                json: {
                  operation: "create",
                  status: "skipped_duplicate",
                  resource: target,
                  changeType,
                  notificationUrl,
                  existing: duplicates,
                  autoRenew,
                  lifetimeMinutes,
                },
              });
              continue;
            }

            const created = await createSubscription.call(this, {
              changeType,
              notificationUrl,
              clientState,
              expirationDateTime: buildExpirationDateTime(lifetimeMinutes),
              resource: target,
              latestSupportedTlsVersion: "v1_2",
              ...(autoRenew && lifecycleNotificationUrl
                ? { lifecycleNotificationUrl }
                : {}),
            });

            returnData.push({
              json: {
                operation: "create",
                status: "created",
                resource: target,
                changeType,
                notificationUrl,
                autoRenew,
                lifetimeMinutes,
                subscription: created,
              },
            });
          }
          continue;
        }

        // ── MESSAGE ──
        if (resource === "message") {
          const config = { mailboxMode: "current" as const };

          // ── list ──
          if (operation === "list") {
            const listFolderId = (
              this.getNodeParameter("listFolderId", itemIndex, "") as string
            ).trim();
            const returnAll = this.getNodeParameter(
              "returnAll",
              itemIndex,
            ) as boolean;
            const limit = this.getNodeParameter(
              "limit",
              itemIndex,
              25,
            ) as number;
            const filter = (
              this.getNodeParameter("filter", itemIndex, "") as string
            ).trim();
            const orderBy = this.getNodeParameter(
              "orderBy",
              itemIndex,
              "receivedDateTime desc",
            ) as string;
            const listSelectRaw = this.getNodeParameter(
              "listSelect",
              itemIndex,
              [],
            );
            const listSelectFields = Array.isArray(listSelectRaw)
              ? (listSelectRaw as string[]).join(",")
              : (listSelectRaw as string)
                  .split(",")
                  .map((s) => s.trim())
                  .filter(Boolean)
                  .join(",");

            const qs: IDataObject = { $orderby: orderBy };
            if (!returnAll) qs.$top = Math.min(limit, 1000);
            if (filter) qs.$filter = filter;
            if (listSelectFields) qs.$select = listSelectFields;

            const messages = await listMessages.call(
              this,
              config,
              listFolderId || undefined,
              qs,
              returnAll,
            );
            returnData.push(...this.helpers.returnJsonArray(messages));
            continue;
          }

          // ── send ──
          if (operation === "send") {
            const sendTo = this.getNodeParameter("sendTo", itemIndex) as string;
            const sendSubject = this.getNodeParameter(
              "sendSubject",
              itemIndex,
            ) as string;
            const sendBody = this.getNodeParameter(
              "sendBody",
              itemIndex,
            ) as string;
            const sendBodyType = this.getNodeParameter(
              "sendBodyType",
              itemIndex,
              "html",
            ) as "html" | "text";
            const sendCc = (
              this.getNodeParameter("sendCc", itemIndex, "") as string
            ).trim();
            const sendBcc = (
              this.getNodeParameter("sendBcc", itemIndex, "") as string
            ).trim();
            const saveToSentItems = this.getNodeParameter(
              "saveToSentItems",
              itemIndex,
              true,
            ) as boolean;
            const addAttachment = this.getNodeParameter(
              "addAttachment",
              itemIndex,
              false,
            ) as boolean;

            const attachments: Array<{
              name: string;
              contentType: string;
              contentBytes: string;
            }> = [];

            if (addAttachment) {
              const binaryFieldName = (
                this.getNodeParameter(
                  "binaryFieldName",
                  itemIndex,
                  "data",
                ) as string
              ).trim();
              const binaryData = this.helpers.assertBinaryData(
                itemIndex,
                binaryFieldName,
              );
              const buffer = await this.helpers.getBinaryDataBuffer(
                itemIndex,
                binaryFieldName,
              );
              attachments.push({
                name: binaryData.fileName || binaryFieldName,
                contentType: binaryData.mimeType || "application/octet-stream",
                contentBytes: buffer.toString("base64"),
              });
            }

            const inlineImagesCollection = this.getNodeParameter(
              "inlineImages",
              itemIndex,
              {},
            ) as {
              image?: Array<{ binaryFieldName: string; contentId: string }>;
            };
            const inlineImageEntries = inlineImagesCollection.image ?? [];
            const inlineImages: Array<{
              name: string;
              contentType: string;
              contentBytes: string;
              contentId: string;
            }> = [];
            for (const entry of inlineImageEntries) {
              const imgField = entry.binaryFieldName.trim() || "data";
              const imgBinary = this.helpers.assertBinaryData(
                itemIndex,
                imgField,
              );
              const imgBuffer = await this.helpers.getBinaryDataBuffer(
                itemIndex,
                imgField,
              );
              inlineImages.push({
                name: imgBinary.fileName || imgField,
                contentType: imgBinary.mimeType || "application/octet-stream",
                contentBytes: imgBuffer.toString("base64"),
                contentId: entry.contentId.trim() || imgField,
              });
            }

            await sendMail.call(this, config, {
              to: sendTo,
              subject: sendSubject,
              body: sendBody,
              bodyType: sendBodyType,
              cc: sendCc || undefined,
              bcc: sendBcc || undefined,
              saveToSentItems,
              attachments: attachments.length > 0 ? attachments : undefined,
              inlineImages: inlineImages.length > 0 ? inlineImages : undefined,
            });
            returnData.push({ json: { success: true, operation: "send" } });
            continue;
          }

          // All remaining message operations require a messageId
          const messageId = this.getNodeParameter(
            "messageId",
            itemIndex,
          ) as string;

          // ── delete ──
          if (operation === "delete") {
            await deleteMessage.call(this, config, messageId);
            returnData.push({
              json: { success: true, operation: "delete", messageId },
            });
            continue;
          }

          // ── move ──
          if (operation === "move") {
            const destinationFolderSource = this.getNodeParameter(
              "destinationFolderSource",
              itemIndex,
              "list",
            ) as "list" | "manual" | "create";

            let destinationId: string;

            if (destinationFolderSource === "create") {
              const newFolderName = (
                this.getNodeParameter("newFolderName", itemIndex, "") as string
              ).trim();
              const created = await createMailFolder.call(
                this,
                config,
                newFolderName,
              );
              destinationId = created.id;
            } else if (destinationFolderSource === "manual") {
              destinationId = this.getNodeParameter(
                "destinationFolderIdManual",
                itemIndex,
                "",
              ) as string;
            } else {
              destinationId = this.getNodeParameter(
                "destinationFolderId",
                itemIndex,
                "",
              ) as string;
            }

            const moved = await moveMessage.call(
              this,
              config,
              messageId,
              destinationId,
            );
            returnData.push({ json: moved });
            continue;
          }

          // ── reply / replyAll ──
          if (operation === "reply" || operation === "replyAll") {
            const replyComment = this.getNodeParameter(
              "replyComment",
              itemIndex,
            ) as string;
            const replyBodyType = this.getNodeParameter(
              "replyBodyType",
              itemIndex,
              "html",
            ) as "html" | "text";

            await replyToMessage.call(
              this,
              config,
              messageId,
              replyComment,
              replyBodyType,
              operation === "replyAll",
            );
            returnData.push({
              json: { success: true, operation, messageId },
            });
            continue;
          }

          if (operation === "get") {
            const expand = this.getNodeParameter(
              "expand",
              itemIndex,
              "",
            ) as string;
            const selectRaw = this.getNodeParameter("select", itemIndex, []);
            const selectFields = Array.isArray(selectRaw)
              ? (selectRaw as string[]).join(",")
              : (selectRaw as string)
                  .split(",")
                  .map((s) => s.trim())
                  .filter(Boolean)
                  .join(",");
            const qs: IDataObject = {};
            if (expand) qs.$expand = expand;
            if (selectFields) qs.$select = selectFields;

            const message = await getMessage.call(
              this,
              config,
              messageId,
              Object.keys(qs).length > 0 ? qs : undefined,
            );
            returnData.push({ json: message });
            continue;
          }

          if (operation === "update") {
            const body: IDataObject = {};

            // Flag status
            const flagStatus = this.getNodeParameter(
              "flagStatus",
              itemIndex,
              "none",
            ) as string;
            if (flagStatus && flagStatus !== "none") {
              body.flag = { flagStatus };
            }

            // Extended properties
            const extProps = this.getNodeParameter(
              "extendedProperties",
              itemIndex,
              {},
            ) as IDataObject;
            const properties = extProps.property as IDataObject[] | undefined;
            if (Array.isArray(properties) && properties.length > 0) {
              body.singleValueExtendedProperties = properties;
            }

            // Additional body fields
            let additionalBody = this.getNodeParameter(
              "additionalBody",
              itemIndex,
              "{}",
            );
            if (typeof additionalBody === "string") {
              try {
                additionalBody = JSON.parse(additionalBody);
              } catch {
                throw new Error("Invalid JSON in Additional Body (JSON) field");
              }
            }
            if (
              additionalBody &&
              typeof additionalBody === "object" &&
              Object.keys(additionalBody as IDataObject).length > 0
            ) {
              Object.assign(body, additionalBody);
            }

            const etag = this.getNodeParameter(
              "ifMatchETag",
              itemIndex,
              "",
            ) as string;

            const updated = await updateMessage.call(
              this,
              config,
              messageId,
              body,
              etag || undefined,
            );
            returnData.push({ json: updated });
            continue;
          }
        }

        // ── ATTACHMENT ──
        if (resource === "attachment") {
          const messageId = this.getNodeParameter(
            "messageId",
            itemIndex,
          ) as string;
          const config = { mailboxMode: "current" as const };

          if (operation === "list") {
            const attachments = await listMessageAttachments.call(
              this,
              config,
              messageId,
            );
            returnData.push(...this.helpers.returnJsonArray(attachments));
            continue;
          }

          if (operation === "download") {
            const attachmentId = this.getNodeParameter(
              "attachmentId",
              itemIndex,
            ) as string;
            const attachment = await getAttachment.call(
              this,
              config,
              messageId,
              attachmentId,
            );

            const binaryPropertyName = "data";
            const fileName = String(
              attachment.name || `attachment_${attachmentId}`,
            );
            const contentType = String(
              attachment.contentType || "application/octet-stream",
            );
            const contentBytes = String(attachment.contentBytes || "");

            returnData.push({
              json: attachment,
              binary: {
                [binaryPropertyName]: await this.helpers.prepareBinaryData(
                  Buffer.from(contentBytes, "base64"),
                  fileName,
                  contentType,
                ),
              },
            });
            continue;
          }
        }
      } catch (error) {
        if (this.continueOnFail()) {
          returnData.push({
            json: { error: (error as Error).message },
            pairedItem: { item: itemIndex },
          });
          continue;
        }

        throw new NodeApiError(this.getNode(), error as JsonObject, {
          itemIndex,
        });
      }
    }

    return [returnData];
  }
}
