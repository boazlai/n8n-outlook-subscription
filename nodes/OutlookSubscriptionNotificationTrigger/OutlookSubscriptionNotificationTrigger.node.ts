import type {
  IBinaryKeyData,
  IDataObject,
  IHookFunctions,
  ILoadOptionsFunctions,
  INodeExecutionData,
  INodePropertyOptions,
  INodeType,
  INodeTypeDescription,
  IWebhookFunctions,
  IWebhookResponseData,
  NodeConnectionType,
} from "n8n-workflow";

import {
  buildExpirationDateTime,
  buildSubscriptionTargets,
  convertBodyToMarkdown,
  createSubscription,
  deleteSubscription,
  evaluateMessageFilter,
  findDuplicateSubscriptions,
  graphApiRequest,
  listSubscriptions,
  loadFolderOptions,
  normalizeChangeTypes,
  renewSubscription,
  resolveMessageByPath,
  GRAPH_SUBSCRIPTION_MINUTES_MAX,
} from "../shared/graph";

type WebhookNotification = {
  changeType?: string;
  clientState?: string;
  subscriptionId?: string;
  resource?: string;
  resourceData?: {
    id?: string;
    "@odata.type"?: string;
    "@odata.id"?: string;
    "@odata.etag"?: string;
  };
  lifecycleEvent?: string;
};

export class OutlookSubscriptionNotificationTrigger implements INodeType {
  description: INodeTypeDescription = {
    displayName: "Outlook Trigger",
    name: "outlookSubscriptionTrigger",
    icon: "file:outlookSubscriptionNotificationTrigger.svg",
    group: ["trigger"],
    version: 1,
    subtitle: '"Trigger: Notification"',
    description:
      "Triggers when a Microsoft Graph Outlook mail notification is received. Automatically creates and manages the subscription.",
    defaults: {
      name: "Outlook Trigger",
    },
    inputs: [],
    outputs: ["main" as NodeConnectionType],
    credentials: [
      {
        name: "microsoftOutlookOAuth2Api",
        required: true,
      },
    ],
    webhooks: [
      {
        name: "default",
        httpMethod: "POST",
        responseMode: "onReceived",
        path: '={{$parameter["webhookPath"]}}',
      },
    ],
    properties: [
      // Auto-generated unique path derived from node ID
      {
        displayName: "Webhook Path",
        name: "webhookPath",
        type: "string",
        default: "={{$nodeId}}",
        required: true,
        description:
          "Auto-generated unique path for this webhook endpoint. Do not change.",
      },

      // ── Folder ──
      {
        displayName: "Folder",
        name: "folderId",
        type: "options",
        typeOptions: {
          loadOptionsMethod: "getMailFolders",
        },
        default: "",
        description:
          "Folder to watch for changes. Leave empty to watch the entire mailbox.",
      },

      // ── Subscribe Subfolders ──
      {
        displayName: "Subscribe Subfolders",
        name: "includeSubfolders",
        type: "boolean",
        default: false,
        description:
          "Whether to also subscribe to all descendant subfolders of the selected folder",
      },

      // ── Change Types ──
      {
        displayName: "Change Types",
        name: "changeTypes",
        type: "multiOptions",
        default: ["created", "updated"],
        options: [
          { name: "Created", value: "created" },
          { name: "Deleted", value: "deleted" },
          { name: "Updated", value: "updated" },
        ],
        description: "The types of changes to subscribe to",
      },

      // ── Auto Renew ──
      {
        displayName: "Auto Renew",
        name: "autoRenew",
        type: "boolean",
        default: true,
        description:
          "Whether to automatically renew the subscription when it is about to expire via Microsoft Graph lifecycle notifications",
      },

      // ── Resolve Full Message ──
      {
        displayName: "Resolve Full Message",
        name: "resolveMessageData",
        type: "boolean",
        default: true,
        description:
          "Whether to fetch the full message content for message notifications (skipped for deleted messages)",
      },

      // ── Options ──
      {
        displayName: "Options",
        name: "options",
        type: "collection",
        placeholder: "Add Option",
        default: {},
        options: [
          {
            displayName: "Client State",
            name: "clientState",
            type: "string",
            default: "",
            placeholder: "my-secret-value",
            description:
              "Shared secret used to validate incoming notifications. Set this when the subscription was created externally (e.g. via the Outlook action node) with a matching client state. Leave empty for self-managed subscriptions — no client state is sent or validated.",
          },
          {
            displayName: "Expand Properties",
            name: "expandProperties",
            type: "string",
            default: "",
            placeholder:
              "singleValueExtendedProperties($filter=id eq 'String {00020329-0000-0000-C000-000000000046} Name contentCheckSum')",
            description: "$expand query parameter when resolving the message",
          },
          {
            displayName: "Include Attachments",
            name: "includeAttachments",
            type: "boolean",
            default: false,
            description:
              "Whether to download all file attachments as binary data (data_0, data_1, …)",
          },
          {
            displayName: "Message Filters",
            name: "messageFilters",
            type: "filter",
            default: {},
            typeOptions: {
              filter: {
                version: 2,
                caseSensitive: false,
                typeValidation: "loose",
              },
            },
            description:
              "Filter messages by field values. Enter a dot-notation field path as the left value (e.g. subject, from.emailAddress.address, changeType, hasAttachments). Only evaluated when Resolve Full Message is on.",
          },
          {
            displayName: "Select Fields",
            name: "selectFields",
            type: "string",
            default: "",
            placeholder: "id,subject,body,from,hasAttachments",
            description: "$select query parameter when resolving the message",
          },
          {
            displayName: "Convert Body to Markdown",
            name: "bodyToMarkdown",
            type: "boolean",
            default: false,
            description:
              "Whether to convert the HTML body content to Markdown. Only applied when the body field is present and its contentType is html. Requires Resolve Full Message to be on.",
          },
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

  webhookMethods = {
    default: {
      async checkExists(this: IHookFunctions): Promise<boolean> {
        const staticData = this.getWorkflowStaticData("node");
        const ids = staticData.subscriptionIds as string[] | undefined;
        if (!ids || ids.length === 0) return false;

        try {
          const existing = await listSubscriptions.call(this);
          const existingIds = new Set(existing.map((s) => s.id));
          return ids.every((id) => existingIds.has(id));
        } catch {
          return false;
        }
      },

      async create(this: IHookFunctions): Promise<boolean> {
        const staticData = this.getWorkflowStaticData("node");

        const notificationUrl = this.getNodeWebhookUrl("default") as string;
        const autoRenew = this.getNodeParameter("autoRenew", true) as boolean;
        const folderId = (
          this.getNodeParameter("folderId", "") as string
        ).trim();
        const includeSubfolders = this.getNodeParameter(
          "includeSubfolders",
          false,
        ) as boolean;
        const changeTypes = this.getNodeParameter("changeTypes", [
          "created",
          "updated",
        ]) as string[];
        const changeType = normalizeChangeTypes(changeTypes);

        // Use user-provided client state if set; otherwise skip entirely
        const options = this.getNodeParameter("options", {}) as IDataObject;
        const clientState =
          (options.clientState as string | undefined)?.trim() || undefined;
        staticData.clientState = clientState;

        const targets = await buildSubscriptionTargets.call(this, {
          mailboxMode: "current",
          entity: "message",
          folderId: folderId || undefined,
          includeSubfolders,
        });

        const existingSubscriptions = await listSubscriptions.call(this);
        const createdIds: string[] = [];

        for (const target of targets) {
          const duplicates = findDuplicateSubscriptions(
            existingSubscriptions,
            target,
            changeType,
            notificationUrl,
          );

          if (duplicates.length > 0) {
            // Reuse the existing subscription
            createdIds.push(duplicates[0].id);
            continue;
          }

          const created = await createSubscription.call(this, {
            changeType,
            notificationUrl,
            ...(clientState ? { clientState } : {}),
            expirationDateTime: buildExpirationDateTime(
              GRAPH_SUBSCRIPTION_MINUTES_MAX,
            ),
            resource: target,
            latestSupportedTlsVersion: "v1_2",
            ...(autoRenew ? { lifecycleNotificationUrl: notificationUrl } : {}),
          });

          createdIds.push(created.id);
        }

        staticData.subscriptionIds = createdIds;
        return true;
      },

      async delete(this: IHookFunctions): Promise<boolean> {
        const staticData = this.getWorkflowStaticData("node");
        const ids = staticData.subscriptionIds as string[] | undefined;

        if (ids && ids.length > 0) {
          for (const id of ids) {
            try {
              await deleteSubscription.call(this, id);
            } catch {
              // Best-effort: subscription may already be gone
            }
          }
        }

        staticData.subscriptionIds = [];
        staticData.clientState = undefined;
        return true;
      },
    },
  };

  async webhook(this: IWebhookFunctions): Promise<IWebhookResponseData> {
    const req = this.getRequestObject();
    const res = this.getResponseObject();

    // MS Graph validation handshake
    if (req.query.validationToken) {
      res
        .status(200)
        .type("text/plain")
        .send(String(req.query.validationToken));
      return { noWebhookResponse: true };
    }

    const staticData = this.getWorkflowStaticData("node");
    const expectedClientState = (
      (staticData.clientState as string) || ""
    ).trim();
    const autoRenew = this.getNodeParameter("autoRenew", true) as boolean;
    const resolveMessageData = this.getNodeParameter(
      "resolveMessageData",
      true,
    ) as boolean;
    const options = this.getNodeParameter("options", {}) as IDataObject;
    const expandProperties = (options.expandProperties as string) || "";
    const selectFields = (options.selectFields as string) || "";
    const includeAttachments = Boolean(options.includeAttachments);
    const messageFilters = (options.messageFilters as IDataObject) || {};
    const bodyToMarkdown = Boolean(options.bodyToMarkdown);

    const notifications = Array.isArray((req.body as IDataObject).value)
      ? ((req.body as IDataObject).value as WebhookNotification[])
      : [];

    const outputItems: INodeExecutionData[] = [];

    for (const notification of notifications) {
      // Validate client state
      if (expectedClientState) {
        if (
          !notification.clientState ||
          notification.clientState !== expectedClientState
        ) {
          res.status(401).send("Invalid client state").end();
          return { noWebhookResponse: true };
        }
      }

      // Handle lifecycle events internally — renew silently, do NOT fire the workflow
      if (notification.lifecycleEvent) {
        if (autoRenew) {
          const ids = (staticData.subscriptionIds as string[]) || [];
          for (const id of ids) {
            try {
              await renewSubscription.call(
                this,
                id,
                GRAPH_SUBSCRIPTION_MINUTES_MAX,
              );
            } catch {
              // Best-effort renewal
            }
          }
        }
        continue;
      }

      let item: IDataObject = {
        type: "notification",
        ...notification,
      };

      const resourcePath =
        notification.resourceData?.["@odata.id"] || notification.resource;
      const cleanResourcePath = (resourcePath || "").replace(/^\/+/, "");
      const isMessage =
        notification.resourceData?.["@odata.type"] ===
        "#Microsoft.Graph.Message";
      const isDeleted = notification.changeType === "deleted";

      if (isMessage && !isDeleted && cleanResourcePath) {
        if (resolveMessageData) {
          const qs: IDataObject = {};
          if (expandProperties) qs.$expand = expandProperties;
          if (selectFields) qs.$select = selectFields;

          const message = await resolveMessageByPath.call(
            this,
            cleanResourcePath,
            Object.keys(qs).length > 0 ? qs : undefined,
          );

          item.message = bodyToMarkdown
            ? convertBodyToMarkdown(message)
            : message;

          if (!evaluateMessageFilter(messageFilters, item)) {
            continue;
          }
        }

        if (includeAttachments) {
          const attachmentsResponse = await graphApiRequest.call(
            this,
            "GET",
            `/v1.0/${cleanResourcePath}/attachments`,
          );

          const rawAttachments: IDataObject[] = Array.isArray(
            attachmentsResponse.value,
          )
            ? attachmentsResponse.value
            : [];

          const binaryData: IBinaryKeyData = {};
          const attachmentsMeta: IDataObject[] = [];
          let binaryIndex = 0;

          for (const attachment of rawAttachments) {
            if (
              attachment["@odata.type"] === "#microsoft.graph.fileAttachment" &&
              attachment.contentBytes
            ) {
              const binaryItem = await this.helpers.prepareBinaryData(
                Buffer.from(attachment.contentBytes as string, "base64"),
                (attachment.name as string) || "attachment",
                (attachment.contentType as string) ||
                  "application/octet-stream",
              );
              binaryData[`data_${binaryIndex++}`] = binaryItem;

              // Strip raw base64 bytes from JSON side
              const { contentBytes: _stripped, ...attachmentMeta } = attachment;
              attachmentsMeta.push(attachmentMeta);
            } else {
              attachmentsMeta.push(attachment);
            }
          }

          item.attachments = attachmentsMeta;

          outputItems.push({
            json: item,
            binary: Object.keys(binaryData).length > 0 ? binaryData : undefined,
          });
          continue;
        }
      }

      outputItems.push({ json: item });
    }

    if (outputItems.length === 0) {
      return { webhookResponse: "OK" };
    }

    return {
      workflowData: [outputItems],
    };
  }
}
