import type {
  IDataObject,
  IHookFunctions,
  INodeExecutionData,
  INodeType,
  INodeTypeDescription,
  IWebhookFunctions,
  IWebhookResponseData,
  NodeConnectionType,
} from "n8n-workflow";

import {
  evaluateMessageFilter,
  extractUserIdFromResource,
  getUser,
  graphApiRequest,
  resolveMessageByPath,
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
    displayName: "Outlook Subscription Notification Trigger",
    name: "outlookSubscriptionNotificationTrigger",
    icon: "file:outlookSubscriptionNotificationTrigger.svg",
    group: ["trigger"],
    version: 1,
    subtitle: "Outlook Notification Listener",
    description:
      "Receive Microsoft Graph Outlook change notifications via webhook. Create subscriptions using the Outlook Subscription node.",
    defaults: {
      name: "Outlook Subscription Notification Trigger",
    },
    inputs: [],
    outputs: ["main" as NodeConnectionType],
    credentials: [
      {
        name: "microsoftOutlookSubscriptionOAuth2Api",
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
      {
        displayName: "Webhook Path",
        name: "webhookPath",
        type: "string",
        default: "outlook-subscription",
        description:
          "URL path for this webhook. Copy the full Webhook URL from above and paste it as the notificationUrl when creating your Graph subscription.",
      },
      {
        displayName: "Client State",
        name: "clientState",
        type: "string",
        default: "",
        description:
          "If set, validates incoming notifications against this secret. Must match the value used when creating the subscription.",
      },
      {
        displayName: "Resolve Full Message",
        name: "resolveMessageData",
        type: "boolean",
        default: true,
        description:
          "Whether to fetch the full message content for message notifications (skipped for deleted messages)",
      },
      {
        displayName: "Message Filters",
        name: "messageFilters",
        type: "filter",
        default: {},
        displayOptions: {
          show: { resolveMessageData: [true] },
        },
        typeOptions: {
          filter: {
            version: 2,
            caseSensitive: false,
            typeValidation: "loose",
          },
        },
        description:
          "Filter messages by field values. Enter a dot-notation field path as the left value (e.g. subject, from.emailAddress.address, changeType, hasAttachments).",
      },
      {
        displayName: "Options",
        name: "options",
        type: "collection",
        placeholder: "Add Option",
        default: {},
        options: [
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
            displayName: "Select Fields",
            name: "selectFields",
            type: "string",
            default: "",
            placeholder: "id,subject,body,from,hasAttachments",
            description: "$select query parameter when resolving the message",
          },
          {
            displayName: "Resolve User Email",
            name: "resolveUserEmail",
            type: "boolean",
            default: false,
            description:
              "Whether to resolve the user ID from the notification to an email address",
          },
          {
            displayName: "Include Attachments",
            name: "includeAttachments",
            type: "boolean",
            default: false,
            description:
              "Whether to fetch the attachment list alongside the resolved message",
          },
        ],
      },
    ],
  };

  webhookMethods = {
    default: {
      async checkExists(this: IHookFunctions): Promise<boolean> {
        return true;
      },
      async create(this: IHookFunctions): Promise<boolean> {
        return true;
      },
      async delete(this: IHookFunctions): Promise<boolean> {
        return true;
      },
    },
  };

  async webhook(this: IWebhookFunctions): Promise<IWebhookResponseData> {
    const req = this.getRequestObject();
    const res = this.getResponseObject();

    if (req.query.validationToken) {
      res
        .status(200)
        .type("text/plain")
        .send(String(req.query.validationToken));
      return { noWebhookResponse: true };
    }

    const expectedClientState = (
      this.getNodeParameter("clientState", "") as string
    ).trim();

    const notifications = Array.isArray((req.body as IDataObject).value)
      ? ((req.body as IDataObject).value as WebhookNotification[])
      : [];

    const resolveMessageData = this.getNodeParameter(
      "resolveMessageData",
      true,
    ) as boolean;
    const messageFilters = this.getNodeParameter(
      "messageFilters",
      {},
    ) as IDataObject;
    const options = this.getNodeParameter("options", {}) as IDataObject;
    const expandProperties = (options.expandProperties as string) || "";
    const selectFields = (options.selectFields as string) || "";
    const resolveUserEmail = Boolean(options.resolveUserEmail);
    const includeAttachments = Boolean(options.includeAttachments);

    const payloads: IDataObject[] = [];

    for (const notification of notifications) {
      if (expectedClientState) {
        if (
          !notification.clientState ||
          notification.clientState !== expectedClientState
        ) {
          res.status(401).send("Invalid client state").end();
          return { noWebhookResponse: true };
        }
      }

      // Skip lifecycle events — handled by the Outlook Subscription Lifecycle Trigger node
      if (notification.lifecycleEvent) {
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

          item.message = message;

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
          item.attachments = Array.isArray(attachmentsResponse.value)
            ? attachmentsResponse.value
            : [];
        }

        if (resolveUserEmail) {
          const userId = extractUserIdFromResource(cleanResourcePath);
          if (userId) {
            const user = await getUser.call(this, userId, "mail,displayName");
            item.user = user;
          }
        }
      }

      payloads.push(item);
    }

    if (payloads.length === 0) {
      return { webhookResponse: "OK" };
    }

    return {
      workflowData: payloads.map((payload) => [
        { json: payload } as INodeExecutionData,
      ]),
    };
  }
}
