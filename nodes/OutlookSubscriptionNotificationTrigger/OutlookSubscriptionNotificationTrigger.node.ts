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

type LifecycleNotification = {
  lifecycleEvent?: string;
  clientState?: string;
  subscriptionId?: string;
  tenantId?: string;
  organizationId?: string;
  expirationDateTime?: string;
};

export class OutlookSubscriptionNotificationTrigger implements INodeType {
  description: INodeTypeDescription = {
    displayName: "Outlook Trigger",
    name: "outlookSubscriptionTrigger",
    icon: "file:outlookSubscriptionNotificationTrigger.svg",
    group: ["trigger"],
    version: 1,
    subtitle:
      '={{$parameter["event"] === "lifecycle" ? "Trigger: Lifecycle" : "Trigger: Notification"}}',
    description:
      "Receive Microsoft Graph Outlook notification and lifecycle webhook events.",
    defaults: {
      name: "Outlook Trigger",
    },
    inputs: [],
    outputs: ["main" as NodeConnectionType],
    credentials: [
      {
        name: "microsoftOutlookOAuth2Api",
        required: false,
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
        displayName: "Trigger On",
        name: "event",
        type: "options",
        noDataExpression: true,
        default: "notification",
        options: [
          { name: "Lifecycle", value: "lifecycle", action: "Lifecycle" },
          {
            name: "Notification",
            value: "notification",
            action: "Notification",
          },
        ],
      },
      {
        displayName: "Webhook Path",
        name: "webhookPath",
        type: "string",
        default: "outlook-subscription",
        description:
          "URL path for this webhook. Use this URL as notificationUrl (Notification operation) or lifecycleNotificationUrl (Lifecycle operation) when creating subscriptions.",
      },
      {
        displayName: "Client State",
        name: "clientState",
        type: "string",
        default: "",
        description:
          "If set, validates incoming webhook events against this secret. Must match the value used when creating the subscription.",
      },
      {
        displayName: "Resolve Full Message",
        name: "resolveMessageData",
        type: "boolean",
        default: true,
        displayOptions: {
          show: { event: ["notification"] },
        },
        description:
          "Whether to fetch the full message content for message notifications (skipped for deleted messages)",
      },
      {
        displayName: "Message Filters",
        name: "messageFilters",
        type: "filter",
        default: {},
        displayOptions: {
          show: { event: ["notification"], resolveMessageData: [true] },
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
        displayOptions: {
          show: { event: ["notification"] },
        },
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
    const event = this.getNodeParameter("event", "notification") as
      | "notification"
      | "lifecycle";

    const notifications = Array.isArray((req.body as IDataObject).value)
      ? ((req.body as IDataObject).value as WebhookNotification[])
      : [];

    if (event === "lifecycle") {
      const payloads: IDataObject[] = [];

      for (const notification of notifications as LifecycleNotification[]) {
        if (!notification.lifecycleEvent) {
          continue;
        }

        if (expectedClientState) {
          if (
            !notification.clientState ||
            notification.clientState !== expectedClientState
          ) {
            res.status(401).send("Invalid client state").end();
            return { noWebhookResponse: true };
          }
        }

        payloads.push({ ...notification });
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
