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

type LifecycleNotification = {
  lifecycleEvent?: string;
  clientState?: string;
  subscriptionId?: string;
  tenantId?: string;
  organizationId?: string;
  expirationDateTime?: string;
};

export class OutlookSubscriptionLifecycleTrigger implements INodeType {
  description: INodeTypeDescription = {
    displayName: "Outlook Subscription Lifecycle Trigger",
    name: "outlookSubscriptionLifecycleTrigger",
    icon: "file:outlookSubscriptionLifecycleTrigger.svg",
    group: ["trigger"],
    version: 1,
    description:
      "Receive Microsoft Graph subscription lifecycle notifications (reauthorizationRequired, subscriptionRemoved, missed). Copy the Webhook URL and use it as lifecycleNotificationUrl when creating your Graph subscription.",
    defaults: {
      name: "Outlook Subscription Lifecycle Trigger",
    },
    inputs: [],
    outputs: ["main" as NodeConnectionType],
    credentials: [],
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
        default: "outlook-lifecycle",
        description:
          "URL path for this webhook. Copy the full Webhook URL from above and paste it as the lifecycleNotificationUrl when creating your Graph subscription.",
      },
      {
        displayName: "Client State",
        name: "clientState",
        type: "string",
        default: "",
        description:
          "The client state secret set when creating your Graph subscription. Used to verify that incoming lifecycle notifications are genuine.",
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

    const notifications = Array.isArray((req.body as IDataObject).value)
      ? ((req.body as IDataObject).value as LifecycleNotification[])
      : [];

    const expectedClientState = (
      this.getNodeParameter("clientState", "") as string
    ).trim();

    const payloads: IDataObject[] = [];

    for (const notification of notifications) {
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
}
