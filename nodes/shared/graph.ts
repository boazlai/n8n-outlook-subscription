import type {
  FilterValue,
  IDataObject,
  IExecuteFunctions,
  IHookFunctions,
  ILoadOptionsFunctions,
  INodeExecutionData,
  INodePropertyOptions,
  IWebhookFunctions,
  JsonObject,
} from "n8n-workflow";
import { NodeApiError, executeFilter } from "n8n-workflow";

export const GRAPH_SUBSCRIPTION_MINUTES_MIN = 45;
export const GRAPH_SUBSCRIPTION_MINUTES_MAX = 10080;

export type NodeContext =
  | IExecuteFunctions
  | IHookFunctions
  | ILoadOptionsFunctions
  | IWebhookFunctions
  | {
      helpers: {
        httpRequestWithAuthentication: (
          credentialType: string,
          options: IDataObject,
        ) => Promise<any>;
      };
      getCredentials: (type: string) => Promise<IDataObject>;
    };

export interface GraphMailFolder extends IDataObject {
  id: string;
  displayName: string;
  childFolderCount?: number;
  parentFolderId?: string;
  path?: string;
  children?: GraphMailFolder[];
}

export interface GraphSubscription extends IDataObject {
  id: string;
  resource: string;
  changeType: string;
  notificationUrl: string;
  lifecycleNotificationUrl?: string;
  clientState?: string;
  expirationDateTime?: string;
  applicationId?: string;
}

export interface MailboxConfig {
  mailboxMode: "current";
}

export interface SubscriptionTargetConfig {
  mailboxMode: "current" | "other";
  otherMailboxEmail?: string;
  entity: "message" | "folder";
  folderId?: string;
  includeSubfolders?: boolean;
}

export interface CreateSubscriptionPayload {
  changeType: string;
  notificationUrl: string;
  lifecycleNotificationUrl?: string;
  clientState: string;
  expirationDateTime: string;
  resource: string;
  latestSupportedTlsVersion?: string;
  includeResourceData?: boolean;
  resourceData?: IDataObject;
  encryptionCertificate?: string;
  encryptionCertificateId?: string;
}

function ensureLeadingSlash(path: string): string {
  return path.startsWith("/") ? path : `/${path}`;
}

export async function graphApiRequest(
  this: NodeContext,
  method: string,
  endpoint: string,
  body?: IDataObject,
  qs?: IDataObject,
  headers?: IDataObject,
): Promise<any> {
  const credentials = await this.getCredentials("microsoftOutlookOAuth2Api");
  const baseUrl = String(
    credentials.graphApiBaseUrl || "https://graph.microsoft.com",
  );

  const options: IDataObject = {
    method,
    baseURL: baseUrl,
    url: ensureLeadingSlash(endpoint),
    json: true,
  };

  if (body && Object.keys(body).length > 0) {
    options.body = body;
  }

  if (qs && Object.keys(qs).length > 0) {
    options.qs = qs;
  }

  if (headers && Object.keys(headers).length > 0) {
    options.headers = headers;
  }

  try {
    return await this.helpers.httpRequestWithAuthentication.call(
      this,
      "microsoftOutlookOAuth2Api",
      options,
    );
  } catch (error) {
    throw new NodeApiError(this as never, error as JsonObject);
  }
}

export function getMailboxBasePath(_config: MailboxConfig): string {
  return "/me";
}

export function clampLifetimeMinutes(requestedMinutes: number): number {
  if (!Number.isFinite(requestedMinutes)) {
    return GRAPH_SUBSCRIPTION_MINUTES_MIN;
  }

  return Math.min(
    GRAPH_SUBSCRIPTION_MINUTES_MAX,
    Math.max(GRAPH_SUBSCRIPTION_MINUTES_MIN, Math.floor(requestedMinutes)),
  );
}

export function buildExpirationDateTime(lifetimeMinutes: number): string {
  const safeLifetime = clampLifetimeMinutes(lifetimeMinutes);
  return new Date(Date.now() + safeLifetime * 60 * 1000).toISOString();
}

export function normalizeChangeTypes(changeTypes: string[]): string {
  return Array.from(new Set(changeTypes.filter(Boolean))).join(",");
}

export function buildClientState(input?: string): string {
  if (input && input.trim()) {
    return input.trim();
  }

  return `n8n-outlook-${Math.random().toString(36).slice(2, 12)}`;
}

export async function listSubscriptions(
  this: NodeContext,
): Promise<GraphSubscription[]> {
  const response = await graphApiRequest.call(
    this,
    "GET",
    "/v1.0/subscriptions",
  );
  return Array.isArray(response.value)
    ? (response.value as GraphSubscription[])
    : [];
}

export async function deleteSubscription(
  this: NodeContext,
  subscriptionId: string,
): Promise<void> {
  await graphApiRequest.call(
    this,
    "DELETE",
    `/v1.0/subscriptions/${subscriptionId}`,
  );
}

export async function renewSubscription(
  this: NodeContext,
  subscriptionId: string,
  lifetimeMinutes: number,
): Promise<GraphSubscription> {
  return await graphApiRequest.call(
    this,
    "PATCH",
    `/v1.0/subscriptions/${subscriptionId}`,
    {
      expirationDateTime: buildExpirationDateTime(lifetimeMinutes),
    },
  );
}

export async function createSubscription(
  this: NodeContext,
  payload: CreateSubscriptionPayload,
): Promise<GraphSubscription> {
  return await graphApiRequest.call(
    this,
    "POST",
    "/v1.0/subscriptions",
    payload,
  );
}

export async function getFolderTree(
  this: NodeContext,
  config: MailboxConfig,
  parentFolderId?: string,
  pathPrefix = "",
): Promise<GraphMailFolder[]> {
  const mailboxBase = getMailboxBasePath(config);
  const endpoint = parentFolderId
    ? `${mailboxBase}/mailFolders/${encodeURIComponent(parentFolderId)}/childFolders`
    : `${mailboxBase}/mailFolders`;

  const response = await graphApiRequest.call(
    this,
    "GET",
    `/v1.0${endpoint}`,
    undefined,
    {
      $top: 100,
      $select: "id,displayName,parentFolderId,childFolderCount",
    },
  );

  const folders = Array.isArray(response.value)
    ? (response.value as GraphMailFolder[])
    : [];
  const result: GraphMailFolder[] = [];

  for (const folder of folders) {
    const currentPath = pathPrefix
      ? `${pathPrefix} / ${folder.displayName}`
      : String(folder.displayName);
    const normalizedFolder: GraphMailFolder = {
      ...folder,
      path: currentPath,
    };

    if ((folder.childFolderCount || 0) > 0) {
      normalizedFolder.children = await getFolderTree.call(
        this,
        config,
        folder.id,
        currentPath,
      );
    }

    result.push(normalizedFolder);
  }

  return result;
}

export function flattenFolders(folders: GraphMailFolder[]): GraphMailFolder[] {
  const flattened: GraphMailFolder[] = [];

  for (const folder of folders) {
    flattened.push(folder);
    if (Array.isArray(folder.children) && folder.children.length > 0) {
      flattened.push(...flattenFolders(folder.children));
    }
  }

  return flattened;
}

export async function loadFolderOptions(
  this: ILoadOptionsFunctions,
): Promise<INodePropertyOptions[]> {
  const config: MailboxConfig = { mailboxMode: "current" };
  const folders = flattenFolders(await getFolderTree.call(this, config));

  return folders.map((folder) => ({
    name: folder.path || folder.displayName,
    value: folder.id,
  }));
}

function buildOtherMailboxSubscriptionTarget(
  otherMailboxEmail: string,
  rawFolderId: string,
  entity: SubscriptionTargetConfig["entity"],
): string {
  const trimmedMailboxEmail = otherMailboxEmail.trim();
  const trimmedFolderId = rawFolderId.trim();

  if (!trimmedMailboxEmail) {
    throw new Error(
      "Other mailbox email is required when mailbox is set to other",
    );
  }

  if (!trimmedFolderId) {
    throw new Error("Folder ID is required when mailbox is set to other");
  }

  const folderPath = `/users/${encodeURIComponent(trimmedMailboxEmail)}/mailFolders/${encodeURIComponent(trimmedFolderId)}`;

  if (entity === "message") {
    return `${folderPath}/messages`;
  }

  return folderPath;
}

export async function buildSubscriptionTargets(
  this: NodeContext,
  config: SubscriptionTargetConfig,
): Promise<string[]> {
  if (config.mailboxMode === "other") {
    return [
      buildOtherMailboxSubscriptionTarget(
        config.otherMailboxEmail || "",
        config.folderId || "",
        config.entity,
      ),
    ];
  }

  const mailboxBase = getMailboxBasePath({ mailboxMode: "current" });
  const selectedFolderId = config.folderId?.trim();

  if (!selectedFolderId) {
    if (config.entity === "message") {
      return [`${mailboxBase}/messages`];
    }

    return [`${mailboxBase}/mailFolders`];
  }

  if (!config.includeSubfolders) {
    return [
      config.entity === "message"
        ? `${mailboxBase}/mailFolders/${selectedFolderId}/messages`
        : `${mailboxBase}/mailFolders/${selectedFolderId}`,
    ];
  }

  const tree = await getFolderTree.call(this, config, selectedFolderId, "");
  const descendants = flattenFolders(tree);
  const folderIds = [
    selectedFolderId,
    ...descendants.map((folder) => folder.id),
  ];

  return folderIds.map((folderId) =>
    config.entity === "message"
      ? `${mailboxBase}/mailFolders/${folderId}/messages`
      : `${mailboxBase}/mailFolders/${folderId}`,
  );
}

export function findDuplicateSubscriptions(
  existingSubscriptions: GraphSubscription[],
  resource: string,
  changeType: string,
  notificationUrl: string,
): GraphSubscription[] {
  return existingSubscriptions.filter(
    (subscription) =>
      subscription.resource === resource &&
      subscription.changeType === changeType &&
      subscription.notificationUrl === notificationUrl,
  );
}

export function shouldRenew(
  expiresAt: string | undefined,
  thresholdMinutes = 15,
): boolean {
  if (!expiresAt) {
    return true;
  }

  const expiresAtTime = new Date(expiresAt).getTime();
  if (Number.isNaN(expiresAtTime)) {
    return true;
  }

  return expiresAtTime - Date.now() <= thresholdMinutes * 60 * 1000;
}

export async function resolveMessage(
  this: NodeContext,
  config: MailboxConfig,
  messageId: string,
): Promise<IDataObject> {
  const mailboxBase = getMailboxBasePath(config);
  return await graphApiRequest.call(
    this,
    "GET",
    `/v1.0${mailboxBase}/messages/${encodeURIComponent(messageId)}`,
  );
}

function resolveFieldPath(obj: IDataObject, path: string): unknown {
  return path.split(".").reduce((current: unknown, key) => {
    if (current === null || current === undefined) return undefined;
    return (current as IDataObject)[key];
  }, obj);
}

export function evaluateMessageFilter(
  filter: FilterValue | IDataObject,
  item: IDataObject,
): boolean {
  const f = filter as FilterValue;
  if (!f?.conditions?.length) return true;

  const resolvedFilter: FilterValue = {
    ...f,
    conditions: f.conditions.map((condition) => {
      const lv = condition.leftValue;
      // If leftValue is a plain string with no expression syntax,
      // treat it as a dot-notation path into the item object.
      if (typeof lv === "string" && lv !== "" && !lv.includes("{{")) {
        return {
          ...condition,
          leftValue: resolveFieldPath(item, lv) as typeof lv,
        };
      }
      return condition;
    }),
  };

  return executeFilter(resolvedFilter);
}

export function toExecutionItems(items: IDataObject[]): INodeExecutionData[] {
  return items.map((item) => ({ json: item }));
}

export async function getMessage(
  this: NodeContext,
  config: MailboxConfig,
  messageId: string,
  qs?: IDataObject,
): Promise<IDataObject> {
  const mailboxBase = getMailboxBasePath(config);
  return await graphApiRequest.call(
    this,
    "GET",
    `/v1.0${mailboxBase}/messages/${encodeURIComponent(messageId)}`,
    undefined,
    qs,
  );
}

export async function updateMessage(
  this: NodeContext,
  config: MailboxConfig,
  messageId: string,
  body: IDataObject,
  etag?: string,
): Promise<IDataObject> {
  const mailboxBase = getMailboxBasePath(config);
  return await graphApiRequest.call(
    this,
    "PATCH",
    `/v1.0${mailboxBase}/messages/${encodeURIComponent(messageId)}`,
    body,
    undefined,
    etag ? { "If-Match": etag } : undefined,
  );
}

export async function listMessageAttachments(
  this: NodeContext,
  config: MailboxConfig,
  messageId: string,
): Promise<IDataObject[]> {
  const mailboxBase = getMailboxBasePath(config);
  const response = await graphApiRequest.call(
    this,
    "GET",
    `/v1.0${mailboxBase}/messages/${encodeURIComponent(messageId)}/attachments`,
  );
  return Array.isArray(response.value) ? (response.value as IDataObject[]) : [];
}

export async function getAttachment(
  this: NodeContext,
  config: MailboxConfig,
  messageId: string,
  attachmentId: string,
): Promise<IDataObject> {
  const mailboxBase = getMailboxBasePath(config);
  return await graphApiRequest.call(
    this,
    "GET",
    `/v1.0${mailboxBase}/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`,
  );
}

export async function resolveMessageByPath(
  this: NodeContext,
  resourcePath: string,
  qs?: IDataObject,
): Promise<IDataObject> {
  const normalizedPath = resourcePath.startsWith("/")
    ? resourcePath
    : `/${resourcePath}`;
  return await graphApiRequest.call(
    this,
    "GET",
    `/v1.0${normalizedPath}`,
    undefined,
    qs,
  );
}

// ── New Tier-1 Message Helpers ─────────────────────────────────────────────

export async function listMessages(
  this: NodeContext,
  config: MailboxConfig,
  folderId: string | undefined,
  qs: IDataObject,
  returnAll: boolean,
): Promise<IDataObject[]> {
  const mailboxBase = getMailboxBasePath(config);
  const basePath = folderId
    ? `${mailboxBase}/mailFolders/${encodeURIComponent(folderId)}/messages`
    : `${mailboxBase}/messages`;

  const results: IDataObject[] = [];
  let endpoint: string | undefined = `/v1.0${basePath}`;
  let firstCall = true;

  while (endpoint) {
    const callQs = firstCall ? qs : undefined;
    firstCall = false;

    const response = await graphApiRequest.call(
      this,
      "GET",
      endpoint,
      undefined,
      callQs,
    );

    const page = Array.isArray(response.value)
      ? (response.value as IDataObject[])
      : [];
    results.push(...page);

    const nextLink = response["@odata.nextLink"] as string | undefined;
    if (returnAll && nextLink) {
      try {
        const parsed = new URL(nextLink);
        endpoint = parsed.pathname + parsed.search;
      } catch {
        endpoint = undefined;
      }
    } else {
      endpoint = undefined;
    }
  }

  return results;
}

export async function moveMessage(
  this: NodeContext,
  config: MailboxConfig,
  messageId: string,
  destinationId: string,
): Promise<IDataObject> {
  const mailboxBase = getMailboxBasePath(config);
  return await graphApiRequest.call(
    this,
    "POST",
    `/v1.0${mailboxBase}/messages/${encodeURIComponent(messageId)}/move`,
    { destinationId },
  );
}

export async function createMailFolder(
  this: NodeContext,
  config: MailboxConfig,
  displayName: string,
  parentFolderId?: string,
): Promise<GraphMailFolder> {
  const mailboxBase = getMailboxBasePath(config);
  const endpoint = parentFolderId
    ? `/v1.0${mailboxBase}/mailFolders/${encodeURIComponent(parentFolderId)}/childFolders`
    : `/v1.0${mailboxBase}/mailFolders`;
  return await graphApiRequest.call(this, "POST", endpoint, { displayName });
}

export async function deleteMessage(
  this: NodeContext,
  config: MailboxConfig,
  messageId: string,
): Promise<void> {
  const mailboxBase = getMailboxBasePath(config);
  await graphApiRequest.call(
    this,
    "DELETE",
    `/v1.0${mailboxBase}/messages/${encodeURIComponent(messageId)}`,
  );
}

function buildRecipientList(raw: string): IDataObject[] {
  return raw
    .split(",")
    .map((addr) => addr.trim())
    .filter(Boolean)
    .map((address) => ({ emailAddress: { address } }));
}

export interface SendMailPayload {
  to: string;
  subject: string;
  body: string;
  bodyType: "html" | "text";
  cc?: string;
  bcc?: string;
  saveToSentItems?: boolean;
  attachments?: Array<{
    name: string;
    contentType: string;
    contentBytes: string;
  }>;
  inlineImages?: Array<{
    name: string;
    contentType: string;
    contentBytes: string;
    contentId: string;
  }>;
}

export async function sendMail(
  this: NodeContext,
  config: MailboxConfig,
  payload: SendMailPayload,
): Promise<void> {
  const mailboxBase = getMailboxBasePath(config);
  const requestBody: IDataObject = {
    message: {
      subject: payload.subject,
      body: {
        contentType: payload.bodyType === "html" ? "HTML" : "Text",
        content: payload.body,
      },
      toRecipients: buildRecipientList(payload.to),
      ...(payload.cc ? { ccRecipients: buildRecipientList(payload.cc) } : {}),
      ...(payload.bcc
        ? { bccRecipients: buildRecipientList(payload.bcc) }
        : {}),
      ...(() => {
        const allAttachments: IDataObject[] = [];
        if (payload.attachments && payload.attachments.length > 0) {
          allAttachments.push(
            ...payload.attachments.map((a) => ({
              "@odata.type": "#microsoft.graph.fileAttachment",
              name: a.name,
              contentType: a.contentType,
              contentBytes: a.contentBytes,
            })),
          );
        }
        if (payload.inlineImages && payload.inlineImages.length > 0) {
          allAttachments.push(
            ...payload.inlineImages.map((img) => ({
              "@odata.type": "#microsoft.graph.fileAttachment",
              name: img.name,
              contentType: img.contentType,
              contentBytes: img.contentBytes,
              contentId: img.contentId,
              isInline: true,
            })),
          );
        }
        return allAttachments.length > 0 ? { attachments: allAttachments } : {};
      })(),
    },
    saveToSentItems: payload.saveToSentItems !== false,
  };

  await graphApiRequest.call(
    this,
    "POST",
    `/v1.0${mailboxBase}/sendMail`,
    requestBody,
  );
}

export async function replyToMessage(
  this: NodeContext,
  config: MailboxConfig,
  messageId: string,
  comment: string,
  bodyType: "html" | "text",
  replyAll: boolean,
): Promise<void> {
  const mailboxBase = getMailboxBasePath(config);
  const action = replyAll ? "replyAll" : "reply";

  const requestBody: IDataObject =
    bodyType === "html"
      ? {
          message: {
            body: { contentType: "HTML", content: comment },
          },
        }
      : { comment };

  await graphApiRequest.call(
    this,
    "POST",
    `/v1.0${mailboxBase}/messages/${encodeURIComponent(messageId)}/${action}`,
    requestBody,
  );
}
