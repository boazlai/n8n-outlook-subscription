# n8n-nodes-outlook-subscription

n8n community nodes for Microsoft Outlook mail automation and change notifications via Microsoft Graph.

## Nodes

### Outlook (single node)

Manage Microsoft Graph subscriptions and perform full Outlook mail operations, plus receive Graph webhook events using the Trigger resource.

**Subscription** resource:

| Operation | Description                                                |
| --------- | ---------------------------------------------------------- |
| Create    | Create one or more subscriptions, with duplicate detection |
| Delete    | Delete a subscription by ID                                |
| List      | List all active subscriptions                              |
| Renew     | Extend a subscription's expiry                             |

**Message** resource:

| Operation | Description                                                                      |
| --------- | -------------------------------------------------------------------------------- |
| Delete    | Soft-delete a message (moves to Deleted Items)                                   |
| Get       | Fetch a message with optional `$select` / `$expand`                              |
| List      | Query messages with filter, ordering, field selection, and auto-pagination       |
| Move      | Move a message to a folder — pick from list, enter an ID, or create a new folder |
| Reply     | Reply to a message (HTML or plain text)                                          |
| Reply All | Reply-all to a message (HTML or plain text)                                      |
| Send      | Send a new email with optional CC, BCC, and a binary attachment                  |
| Update    | Set flag status, single-value extended properties, or any arbitrary PATCH fields |

**Attachment** resource:

| Operation | Description                           |
| --------- | ------------------------------------- |
| Download  | Download an attachment as binary data |
| List      | List all attachments on a message     |

**Trigger** resource:

| Operation    | Description                                                                                       |
| ------------ | ------------------------------------------------------------------------------------------------- |
| Notification | Receives Microsoft Graph change notifications via webhook with optional message resolve + filters |
| Lifecycle    | Receives lifecycle notifications (e.g. `subscriptionRemoved`, `reauthorizationRequired`)          |

Trigger capabilities:

- Automatically handles Graph validation token handshakes.
- Optionally validates incoming webhook payloads with client state.
- Optionally resolves full message payloads and attachments for notification events.

---

## Credentials

Use n8n's built-in **Microsoft Outlook OAuth2 API** credential (OAuth 2.0, delegated).

This package does not ship its own Outlook credential anymore. The exact Microsoft Graph scope bundle now comes from the n8n version hosting the node.

To access a shared mailbox for subscription creation, the signed-in account must have the necessary Exchange Online delegate permissions. Other-mailbox subscriptions use the other mailbox email plus a folder ID.

---

## Node Architecture

| Use case                                    | Configuration                                    |
| ------------------------------------------- | ------------------------------------------------ |
| n8n receives mail notifications in-workflow | `Resource = Trigger`, `Operation = Notification` |
| Handle subscription lifecycle events        | `Resource = Trigger`, `Operation = Lifecycle`    |
| An external service receives notifications  | `Resource = Subscription`, `Operation = Create`  |
| Perform mail operations (send, move, list…) | `Resource = Message/Attachment`                  |

---

## Notes

- Subscription lifetime is clamped to Microsoft Graph limits (45 min – 7 days).
- Duplicate subscriptions are detected and skipped automatically during Create.
- Message List supports full OData `$filter`, ordering, field selection, and auto-pagination via `@odata.nextLink`.
- Move to new folder: selecting "Create New Folder" creates the folder then moves the message atomically in two Graph calls.
- Send with attachment: enable **Add Attachment** and supply the binary field name from an upstream node.
