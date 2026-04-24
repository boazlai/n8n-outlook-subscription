# n8n-nodes-outlook-subscription

n8n community nodes for Microsoft Outlook mail automation and change notifications via Microsoft Graph.

## Nodes

### Outlook Subscription (action node)

Manage Microsoft Graph subscriptions and perform full Outlook mail operations.

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

**User** resource:

| Operation | Description                       |
| --------- | --------------------------------- |
| Get       | Fetch a user profile by ID or UPN |

---

### Outlook Subscription Notification Trigger

Receives Microsoft Graph change notifications via webhook.

- Automatically handles Graph validation token handshakes.
- Resolves the full message for `message` notifications (configurable).
- Applies client-side message filters (subject, sender, custom fields).
- Opportunistically renews the subscription on delivery.
- Supports lifecycle notification URL for managed renewal.

---

### Outlook Subscription Lifecycle Trigger

Receives Microsoft Graph lifecycle notifications (e.g. `subscriptionRemoved`, `reauthorizationRequired`) and surfaces them as n8n items for downstream handling.

---

## Credentials

Use the **Microsoft Outlook Subscription OAuth2 API** credential (OAuth 2.0, delegated).

Required Microsoft Graph scopes:

| Scope                  | Required for                      |
| ---------------------- | --------------------------------- |
| `openid`               | Authentication                    |
| `offline_access`       | Token refresh                     |
| `Mail.Read`            | Read messages, subscriptions      |
| `Mail.ReadWrite`       | Update, delete, move, send, reply |
| `MailboxSettings.Read` | Folder operations                 |
| `User.Read.All`        | User Get operation                |

To access a shared mailbox or another user's mailbox, the signed-in account must have the necessary Exchange Online delegate permissions.

---

## Node Architecture

| Use case                                    | Node to use                                                    |
| ------------------------------------------- | -------------------------------------------------------------- |
| n8n receives mail notifications in-workflow | `Outlook Subscription Notification Trigger`                    |
| An external service receives notifications  | `Outlook Subscription` (create operation, supply your own URL) |
| Perform mail operations (send, move, list…) | `Outlook Subscription`                                         |
| Handle subscription lifecycle events        | `Outlook Subscription Lifecycle Trigger`                       |

---

## Notes

- Subscription lifetime is clamped to Microsoft Graph limits (45 min – 7 days).
- Duplicate subscriptions are detected and skipped automatically during Create.
- Message List supports full OData `$filter`, ordering, field selection, and auto-pagination via `@odata.nextLink`.
- Move to new folder: selecting "Create New Folder" creates the folder then moves the message atomically in two Graph calls.
- Send with attachment: enable **Add Attachment** and supply the binary field name from an upstream node.
