import type { ICredentialType, INodeProperties } from "n8n-workflow";

const scopes = [
  "openid",
  "offline_access",
  "Mail.Read",
  "Mail.ReadWrite",
  "MailboxSettings.Read",
  "User.Read.All",
];

export class MicrosoftOutlookSubscriptionOAuth2Api implements ICredentialType {
  name = "microsoftOutlookSubscriptionOAuth2Api";

  extends = ["microsoftOAuth2Api"];

  displayName = "Microsoft Outlook Subscription OAuth2 API";

  documentationUrl = "microsoft";

  properties: INodeProperties[] = [
    {
      displayName: "Scope",
      name: "scope",
      type: "hidden",
      default: scopes.join(" "),
    },
  ];
}
