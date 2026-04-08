import { config } from "dotenv";

config();

const required = [
  "MicrosoftAppId",
  "MicrosoftAppPassword",
  "AzureAdTenantId",
  "AzureAdClientId",
  "AzureAdClientSecret"
];

for (const key of required) {
  if (!process.env[key]) {
    process.env[key] = "";
  }
}

export const env = {
  port: Number(process.env.PORT ?? 3978),
  microsoftAppId: process.env.MicrosoftAppId ?? "",
  microsoftAppPassword: process.env.MicrosoftAppPassword ?? "",
  microsoftAppTenantId: process.env.MicrosoftAppTenantId ?? "",
  azureAdTenantId: process.env.AzureAdTenantId ?? "",
  azureAdClientId: process.env.AzureAdClientId ?? "",
  azureAdClientSecret: process.env.AzureAdClientSecret ?? "",
  graphScopes: (process.env.GraphScopes ?? "https://graph.microsoft.com/.default")
    .split(",")
    .map((scope) => scope.trim())
    .filter(Boolean),
  apiAudience: process.env.ApiAudience ?? process.env.AzureAdClientId ?? "",
  apiRequiredReadScopes: (
    process.env.ApiRequiredReadScopes ?? "Tasks.Read Tasks.ReadWrite access_as_user"
  )
    .split(/[,\s]+/)
    .map((scope) => scope.trim())
    .filter(Boolean),
  apiRequiredWriteScopes: (
    process.env.ApiRequiredWriteScopes ?? "Tasks.ReadWrite access_as_user"
  )
    .split(/[,\s]+/)
    .map((scope) => scope.trim())
    .filter(Boolean),
  apiRequiredReadRoles: (process.env.ApiRequiredReadRoles ?? "Task.Sync.Read Task.Sync.ReadWrite")
    .split(/[,\s]+/)
    .map((role) => role.trim())
    .filter(Boolean),
  apiRequiredWriteRoles: (process.env.ApiRequiredWriteRoles ?? "Task.Sync.ReadWrite")
    .split(/[,\s]+/)
    .map((role) => role.trim())
    .filter(Boolean),
  taskStateFilePath: process.env.TaskStateFilePath ?? "data/task-state.json",
  conversationReferenceFilePath:
    process.env.ConversationReferenceFilePath ?? "data/conversation-references.json",
  nodeEnv: process.env.NODE_ENV ?? "development"
};
