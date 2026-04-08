# Teams Task Manager Bot

Monorepo for a Microsoft Teams task manager bot that coordinates Planner and To Do actions through Microsoft Graph, Azure AD, and a React configuration tab.

## Workspace layout

- `apps/bot`: Azure Bot Service compatible backend in TypeScript.
- `apps/config-tab`: React tab for workspace configuration and service status.

## Planned increments

1. Base bot and React workspace scaffolding.
2. Microsoft Graph auth and task orchestration services.
3. Adaptive card task flows for create, assign, and update.
4. React configuration tab wired to the backend APIs.

## Development

```bash
npm install
npm run build
npm test
```

## Environment

Backend configuration is driven by these variables:

- `MicrosoftAppId`
- `MicrosoftAppPassword`
- `MicrosoftAppTenantId`
- `AzureAdTenantId`
- `AzureAdClientId`
- `AzureAdClientSecret`
- `GraphScopes`
- `ApiAudience`
- `ApiRequiredReadScopes`
- `ApiRequiredWriteScopes`
- `ApiRequiredReadRoles`
- `ApiRequiredWriteRoles`
- `TaskStateFilePath`
- `ConversationReferenceFilePath`

Task commands also expect request payloads to include a `planId`, `bucketId`, and assignee information so Planner and To Do operations can be synchronized.

Frontend configuration:

- `VITE_API_BASE_URL`: Base URL for the bot backend, for example `http://localhost:3978`.

The synchronized task mapping is persisted to disk. By default the bot writes to `data/task-state.json`, and you can override that path with `TaskStateFilePath`.

The task REST API now requires an Azure AD bearer token. The backend validates issuer and audience, then authorizes requests using configured scopes or app roles.

Teams conversation references are also persisted to disk so the bot can send proactive channel notifications after task state changes. The default file is `data/conversation-references.json`.
