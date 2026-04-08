# AGENTS.md

This file is the repository handoff and operating context for Codex agents, terminal-based contributors, and other automated coding workflows.

## Project Identity

- Name: `teams-task-manager-bot`
- Stack: TypeScript, Microsoft Graph API, Azure Bot Service patterns, React, Azure Active Directory, REST APIs
- Repo shape: npm workspaces monorepo
- Primary goal: let Microsoft Teams users create, assign, and update Planner tasks from Teams, while synchronizing task state into Microsoft To Do and exposing a React configuration tab for service status and task routing

## Current Repository Structure

- `package.json`
  Root npm workspace config and shared scripts
- `tsconfig.base.json`
  Shared TypeScript compiler settings
- `apps/bot`
  Backend bot service and task sync API
- `apps/config-tab`
  React configuration dashboard for the Teams tab

## Backend Scope

The backend in `apps/bot` currently contains these responsibilities:

- Azure Bot Service compatible Express host
- Teams bot message handling
- Adaptive card generation and submit handling
- Azure AD bearer-token validation for REST access
- Microsoft Graph client creation using Azure AD app credentials or on-behalf-of flow
- Planner task create, update, and assign operations
- Planner task complete and delete operations
- To Do task create, update, and delete operations
- File-backed synchronized task state tracking
- File-backed Teams conversation reference tracking for proactive notifications
- REST API for listing, creating, updating, and reassigning synchronized tasks

Important files:

- `apps/bot/src/index.ts`
  App composition root
- `apps/bot/src/bot/TaskManagerBot.ts`
  Teams command handling and adaptive card submit handling
- `apps/bot/src/bot/conversationReferenceStore.ts`
  Persistent Teams conversation references for proactive notifications
- `apps/bot/src/cards/taskCards.ts`
  Adaptive card payload builders
- `apps/bot/src/api/taskRoutes.ts`
  REST routes for task orchestration
- `apps/bot/src/graph/graphClientFactory.ts`
  Azure AD / Graph token acquisition
- `apps/bot/src/graph/plannerService.ts`
  Planner integration
- `apps/bot/src/graph/todoService.ts`
  To Do integration
- `apps/bot/src/services/taskSyncService.ts`
  Cross-service orchestration and conflict retry behavior
- `apps/bot/src/services/taskStateStore.ts`
  Current file-backed state store
- `apps/bot/src/services/taskNotificationService.ts`
  Proactive Teams notification dispatch

## Frontend Scope

The frontend in `apps/config-tab` currently contains these responsibilities:

- Dashboard view for backend health and synchronized task count
- Form for creating synchronized tasks against the backend API
- UI that lists synchronized tasks returned by the backend
- Backend API client helpers

Important files:

- `apps/config-tab/src/App.tsx`
  Main dashboard UI and create-task flow
- `apps/config-tab/src/api.ts`
  Fetch wrapper for backend routes
- `apps/config-tab/src/types.ts`
  Shared frontend task and health types
- `apps/config-tab/src/styles.css`
  Dashboard styling

## User-Facing Behavior Implemented

Bot commands supported right now:

- `task create`
  Opens an adaptive card for synchronized task creation
- `task update <plannerTaskId>`
  Opens an adaptive card for updating a synchronized task
- `task assign <plannerTaskId>`
  Opens an adaptive card for reassigning a synchronized task
- `task complete <plannerTaskId>`
  Completes the synchronized task in Planner and To Do
- `task delete <plannerTaskId>`
  Deletes the synchronized task from Planner and To Do
- `task list`
  Posts summary cards for the tasks currently stored on disk-backed state

Proactive behavior supported right now:

- REST-driven create, update, and assign operations can post summary cards back into a stored Teams channel reference
- Interactive bot card submissions suppress the proactive send because they already reply in-turn

Lifecycle behavior supported right now:

- Completing a task updates both Planner and To Do to completed state
- Deleting a task removes both the Planner task and the linked To Do task
- Reassigning a task creates the replacement To Do task and removes the old assignee's To Do task

REST endpoints supported right now:

- `GET /health`
- `GET /api/tasks`
- `POST /api/tasks`
- `PATCH /api/tasks/:plannerTaskId`
- `POST /api/tasks/:plannerTaskId/assign`
- `POST /api/tasks/:plannerTaskId/complete`
- `DELETE /api/tasks/:plannerTaskId`

## Runtime and Environment

Backend environment variables:

- `PORT`
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

Frontend environment variables:

- `VITE_API_BASE_URL`

Notes:

- `GraphScopes` defaults to `https://graph.microsoft.com/.default`
- `ApiAudience` defaults to `AzureAdClientId`
- The config tab defaults its backend target to `http://localhost:3978`
- The create-task form converts browser-local datetime input into ISO-8601 before sending it to the backend
- The synchronized task mapping defaults to `data/task-state.json`
- The conversation reference store defaults to `data/conversation-references.json`

## Build and Verification

Standard commands:

```bash
npm install
npm run build
npm test
```

Workspace-specific commands:

```bash
npm run dev:bot
npm run dev:tab
```

Current known-good verification:

- `npm install` completed successfully
- `npm run build` completed successfully for both workspaces
- `npm test` runs the bot workspace test suite with Vitest

## Architectural Constraints and Assumptions

- The current synchronized task store is file-backed JSON by default.
  Restarting the backend preserves tracked task mappings between Planner and To Do as long as `TaskStateFilePath` points to persistent storage.
- The task REST API now expects valid Azure AD bearer tokens and enforces read/write access by scopes or app roles.
- The bot can proactively notify a Teams channel only after a conversation reference has been captured for that team/channel.
- Reassignment now actively removes the old To Do item to reduce duplicate tasks.
- Planner concurrency is handled with a retry path on version conflict via ETag refresh.
- The backend is structured to support on-behalf-of Graph access, but real deployment still requires Azure AD app registration, permissions, and Teams/Bot configuration.
- The React tab is a functional control panel, not yet a fully Teams SDK-aware production tab experience.

## Known Gaps / Likely Next Work

Highest-value next steps:

- Add Microsoft Teams SDK integration inside the React tab
- Add deployment assets such as Teams app manifest, Azure deployment config, and CI pipeline

## Git Context

Recent local commits at the time this file was added:

- `3ea1179` `chore: scaffold teams task manager workspace`
- `a574910` `feat: add graph task synchronization services`
- `db0f957` `feat: add teams adaptive card task workflows`
- `e39835d` `feat: add react configuration dashboard`

Remote context:

- Intended origin: `https://github.com/akurtle/teams-task-manager-bot.git`
- Push attempts have failed with `Repository not found`
- Treat the local repo as the source of truth until the correct remote URL or credentials are provided

## Guidance For Future Agents

- Read this file first, then inspect the composition roots:
  `apps/bot/src/index.ts` and `apps/config-tab/src/App.tsx`
- Do not assume persistent storage exists; it does not
- Do not assume Graph calls are fully integration-tested; they compile, but live credentials and tenant setup are still required
- Preserve the split between Graph adapters, orchestration services, bot handlers, and frontend API client code
- If adding features, keep commits scoped by slice so they can be reviewed independently
