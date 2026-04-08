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
- Microsoft Graph client creation using Azure AD app credentials or on-behalf-of flow
- Planner task create, update, and assign operations
- To Do task create and update operations
- In-memory synchronized task state tracking
- REST API for listing, creating, updating, and reassigning synchronized tasks

Important files:

- `apps/bot/src/index.ts`
  App composition root
- `apps/bot/src/bot/TaskManagerBot.ts`
  Teams command handling and adaptive card submit handling
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
  Current in-memory state store

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
- `task list`
  Posts summary cards for the tasks currently tracked in memory

REST endpoints supported right now:

- `GET /health`
- `GET /api/tasks`
- `POST /api/tasks`
- `PATCH /api/tasks/:plannerTaskId`
- `POST /api/tasks/:plannerTaskId/assign`

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

Frontend environment variables:

- `VITE_API_BASE_URL`

Notes:

- `GraphScopes` defaults to `https://graph.microsoft.com/.default`
- The config tab defaults its backend target to `http://localhost:3978`
- The create-task form converts browser-local datetime input into ISO-8601 before sending it to the backend

## Build and Verification

Standard commands:

```bash
npm install
npm run build
```

Workspace-specific commands:

```bash
npm run dev:bot
npm run dev:tab
```

Current known-good verification:

- `npm install` completed successfully
- `npm run build` completed successfully for both workspaces

## Architectural Constraints and Assumptions

- The current synchronized task store is in-memory only.
  Restarting the backend loses tracked task mappings between Planner and To Do.
- Planner concurrency is handled with a retry path on version conflict via ETag refresh.
- To Do reassignment currently creates a new task in the target user/list context and updates the local mapping.
- The backend is structured to support on-behalf-of Graph access, but real deployment still requires Azure AD app registration, permissions, and Teams/Bot configuration.
- The React tab is a functional control panel, not yet a fully Teams SDK-aware production tab experience.

## Known Gaps / Likely Next Work

Highest-value next steps:

- Persist synchronized task mappings in durable storage instead of memory
- Add authentication and authorization around the REST API
- Add Microsoft Teams SDK integration inside the React tab
- Add outbound bot notifications for state changes triggered outside the current process
- Add tests for Graph service adapters, route validation, and bot command handling
- Add deployment assets such as Teams app manifest, Azure deployment config, and CI pipeline
- Handle deletion and completion synchronization more explicitly across Planner and To Do

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
