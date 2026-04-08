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
```
