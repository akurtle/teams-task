import express, { NextFunction, Request, Response } from "express";
import { BotFrameworkAdapter } from "botbuilder";
import cors from "cors";
import { env } from "./config/env";
import { TaskManagerBot } from "./bot/TaskManagerBot";
import { createTaskRouter } from "./api/taskRoutes";
import { GraphClientFactory } from "./graph/graphClientFactory";
import { PlannerService } from "./graph/plannerService";
import { TodoService } from "./graph/todoService";
import { TaskStateStore } from "./services/taskStateStore";
import { TaskSyncService } from "./services/taskSyncService";

void start();

async function start() {
  const adapter = new BotFrameworkAdapter({
    appId: env.microsoftAppId,
    appPassword: env.microsoftAppPassword
  });

  adapter.onTurnError = async (context, error) => {
    console.error("[onTurnError]", error);
    await context.sendActivity("The bot hit an unexpected error.");
  };

  const app = express();
  const graphClientFactory = new GraphClientFactory(env);
  const plannerService = new PlannerService(graphClientFactory);
  const todoService = new TodoService(graphClientFactory);
  const stateStore = new TaskStateStore(env.taskStateFilePath);
  const taskSyncService = new TaskSyncService(plannerService, todoService, stateStore);
  const bot = new TaskManagerBot(taskSyncService);

  app.use(cors());
  app.use(express.json());

  app.get("/health", async (_request: Request, response: Response, next: NextFunction) => {
    try {
      response.json({
        ok: true,
        service: "teams-task-manager-bot",
        mode: env.nodeEnv,
        syncedTasks: (await stateStore.list()).length
      });
    } catch (error) {
      next(error);
    }
  });

  app.use("/api/tasks", createTaskRouter(taskSyncService));

  app.post("/api/messages", async (request: Request, response: Response) => {
    await adapter.processActivity(request, response, async (context) => bot.run(context));
  });

  app.listen(env.port, () => {
    console.log(`Bot service listening on port ${env.port}`);
  });
}
