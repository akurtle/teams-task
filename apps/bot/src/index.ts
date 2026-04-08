import express, { Request, Response } from "express";
import { BotFrameworkAdapter } from "botbuilder";
import { env } from "./config/env";
import { TaskManagerBot } from "./bot/TaskManagerBot";

const adapter = new BotFrameworkAdapter({
  appId: env.microsoftAppId,
  appPassword: env.microsoftAppPassword
});

adapter.onTurnError = async (context, error) => {
  console.error("[onTurnError]", error);
  await context.sendActivity("The bot hit an unexpected error.");
};

const bot = new TaskManagerBot();
const app = express();

app.use(express.json());

app.get("/health", (_request: Request, response: Response) => {
  response.json({
    ok: true,
    service: "teams-task-manager-bot",
    mode: env.nodeEnv
  });
});

app.post("/api/messages", async (request: Request, response: Response) => {
  await adapter.processActivity(request, response, async (context) => bot.run(context));
});

app.listen(env.port, () => {
  console.log(`Bot service listening on port ${env.port}`);
});
