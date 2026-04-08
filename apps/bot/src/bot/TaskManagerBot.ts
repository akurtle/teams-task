import { TeamsActivityHandler, TurnContext } from "botbuilder";

export class TaskManagerBot extends TeamsActivityHandler {
  public constructor() {
    super();

    this.onMessage(async (context, next) => {
      const text = TurnContext.removeRecipientMention(context.activity)?.trim() ?? "";

      await context.sendActivity(
        text
          ? `Task Manager Bot received "${text}". Task workflows will be added in the next slices.`
          : "Task Manager Bot is online."
      );

      await next();
    });
  }
}
