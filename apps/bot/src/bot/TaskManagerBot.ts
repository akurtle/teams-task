import {
  CardFactory,
  MessageFactory,
  TeamsActivityHandler,
  TurnContext
} from "botbuilder";
import {
  buildAssignTaskCard,
  buildCreateTaskCard,
  buildTaskSummaryCard,
  buildUpdateTaskCard
} from "../cards/taskCards";
import { TaskSyncService } from "../services/taskSyncService";

export class TaskManagerBot extends TeamsActivityHandler {
  public constructor(private readonly taskSyncService: TaskSyncService) {
    super();

    this.onMessage(async (context, next) => {
      const cardCommand = context.activity.value?.command;
      if (typeof cardCommand === "string") {
        await this.handleCardCommand(context, cardCommand);
        await next();
        return;
      }

      const text = TurnContext.removeRecipientMention(context.activity)?.trim() ?? "";
      const normalized = text.toLowerCase();

      if (normalized === "task create") {
        await context.sendActivity(
          MessageFactory.attachment(CardFactory.adaptiveCard(buildCreateTaskCard()))
        );
        await next();
        return;
      }

      if (normalized.startsWith("task update ")) {
        const plannerTaskId = text.split(/\s+/).at(-1) ?? "";

        await context.sendActivity(
          MessageFactory.attachment(CardFactory.adaptiveCard(buildUpdateTaskCard(plannerTaskId)))
        );
        await next();
        return;
      }

      if (normalized.startsWith("task assign ")) {
        const plannerTaskId = text.split(/\s+/).at(-1) ?? "";

        await context.sendActivity(
          MessageFactory.attachment(CardFactory.adaptiveCard(buildAssignTaskCard(plannerTaskId)))
        );
        await next();
        return;
      }

      if (normalized === "task list") {
        const records = await this.taskSyncService.listState();

        if (!records.length) {
          await context.sendActivity("No synchronized tasks are currently stored.");
          await next();
          return;
        }

        for (const record of records) {
          await context.sendActivity(
            MessageFactory.attachment(CardFactory.adaptiveCard(buildTaskSummaryCard(record)))
          );
        }

        await next();
        return;
      }

      await context.sendActivity(
        "Commands: `task create`, `task update <plannerTaskId>`, `task assign <plannerTaskId>`, `task list`."
      );

      await next();
    });
  }

  private async handleCardCommand(context: TurnContext, command: string) {
    switch (command) {
      case "create-task": {
        const payload = {
          title: String(context.activity.value.title ?? ""),
          planId: String(context.activity.value.planId ?? ""),
          bucketId: String(context.activity.value.bucketId ?? ""),
          assigneeUserId: String(context.activity.value.assigneeUserId ?? ""),
          assigneeTodoListId: optionalString(context.activity.value.assigneeTodoListId),
          description: optionalString(context.activity.value.description),
          dueDateTime: optionalString(context.activity.value.dueDateTime),
          percentComplete: optionalNumber(context.activity.value.percentComplete),
          teamId: String(context.activity.channelData?.team?.id ?? ""),
          channelId: String(context.activity.channelData?.channel?.id ?? "")
        };

        const record = await this.taskSyncService.createTask(payload);
        await context.sendActivity(
          MessageFactory.attachment(CardFactory.adaptiveCard(buildTaskSummaryCard(record)))
        );
        return;
      }

      case "update-task": {
        const plannerTaskId = String(context.activity.value.plannerTaskId ?? "");
        const record = await this.taskSyncService.updateTask(plannerTaskId, {
          title: optionalString(context.activity.value.title),
          description: optionalString(context.activity.value.description),
          dueDateTime: nullableString(context.activity.value.dueDateTime),
          percentComplete: optionalNumber(context.activity.value.percentComplete),
          versionTag: optionalString(context.activity.value.versionTag)
        });

        await context.sendActivity(
          MessageFactory.attachment(CardFactory.adaptiveCard(buildTaskSummaryCard(record)))
        );
        return;
      }

      case "assign-task": {
        const plannerTaskId = String(context.activity.value.plannerTaskId ?? "");
        const record = await this.taskSyncService.assignTask(plannerTaskId, {
          assigneeUserId: String(context.activity.value.assigneeUserId ?? ""),
          assigneeTodoListId: optionalString(context.activity.value.assigneeTodoListId),
          versionTag: optionalString(context.activity.value.versionTag)
        });

        await context.sendActivity(
          MessageFactory.attachment(CardFactory.adaptiveCard(buildTaskSummaryCard(record)))
        );
        return;
      }

      default:
        await context.sendActivity("Unsupported adaptive card command.");
    }
  }
}

function optionalString(value: unknown) {
  return typeof value === "string" && value.trim().length > 0 ? value.trim() : undefined;
}

function nullableString(value: unknown) {
  if (typeof value !== "string") {
    return undefined;
  }

  return value.trim().length > 0 ? value.trim() : null;
}

function optionalNumber(value: unknown) {
  if (typeof value === "number") {
    return value;
  }

  if (typeof value === "string" && value.trim().length > 0) {
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : undefined;
  }

  return undefined;
}
