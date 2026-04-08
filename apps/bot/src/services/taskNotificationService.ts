import { CardFactory, ConversationReference, TurnContext } from "botbuilder";
import { BotFrameworkAdapter } from "botbuilder";
import { ConversationReferenceStore } from "../bot/conversationReferenceStore";
import { buildTaskSummaryCard } from "../cards/taskCards";
import { SyncedTaskRecord, TaskChangeType, TaskNotifier } from "./types";

export class TaskNotificationService implements TaskNotifier {
  public constructor(
    private readonly adapter: BotFrameworkAdapter,
    private readonly appId: string,
    private readonly conversationReferenceStore: ConversationReferenceStore
  ) {}

  public async notifyTaskChanged(record: SyncedTaskRecord, changeType: TaskChangeType) {
    const reference = await this.conversationReferenceStore.get(record.teamId, record.channelId);

    if (!reference) {
      return;
    }

    await this.adapter.continueConversationAsync(
      this.appId,
      reference as Partial<ConversationReference>,
      async (context) => {
        await context.sendActivity({
          type: "message",
          text: toNotificationText(changeType, record.title),
          attachments: [CardFactory.adaptiveCard(buildTaskSummaryCard(record))]
        });
      }
    );
  }
}

function toNotificationText(changeType: TaskChangeType, title: string) {
  switch (changeType) {
    case "created":
      return `Task created: ${title}`;
    case "updated":
      return `Task updated: ${title}`;
    case "assigned":
      return `Task reassigned: ${title}`;
    case "completed":
      return `Task completed: ${title}`;
    case "deleted":
      return `Task deleted: ${title}`;
  }
}
