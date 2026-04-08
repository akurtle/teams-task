import { TestAdapter } from "botbuilder";
import { describe, expect, it, vi } from "vitest";
import { TaskManagerBot } from "./TaskManagerBot";

describe("TaskManagerBot", () => {
  it("returns a create-task adaptive card", async () => {
    const taskSyncService = {
      listState: vi.fn().mockResolvedValue([])
    };
    const conversationReferenceStore = {
      upsert: vi.fn().mockResolvedValue(undefined)
    };
    const bot = new TaskManagerBot(taskSyncService as never, conversationReferenceStore as never);
    const adapter = new TestAdapter(async (context) => bot.run(context));

    await adapter.send("task create").assertReply((activity) => {
      expect(activity.attachments).toHaveLength(1);
      expect(activity.attachments?.[0]?.contentType).toBe("application/vnd.microsoft.card.adaptive");
    });
  });

  it("lists stored synchronized tasks", async () => {
    const taskSyncService = {
      listState: vi.fn().mockResolvedValue([
        {
          plannerTaskId: "planner-1",
          todoTaskId: "todo-1",
          title: "Launch prep",
          description: "Coordinate launch",
          percentComplete: 0,
          planId: "plan-1",
          bucketId: "bucket-1",
          assigneeUserId: "user-1",
          assigneeTodoListId: "list-1",
          teamId: "team-1",
          channelId: "channel-1"
        }
      ])
    };
    const conversationReferenceStore = {
      upsert: vi.fn().mockResolvedValue(undefined)
    };
    const bot = new TaskManagerBot(taskSyncService as never, conversationReferenceStore as never);
    const adapter = new TestAdapter(async (context) => bot.run(context));

    await adapter.send("task list").assertReply((activity) => {
      const card = activity.attachments?.[0]?.content as { body?: Array<{ text?: string }> };
      expect(card.body?.[0]?.text).toBe("Launch prep");
    });
  });
});
