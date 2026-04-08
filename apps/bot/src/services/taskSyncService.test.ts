import { describe, expect, it, vi } from "vitest";
import { TaskSyncService } from "./taskSyncService";
import { SyncedTaskRecord } from "./types";

describe("TaskSyncService", () => {
  it("creates linked Planner and To Do records", async () => {
    const plannerService = {
      createTask: vi.fn().mockResolvedValue({
        plannerTaskId: "planner-1",
        title: "Ship release",
        planId: "plan-1",
        bucketId: "bucket-1",
        percentComplete: 0,
        versionTag: "v1"
      })
    };
    const todoService = {
      createTask: vi.fn().mockResolvedValue({
        todoTaskId: "todo-1",
        assigneeUserId: "user-1",
        assigneeTodoListId: "list-1",
        title: "Ship release",
        status: "notStarted"
      })
    };
    const taskStateStore = createMemoryStore();
    const service = new TaskSyncService(
      plannerService as never,
      todoService as never,
      taskStateStore as never
    );

    const result = await service.createTask({
      title: "Ship release",
      planId: "plan-1",
      bucketId: "bucket-1",
      assigneeUserId: "user-1",
      description: "Coordinate launch",
      teamId: "team-1",
      channelId: "channel-1"
    });

    expect(result).toMatchObject({
      plannerTaskId: "planner-1",
      todoTaskId: "todo-1",
      assigneeTodoListId: "list-1"
    });
    expect(plannerService.createTask).toHaveBeenCalledOnce();
    expect(todoService.createTask).toHaveBeenCalledOnce();
  });

  it("retries planner updates after version conflicts", async () => {
    const staleError = Object.assign(new Error("stale"), { statusCode: 412 });
    const plannerService = {
      updateTask: vi
        .fn()
        .mockRejectedValueOnce(staleError)
        .mockResolvedValueOnce({
          plannerTaskId: "planner-1",
          title: "Ship release",
          planId: "plan-1",
          bucketId: "bucket-1",
          percentComplete: 100,
          versionTag: "v2"
        }),
      getTask: vi.fn().mockResolvedValue({
        plannerTaskId: "planner-1",
        title: "Ship release",
        planId: "plan-1",
        bucketId: "bucket-1",
        percentComplete: 10,
        versionTag: "fresh"
      })
    };
    const todoService = {
      updateTask: vi.fn().mockResolvedValue({
        todoTaskId: "todo-1",
        assigneeUserId: "user-1",
        assigneeTodoListId: "list-1",
        title: "Ship release",
        status: "completed"
      })
    };
    const taskStateStore = createMemoryStore({
      plannerTaskId: "planner-1",
      todoTaskId: "todo-1",
      title: "Ship release",
      percentComplete: 10,
      planId: "plan-1",
      bucketId: "bucket-1",
      assigneeUserId: "user-1",
      assigneeTodoListId: "list-1",
      teamId: "team-1",
      channelId: "channel-1",
      versionTag: "stale"
    });
    const service = new TaskSyncService(
      plannerService as never,
      todoService as never,
      taskStateStore as never
    );

    const result = await service.updateTask("planner-1", {
      percentComplete: 100
    });

    expect(plannerService.updateTask).toHaveBeenCalledTimes(2);
    expect(plannerService.getTask).toHaveBeenCalledOnce();
    expect(todoService.updateTask).toHaveBeenCalledOnce();
    expect(result.versionTag).toBe("v2");
    expect(result.percentComplete).toBe(100);
  });
});

function createMemoryStore(seed?: SyncedTaskRecord) {
  const records = new Map<string, SyncedTaskRecord>();

  if (seed) {
    records.set(seed.plannerTaskId, seed);
  }

  return {
    upsert: vi.fn(async (record: SyncedTaskRecord) => {
      records.set(record.plannerTaskId, record);
      return record;
    }),
    get: vi.fn(async (plannerTaskId: string) => records.get(plannerTaskId)),
    list: vi.fn(async () => [...records.values()])
  };
}
