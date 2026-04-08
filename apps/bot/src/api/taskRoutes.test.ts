import express from "express";
import request from "supertest";
import { describe, expect, it, vi } from "vitest";
import { createTaskRouter, TaskApiAccessGuards } from "./taskRoutes";

const allowAllAccess: TaskApiAccessGuards = {
  requireReadAccess: (_request, _response, next) => next(),
  requireWriteAccess: (_request, _response, next) => next()
};

describe("taskRoutes", () => {
  it("lists synchronized tasks", async () => {
    const taskSyncService = {
      listState: vi.fn().mockResolvedValue([
        {
          plannerTaskId: "planner-1",
          todoTaskId: "todo-1",
          title: "Launch prep",
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

    const app = createApp(taskSyncService);
    const response = await request(app).get("/api/tasks");

    expect(response.status).toBe(200);
    expect(response.body).toHaveLength(1);
  });

  it("rejects invalid create payloads", async () => {
    const taskSyncService = {
      listState: vi.fn().mockResolvedValue([]),
      createTask: vi.fn()
    };

    const app = createApp(taskSyncService);
    const response = await request(app).post("/api/tasks").send({
      bucketId: "bucket-1",
      planId: "plan-1",
      assigneeUserId: "user-1",
      teamId: "team-1",
      channelId: "channel-1"
    });

    expect(response.status).toBe(400);
    expect(response.body.error).toBe("ValidationError");
    expect(taskSyncService.createTask).not.toHaveBeenCalled();
  });

  it("routes complete requests to the sync service", async () => {
    const taskSyncService = {
      listState: vi.fn().mockResolvedValue([]),
      completeTask: vi.fn().mockResolvedValue({
        plannerTaskId: "planner-1",
        todoTaskId: "todo-1",
        title: "Launch prep",
        percentComplete: 100,
        planId: "plan-1",
        bucketId: "bucket-1",
        assigneeUserId: "user-1",
        assigneeTodoListId: "list-1",
        teamId: "team-1",
        channelId: "channel-1"
      })
    };

    const app = createApp(taskSyncService);
    const response = await request(app).post("/api/tasks/planner-1/complete").send({});

    expect(response.status).toBe(200);
    expect(taskSyncService.completeTask).toHaveBeenCalledWith("planner-1", undefined);
  });
});

function createApp(taskSyncService: object) {
  const app = express();
  app.use(express.json());
  app.use("/api/tasks", createTaskRouter(taskSyncService as never, allowAllAccess));
  return app;
}
