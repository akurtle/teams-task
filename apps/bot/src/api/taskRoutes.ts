import { NextFunction, Request, Response, Router } from "express";
import { z } from "zod";
import { TaskSyncService } from "../services/taskSyncService";
import {
  PlannerTaskMutation,
  PlannerTaskUpsertInput,
  TaskAssignmentInput
} from "../services/types";

const createTaskSchema = z.object({
  title: z.string().min(1),
  planId: z.string().min(1),
  bucketId: z.string().min(1),
  assigneeUserId: z.string().min(1),
  assigneeTodoListId: z.string().min(1).optional(),
  description: z.string().optional(),
  dueDateTime: z.string().datetime().optional(),
  percentComplete: z.number().min(0).max(100).optional(),
  channelId: z.string().min(1),
  teamId: z.string().min(1)
});

const updateTaskSchema = z.object({
  title: z.string().min(1).optional(),
  description: z.string().optional(),
  dueDateTime: z.string().datetime().optional().nullable(),
  percentComplete: z.number().min(0).max(100).optional(),
  versionTag: z.string().min(1).optional()
});

const assignTaskSchema = z.object({
  assigneeUserId: z.string().min(1),
  assigneeTodoListId: z.string().min(1).optional(),
  versionTag: z.string().min(1).optional()
});

export function createTaskRouter(taskSyncService: TaskSyncService) {
  const router = Router();

  router.get("/", async (_request, response, next) => {
    try {
      response.json(await taskSyncService.listState());
    } catch (error) {
      next(error);
    }
  });

  router.post("/", async (request, response, next) => {
    try {
      const payload = createTaskSchema.parse(request.body);
      const result = await taskSyncService.createTask(
        payload satisfies PlannerTaskUpsertInput,
        extractBearerToken(request.header("authorization"))
      );

      response.status(201).json(result);
    } catch (error) {
      next(error);
    }
  });

  router.patch("/:plannerTaskId", async (request, response, next) => {
    try {
      const payload = updateTaskSchema.parse(request.body);
      const result = await taskSyncService.updateTask(
        request.params.plannerTaskId,
        payload satisfies PlannerTaskMutation,
        extractBearerToken(request.header("authorization"))
      );

      response.json(result);
    } catch (error) {
      next(error);
    }
  });

  router.post("/:plannerTaskId/assign", async (request, response, next) => {
    try {
      const payload = assignTaskSchema.parse(request.body);
      const result = await taskSyncService.assignTask(
        request.params.plannerTaskId,
        payload satisfies TaskAssignmentInput,
        extractBearerToken(request.header("authorization"))
      );

      response.json(result);
    } catch (error) {
      next(error);
    }
  });

  router.use((error: unknown, _request: Request, response: Response, _next: NextFunction) => {
    if (error instanceof z.ZodError) {
      response.status(400).json({
        error: "ValidationError",
        details: error.flatten()
      });
      return;
    }

    const statusCode =
      typeof error === "object" &&
      error !== null &&
      "statusCode" in error &&
      typeof error.statusCode === "number"
        ? error.statusCode
        : 500;

    const message =
      error instanceof Error ? error.message : "Unexpected task synchronization error.";

    response.status(statusCode).json({
      error: "TaskSyncError",
      message
    });
  });

  return router;
}

function extractBearerToken(headerValue?: string) {
  if (!headerValue?.startsWith("Bearer ")) {
    return undefined;
  }

  return headerValue.slice("Bearer ".length);
}
