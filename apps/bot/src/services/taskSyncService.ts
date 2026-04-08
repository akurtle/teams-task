import { PlannerService } from "../graph/plannerService";
import { TodoService } from "../graph/todoService";
import { TaskStateStore } from "./taskStateStore";
import {
  PlannerTaskMutation,
  PlannerTaskUpsertInput,
  SyncedTaskRecord,
  TaskAssignmentInput
} from "./types";

export class TaskSyncService {
  public constructor(
    private readonly plannerService: PlannerService,
    private readonly todoService: TodoService,
    private readonly taskStateStore: TaskStateStore
  ) {}

  public async listState() {
    return this.taskStateStore.list();
  }

  public async createTask(input: PlannerTaskUpsertInput, userAssertion?: string) {
    const plannerTask = await this.plannerService.createTask(input, userAssertion);
    const todoTask = await this.todoService.createTask(
      {
        assigneeUserId: input.assigneeUserId,
        assigneeTodoListId: input.assigneeTodoListId,
        title: input.title,
        description: input.description,
        dueDateTime: input.dueDateTime,
        percentComplete: input.percentComplete
      },
      userAssertion
    );

    return this.taskStateStore.upsert({
      plannerTaskId: plannerTask.plannerTaskId,
      todoTaskId: todoTask.todoTaskId,
      planId: input.planId,
      bucketId: input.bucketId,
      title: input.title,
      description: input.description,
      dueDateTime: input.dueDateTime,
      percentComplete: input.percentComplete ?? 0,
      assigneeUserId: input.assigneeUserId,
      assigneeTodoListId: todoTask.assigneeTodoListId,
      teamId: input.teamId,
      channelId: input.channelId,
      versionTag: plannerTask.versionTag
    });
  }

  public async updateTask(
    plannerTaskId: string,
    mutation: PlannerTaskMutation,
    userAssertion?: string
  ) {
    const currentRecord = await this.requireRecord(plannerTaskId);

    const plannerTask = await this.retryOnVersionConflict(
      () => this.plannerService.updateTask(plannerTaskId, mutation, userAssertion),
      async () => {
        const latestTask = await this.plannerService.getTask(plannerTaskId, userAssertion);
        currentRecord.versionTag = latestTask.versionTag;
      }
    );

    await this.todoService.updateTask(
      currentRecord.todoTaskId,
      {
        title: mutation.title,
        description: mutation.description,
        dueDateTime: mutation.dueDateTime,
        percentComplete: mutation.percentComplete
      },
      currentRecord.assigneeUserId,
      currentRecord.assigneeTodoListId,
      userAssertion
    );

    return this.taskStateStore.upsert({
      ...currentRecord,
      title: plannerTask.title,
      description:
        mutation.description !== undefined ? mutation.description : currentRecord.description,
      dueDateTime:
        mutation.dueDateTime !== undefined
          ? mutation.dueDateTime ?? undefined
          : currentRecord.dueDateTime,
      percentComplete: mutation.percentComplete ?? currentRecord.percentComplete,
      versionTag: plannerTask.versionTag
    });
  }

  public async assignTask(
    plannerTaskId: string,
    assignment: TaskAssignmentInput,
    userAssertion?: string
  ) {
    const currentRecord = await this.requireRecord(plannerTaskId);

    const plannerTask = await this.retryOnVersionConflict(
      () => this.plannerService.assignTask(plannerTaskId, assignment, userAssertion),
      async () => {
        const latestTask = await this.plannerService.getTask(plannerTaskId, userAssertion);
        currentRecord.versionTag = latestTask.versionTag;
      }
    );

    const todoTask = await this.todoService.rehomeTask(
      currentRecord.todoTaskId,
      assignment.assigneeUserId,
      currentRecord.title,
      assignment.assigneeTodoListId,
      currentRecord.description,
      currentRecord.dueDateTime,
      currentRecord.percentComplete,
      userAssertion
    );

    return this.taskStateStore.upsert({
      ...currentRecord,
      assigneeUserId: assignment.assigneeUserId,
      assigneeTodoListId: todoTask.assigneeTodoListId,
      todoTaskId: todoTask.todoTaskId,
      versionTag: plannerTask.versionTag
    });
  }

  private async requireRecord(plannerTaskId: string): Promise<SyncedTaskRecord> {
    const record = await this.taskStateStore.get(plannerTaskId);

    if (!record) {
      const error = new Error(`No synchronized task state exists for planner task ${plannerTaskId}.`);
      (error as Error & { statusCode: number }).statusCode = 404;
      throw error;
    }

    return record;
  }

  private async retryOnVersionConflict<T>(
    action: () => Promise<T>,
    refresh: () => Promise<void>
  ) {
    try {
      return await action();
    } catch (error) {
      const statusCode =
        typeof error === "object" &&
        error !== null &&
        "statusCode" in error &&
        typeof error.statusCode === "number"
          ? error.statusCode
          : undefined;

      if (statusCode !== 409 && statusCode !== 412) {
        throw error;
      }

      await refresh();
      return action();
    }
  }
}
