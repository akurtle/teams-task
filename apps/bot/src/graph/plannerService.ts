import { GraphClientFactory } from "./graphClientFactory";
import {
  PlannerTaskMutation,
  PlannerTaskRecord,
  PlannerTaskUpsertInput,
  TaskAssignmentInput
} from "../services/types";

type PlannerTaskResponse = {
  id: string;
  title: string;
  planId: string;
  bucketId: string;
  percentComplete?: number;
  dueDateTime?: string;
  assignments?: Record<string, unknown>;
  ["@odata.etag"]?: string;
};

type PlannerTaskDetailsResponse = {
  description?: string;
  ["@odata.etag"]?: string;
};

export class PlannerService {
  public constructor(private readonly graphClientFactory: GraphClientFactory) {}

  public async createTask(input: PlannerTaskUpsertInput, userAssertion?: string) {
    const client = await this.graphClientFactory.createClient(userAssertion);
    const createdTask = (await this.graphClientFactory.request(client, "/planner/tasks").post({
      title: input.title,
      planId: input.planId,
      bucketId: input.bucketId,
      dueDateTime: input.dueDateTime,
      percentComplete: input.percentComplete ?? 0,
      assignments: buildAssignments(input.assigneeUserId)
    })) as PlannerTaskResponse;

    if (input.description) {
      await this.updateTaskDetails(createdTask.id, input.description, userAssertion);
    }

    return toPlannerTaskRecord(createdTask, input.description);
  }

  public async updateTask(
    plannerTaskId: string,
    mutation: PlannerTaskMutation,
    userAssertion?: string
  ) {
    const client = await this.graphClientFactory.createClient(userAssertion);
    const currentTask = await this.getTask(plannerTaskId, userAssertion);
    const etag = mutation.versionTag ?? currentTask.versionTag ?? "*";

    const updatedTask = (await this.graphClientFactory
      .request(client, `/planner/tasks/${plannerTaskId}`)
      .header("If-Match", etag)
      .patch({
        title: mutation.title ?? currentTask.title,
        dueDateTime:
          mutation.dueDateTime === null ? null : mutation.dueDateTime ?? currentTask.dueDateTime,
        percentComplete: mutation.percentComplete ?? currentTask.percentComplete ?? 0
      })) as PlannerTaskResponse;

    if (mutation.description !== undefined) {
      await this.updateTaskDetails(plannerTaskId, mutation.description, userAssertion);
    }

    const details = await this.getTaskDetails(plannerTaskId, userAssertion);

    return {
      ...toPlannerTaskRecord(updatedTask, details.description),
      description: details.description
    };
  }

  public async assignTask(
    plannerTaskId: string,
    assignment: TaskAssignmentInput,
    userAssertion?: string
  ) {
    const client = await this.graphClientFactory.createClient(userAssertion);
    const currentTask = await this.getTask(plannerTaskId, userAssertion);
    const etag = assignment.versionTag ?? currentTask.versionTag ?? "*";

    const updatedTask = (await this.graphClientFactory
      .request(client, `/planner/tasks/${plannerTaskId}`)
      .header("If-Match", etag)
      .patch({
        assignments: buildAssignments(assignment.assigneeUserId)
      })) as PlannerTaskResponse;

    const details = await this.getTaskDetails(plannerTaskId, userAssertion);
    return toPlannerTaskRecord(updatedTask, details.description);
  }

  public async getTask(plannerTaskId: string, userAssertion?: string) {
    const client = await this.graphClientFactory.createClient(userAssertion);
    const task = (await this.graphClientFactory
      .request(client, `/planner/tasks/${plannerTaskId}`)
      .get()) as PlannerTaskResponse;
    const details = await this.getTaskDetails(plannerTaskId, userAssertion);

    return toPlannerTaskRecord(task, details.description);
  }

  private async getTaskDetails(plannerTaskId: string, userAssertion?: string) {
    const client = await this.graphClientFactory.createClient(userAssertion);
    const details = (await this.graphClientFactory
      .request(client, `/planner/tasks/${plannerTaskId}/details`)
      .get()) as PlannerTaskDetailsResponse;

    return details;
  }

  private async updateTaskDetails(
    plannerTaskId: string,
    description: string,
    userAssertion?: string
  ) {
    const client = await this.graphClientFactory.createClient(userAssertion);
    const details = await this.getTaskDetails(plannerTaskId, userAssertion);

    await this.graphClientFactory
      .request(client, `/planner/tasks/${plannerTaskId}/details`)
      .header("If-Match", details["@odata.etag"] ?? "*")
      .patch({
        description
      });
  }
}

function buildAssignments(userId: string) {
  return {
    [userId]: {
      "@odata.type": "#microsoft.graph.plannerAssignment",
      orderHint: " !"
    }
  };
}

function toPlannerTaskRecord(
  task: PlannerTaskResponse,
  description?: string
): PlannerTaskRecord {
  return {
    plannerTaskId: task.id,
    title: task.title,
    planId: task.planId,
    bucketId: task.bucketId,
    description,
    dueDateTime: task.dueDateTime,
    percentComplete: task.percentComplete ?? 0,
    versionTag: task["@odata.etag"]
  };
}
