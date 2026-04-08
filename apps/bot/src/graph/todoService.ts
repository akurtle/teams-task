import { GraphClientFactory } from "./graphClientFactory";
import {
  TodoTaskMutation,
  TodoTaskRecord,
  TodoTaskUpsertInput
} from "../services/types";

type TodoTaskResponse = {
  id: string;
  title: string;
  body?: {
    content?: string;
  };
  dueDateTime?: {
    dateTime?: string;
    timeZone?: string;
  };
  status?: "notStarted" | "inProgress" | "completed";
};

type TodoListResponse = {
  value: Array<{
    id: string;
    displayName: string;
    wellknownListName?: string;
  }>;
};

export class TodoService {
  public constructor(private readonly graphClientFactory: GraphClientFactory) {}

  public async createTask(input: TodoTaskUpsertInput, userAssertion?: string) {
    const client = await this.graphClientFactory.createClient(userAssertion);
    const listId = await this.resolveTodoListId(
      input.assigneeUserId,
      input.assigneeTodoListId,
      userAssertion
    );

    const createdTask = (await this.graphClientFactory
      .request(client, `/users/${input.assigneeUserId}/todo/lists/${listId}/tasks`)
      .post({
        title: input.title,
        body: input.description
          ? {
              contentType: "text",
              content: input.description
            }
          : undefined,
        dueDateTime: input.dueDateTime
          ? {
              dateTime: input.dueDateTime,
              timeZone: "UTC"
            }
          : undefined,
        status: toTodoStatus(input.percentComplete ?? 0)
      })) as TodoTaskResponse;

    return toTodoTaskRecord(createdTask, listId, input.assigneeUserId);
  }

  public async updateTask(
    todoTaskId: string,
    mutation: TodoTaskMutation,
    assigneeUserId: string,
    assigneeTodoListId?: string,
    userAssertion?: string
  ) {
    const client = await this.graphClientFactory.createClient(userAssertion);
    const listId = await this.resolveTodoListId(
      assigneeUserId,
      assigneeTodoListId,
      userAssertion
    );

    const updatedTask = (await this.graphClientFactory
      .request(client, `/users/${assigneeUserId}/todo/lists/${listId}/tasks/${todoTaskId}`)
      .patch({
        title: mutation.title,
        body:
          mutation.description !== undefined
            ? {
                contentType: "text",
                content: mutation.description
              }
            : undefined,
        dueDateTime:
          mutation.dueDateTime === undefined
            ? undefined
            : mutation.dueDateTime === null
              ? null
              : {
                  dateTime: mutation.dueDateTime,
                  timeZone: "UTC"
                },
        status:
          mutation.percentComplete === undefined
            ? undefined
            : toTodoStatus(mutation.percentComplete)
      })) as TodoTaskResponse;

    return toTodoTaskRecord(updatedTask, listId, assigneeUserId);
  }

  public async rehomeTask(
    todoTaskId: string,
    assigneeUserId: string,
    title: string,
    assigneeTodoListId?: string,
    description?: string,
    dueDateTime?: string,
    percentComplete?: number,
    userAssertion?: string
  ) {
    return this.createTask(
      {
        assigneeUserId,
        assigneeTodoListId,
        title,
        description,
        dueDateTime,
        percentComplete
      },
      userAssertion
    );
  }

  public async deleteTask(
    todoTaskId: string,
    assigneeUserId: string,
    assigneeTodoListId?: string,
    userAssertion?: string
  ) {
    const client = await this.graphClientFactory.createClient(userAssertion);
    const listId = await this.resolveTodoListId(
      assigneeUserId,
      assigneeTodoListId,
      userAssertion
    );

    await this.graphClientFactory
      .request(client, `/users/${assigneeUserId}/todo/lists/${listId}/tasks/${todoTaskId}`)
      .delete();
  }

  private async resolveTodoListId(
    userId: string,
    preferredListId: string | undefined,
    userAssertion?: string
  ) {
    if (preferredListId) {
      return preferredListId;
    }

    const client = await this.graphClientFactory.createClient(userAssertion);
    const lists = (await this.graphClientFactory
      .request(client, `/users/${userId}/todo/lists`)
      .get()) as TodoListResponse;

    const defaultList = lists.value.find((list) => list.wellknownListName === "defaultList");

    return defaultList?.id ?? lists.value[0]?.id ?? "Tasks";
  }
}

function toTodoStatus(percentComplete: number) {
  if (percentComplete >= 100) {
    return "completed";
  }

  if (percentComplete > 0) {
    return "inProgress";
  }

  return "notStarted";
}

function toTodoTaskRecord(
  task: TodoTaskResponse,
  listId: string,
  userId: string
): TodoTaskRecord {
  return {
    todoTaskId: task.id,
    assigneeUserId: userId,
    assigneeTodoListId: listId,
    title: task.title,
    description: task.body?.content,
    dueDateTime: task.dueDateTime?.dateTime,
    status: task.status ?? "notStarted"
  };
}
