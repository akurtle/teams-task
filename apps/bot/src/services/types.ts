export type PlannerTaskUpsertInput = {
  title: string;
  planId: string;
  bucketId: string;
  assigneeUserId: string;
  assigneeTodoListId?: string;
  description?: string;
  dueDateTime?: string;
  percentComplete?: number;
  teamId: string;
  channelId: string;
};

export type PlannerTaskMutation = {
  title?: string;
  description?: string;
  dueDateTime?: string | null;
  percentComplete?: number;
  versionTag?: string;
};

export type TaskAssignmentInput = {
  assigneeUserId: string;
  assigneeTodoListId?: string;
  versionTag?: string;
};

export type TodoTaskUpsertInput = {
  assigneeUserId: string;
  assigneeTodoListId?: string;
  title: string;
  description?: string;
  dueDateTime?: string;
  percentComplete?: number;
};

export type TodoTaskMutation = {
  title?: string;
  description?: string;
  dueDateTime?: string | null;
  percentComplete?: number;
};

export type PlannerTaskRecord = {
  plannerTaskId: string;
  title: string;
  planId: string;
  bucketId: string;
  description?: string;
  dueDateTime?: string;
  percentComplete: number;
  versionTag?: string;
};

export type TodoTaskRecord = {
  todoTaskId: string;
  assigneeUserId: string;
  assigneeTodoListId: string;
  title: string;
  description?: string;
  dueDateTime?: string;
  status: "notStarted" | "inProgress" | "completed";
};

export type SyncedTaskRecord = {
  plannerTaskId: string;
  todoTaskId: string;
  title: string;
  description?: string;
  dueDateTime?: string;
  percentComplete: number;
  planId: string;
  bucketId: string;
  assigneeUserId: string;
  assigneeTodoListId: string;
  teamId: string;
  channelId: string;
  versionTag?: string;
};
