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

export type CreateTaskPayload = {
  title: string;
  description?: string;
  planId: string;
  bucketId: string;
  assigneeUserId: string;
  assigneeTodoListId?: string;
  dueDateTime?: string;
  percentComplete?: number;
  teamId: string;
  channelId: string;
};

export type ServiceHealth = {
  ok: boolean;
  service: string;
  mode: string;
  syncedTasks: number;
};
