import { SyncedTaskRecord } from "./types";

export class TaskStateStore {
  private readonly records = new Map<string, SyncedTaskRecord>();

  public upsert(record: SyncedTaskRecord) {
    this.records.set(record.plannerTaskId, record);
    return record;
  }

  public get(plannerTaskId: string) {
    return this.records.get(plannerTaskId);
  }

  public list() {
    return [...this.records.values()];
  }
}
