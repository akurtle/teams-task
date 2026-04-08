import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { SyncedTaskRecord } from "./types";

type PersistedTaskState = {
  records: SyncedTaskRecord[];
};

export class TaskStateStore {
  private readonly records = new Map<string, SyncedTaskRecord>();
  private readonly storagePath: string;
  private readonly initialized: Promise<void>;
  private writeChain = Promise.resolve();

  public constructor(storagePath: string) {
    this.storagePath = path.resolve(storagePath);
    this.initialized = this.load();
  }

  public async upsert(record: SyncedTaskRecord) {
    await this.initialized;
    this.records.set(record.plannerTaskId, record);
    await this.persist();
    return record;
  }

  public async get(plannerTaskId: string) {
    await this.initialized;
    return this.records.get(plannerTaskId);
  }

  public async list() {
    await this.initialized;
    return [...this.records.values()];
  }

  public async delete(plannerTaskId: string) {
    await this.initialized;
    const deleted = this.records.delete(plannerTaskId);

    if (deleted) {
      await this.persist();
    }

    return deleted;
  }

  private async load() {
    await mkdir(path.dirname(this.storagePath), { recursive: true });

    try {
      const content = await readFile(this.storagePath, "utf-8");
      const parsed = JSON.parse(content) as PersistedTaskState;

      for (const record of parsed.records ?? []) {
        this.records.set(record.plannerTaskId, record);
      }
    } catch (error) {
      if (!isFileMissingError(error)) {
        throw error;
      }

      await this.persist();
    }
  }

  private async persist() {
    this.writeChain = this.writeChain.then(async () => {
      const payload: PersistedTaskState = {
        records: [...this.records.values()]
      };

      await writeFile(this.storagePath, JSON.stringify(payload, null, 2), "utf-8");
    });

    await this.writeChain;
  }
}

function isFileMissingError(error: unknown) {
  return (
    typeof error === "object" &&
    error !== null &&
    "code" in error &&
    error.code === "ENOENT"
  );
}
