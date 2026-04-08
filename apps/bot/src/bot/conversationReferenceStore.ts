import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { ConversationReference } from "botbuilder";

type StoredConversationReference = {
  teamId: string;
  channelId: string;
  reference: Partial<ConversationReference>;
};

type PersistedConversationReferences = {
  records: StoredConversationReference[];
};

export class ConversationReferenceStore {
  private readonly storagePath: string;
  private readonly initialized: Promise<void>;
  private readonly records = new Map<string, StoredConversationReference>();
  private writeChain = Promise.resolve();

  public constructor(storagePath: string) {
    this.storagePath = path.resolve(storagePath);
    this.initialized = this.load();
  }

  public async upsert(
    teamId: string,
    channelId: string,
    reference: Partial<ConversationReference>
  ) {
    await this.initialized;
    this.records.set(buildKey(teamId, channelId), {
      teamId,
      channelId,
      reference
    });
    await this.persist();
  }

  public async get(teamId: string, channelId: string) {
    await this.initialized;
    return this.records.get(buildKey(teamId, channelId))?.reference;
  }

  private async load() {
    await mkdir(path.dirname(this.storagePath), { recursive: true });

    try {
      const content = await readFile(this.storagePath, "utf-8");
      const parsed = JSON.parse(content) as PersistedConversationReferences;

      for (const record of parsed.records ?? []) {
        this.records.set(buildKey(record.teamId, record.channelId), record);
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
      const payload: PersistedConversationReferences = {
        records: [...this.records.values()]
      };

      await writeFile(this.storagePath, JSON.stringify(payload, null, 2), "utf-8");
    });

    await this.writeChain;
  }
}

function buildKey(teamId: string, channelId: string) {
  return `${teamId}:${channelId}`;
}

function isFileMissingError(error: unknown) {
  return (
    typeof error === "object" &&
    error !== null &&
    "code" in error &&
    error.code === "ENOENT"
  );
}
