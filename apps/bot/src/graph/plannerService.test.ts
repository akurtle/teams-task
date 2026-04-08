import { describe, expect, it } from "vitest";
import { PlannerService } from "./plannerService";

describe("PlannerService", () => {
  it("creates planner assignments and updates details when a description is supplied", async () => {
    const requests: Array<{ path: string; action: string; payload?: unknown; header?: string }> = [];
    const graphClientFactory = {
      createClient: async () => ({ client: true }),
      request: (_client: unknown, path: string) => createRequest(path, requests)
    };
    const plannerService = new PlannerService(graphClientFactory as never);

    const record = await plannerService.createTask({
      title: "Launch prep",
      planId: "plan-1",
      bucketId: "bucket-1",
      assigneeUserId: "user-1",
      description: "Coordinate the rollout",
      teamId: "team-1",
      channelId: "channel-1"
    });

    expect(record.plannerTaskId).toBe("planner-1");
    expect(requests[0]).toMatchObject({
      path: "/planner/tasks",
      action: "post"
    });
    expect(requests[0]?.payload).toMatchObject({
      assignments: {
        "user-1": {
          "@odata.type": "#microsoft.graph.plannerAssignment"
        }
      }
    });
    expect(requests.some((request) => request.path === "/planner/tasks/planner-1/details")).toBe(
      true
    );
  });

  it("uses If-Match when updating a planner task", async () => {
    const requests: Array<{ path: string; action: string; payload?: unknown; header?: string }> = [];
    const graphClientFactory = {
      createClient: async () => ({ client: true }),
      request: (_client: unknown, path: string) => createRequest(path, requests)
    };
    const plannerService = new PlannerService(graphClientFactory as never);

    const record = await plannerService.updateTask("planner-1", {
      percentComplete: 50,
      versionTag: "etag-1"
    });

    expect(record.percentComplete).toBe(50);
    expect(
      requests.find((request) => request.path === "/planner/tasks/planner-1" && request.action === "patch")
    ).toMatchObject({
      header: "etag-1"
    });
  });
});

function createRequest(
  path: string,
  requests: Array<{ path: string; action: string; payload?: unknown; header?: string }>
) {
  return {
    header(name: string, value: string) {
      if (name === "If-Match") {
        requests.push({ path, action: "header", header: value });
      }

      return this;
    },
    async get() {
      requests.push({ path, action: "get" });

      if (path.endsWith("/details")) {
        return {
          description: "Existing detail",
          "@odata.etag": "detail-etag"
        };
      }

      return {
        id: "planner-1",
        title: "Launch prep",
        planId: "plan-1",
        bucketId: "bucket-1",
        percentComplete: 10,
        "@odata.etag": "etag-1"
      };
    },
    async post(payload: unknown) {
      requests.push({ path, action: "post", payload });
      return {
        id: "planner-1",
        title: "Launch prep",
        planId: "plan-1",
        bucketId: "bucket-1",
        percentComplete: 0,
        "@odata.etag": "etag-created"
      };
    },
    async patch(payload: unknown) {
      const priorHeader = [...requests].reverse().find(
        (request) => request.path === path && request.action === "header"
      )?.header;
      requests.push({ path, action: "patch", payload, header: priorHeader });
      return {
        id: "planner-1",
        title: "Launch prep",
        planId: "plan-1",
        bucketId: "bucket-1",
        percentComplete: 50,
        "@odata.etag": "etag-2"
      };
    }
  };
}
