import { SyncedTaskRecord } from "../services/types";

export function buildCreateTaskCard() {
  return {
    type: "AdaptiveCard",
    version: "1.5",
    body: [
      {
        type: "TextBlock",
        text: "Create synchronized task",
        weight: "Bolder",
        size: "Medium"
      },
      {
        type: "Input.Text",
        id: "title",
        label: "Title",
        isRequired: true,
        errorMessage: "Title is required."
      },
      {
        type: "Input.Text",
        id: "description",
        label: "Description",
        isMultiline: true
      },
      {
        type: "Input.Text",
        id: "planId",
        label: "Planner plan ID",
        isRequired: true,
        errorMessage: "Plan ID is required."
      },
      {
        type: "Input.Text",
        id: "bucketId",
        label: "Planner bucket ID",
        isRequired: true,
        errorMessage: "Bucket ID is required."
      },
      {
        type: "Input.Text",
        id: "assigneeUserId",
        label: "Assignee Entra user ID",
        isRequired: true,
        errorMessage: "Assignee user ID is required."
      },
      {
        type: "Input.Text",
        id: "assigneeTodoListId",
        label: "Optional To Do list ID"
      },
      {
        type: "Input.Text",
        id: "dueDateTime",
        label: "Due date time (ISO-8601 UTC)",
        placeholder: "2026-04-30T18:00:00Z"
      },
      {
        type: "Input.Number",
        id: "percentComplete",
        label: "Percent complete",
        min: 0,
        max: 100,
        value: 0
      }
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "Create task",
        data: {
          command: "create-task"
        }
      }
    ]
  };
}

export function buildUpdateTaskCard(plannerTaskId: string) {
  return {
    type: "AdaptiveCard",
    version: "1.5",
    body: [
      {
        type: "TextBlock",
        text: "Update synchronized task",
        weight: "Bolder",
        size: "Medium"
      },
      {
        type: "Input.Text",
        id: "plannerTaskId",
        label: "Planner task ID",
        value: plannerTaskId,
        isRequired: true,
        errorMessage: "Planner task ID is required."
      },
      {
        type: "Input.Text",
        id: "title",
        label: "Title"
      },
      {
        type: "Input.Text",
        id: "description",
        label: "Description",
        isMultiline: true
      },
      {
        type: "Input.Text",
        id: "dueDateTime",
        label: "Due date time (leave blank to keep current)",
        placeholder: "2026-04-30T18:00:00Z"
      },
      {
        type: "Input.Number",
        id: "percentComplete",
        label: "Percent complete",
        min: 0,
        max: 100
      },
      {
        type: "Input.Text",
        id: "versionTag",
        label: "Optional version tag"
      }
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "Update task",
        data: {
          command: "update-task"
        }
      }
    ]
  };
}

export function buildAssignTaskCard(plannerTaskId: string) {
  return {
    type: "AdaptiveCard",
    version: "1.5",
    body: [
      {
        type: "TextBlock",
        text: "Assign synchronized task",
        weight: "Bolder",
        size: "Medium"
      },
      {
        type: "Input.Text",
        id: "plannerTaskId",
        label: "Planner task ID",
        value: plannerTaskId,
        isRequired: true,
        errorMessage: "Planner task ID is required."
      },
      {
        type: "Input.Text",
        id: "assigneeUserId",
        label: "New assignee Entra user ID",
        isRequired: true,
        errorMessage: "Assignee user ID is required."
      },
      {
        type: "Input.Text",
        id: "assigneeTodoListId",
        label: "Optional To Do list ID"
      },
      {
        type: "Input.Text",
        id: "versionTag",
        label: "Optional version tag"
      }
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "Reassign task",
        data: {
          command: "assign-task"
        }
      }
    ]
  };
}

export function buildTaskSummaryCard(record: SyncedTaskRecord) {
  return {
    type: "AdaptiveCard",
    version: "1.5",
    body: [
      {
        type: "TextBlock",
        text: record.title,
        weight: "Bolder",
        size: "Large"
      },
      {
        type: "FactSet",
        facts: [
          {
            title: "Planner",
            value: record.plannerTaskId
          },
          {
            title: "To Do",
            value: record.todoTaskId
          },
          {
            title: "Assignee",
            value: record.assigneeUserId
          },
          {
            title: "Progress",
            value: `${record.percentComplete}%`
          },
          {
            title: "Due",
            value: record.dueDateTime ?? "Not set"
          }
        ]
      },
      {
        type: "TextBlock",
        text: record.description ?? "No description provided.",
        wrap: true,
        spacing: "Medium"
      }
    ],
    actions: [
      {
        type: "Action.Submit",
        title: "Update",
        data: {
          command: "update-task",
          plannerTaskId: record.plannerTaskId,
          versionTag: record.versionTag
        }
      },
      {
        type: "Action.Submit",
        title: "Assign",
        data: {
          command: "assign-task",
          plannerTaskId: record.plannerTaskId,
          versionTag: record.versionTag
        }
      }
    ]
  };
}
