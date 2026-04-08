import { FormEvent, startTransition, useEffect, useState } from "react";
import { createTask, fetchHealth, fetchTasks } from "./api";
import { CreateTaskPayload, ServiceHealth, SyncedTaskRecord } from "./types";

const initialForm: CreateTaskPayload = {
  title: "",
  description: "",
  planId: "",
  bucketId: "",
  assigneeUserId: "",
  assigneeTodoListId: "",
  dueDateTime: "",
  percentComplete: 0,
  teamId: "",
  channelId: ""
};

export function App() {
  const [health, setHealth] = useState<ServiceHealth | null>(null);
  const [tasks, setTasks] = useState<SyncedTaskRecord[]>([]);
  const [form, setForm] = useState<CreateTaskPayload>(initialForm);
  const [submitting, setSubmitting] = useState(false);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    void refresh();
  }, []);

  async function refresh() {
    setLoading(true);
    setError(null);

    try {
      const [nextHealth, nextTasks] = await Promise.all([fetchHealth(), fetchTasks()]);

      startTransition(() => {
        setHealth(nextHealth);
        setTasks(nextTasks);
      });
    } catch (refreshError) {
      setError(
        refreshError instanceof Error
          ? refreshError.message
          : "Unable to load Teams Task Manager status."
      );
    } finally {
      setLoading(false);
    }
  }

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setSubmitting(true);
    setError(null);

    try {
      const createdTask = await createTask({
        ...form,
        description: form.description || undefined,
        assigneeTodoListId: form.assigneeTodoListId || undefined,
        dueDateTime: form.dueDateTime ? new Date(form.dueDateTime).toISOString() : undefined,
        percentComplete: Number(form.percentComplete ?? 0)
      });

      startTransition(() => {
        setTasks((current) => [createdTask, ...current]);
        setHealth((current) =>
          current
            ? {
                ...current,
                syncedTasks: current.syncedTasks + 1
              }
            : current
        );
        setForm(initialForm);
      });
    } catch (submitError) {
      setError(
        submitError instanceof Error ? submitError.message : "Unable to create synchronized task."
      );
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <main className="app-shell">
      <section className="hero">
        <div>
          <p className="eyebrow">Teams Task Manager</p>
          <h1>Planner, To Do, and Teams orchestration from one control tab.</h1>
          <p className="lead">
            Configure task routing, inspect sync health, and create channel-bound tasks without
            leaving Microsoft Teams.
          </p>
        </div>

        <div className="hero-metrics">
          <Metric label="Service" value={health?.service ?? "Loading"} />
          <Metric label="Mode" value={health?.mode ?? "Loading"} />
          <Metric label="Synced tasks" value={String(health?.syncedTasks ?? tasks.length)} />
        </div>
      </section>

      {error ? <p className="error-banner">{error}</p> : null}

      <section className="dashboard-grid">
        <article className="panel panel-form">
          <div className="panel-header">
            <p className="panel-kicker">Create</p>
            <h2>New synchronized task</h2>
          </div>

          <form className="task-form" onSubmit={handleSubmit}>
            <label>
              <span>Title</span>
              <input
                value={form.title}
                onChange={(event) => setForm({ ...form, title: event.target.value })}
                required
              />
            </label>
            <label>
              <span>Description</span>
              <textarea
                rows={4}
                value={form.description}
                onChange={(event) => setForm({ ...form, description: event.target.value })}
              />
            </label>
            <div className="form-row">
              <label>
                <span>Plan ID</span>
                <input
                  value={form.planId}
                  onChange={(event) => setForm({ ...form, planId: event.target.value })}
                  required
                />
              </label>
              <label>
                <span>Bucket ID</span>
                <input
                  value={form.bucketId}
                  onChange={(event) => setForm({ ...form, bucketId: event.target.value })}
                  required
                />
              </label>
            </div>
            <div className="form-row">
              <label>
                <span>Assignee user ID</span>
                <input
                  value={form.assigneeUserId}
                  onChange={(event) => setForm({ ...form, assigneeUserId: event.target.value })}
                  required
                />
              </label>
              <label>
                <span>To Do list ID</span>
                <input
                  value={form.assigneeTodoListId}
                  onChange={(event) =>
                    setForm({ ...form, assigneeTodoListId: event.target.value })
                  }
                />
              </label>
            </div>
            <div className="form-row">
              <label>
                <span>Due time</span>
                <input
                  type="datetime-local"
                  value={form.dueDateTime}
                  onChange={(event) => setForm({ ...form, dueDateTime: event.target.value })}
                />
              </label>
              <label>
                <span>Percent complete</span>
                <input
                  type="number"
                  min={0}
                  max={100}
                  value={form.percentComplete}
                  onChange={(event) =>
                    setForm({ ...form, percentComplete: Number(event.target.value) })
                  }
                />
              </label>
            </div>
            <div className="form-row">
              <label>
                <span>Team ID</span>
                <input
                  value={form.teamId}
                  onChange={(event) => setForm({ ...form, teamId: event.target.value })}
                  required
                />
              </label>
              <label>
                <span>Channel ID</span>
                <input
                  value={form.channelId}
                  onChange={(event) => setForm({ ...form, channelId: event.target.value })}
                  required
                />
              </label>
            </div>

            <div className="form-actions">
              <button type="submit" disabled={submitting}>
                {submitting ? "Creating..." : "Create synced task"}
              </button>
              <button type="button" className="secondary" onClick={() => void refresh()}>
                {loading ? "Refreshing..." : "Refresh"}
              </button>
            </div>
          </form>
        </article>

        <article className="panel panel-status">
          <div className="panel-header">
            <p className="panel-kicker">Flow</p>
            <h2>Task routing map</h2>
          </div>

          <div className="flow-steps">
            <div>
              <strong>1.</strong>
              <span>Teams command or adaptive card submits task metadata.</span>
            </div>
            <div>
              <strong>2.</strong>
              <span>Backend acquires Azure AD Graph token and writes Planner state.</span>
            </div>
            <div>
              <strong>3.</strong>
              <span>Matching To Do task is created or updated for the assignee.</span>
            </div>
            <div>
              <strong>4.</strong>
              <span>Bot posts a summary card back into the channel.</span>
            </div>
          </div>
        </article>
      </section>

      <section className="task-board">
        <div className="panel-header">
          <p className="panel-kicker">State</p>
          <h2>Recent synchronized tasks</h2>
        </div>

        <div className="task-grid">
          {tasks.map((task) => (
            <article className="task-card" key={task.plannerTaskId}>
              <header>
                <h3>{task.title}</h3>
                <span>{task.percentComplete}%</span>
              </header>
              <p>{task.description || "No description provided."}</p>
              <dl>
                <div>
                  <dt>Planner</dt>
                  <dd>{task.plannerTaskId}</dd>
                </div>
                <div>
                  <dt>To Do</dt>
                  <dd>{task.todoTaskId}</dd>
                </div>
                <div>
                  <dt>Assignee</dt>
                  <dd>{task.assigneeUserId}</dd>
                </div>
                <div>
                  <dt>Due</dt>
                  <dd>{task.dueDateTime || "Not set"}</dd>
                </div>
              </dl>
            </article>
          ))}

          {!tasks.length && !loading ? (
            <article className="task-card empty-state">
              <h3>No synchronized tasks yet</h3>
              <p>Create a task here or use the Teams bot command `task create`.</p>
            </article>
          ) : null}
        </div>
      </section>
    </main>
  );
}

function Metric({ label, value }: { label: string; value: string }) {
  return (
    <div className="metric-card">
      <span>{label}</span>
      <strong>{value}</strong>
    </div>
  );
}
