const features = [
  "Planner and To Do orchestration",
  "Adaptive card task actions in Teams",
  "Azure AD secured configuration",
  "Realtime task status notifications"
];

export function App() {
  return (
    <main className="app-shell">
      <section className="hero">
        <p className="eyebrow">Teams Task Manager</p>
        <h1>Configuration tab for cross-service task orchestration.</h1>
        <p className="lead">
          This tab will manage channel mappings, planner plans, notification preferences,
          and service health for the Teams Task Manager Bot.
        </p>
      </section>

      <section className="feature-grid">
        {features.map((feature) => (
          <article className="feature-card" key={feature}>
            <h2>{feature}</h2>
            <p>Implementation is staged so the bot, Graph sync layer, and tab ship independently.</p>
          </article>
        ))}
      </section>
    </main>
  );
}
