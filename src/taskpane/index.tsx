import { createRoot } from "react-dom/client";
import App from "./components/App";

/* The task pane only renders once Office.js has initialised and confirmed
 * the host is PowerPoint. */
Office.onReady((info) => {
  const container = document.getElementById("root");
  if (!container) return;
  const root = createRoot(container);

  if (info.host === Office.HostType.PowerPoint) {
    root.render(<App />);
  } else {
    root.render(
      <p style={{ font: "14px Segoe UI", padding: 16 }}>
        MacTools runs in PowerPoint only.
      </p>,
    );
  }
});
