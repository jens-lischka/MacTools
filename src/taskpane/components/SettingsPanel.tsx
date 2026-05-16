import * as React from "react";
import { Field, Input, SpinButton, Divider, Caption1 } from "@fluentui/react-components";
import { getSetting, setSetting } from "../../lib/settings";
import { useActions } from "./ActionContext";
import { ToolButton } from "./ToolButton";
import { usePanelStyles } from "./panelStyles";
import { spinHandler } from "./spin";

interface SlideSize {
  width: number;
  height: number;
}

const DEFAULT_SLIDE_SIZE: SlideSize = { width: 960, height: 540 };

export const SettingsPanel: React.FC = () => {
  const styles = usePanelStyles();
  const { run, busy } = useActions();
  const [width, setWidth] = React.useState(DEFAULT_SLIDE_SIZE.width);
  const [height, setHeight] = React.useState(DEFAULT_SLIDE_SIZE.height);
  const [initials, setInitials] = React.useState("");

  React.useEffect(() => {
    let cancelled = false;
    void getSetting<SlideSize>("slideSize", DEFAULT_SLIDE_SIZE).then((s) => {
      if (cancelled) return;
      setWidth(s.width);
      setHeight(s.height);
    });
    void getSetting<string>("stickyNote:initials", "").then((value) => {
      if (!cancelled) setInitials(value);
    });
    return () => {
      cancelled = true;
    };
  }, []);

  return (
    <div className={styles.section}>
      <Caption1>
        Slide size is used by the &ldquo;align / stretch to slide&rdquo; tools.
        It is stored with the presentation.
      </Caption1>
      <Field label="Slide width &amp; height (pt)">
        <div className={styles.toolbar}>
          <SpinButton
            className={styles.spin}
            min={1}
            step={10}
            value={width}
            onChange={spinHandler(setWidth)}
          />
          <SpinButton
            className={styles.spin}
            min={1}
            step={10}
            value={height}
            onChange={spinHandler(setHeight)}
          />
          <ToolButton
            icon="Save"
            label="Save slide size"
            disabled={busy}
            onClick={() =>
              run("Save slide size", () =>
                setSetting("slideSize", { width, height }, "document"),
              )
            }
          />
        </div>
      </Field>

      <Divider />

      <Field label="Default initials (for sticky notes)">
        <div className={styles.toolbar}>
          <Input
            className={styles.grow}
            value={initials}
            placeholder="e.g. JL"
            onChange={(_, d) => setInitials(d.value)}
          />
          <ToolButton
            icon="Save"
            label="Save initials"
            disabled={busy}
            onClick={() =>
              run("Save initials", () =>
                setSetting("stickyNote:initials", initials, "roaming"),
              )
            }
          />
        </div>
      </Field>

      <Divider />

      <Field label="Reset">
        <div className={styles.toolbar}>
          <ToolButton
            icon="Clear"
            label="Reset slide size to default"
            disabled={busy}
            onClick={() =>
              run("Reset slide size", async () => {
                await setSetting("slideSize", DEFAULT_SLIDE_SIZE, "document");
                setWidth(DEFAULT_SLIDE_SIZE.width);
                setHeight(DEFAULT_SLIDE_SIZE.height);
              })
            }
          />
        </div>
      </Field>
    </div>
  );
};
