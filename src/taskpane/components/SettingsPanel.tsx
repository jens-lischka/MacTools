import * as React from "react";
import {
  Field,
  Input,
  SpinButton,
  Divider,
  Caption1,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { getSetting, setSetting } from "../../lib/settings";
import { useActions } from "./ActionContext";
import { ToolButton } from "./ToolButton";

const useStyles = makeStyles({
  section: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
  toolbar: {
    display: "flex",
    flexWrap: "wrap",
    gap: tokens.spacingHorizontalS,
    alignItems: "center",
  },
  grow: { flexGrow: 1 },
  spin: { width: "96px" },
});

interface SlideSize {
  width: number;
  height: number;
}

const DEFAULT_SLIDE_SIZE: SlideSize = { width: 960, height: 540 };

function spinValue(value: number | undefined, displayValue: string | undefined): number | null {
  const next = value ?? Number(displayValue);
  return Number.isFinite(next) ? (next as number) : null;
}

export const SettingsPanel: React.FC = () => {
  const styles = useStyles();
  const { run, busy } = useActions();
  const [width, setWidth] = React.useState(DEFAULT_SLIDE_SIZE.width);
  const [height, setHeight] = React.useState(DEFAULT_SLIDE_SIZE.height);
  const [initials, setInitials] = React.useState("");

  React.useEffect(() => {
    void getSetting<SlideSize>("slideSize", DEFAULT_SLIDE_SIZE).then((s) => {
      setWidth(s.width);
      setHeight(s.height);
    });
    void getSetting<string>("stickyNote:initials", "").then(setInitials);
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
            onChange={(_, d) => {
              const v = spinValue(d.value ?? undefined, d.displayValue);
              if (v !== null) setWidth(v);
            }}
          />
          <SpinButton
            className={styles.spin}
            min={1}
            step={10}
            value={height}
            onChange={(_, d) => {
              const v = spinValue(d.value ?? undefined, d.displayValue);
              if (v !== null) setHeight(v);
            }}
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
