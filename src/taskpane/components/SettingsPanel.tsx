import * as React from "react";
import {
  Button,
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
import { ToolIcon } from "./ToolIcon";

const useStyles = makeStyles({
  section: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
  row: { display: "flex", gap: tokens.spacingHorizontalS, flexWrap: "wrap" },
  grow: { flexGrow: 1 },
});

interface SlideSize {
  width: number;
  height: number;
}

const DEFAULT_SLIDE_SIZE: SlideSize = { width: 960, height: 540 };
const SUPPORT_URL = "https://jens-lischka.github.io/MacTools/";

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
      <Field label="Slide width (pt)">
        <SpinButton
          min={1}
          step={10}
          value={width}
          onChange={(_, d) => {
            const v = spinValue(d.value ?? undefined, d.displayValue);
            if (v !== null) setWidth(v);
          }}
        />
      </Field>
      <Field label="Slide height (pt)">
        <div className={styles.row}>
          <SpinButton
            className={styles.grow}
            min={1}
            step={10}
            value={height}
            onChange={(_, d) => {
              const v = spinValue(d.value ?? undefined, d.displayValue);
              if (v !== null) setHeight(v);
            }}
          />
          <Button
            disabled={busy}
            icon={<ToolIcon name="Save" />}
            onClick={() =>
              run("Save slide size", () =>
                setSetting("slideSize", { width, height }, "document"),
              )
            }
          >
            Save
          </Button>
        </div>
      </Field>

      <Divider />

      <Field label="Default initials (for sticky notes)">
        <div className={styles.row}>
          <Input
            className={styles.grow}
            value={initials}
            placeholder="e.g. JL"
            onChange={(_, d) => setInitials(d.value)}
          />
          <Button
            disabled={busy}
            icon={<ToolIcon name="Save" />}
            onClick={() =>
              run("Save initials", () =>
                setSetting("stickyNote:initials", initials, "roaming"),
              )
            }
          >
            Save
          </Button>
        </div>
      </Field>

      <Divider />

      <div className={styles.row}>
        <Button
          className={styles.grow}
          disabled={busy}
          icon={<ToolIcon name="Clear" />}
          onClick={() =>
            run("Reset slide size", async () => {
              await setSetting("slideSize", DEFAULT_SLIDE_SIZE, "document");
              setWidth(DEFAULT_SLIDE_SIZE.width);
              setHeight(DEFAULT_SLIDE_SIZE.height);
            })
          }
        >
          Reset Defaults
        </Button>
        <Button
          className={styles.grow}
          appearance="secondary"
          icon={<ToolIcon name="OnlineHelp" />}
          onClick={() => window.open(SUPPORT_URL, "_blank", "noopener")}
        >
          Online Help
        </Button>
      </div>
    </div>
  );
};
