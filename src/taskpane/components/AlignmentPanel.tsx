import * as React from "react";
import {
  Button,
  Field,
  Radio,
  RadioGroup,
  Caption1,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import {
  align,
  distribute,
  type AlignEdge,
  type AlignTarget,
} from "../../lib/alignment";
import { useActions } from "./ActionContext";

const useStyles = makeStyles({
  grid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr 1fr",
    gap: tokens.spacingHorizontalS,
    marginTop: tokens.spacingVerticalM,
  },
  wide: { gridColumn: "span 3" },
  section: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
});

const ALIGN_BUTTONS: { edge: AlignEdge; label: string }[] = [
  { edge: "left", label: "Left" },
  { edge: "centerH", label: "Center" },
  { edge: "right", label: "Right" },
  { edge: "top", label: "Top" },
  { edge: "middle", label: "Middle" },
  { edge: "bottom", label: "Bottom" },
];

/** Custom UI for the Alignment category — needs a target selector that a
 *  plain button list cannot express. */
export const AlignmentPanel: React.FC = () => {
  const styles = useStyles();
  const { run, busy } = useActions();
  const [target, setTarget] = React.useState<AlignTarget>("selection");

  return (
    <div className={styles.section}>
      <Field label="Align relative to">
        <RadioGroup
          layout="horizontal-stacked"
          value={target}
          onChange={(_, d) => setTarget(d.value as AlignTarget)}
        >
          <Radio value="selection" label="Selection" />
          <Radio value="first" label="First" />
          <Radio value="last" label="Last" />
          <Radio value="slide" label="Slide" />
        </RadioGroup>
      </Field>
      <Caption1>
        Replaces the legacy guide / modifier-key Power Align modes, which the
        cross-platform API does not expose.
      </Caption1>

      <div className={styles.grid}>
        {ALIGN_BUTTONS.map((b) => (
          <Button
            key={b.edge}
            disabled={busy}
            onClick={() => run(`Align ${b.label}`, () => align(b.edge, target))}
          >
            {b.label}
          </Button>
        ))}
        <Button
          className={styles.wide}
          disabled={busy}
          onClick={() =>
            run("Distribute horizontally", () => distribute("horizontal"))
          }
        >
          Distribute Horizontally
        </Button>
        <Button
          className={styles.wide}
          disabled={busy}
          onClick={() =>
            run("Distribute vertically", () => distribute("vertical"))
          }
        >
          Distribute Vertically
        </Button>
      </div>
    </div>
  );
};
