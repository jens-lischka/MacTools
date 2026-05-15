import * as React from "react";
import {
  Button,
  Field,
  Radio,
  RadioGroup,
  SpinButton,
  Divider,
  Caption1,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import {
  matchSize,
  scaleShapes,
  stretchToEdge,
  straightenLines,
  type MatchDimension,
  type Reference,
  type StretchEdge,
  type StretchTarget,
} from "../../lib/sizePosition";
import { useActions } from "./ActionContext";

const useStyles = makeStyles({
  section: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
  row: {
    display: "flex",
    gap: tokens.spacingHorizontalS,
    flexWrap: "wrap",
  },
  grow: { flexGrow: 1 },
});

const MATCH: { dimension: MatchDimension; label: string }[] = [
  { dimension: "width", label: "Width" },
  { dimension: "height", label: "Height" },
  { dimension: "both", label: "Both" },
];

const STRETCH: { edge: StretchEdge; label: string }[] = [
  { edge: "left", label: "Left" },
  { edge: "right", label: "Right" },
  { edge: "top", label: "Top" },
  { edge: "bottom", label: "Bottom" },
];

export const SizePositionPanel: React.FC = () => {
  const styles = useStyles();
  const { run, busy } = useActions();
  const [reference, setReference] = React.useState<Reference>("last");
  const [stretchTarget, setStretchTarget] = React.useState<StretchTarget>("selection");
  const [scale, setScale] = React.useState(100);

  return (
    <div className={styles.section}>
      <Field label="Match size — reference shape">
        <RadioGroup
          layout="horizontal"
          value={reference}
          onChange={(_, d) => setReference(d.value as Reference)}
        >
          <Radio value="first" label="First selected" />
          <Radio value="last" label="Last selected" />
        </RadioGroup>
      </Field>
      <div className={styles.row}>
        {MATCH.map((m) => (
          <Button
            key={m.dimension}
            className={styles.grow}
            disabled={busy}
            onClick={() =>
              run(`Match ${m.label.toLowerCase()}`, () =>
                matchSize(m.dimension, reference),
              )
            }
          >
            {m.label}
          </Button>
        ))}
      </div>

      <Divider />

      <Field label="Scale (%)">
        <div className={styles.row}>
          <SpinButton
            className={styles.grow}
            min={1}
            max={1000}
            step={5}
            value={scale}
            onChange={(_, d) => {
              const next = d.value ?? Number(d.displayValue);
              if (Number.isFinite(next)) setScale(next as number);
            }}
          />
          <Button
            disabled={busy}
            onClick={() => run(`Scale to ${scale}%`, () => scaleShapes(scale))}
          >
            Apply
          </Button>
        </div>
      </Field>

      <Divider />

      <Field label="Stretch to edge of">
        <RadioGroup
          layout="horizontal"
          value={stretchTarget}
          onChange={(_, d) => setStretchTarget(d.value as StretchTarget)}
        >
          <Radio value="selection" label="Selection" />
          <Radio value="slide" label="Slide" />
        </RadioGroup>
      </Field>
      <div className={styles.row}>
        {STRETCH.map((s) => (
          <Button
            key={s.edge}
            className={styles.grow}
            disabled={busy}
            onClick={() =>
              run(`Stretch ${s.label.toLowerCase()}`, () =>
                stretchToEdge(s.edge, stretchTarget),
              )
            }
          >
            {s.label}
          </Button>
        ))}
      </div>

      <Divider />

      <Button
        disabled={busy}
        onClick={() => run("Straighten lines", () => straightenLines())}
      >
        Straighten Lines
      </Button>
      <Caption1>
        Flattens selected line shapes to be perfectly horizontal or vertical.
      </Caption1>
    </div>
  );
};
