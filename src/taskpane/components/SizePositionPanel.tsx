import * as React from "react";
import {
  Field,
  Radio,
  RadioGroup,
  SpinButton,
  Divider,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import {
  matchSize,
  scaleShapes,
  stretchToEdge,
  straightenLines,
  fillGap,
  unifyShapes,
  type MatchDimension,
  type Reference,
  type StretchEdge,
  type StretchTarget,
} from "../../lib/sizePosition";
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
  spin: { width: "96px" },
});

const MATCH: { dimension: MatchDimension; label: string; icon: string }[] = [
  { dimension: "width", label: "Match Width", icon: "MatchWidth" },
  { dimension: "height", label: "Match Height", icon: "MatchHeight" },
  { dimension: "both", label: "Match Size", icon: "MatchSize" },
];

const STRETCH: { edge: StretchEdge; label: string; icon: string }[] = [
  { edge: "left", label: "Stretch Left", icon: "StretchLeft" },
  { edge: "right", label: "Stretch Right", icon: "StretchRight" },
  { edge: "top", label: "Stretch Top", icon: "StretchTop" },
  { edge: "bottom", label: "Stretch Bottom", icon: "StretchBottom" },
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
      <div className={styles.toolbar}>
        {MATCH.map((m) => (
          <ToolButton
            key={m.dimension}
            icon={m.icon}
            label={m.label}
            disabled={busy}
            onClick={() => run(m.label, () => matchSize(m.dimension, reference))}
          />
        ))}
      </div>

      <Divider />

      <Field label="Scale (%)">
        <div className={styles.toolbar}>
          <SpinButton
            className={styles.spin}
            min={1}
            max={1000}
            step={5}
            value={scale}
            onChange={(_, d) => {
              const next = d.value ?? Number(d.displayValue);
              if (Number.isFinite(next)) setScale(next as number);
            }}
          />
          <ToolButton
            icon="ScaleToValue"
            label="Apply scale"
            disabled={busy}
            onClick={() => run(`Scale to ${scale}%`, () => scaleShapes(scale))}
          />
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
      <div className={styles.toolbar}>
        {STRETCH.map((s) => (
          <ToolButton
            key={s.edge}
            icon={s.icon}
            label={s.label}
            disabled={busy}
            onClick={() => run(s.label, () => stretchToEdge(s.edge, stretchTarget))}
          />
        ))}
      </div>

      <Divider />

      <Field label="Close gaps / unify / straighten">
        <div className={styles.toolbar}>
          <ToolButton
            icon="FillHorizontalGap"
            label="Fill horizontal gaps"
            disabled={busy}
            onClick={() => run("Fill horizontal gaps", () => fillGap("horizontal"))}
          />
          <ToolButton
            icon="FillVerticalGap"
            label="Fill vertical gaps"
            disabled={busy}
            onClick={() => run("Fill vertical gaps", () => fillGap("vertical"))}
          />
          <ToolButton
            icon="UnifyCorners"
            label="Unify size"
            disabled={busy}
            onClick={() => run("Unify shapes", () => unifyShapes())}
          />
          <ToolButton
            icon="StraightenLine"
            label="Straighten lines"
            disabled={busy}
            onClick={() => run("Straighten lines", () => straightenLines())}
          />
        </div>
      </Field>
    </div>
  );
};
