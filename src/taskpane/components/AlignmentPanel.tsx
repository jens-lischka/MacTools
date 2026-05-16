import * as React from "react";
import {
  Field,
  Radio,
  RadioGroup,
  Caption1,
} from "@fluentui/react-components";
import {
  align,
  distribute,
  type AlignEdge,
  type AlignTarget,
} from "../../lib/alignment";
import { useActions } from "./ActionContext";
import { ToolButton } from "./ToolButton";
import { usePanelStyles } from "./panelStyles";


const ALIGN_BUTTONS: { edge: AlignEdge; label: string; icon: string }[] = [
  { edge: "left", label: "Align Left", icon: "AlignLeft" },
  { edge: "centerH", label: "Align Center", icon: "AlignCenter" },
  { edge: "right", label: "Align Right", icon: "AlignRight" },
  { edge: "top", label: "Align Top", icon: "AlignTop" },
  { edge: "middle", label: "Align Middle", icon: "AlignMiddle" },
  { edge: "bottom", label: "Align Bottom", icon: "AlignBottom" },
];

export const AlignmentPanel: React.FC = () => {
  const styles = usePanelStyles();
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

      <Field label="Align">
        <div className={styles.toolbar}>
          {ALIGN_BUTTONS.map((b) => (
            <ToolButton
              key={b.edge}
              icon={b.icon}
              label={b.label}
              disabled={busy}
              onClick={() => run(b.label, () => align(b.edge, target))}
            />
          ))}
        </div>
      </Field>

      <Field label="Distribute">
        <div className={styles.toolbar}>
          <ToolButton
            icon="DistributeHorizontally"
            label="Distribute Horizontally"
            disabled={busy}
            onClick={() =>
              run("Distribute horizontally", () => distribute("horizontal"))
            }
          />
          <ToolButton
            icon="DistributeVertically"
            label="Distribute Vertically"
            disabled={busy}
            onClick={() =>
              run("Distribute vertically", () => distribute("vertical"))
            }
          />
        </div>
      </Field>

      <Caption1>
        Power Align&apos;s guide / modifier-key modes are replaced by the
        &ldquo;align relative to&rdquo; choice above.
      </Caption1>
    </div>
  );
};
