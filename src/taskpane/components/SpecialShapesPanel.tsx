import * as React from "react";
import { Button, Caption1, makeStyles, tokens } from "@fluentui/react-components";
import { insertSpecialShape, type SpecialKind } from "../../lib/specialShapes";
import { useActions } from "./ActionContext";
import { ToolIcon } from "./ToolIcon";

const useStyles = makeStyles({
  section: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
  grid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: tokens.spacingHorizontalS,
  },
});

const KINDS: { kind: SpecialKind; label: string; icon: string }[] = [
  { kind: "title", label: "Slide Title", icon: "SlideTitle" },
  { kind: "conclusion", label: "Conclusion", icon: "Conclusion" },
  { kind: "footnote", label: "Footnote", icon: "Footnote" },
  { kind: "ghost", label: "Ghost", icon: "Ghost" },
  { kind: "label", label: "Label", icon: "Labels" },
];

export const SpecialShapesPanel: React.FC = () => {
  const styles = useStyles();
  const { run, busy } = useActions();

  return (
    <div className={styles.section}>
      <Caption1>Inserts a pre-formatted text box on the slide in view.</Caption1>
      <div className={styles.grid}>
        {KINDS.map((k) => (
          <Button
            key={k.kind}
            disabled={busy}
            icon={<ToolIcon name={k.icon} />}
            onClick={() => run(`Insert ${k.label.toLowerCase()}`, () => insertSpecialShape(k.kind))}
          >
            {k.label}
          </Button>
        ))}
      </div>
    </div>
  );
};
