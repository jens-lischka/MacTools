import * as React from "react";
import { Button, Caption1, makeStyles, tokens } from "@fluentui/react-components";
import { insertSpecialShape, type SpecialKind } from "../../lib/specialShapes";
import { useActions } from "./ActionContext";

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

const KINDS: { kind: SpecialKind; label: string }[] = [
  { kind: "title", label: "Slide Title" },
  { kind: "conclusion", label: "Conclusion" },
  { kind: "footnote", label: "Footnote" },
  { kind: "ghost", label: "Ghost" },
  { kind: "label", label: "Label" },
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
            onClick={() => run(`Insert ${k.label.toLowerCase()}`, () => insertSpecialShape(k.kind))}
          >
            {k.label}
          </Button>
        ))}
      </div>
    </div>
  );
};
