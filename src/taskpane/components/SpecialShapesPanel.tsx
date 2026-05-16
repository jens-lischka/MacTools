import * as React from "react";
import { Caption1, makeStyles, tokens } from "@fluentui/react-components";
import { insertSpecialShape, type SpecialKind } from "../../lib/specialShapes";
import { useActions } from "./ActionContext";
import { ToolButton } from "./ToolButton";

const useStyles = makeStyles({
  section: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
  toolbar: { display: "flex", flexWrap: "wrap", gap: tokens.spacingHorizontalS },
});

const KINDS: { kind: SpecialKind; label: string; icon: string }[] = [
  { kind: "title", label: "Insert Slide Title", icon: "SlideTitle" },
  { kind: "conclusion", label: "Insert Conclusion", icon: "Conclusion" },
  { kind: "footnote", label: "Insert Footnote", icon: "Footnote" },
  { kind: "ghost", label: "Insert Ghost", icon: "Ghost" },
  { kind: "label", label: "Insert Label", icon: "Labels" },
];

export const SpecialShapesPanel: React.FC = () => {
  const styles = useStyles();
  const { run, busy } = useActions();

  return (
    <div className={styles.section}>
      <Caption1>Inserts a pre-formatted text box on the slide in view.</Caption1>
      <div className={styles.toolbar}>
        {KINDS.map((k) => (
          <ToolButton
            key={k.kind}
            icon={k.icon}
            label={k.label}
            disabled={busy}
            onClick={() => run(k.label, () => insertSpecialShape(k.kind))}
          />
        ))}
      </div>
    </div>
  );
};
