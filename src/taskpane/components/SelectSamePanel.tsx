import * as React from "react";
import {
  Button,
  Caption1,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { selectSame, type SameProperty } from "../../lib/selectSame";
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

const PROPERTIES: { property: SameProperty; label: string }[] = [
  { property: "fillColor", label: "Fill Colour" },
  { property: "lineColor", label: "Outline Colour" },
  { property: "lineWeight", label: "Outline Weight" },
  { property: "shapeType", label: "Shape Type" },
  { property: "size", label: "Size" },
];

export const SelectSamePanel: React.FC = () => {
  const styles = useStyles();
  const { run, busy } = useActions();

  return (
    <div className={styles.section}>
      <Caption1>
        Select one reference shape, then pick a property — every matching shape
        on the current slide is selected.
      </Caption1>
      <div className={styles.grid}>
        {PROPERTIES.map((p) => (
          <Button
            key={p.property}
            disabled={busy}
            onClick={() =>
              run(`Select same ${p.label.toLowerCase()}`, () =>
                selectSame(p.property),
              )
            }
          >
            {p.label}
          </Button>
        ))}
      </div>
    </div>
  );
};
