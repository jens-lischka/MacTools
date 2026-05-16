import * as React from "react";
import {
  Button,
  Field,
  Divider,
  Caption1,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import {
  insertShape,
  insertLine,
  insertTextBox,
  insertNumberedCircle,
  insertConnector,
  groupAsLayout,
} from "../../lib/shapes";
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

export const ShapeCreationPanel: React.FC = () => {
  const styles = useStyles();
  const { run, busy } = useActions();

  // Built inside the component so the PowerPoint enum is read after Office.js
  // has loaded, not at module-evaluation time.
  const SHAPES: { label: string; type: PowerPoint.GeometricShapeType }[] = [
    { label: "Rectangle", type: PowerPoint.GeometricShapeType.rectangle },
    { label: "Rounded Rectangle", type: PowerPoint.GeometricShapeType.roundRectangle },
    { label: "Oval", type: PowerPoint.GeometricShapeType.ellipse },
    { label: "Triangle", type: PowerPoint.GeometricShapeType.triangle },
    { label: "Right Arrow", type: PowerPoint.GeometricShapeType.rightArrow },
    { label: "Chevron", type: PowerPoint.GeometricShapeType.chevron },
  ];

  return (
    <div className={styles.section}>
      <Caption1>New shapes are added to the slide in view.</Caption1>
      <div className={styles.grid}>
        {SHAPES.map((s) => (
          <Button
            key={s.label}
            disabled={busy}
            onClick={() => run(`Insert ${s.label.toLowerCase()}`, () => insertShape(s.type))}
          >
            {s.label}
          </Button>
        ))}
        <Button disabled={busy} onClick={() => run("Insert line", () => insertLine())}>
          Line
        </Button>
        <Button
          disabled={busy}
          onClick={() => run("Insert text box", () => insertTextBox())}
        >
          Text Box
        </Button>
        <Button
          disabled={busy}
          onClick={() =>
            run("Insert numbered circle", () => insertNumberedCircle("1"))
          }
        >
          Numbered Circle
        </Button>
        <Button
          disabled={busy}
          onClick={() => run("Connect two shapes", () => insertConnector())}
        >
          Connect Two Shapes
        </Button>
      </div>

      <Divider />

      <Field label="Lay selected shapes out as">
        <div className={styles.grid}>
          <Button
            disabled={busy}
            onClick={() => run("Group as row", () => groupAsLayout("row"))}
          >
            Row
          </Button>
          <Button
            disabled={busy}
            onClick={() =>
              run("Group as column", () => groupAsLayout("column"))
            }
          >
            Column
          </Button>
        </div>
      </Field>
    </div>
  );
};
