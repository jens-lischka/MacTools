import * as React from "react";
import { Field, Divider, Caption1, makeStyles, tokens } from "@fluentui/react-components";
import {
  insertShape,
  insertLine,
  insertTextBox,
  insertNumberedCircle,
  insertConnector,
  groupAsLayout,
} from "../../lib/shapes";
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

export const ShapeCreationPanel: React.FC = () => {
  const styles = useStyles();
  const { run, busy } = useActions();

  // Built inside the component so the PowerPoint enum is read after Office.js
  // has loaded, not at module-evaluation time.
  const SHAPES: { label: string; type: PowerPoint.GeometricShapeType; icon: string }[] = [
    { label: "Rectangle", type: PowerPoint.GeometricShapeType.rectangle, icon: "ShapesRectangle" },
    { label: "Rounded Rectangle", type: PowerPoint.GeometricShapeType.roundRectangle, icon: "ShapesRoundedRectangle" },
    { label: "Oval", type: PowerPoint.GeometricShapeType.ellipse, icon: "ShapesOval" },
    { label: "Triangle", type: PowerPoint.GeometricShapeType.triangle, icon: "ShapesLargeCaret" },
    { label: "Right Arrow", type: PowerPoint.GeometricShapeType.rightArrow, icon: "SymbolsArrowRight" },
    { label: "Chevron", type: PowerPoint.GeometricShapeType.chevron, icon: "ShapesChevron1" },
  ];

  return (
    <div className={styles.section}>
      <Caption1>New shapes are added to the slide in view.</Caption1>
      <Field label="Insert">
        <div className={styles.toolbar}>
          {SHAPES.map((s) => (
            <ToolButton
              key={s.label}
              icon={s.icon}
              label={`Insert ${s.label}`}
              disabled={busy}
              onClick={() => run(`Insert ${s.label.toLowerCase()}`, () => insertShape(s.type))}
            />
          ))}
          <ToolButton
            icon="ShapesLine"
            label="Insert line"
            disabled={busy}
            onClick={() => run("Insert line", () => insertLine())}
          />
          <ToolButton
            icon="InsertTextBox"
            label="Insert text box"
            disabled={busy}
            onClick={() => run("Insert text box", () => insertTextBox())}
          />
          <ToolButton
            icon="ShapesCircle"
            label="Insert numbered circle"
            disabled={busy}
            onClick={() => run("Insert numbered circle", () => insertNumberedCircle("1"))}
          />
          <ToolButton
            icon="ConnectObjects"
            label="Connect two shapes"
            disabled={busy}
            onClick={() => run("Connect two shapes", () => insertConnector())}
          />
        </div>
      </Field>

      <Divider />

      <Field label="Lay selected shapes out as">
        <div className={styles.toolbar}>
          <ToolButton
            icon="GroupAsRows"
            label="Lay out as a row"
            disabled={busy}
            onClick={() => run("Group as row", () => groupAsLayout("row"))}
          />
          <ToolButton
            icon="GroupAsColumns"
            label="Lay out as a column"
            disabled={busy}
            onClick={() => run("Group as column", () => groupAsLayout("column"))}
          />
        </div>
      </Field>
    </div>
  );
};
