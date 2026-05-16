import * as React from "react";
import { Field, Divider, Caption1, makeStyles, tokens } from "@fluentui/react-components";
import { selectSame, type SameProperty } from "../../lib/selectSame";
import { setSelectedShapesVisible, showAllShapes } from "../../lib/visibility";
import { isApiSupported } from "../../lib/capabilities";
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

const STYLE_PROPS: { property: SameProperty; label: string; icon: string }[] = [
  { property: "fillColor", label: "Same Fill Colour", icon: "SelectSameFill" },
  { property: "lineColor", label: "Same Outline Colour", icon: "SelectSameOutline" },
  { property: "lineWeight", label: "Same Outline Weight", icon: "SelectSameOutlineWeight" },
  { property: "fontName", label: "Same Font", icon: "SelectSameFontName" },
  { property: "shapeType", label: "Same Shape Type", icon: "SelectSameType" },
  { property: "size", label: "Same Size", icon: "SelectSameSize" },
];

const POSITION_PROPS: { property: SameProperty; label: string; icon: string }[] = [
  { property: "positionTop", label: "Same Top Edge", icon: "SelectSamePositionTop" },
  { property: "positionLeft", label: "Same Left Edge", icon: "SelectSamePositionLeft" },
  { property: "positionRight", label: "Same Right Edge", icon: "SelectSamePositionRight" },
  { property: "positionBottom", label: "Same Bottom Edge", icon: "SelectSamePositionBottom" },
];

export const SelectSamePanel: React.FC = () => {
  const styles = useStyles();
  const { run, busy } = useActions();
  const visibilitySupported = isApiSupported("1.10");

  const toolButton = (property: SameProperty, label: string, icon: string) => (
    <ToolButton
      key={property}
      icon={icon}
      label={label}
      disabled={busy}
      onClick={() => run(label, () => selectSame(property))}
    />
  );

  return (
    <div className={styles.section}>
      <Caption1>
        Select one reference shape, then pick a property — every matching shape
        on the current slide is selected.
      </Caption1>
      <Field label="Style &amp; size">
        <div className={styles.toolbar}>
          {STYLE_PROPS.map((p) => toolButton(p.property, p.label, p.icon))}
        </div>
      </Field>
      <Field label="Position (matching edge)">
        <div className={styles.toolbar}>
          {POSITION_PROPS.map((p) => toolButton(p.property, p.label, p.icon))}
        </div>
      </Field>

      <Divider />

      <Field label="Visibility">
        <div className={styles.toolbar}>
          <ToolButton
            icon="HideObject"
            label={
              visibilitySupported
                ? "Hide selected shapes"
                : "Hide shapes — needs PowerPoint API 1.10"
            }
            disabled={busy || !visibilitySupported}
            onClick={() =>
              run("Hide selected shapes", () => setSelectedShapesVisible(false))
            }
          />
          <ToolButton
            icon="ShowAll"
            label={
              visibilitySupported
                ? "Show all shapes"
                : "Show all — needs PowerPoint API 1.10"
            }
            disabled={busy || !visibilitySupported}
            onClick={() => run("Show all shapes", () => showAllShapes())}
          />
        </div>
      </Field>
    </div>
  );
};
