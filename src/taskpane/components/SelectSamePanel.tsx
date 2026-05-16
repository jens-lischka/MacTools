import * as React from "react";
import {
  Button,
  Field,
  Divider,
  Caption1,
  Tooltip,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { selectSame, type SameProperty } from "../../lib/selectSame";
import { setSelectedShapesVisible, showAllShapes } from "../../lib/visibility";
import { isApiSupported } from "../../lib/capabilities";
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

const STYLE_PROPS: { property: SameProperty; label: string; icon: string }[] = [
  { property: "fillColor", label: "Fill Colour", icon: "SelectSameFill" },
  { property: "lineColor", label: "Outline Colour", icon: "SelectSameOutline" },
  { property: "lineWeight", label: "Outline Weight", icon: "SelectSameOutlineWeight" },
  { property: "fontName", label: "Font", icon: "SelectSameFontName" },
  { property: "shapeType", label: "Shape Type", icon: "SelectSameType" },
  { property: "size", label: "Size", icon: "SelectSameSize" },
];

const POSITION_PROPS: { property: SameProperty; label: string; icon: string }[] = [
  { property: "positionTop", label: "Top", icon: "SelectSamePositionTop" },
  { property: "positionLeft", label: "Left", icon: "SelectSamePositionLeft" },
  { property: "positionRight", label: "Right", icon: "SelectSamePositionRight" },
  { property: "positionBottom", label: "Bottom", icon: "SelectSamePositionBottom" },
];

export const SelectSamePanel: React.FC = () => {
  const styles = useStyles();
  const { run, busy } = useActions();
  const visibilitySupported = isApiSupported("1.10");

  const buttonFor = (property: SameProperty, label: string, icon: string) => (
    <Button
      key={property}
      disabled={busy}
      icon={<ToolIcon name={icon} />}
      onClick={() =>
        run(`Select same ${label.toLowerCase()}`, () => selectSame(property))
      }
    >
      {label}
    </Button>
  );

  return (
    <div className={styles.section}>
      <Caption1>
        Select one reference shape, then pick a property — every matching shape
        on the current slide is selected.
      </Caption1>
      <Field label="Style &amp; size">
        <div className={styles.grid}>
          {STYLE_PROPS.map((p) => buttonFor(p.property, p.label, p.icon))}
        </div>
      </Field>
      <Field label="Position (matching edge)">
        <div className={styles.grid}>
          {POSITION_PROPS.map((p) => buttonFor(p.property, p.label, p.icon))}
        </div>
      </Field>

      <Divider />

      <Field label="Visibility">
        <Tooltip
          content={
            visibilitySupported
              ? "Hide or show shapes."
              : "Needs PowerPoint API 1.10, which this client does not support."
          }
          relationship="description"
        >
          <div className={styles.grid}>
            <Button
              disabled={busy || !visibilitySupported}
              icon={<ToolIcon name="HideObject" />}
              onClick={() =>
                run("Hide selected shapes", () =>
                  setSelectedShapesVisible(false),
                )
              }
            >
              Hide Selected
            </Button>
            <Button
              disabled={busy || !visibilitySupported}
              icon={<ToolIcon name="ShowAll" />}
              onClick={() => run("Show all shapes", () => showAllShapes())}
            >
              Show All
            </Button>
          </div>
        </Tooltip>
      </Field>
    </div>
  );
};
