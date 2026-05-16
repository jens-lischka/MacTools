import * as React from "react";
import { Field, Divider, Caption1, makeStyles, tokens } from "@fluentui/react-components";
import {
  swapPosition,
  swapSize,
  swapFillAndOutline,
  pickUpSizePosition,
  applySizePosition,
} from "../../lib/transform";
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

export const TransformPanel: React.FC = () => {
  const styles = useStyles();
  const { run, busy } = useActions();

  return (
    <div className={styles.section}>
      <Field label="Swap (exactly two shapes)">
        <div className={styles.toolbar}>
          <ToolButton
            icon="SwapPosition"
            label="Swap position"
            disabled={busy}
            onClick={() => run("Swap position", () => swapPosition())}
          />
          <ToolButton
            icon="MatchSize"
            label="Swap size"
            disabled={busy}
            onClick={() => run("Swap size", () => swapSize())}
          />
          <ToolButton
            icon="SwapFillAndOutline"
            label="Swap fill & outline"
            disabled={busy}
            onClick={() => run("Swap fill & outline", () => swapFillAndOutline())}
          />
        </div>
      </Field>

      <Divider />

      <Field label="Pick up &amp; apply size and position">
        <div className={styles.toolbar}>
          <ToolButton
            icon="SizePositionCopy"
            label="Pick up size & position"
            disabled={busy}
            onClick={() => run("Pick up size & position", () => pickUpSizePosition())}
          />
          <ToolButton
            icon="SizePositionPaste"
            label="Apply size & position"
            disabled={busy}
            onClick={() => run("Apply size & position", () => applySizePosition())}
          />
        </div>
      </Field>
      <Caption1>
        Pick Up stores one shape&apos;s size and position; Apply sets it on
        every selected shape. The pickup is remembered across documents.
      </Caption1>
    </div>
  );
};
