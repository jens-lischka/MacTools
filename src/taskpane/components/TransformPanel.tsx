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
  swapPosition,
  swapSize,
  swapFillAndOutline,
  pickUpSizePosition,
  applySizePosition,
} from "../../lib/transform";
import { useActions } from "./ActionContext";

const useStyles = makeStyles({
  section: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
  row: { display: "flex", gap: tokens.spacingHorizontalS, flexWrap: "wrap" },
  grow: { flexGrow: 1 },
});

export const TransformPanel: React.FC = () => {
  const styles = useStyles();
  const { run, busy } = useActions();

  return (
    <div className={styles.section}>
      <Field label="Swap (exactly two shapes)">
        <div className={styles.row}>
          <Button
            className={styles.grow}
            disabled={busy}
            onClick={() => run("Swap position", () => swapPosition())}
          >
            Swap Position
          </Button>
          <Button
            className={styles.grow}
            disabled={busy}
            onClick={() => run("Swap size", () => swapSize())}
          >
            Swap Size
          </Button>
          <Button
            className={styles.grow}
            disabled={busy}
            onClick={() =>
              run("Swap fill & outline", () => swapFillAndOutline())
            }
          >
            Swap Fill &amp; Outline
          </Button>
        </div>
      </Field>

      <Divider />

      <Field label="Pick up &amp; apply size and position">
        <div className={styles.row}>
          <Button
            className={styles.grow}
            disabled={busy}
            onClick={() => run("Pick up size & position", () => pickUpSizePosition())}
          >
            Pick Up
          </Button>
          <Button
            className={styles.grow}
            disabled={busy}
            onClick={() => run("Apply size & position", () => applySizePosition())}
          >
            Apply
          </Button>
        </div>
      </Field>
      <Caption1>
        Pick Up stores one shape&apos;s size and position; Apply sets it on
        every selected shape. The pickup is remembered across documents.
      </Caption1>
    </div>
  );
};
