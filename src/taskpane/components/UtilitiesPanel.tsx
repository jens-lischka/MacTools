import * as React from "react";
import {
  Button,
  Input,
  SpinButton,
  Field,
  Divider,
  Caption1,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { replaceFonts, insertCagr, getFileSize } from "../../lib/utilities";
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

function spinValue(value: number | undefined, displayValue: string | undefined): number | null {
  const next = value ?? Number(displayValue);
  return Number.isFinite(next) ? (next as number) : null;
}

export const UtilitiesPanel: React.FC = () => {
  const styles = useStyles();
  const { run, busy } = useActions();
  const [fromFont, setFromFont] = React.useState("");
  const [toFont, setToFont] = React.useState("");
  const [startValue, setStartValue] = React.useState(100);
  const [endValue, setEndValue] = React.useState(200);
  const [periods, setPeriods] = React.useState(5);
  const [fileSize, setFileSize] = React.useState("");
  const [checkingSize, setCheckingSize] = React.useState(false);

  const checkFileSize = async () => {
    setCheckingSize(true);
    try {
      setFileSize(await getFileSize());
    } catch (err) {
      setFileSize(err instanceof Error ? err.message : String(err));
    } finally {
      setCheckingSize(false);
    }
  };

  return (
    <div className={styles.section}>
      <Field label="Replace font — from (blank = all fonts)">
        <Input
          value={fromFont}
          placeholder="e.g. Calibri"
          onChange={(_, d) => setFromFont(d.value)}
        />
      </Field>
      <Field label="Replace font — to">
        <div className={styles.row}>
          <Input
            className={styles.grow}
            value={toFont}
            placeholder="e.g. Arial"
            onChange={(_, d) => setToFont(d.value)}
          />
          <Button
            disabled={busy}
            onClick={() =>
              run("Replace fonts", async () => {
                const count = await replaceFonts(fromFont, toFont);
                if (count === 0) throw new Error("No shapes used that font.");
              })
            }
          >
            Replace
          </Button>
        </div>
      </Field>
      <Caption1>Replaces the font across the whole presentation.</Caption1>

      <Divider />

      <Field label="CAGR — start value">
        <SpinButton
          min={0}
          step={10}
          value={startValue}
          onChange={(_, d) => {
            const v = spinValue(d.value ?? undefined, d.displayValue);
            if (v !== null) setStartValue(v);
          }}
        />
      </Field>
      <Field label="CAGR — end value">
        <SpinButton
          min={0}
          step={10}
          value={endValue}
          onChange={(_, d) => {
            const v = spinValue(d.value ?? undefined, d.displayValue);
            if (v !== null) setEndValue(v);
          }}
        />
      </Field>
      <Field label="CAGR — number of periods">
        <div className={styles.row}>
          <SpinButton
            className={styles.grow}
            min={1}
            step={1}
            value={periods}
            onChange={(_, d) => {
              const v = spinValue(d.value ?? undefined, d.displayValue);
              if (v !== null) setPeriods(v);
            }}
          />
          <Button
            disabled={busy}
            onClick={() =>
              run("Insert CAGR", () =>
                insertCagr(startValue, endValue, periods),
              )
            }
          >
            Insert
          </Button>
        </div>
      </Field>

      <Divider />

      <Field label="Presentation file size">
        <div className={styles.row}>
          <Button
            className={styles.grow}
            disabled={busy || checkingSize}
            onClick={() => void checkFileSize()}
          >
            Check File Size
          </Button>
          {fileSize && <Caption1>{fileSize}</Caption1>}
        </div>
      </Field>
    </div>
  );
};
