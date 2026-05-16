import * as React from "react";
import {
  Combobox,
  Option,
  SpinButton,
  Field,
  Divider,
  Caption1,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import {
  replaceFonts,
  insertCagr,
  getFileSize,
  getUsedFonts,
} from "../../lib/utilities";
import { useActions } from "./ActionContext";
import { ToolButton } from "./ToolButton";

const useStyles = makeStyles({
  section: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
  toolbar: {
    display: "flex",
    flexWrap: "wrap",
    gap: tokens.spacingHorizontalS,
    alignItems: "center",
  },
  grow: { flexGrow: 1, minWidth: "140px" },
  spin: { width: "96px" },
});

/** Common fonts offered in "Replace with" — the JS API cannot list the
 *  system-installed fonts, so this is a curated set; any font may be typed. */
const COMMON_FONTS = [
  "Aptos", "Arial", "Calibri", "Cambria", "Candara", "Comic Sans MS",
  "Consolas", "Constantia", "Corbel", "Courier New", "Franklin Gothic",
  "Garamond", "Georgia", "Helvetica", "Lucida Sans", "Noto Sans",
  "Segoe UI", "Tahoma", "Times New Roman", "Trebuchet MS", "Verdana",
];

function spinValue(value: number | undefined, displayValue: string | undefined): number | null {
  const next = value ?? Number(displayValue);
  return Number.isFinite(next) ? (next as number) : null;
}

export const UtilitiesPanel: React.FC = () => {
  const styles = useStyles();
  const { run, busy } = useActions();
  const [fromFont, setFromFont] = React.useState("");
  const [toFont, setToFont] = React.useState("");
  const [usedFonts, setUsedFonts] = React.useState<string[]>([]);
  const [startValue, setStartValue] = React.useState(100);
  const [endValue, setEndValue] = React.useState(200);
  const [periods, setPeriods] = React.useState(5);
  const [fileSize, setFileSize] = React.useState("");
  const [checkingSize, setCheckingSize] = React.useState(false);

  React.useEffect(() => {
    void getUsedFonts().then(setUsedFonts).catch(() => setUsedFonts([]));
  }, []);

  const replaceWithOptions = React.useMemo(
    () => Array.from(new Set([...usedFonts, ...COMMON_FONTS])).sort((a, b) => a.localeCompare(b)),
    [usedFonts],
  );

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
        <Combobox
          className={styles.grow}
          freeform
          placeholder="Fonts used in this presentation"
          value={fromFont}
          onChange={(ev) => setFromFont(ev.target.value)}
          onOptionSelect={(_, d) => setFromFont(d.optionText ?? "")}
        >
          {usedFonts.map((font) => (
            <Option key={font}>{font}</Option>
          ))}
        </Combobox>
      </Field>
      <Field label="Replace font — with">
        <div className={styles.toolbar}>
          <Combobox
            className={styles.grow}
            freeform
            placeholder="Pick a font or type any name"
            value={toFont}
            onChange={(ev) => setToFont(ev.target.value)}
            onOptionSelect={(_, d) => setToFont(d.optionText ?? "")}
          >
            {replaceWithOptions.map((font) => (
              <Option key={font}>{font}</Option>
            ))}
          </Combobox>
          <ToolButton
            icon="ReplaceFonts"
            label="Replace fonts"
            disabled={busy}
            onClick={() =>
              run("Replace fonts", async () => {
                const count = await replaceFonts(fromFont, toFont);
                if (count === 0) throw new Error("No shapes used that font.");
              })
            }
          />
        </div>
      </Field>
      <Caption1>
        &ldquo;From&rdquo; lists the fonts used in this presentation. The
        system&apos;s installed fonts cannot be listed by the API, so
        &ldquo;with&rdquo; offers common fonts — type any other name to use it.
      </Caption1>

      <Divider />

      <Field label="CAGR — start, end &amp; periods">
        <div className={styles.toolbar}>
          <SpinButton
            className={styles.spin}
            min={0}
            step={10}
            value={startValue}
            onChange={(_, d) => {
              const v = spinValue(d.value ?? undefined, d.displayValue);
              if (v !== null) setStartValue(v);
            }}
          />
          <SpinButton
            className={styles.spin}
            min={0}
            step={10}
            value={endValue}
            onChange={(_, d) => {
              const v = spinValue(d.value ?? undefined, d.displayValue);
              if (v !== null) setEndValue(v);
            }}
          />
          <SpinButton
            className={styles.spin}
            min={1}
            step={1}
            value={periods}
            onChange={(_, d) => {
              const v = spinValue(d.value ?? undefined, d.displayValue);
              if (v !== null) setPeriods(v);
            }}
          />
          <ToolButton
            icon="CAGR"
            label="Insert CAGR"
            disabled={busy}
            onClick={() =>
              run("Insert CAGR", () => insertCagr(startValue, endValue, periods))
            }
          />
        </div>
      </Field>

      <Divider />

      <Field label="Presentation file size">
        <div className={styles.toolbar}>
          <ToolButton
            icon="FileSize"
            label="Check file size"
            disabled={busy || checkingSize}
            onClick={() => void checkFileSize()}
          />
          {fileSize && <Caption1>{fileSize}</Caption1>}
        </div>
      </Field>
    </div>
  );
};
