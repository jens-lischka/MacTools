import * as React from "react";
import {
  Button,
  Field,
  SpinButton,
  Divider,
  Caption1,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import {
  setAutoSize,
  setWordWrap,
  setMargins,
  setParagraphAlignment,
  clearLineBreaks,
  clearText,
  setBullets,
  applyTextStyle,
  mergeText,
  splitText,
  insertSpecialCharacter,
  type Margins,
  type TextStyle,
} from "../../lib/text";
import { useActions } from "./ActionContext";
import { ToolIcon } from "./ToolIcon";

const useStyles = makeStyles({
  section: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
  row: { display: "flex", gap: tokens.spacingHorizontalS, flexWrap: "wrap" },
  grow: { flexGrow: 1 },
});

const all = (v: number): Margins => ({ left: v, right: v, top: v, bottom: v });

const MARGIN_PRESETS: { label: string; margins: Margins; icon: string }[] = [
  { label: "None", margins: all(0), icon: "SetMarginsNone" },
  { label: "Narrow", margins: all(3.6), icon: "SetMarginsNarrow" },
  { label: "Normal", margins: { left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 }, icon: "SetMarginsNormal" },
  { label: "Wide", margins: all(14.4), icon: "SetMarginsWide" },
];

const STYLES: { label: string; style: TextStyle; icon: string }[] = [
  { label: "Heading", style: { size: 28, bold: true }, icon: "Heading1Text" },
  { label: "Subheading", style: { size: 20, bold: true }, icon: "Subheading1Text" },
  { label: "Body", style: { size: 14, bold: false }, icon: "Body1Text" },
];

const SPECIAL_CHARS = ["—", "–", "•", "→", "←", "↑", "↓", "×", "✓", "€"];

export const TextPanel: React.FC = () => {
  const styles = useStyles();
  const { run, busy } = useActions();
  const [customMargin, setCustomMargin] = React.useState(6);

  // Built inside the component so the PowerPoint enum is read after Office.js
  // has loaded, not at module-evaluation time.
  const ALIGNMENTS: { label: string; value: PowerPoint.ParagraphHorizontalAlignment; icon: string }[] = [
    { label: "Left", value: PowerPoint.ParagraphHorizontalAlignment.left, icon: "AlignLeft" },
    { label: "Center", value: PowerPoint.ParagraphHorizontalAlignment.center, icon: "AlignCenter" },
    { label: "Right", value: PowerPoint.ParagraphHorizontalAlignment.right, icon: "AlignRight" },
    { label: "Justify", value: PowerPoint.ParagraphHorizontalAlignment.justify, icon: "AlignInGrid" },
  ];

  return (
    <div className={styles.section}>
      <Field label="Autofit">
        <div className={styles.row}>
          <Button
            className={styles.grow}
            disabled={busy}
            icon={<ToolIcon name="ResizeShapeToFitTextOn" />}
            onClick={() =>
              run("Resize shape to fit text", () =>
                setAutoSize(PowerPoint.ShapeAutoSize.autoSizeShapeToFitText),
              )
            }
          >
            Shape to fit text
          </Button>
          <Button
            className={styles.grow}
            disabled={busy}
            icon={<ToolIcon name="ResizeShapeToFitTextMix" />}
            onClick={() =>
              run("Autofit off", () =>
                setAutoSize(PowerPoint.ShapeAutoSize.autoSizeNone),
              )
            }
          >
            Off
          </Button>
        </div>
      </Field>

      <Field label="Word wrap">
        <div className={styles.row}>
          <Button
            className={styles.grow}
            disabled={busy}
            icon={<ToolIcon name="WrapTextOn" />}
            onClick={() => run("Word wrap on", () => setWordWrap(true))}
          >
            On
          </Button>
          <Button
            className={styles.grow}
            disabled={busy}
            icon={<ToolIcon name="WrapTextMix" />}
            onClick={() => run("Word wrap off", () => setWordWrap(false))}
          >
            Off
          </Button>
        </div>
      </Field>

      <Divider />

      <Field label="Text-box margins">
        <div className={styles.row}>
          {MARGIN_PRESETS.map((p) => (
            <Button
              key={p.label}
              className={styles.grow}
              disabled={busy}
              icon={<ToolIcon name={p.icon} />}
              onClick={() =>
                run(`Margins — ${p.label}`, () => setMargins(p.margins))
              }
            >
              {p.label}
            </Button>
          ))}
        </div>
      </Field>
      <Field label="Custom margin, all sides (pt)">
        <div className={styles.row}>
          <SpinButton
            className={styles.grow}
            min={0}
            max={200}
            step={1}
            value={customMargin}
            onChange={(_, d) => {
              const next = d.value ?? Number(d.displayValue);
              if (Number.isFinite(next)) setCustomMargin(next as number);
            }}
          />
          <Button
            disabled={busy}
            icon={<ToolIcon name="SetMargins" />}
            onClick={() =>
              run(`Margins — ${customMargin}pt`, () =>
                setMargins(all(customMargin)),
              )
            }
          >
            Apply
          </Button>
        </div>
      </Field>

      <Divider />

      <Field label="Paragraph alignment">
        <div className={styles.row}>
          {ALIGNMENTS.map((a) => (
            <Button
              key={a.label}
              className={styles.grow}
              disabled={busy}
              icon={<ToolIcon name={a.icon} />}
              onClick={() =>
                run(`Align text ${a.label.toLowerCase()}`, () =>
                  setParagraphAlignment(a.value),
                )
              }
            >
              {a.label}
            </Button>
          ))}
        </div>
      </Field>

      <Divider />

      <div className={styles.row}>
        <Button
          className={styles.grow}
          disabled={busy}
          icon={<ToolIcon name="ClearLineBreaks" />}
          onClick={() => run("Clear line breaks", () => clearLineBreaks())}
        >
          Clear Line Breaks
        </Button>
        <Button
          className={styles.grow}
          disabled={busy}
          icon={<ToolIcon name="DeleteText" />}
          onClick={() => run("Clear text", () => clearText())}
        >
          Clear Text
        </Button>
      </div>
      <Caption1>
        Clear Line Breaks joins multi-line text into one line; Clear Text
        empties the selected shapes.
      </Caption1>

      <Divider />

      <Field label="Bullets">
        <div className={styles.row}>
          <Button
            className={styles.grow}
            disabled={busy}
            icon={<ToolIcon name="FixBullets" />}
            onClick={() => run("Bullets on", () => setBullets(true))}
          >
            On
          </Button>
          <Button
            className={styles.grow}
            disabled={busy}
            icon={<ToolIcon name="NoBullets" />}
            onClick={() => run("Bullets off", () => setBullets(false))}
          >
            Off
          </Button>
        </div>
      </Field>

      <Field label="Text style">
        <div className={styles.row}>
          {STYLES.map((s) => (
            <Button
              key={s.label}
              className={styles.grow}
              disabled={busy}
              icon={<ToolIcon name={s.icon} />}
              onClick={() =>
                run(`Apply ${s.label} style`, () => applyTextStyle(s.style))
              }
            >
              {s.label}
            </Button>
          ))}
        </div>
      </Field>

      <Divider />

      <div className={styles.row}>
        <Button
          className={styles.grow}
          disabled={busy}
          icon={<ToolIcon name="MergeText" />}
          onClick={() => run("Merge text", () => mergeText())}
        >
          Merge Text
        </Button>
        <Button
          className={styles.grow}
          disabled={busy}
          icon={<ToolIcon name="SplitText" />}
          onClick={() => run("Split text", () => splitText())}
        >
          Split Text
        </Button>
      </div>

      <Field label="Insert special character (at the text cursor)">
        <div className={styles.row}>
          {SPECIAL_CHARS.map((ch) => (
            <Button
              key={ch}
              disabled={busy}
              onClick={() =>
                run(`Insert ${ch}`, () => insertSpecialCharacter(ch))
              }
            >
              {ch}
            </Button>
          ))}
        </div>
      </Field>
    </div>
  );
};
