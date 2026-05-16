import * as React from "react";
import { Button, Field, SpinButton, Divider, makeStyles } from "@fluentui/react-components";
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
import { ToolButton } from "./ToolButton";
import { usePanelStyles } from "./panelStyles";
import { spinHandler } from "./spin";

const useTextStyles = makeStyles({
  squareButton: {
    minWidth: "36px",
    maxWidth: "36px",
    width: "36px",
    height: "36px",
    padding: "0",
  },
});

const all = (v: number): Margins => ({ left: v, right: v, top: v, bottom: v });

const MARGIN_PRESETS: { label: string; margins: Margins; icon: string }[] = [
  { label: "No margins", margins: all(0), icon: "SetMarginsNone" },
  { label: "Narrow margins", margins: all(3.6), icon: "SetMarginsNarrow" },
  { label: "Normal margins", margins: { left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 }, icon: "SetMarginsNormal" },
  { label: "Wide margins", margins: all(14.4), icon: "SetMarginsWide" },
];

const STYLES: { label: string; style: TextStyle; icon: string }[] = [
  { label: "Heading style", style: { size: 28, bold: true }, icon: "Heading1Text" },
  { label: "Subheading style", style: { size: 20, bold: true }, icon: "Subheading1Text" },
  { label: "Body style", style: { size: 14, bold: false }, icon: "Body1Text" },
];

const SPECIAL_CHARS = ["—", "–", "•", "→", "←", "↑", "↓", "×", "✓", "€"];

export const TextPanel: React.FC = () => {
  const styles = usePanelStyles();
  const textStyles = useTextStyles();
  const { run, busy } = useActions();
  const [customMargin, setCustomMargin] = React.useState(6);

  // Built inside the component so the PowerPoint enum is read after Office.js
  // has loaded, not at module-evaluation time.
  const ALIGNMENTS: { label: string; value: PowerPoint.ParagraphHorizontalAlignment; icon: string }[] = [
    { label: "Align text left", value: PowerPoint.ParagraphHorizontalAlignment.left, icon: "AlignLeft" },
    { label: "Align text center", value: PowerPoint.ParagraphHorizontalAlignment.center, icon: "AlignCenter" },
    { label: "Align text right", value: PowerPoint.ParagraphHorizontalAlignment.right, icon: "AlignRight" },
    { label: "Justify text", value: PowerPoint.ParagraphHorizontalAlignment.justify, icon: "AlignInGrid" },
  ];

  return (
    <div className={styles.section}>
      <Field label="Autofit">
        <div className={styles.toolbar}>
          <ToolButton
            icon="ResizeShapeToFitTextOn"
            label="Resize shape to fit text"
            disabled={busy}
            onClick={() =>
              run("Resize shape to fit text", () =>
                setAutoSize(PowerPoint.ShapeAutoSize.autoSizeShapeToFitText),
              )
            }
          />
          <ToolButton
            icon="ResizeShapeToFitTextMix"
            label="Turn autofit off"
            disabled={busy}
            onClick={() =>
              run("Autofit off", () =>
                setAutoSize(PowerPoint.ShapeAutoSize.autoSizeNone),
              )
            }
          />
        </div>
      </Field>

      <Field label="Word wrap">
        <div className={styles.toolbar}>
          <ToolButton
            icon="WrapTextOn"
            label="Word wrap on"
            disabled={busy}
            onClick={() => run("Word wrap on", () => setWordWrap(true))}
          />
          <ToolButton
            icon="WrapTextMix"
            label="Word wrap off"
            disabled={busy}
            onClick={() => run("Word wrap off", () => setWordWrap(false))}
          />
        </div>
      </Field>

      <Divider />

      <Field label="Text-box margins">
        <div className={styles.toolbar}>
          {MARGIN_PRESETS.map((p) => (
            <ToolButton
              key={p.label}
              icon={p.icon}
              label={p.label}
              disabled={busy}
              onClick={() => run(p.label, () => setMargins(p.margins))}
            />
          ))}
        </div>
      </Field>
      <Field label="Custom margin, all sides (pt)">
        <div className={styles.toolbar}>
          <SpinButton
            className={styles.spin}
            min={0}
            max={200}
            step={1}
            value={customMargin}
            onChange={spinHandler(setCustomMargin)}
          />
          <ToolButton
            icon="SetMargins"
            label="Apply custom margin"
            disabled={busy}
            onClick={() =>
              run(`Margins — ${customMargin}pt`, () => setMargins(all(customMargin)))
            }
          />
        </div>
      </Field>

      <Divider />

      <Field label="Paragraph alignment">
        <div className={styles.toolbar}>
          {ALIGNMENTS.map((a) => (
            <ToolButton
              key={a.label}
              icon={a.icon}
              label={a.label}
              disabled={busy}
              onClick={() => run(a.label, () => setParagraphAlignment(a.value))}
            />
          ))}
        </div>
      </Field>

      <Divider />

      <Field label="Clean up text">
        <div className={styles.toolbar}>
          <ToolButton
            icon="ClearLineBreaks"
            label="Clear line breaks"
            disabled={busy}
            onClick={() => run("Clear line breaks", () => clearLineBreaks())}
          />
          <ToolButton
            icon="DeleteText"
            label="Clear text"
            disabled={busy}
            onClick={() => run("Clear text", () => clearText())}
          />
        </div>
      </Field>

      <Divider />

      <Field label="Bullets &amp; styles">
        <div className={styles.toolbar}>
          <ToolButton
            icon="FixBullets"
            label="Bullets on"
            disabled={busy}
            onClick={() => run("Bullets on", () => setBullets(true))}
          />
          <ToolButton
            icon="NoBullets"
            label="Bullets off"
            disabled={busy}
            onClick={() => run("Bullets off", () => setBullets(false))}
          />
          {STYLES.map((s) => (
            <ToolButton
              key={s.label}
              icon={s.icon}
              label={s.label}
              disabled={busy}
              onClick={() => run(s.label, () => applyTextStyle(s.style))}
            />
          ))}
        </div>
      </Field>

      <Divider />

      <Field label="Merge / split">
        <div className={styles.toolbar}>
          <ToolButton
            icon="MergeText"
            label="Merge text"
            disabled={busy}
            onClick={() => run("Merge text", () => mergeText())}
          />
          <ToolButton
            icon="SplitText"
            label="Split text"
            disabled={busy}
            onClick={() => run("Split text", () => splitText())}
          />
        </div>
      </Field>

      <Field label="Insert special character (at the text cursor)">
        <div className={styles.toolbar}>
          {SPECIAL_CHARS.map((ch) => (
            <Button
              key={ch}
              className={textStyles.squareButton}
              disabled={busy}
              aria-label={`Insert ${ch}`}
              onClick={() => run(`Insert ${ch}`, () => insertSpecialCharacter(ch))}
            >
              {ch}
            </Button>
          ))}
        </div>
      </Field>
    </div>
  );
};
