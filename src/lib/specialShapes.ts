/**
 * Special shapes — Phase 5.
 *
 * Inserts pre-formatted text boxes (title, conclusion, footnote, ghost,
 * label). Positions assume a 960×540 pt slide; the user repositions as
 * needed. The legacy add-in drove these from template settings — a richer
 * settings-driven version can replace these static defaults later.
 */

import { getActiveSlide } from "./powerpoint";

export type SpecialKind = "title" | "conclusion" | "footnote" | "ghost" | "label";

interface SpecialSpec {
  text: string;
  left: number;
  top: number;
  width: number;
  height: number;
  size: number;
  bold: boolean;
}

const SPECS: Record<SpecialKind, SpecialSpec> = {
  title: { text: "Slide title", left: 40, top: 30, width: 880, height: 60, size: 28, bold: true },
  conclusion: { text: "Key takeaway", left: 40, top: 470, width: 880, height: 50, size: 18, bold: true },
  footnote: { text: "Source: …", left: 40, top: 516, width: 880, height: 18, size: 9, bold: false },
  ghost: { text: "Ghost", left: 360, top: 230, width: 240, height: 80, size: 14, bold: false },
  label: { text: "Label", left: 120, top: 120, width: 160, height: 28, size: 12, bold: false },
};

/** Insert a pre-formatted text box of the given kind on the slide in view. */
export async function insertSpecialShape(kind: SpecialKind): Promise<void> {
  await PowerPoint.run(async (context) => {
    const slide = await getActiveSlide(context);
    const spec = SPECS[kind];
    const box = slide.shapes.addTextBox(spec.text, {
      left: spec.left,
      top: spec.top,
      width: spec.width,
      height: spec.height,
    });
    const font = box.textFrame.textRange.font;
    font.size = spec.size;
    font.bold = spec.bold;
    await context.sync();
  });
}
