/**
 * Utility operations — Phase 4.
 */

/** True for shape types that carry a text frame. A function (not a
 *  module-level constant) so the `PowerPoint` enum is read lazily. */
function isTextShape(type: PowerPoint.Shape["type"]): boolean {
  return (
    type === PowerPoint.ShapeType.geometricShape ||
    type === PowerPoint.ShapeType.textBox ||
    type === PowerPoint.ShapeType.placeholder
  );
}

/**
 * Replace a font across the whole presentation. When `fromFont` is blank,
 * every text shape is switched to `toFont`. Returns the number of shapes
 * changed.
 */
export async function replaceFonts(
  fromFont: string,
  toFont: string,
): Promise<number> {
  const target = toFont.trim();
  if (!target) throw new Error("Enter the replacement font name.");
  const source = fromFont.trim();

  let changed = 0;
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    const shapeCollections = slides.items.map((slide) => {
      const shapes = slide.shapes;
      shapes.load("items/type");
      return shapes;
    });
    await context.sync();

    const fonts: PowerPoint.ShapeFont[] = [];
    shapeCollections.forEach((shapes) => {
      shapes.items
        .filter((s) => isTextShape(s.type))
        .forEach((s) => {
          const font = s.textFrame.textRange.font;
          font.load("name");
          fonts.push(font);
        });
    });
    await context.sync();

    fonts.forEach((font) => {
      if (!source || font.name === source) {
        font.name = target;
        changed += 1;
      }
    });
    await context.sync();
  });
  return changed;
}

/** Distinct font names used by text shapes across the presentation, sorted. */
export async function getUsedFonts(): Promise<string[]> {
  const names = new Set<string>();
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    const shapeCollections = slides.items.map((slide) => {
      const shapes = slide.shapes;
      shapes.load("items/type");
      return shapes;
    });
    await context.sync();

    const fonts: PowerPoint.ShapeFont[] = [];
    shapeCollections.forEach((shapes) => {
      shapes.items
        .filter((s) => isTextShape(s.type))
        .forEach((s) => {
          const font = s.textFrame.textRange.font;
          font.load("name");
          fonts.push(font);
        });
    });
    await context.sync();

    fonts.forEach((font) => {
      if (font.name) names.add(font.name);
    });
  });
  return [...names].sort((a, b) => a.localeCompare(b));
}

/** Return the current presentation's file size as a human-readable string. */
export async function getFileSize(): Promise<string> {
  return new Promise<string>((resolve, reject) => {
    Office.context.document.getFileAsync(Office.FileType.Compressed, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reject(result.error);
        return;
      }
      const file = result.value;
      const bytes = file.size;
      file.closeAsync(() => undefined);
      const mb = bytes / (1024 * 1024);
      resolve(
        mb >= 1 ? `${mb.toFixed(2)} MB` : `${(bytes / 1024).toFixed(0)} KB`,
      );
    });
  });
}

/**
 * Compute the compound annual growth rate and insert it as a text box on the
 * slide in view.
 */
export async function insertCagr(
  startValue: number,
  endValue: number,
  periods: number,
): Promise<void> {
  if (!(startValue > 0) || !(endValue > 0) || !(periods > 0)) {
    throw new Error("Start value, end value and periods must all be positive.");
  }
  const cagr = Math.pow(endValue / startValue, 1 / periods) - 1;
  const text = `CAGR  ${(cagr * 100).toFixed(1)}%`;

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.getSelectedSlides();
    slides.load("items/id");
    await context.sync();
    if (slides.items.length === 0) throw new Error("Open a slide first.");

    slides.items[0].shapes.addTextBox(text, {
      left: 120,
      top: 120,
      width: 200,
      height: 50,
    });
    await context.sync();
  });
}
