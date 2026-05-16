/**
 * Slide-level operations — Phase 4.
 *
 * Title detection uses the shape name (PowerPoint names title placeholders
 * "Title …"). This is a heuristic, but it is throw-free — unlike reading
 * `placeholderFormat` on shapes that are not placeholders.
 */

function cleanLine(text: string): string {
  return text.replace(/[\r\n\v\f]+/g, " ").trim();
}

/** The cleaned title text of every slide, in order. */
export async function getSlideTitles(): Promise<string[]> {
  const titles: string[] = [];
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    const shapeCollections = slides.items.map((slide) => {
      const shapes = slide.shapes;
      shapes.load("items/name");
      return shapes;
    });
    await context.sync();

    const titleRanges = shapeCollections.map((shapes) => {
      const titleShape = shapes.items.find((s) => /^title/i.test(s.name));
      return titleShape ? titleShape.textFrame.textRange : null;
    });
    titleRanges.forEach((range) => range?.load("text"));
    await context.sync();

    titleRanges.forEach((range, index) => {
      const text = range ? cleanLine(range.text) : "";
      titles.push(text || `(Slide ${index + 1} — untitled)`);
    });
  });
  return titles;
}

/** Insert a numbered table of contents as a text box on the slide in view. */
export async function insertTableOfContents(): Promise<void> {
  const titles = await getSlideTitles();
  if (titles.length === 0) throw new Error("The presentation has no slides.");

  const body = titles.map((title, i) => `${i + 1}.  ${title}`).join("\n");

  await PowerPoint.run(async (context) => {
    const slides = context.presentation.getSelectedSlides();
    slides.load("items/id");
    await context.sync();
    if (slides.items.length === 0) throw new Error("Open a slide first.");

    slides.items[0].shapes.addTextBox(body, {
      left: 60,
      top: 60,
      width: 600,
      height: 400,
    });
    await context.sync();
  });
}
