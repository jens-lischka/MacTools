/**
 * Alignment & distribution operations — the Phase 2 reference implementation.
 *
 * Power Align in the legacy add-in relied on slide guides and live Ctrl/Alt/
 * Shift modifier state, neither of which Office.js exposes. Per the agreed
 * design, the guide/modifier modes are replaced by an explicit in-pane
 * "target" choice: align relative to the selection, the first or last
 * selected shape, or the slide.
 */

import { withSelectedShapes, unionBox, type Box, type SelectedShape } from "./powerpoint";
import { getSetting } from "./settings";

export type AlignTarget = "selection" | "first" | "last" | "slide";
export type AlignEdge = "left" | "centerH" | "right" | "top" | "middle" | "bottom";
export type DistributeAxis = "horizontal" | "vertical";

interface SlideSize {
  width: number;
  height: number;
}

/** Default widescreen 16:9 slide in points (13.333in x 7.5in). */
const DEFAULT_SLIDE_SIZE: SlideSize = { width: 960, height: 540 };

function targetBox(
  target: AlignTarget,
  geometry: SelectedShape[],
  slideSize: SlideSize,
): Box {
  switch (target) {
    case "first":
      return geometry[0];
    case "last":
      return geometry[geometry.length - 1];
    case "slide":
      return { left: 0, top: 0, width: slideSize.width, height: slideSize.height };
    case "selection":
    default:
      return unionBox(geometry);
  }
}

/** Align every selected shape to one edge of the resolved target box. */
export async function align(edge: AlignEdge, target: AlignTarget): Promise<void> {
  const slideSize = await getSetting<SlideSize>("slideSize", DEFAULT_SLIDE_SIZE);
  const minShapes = target === "slide" ? 1 : 2;

  await withSelectedShapes(minShapes, (shapes, geometry) => {
    const box = targetBox(target, geometry, slideSize);
    shapes.forEach((shape, i) => {
      const g = geometry[i];
      switch (edge) {
        case "left":
          shape.left = box.left;
          break;
        case "centerH":
          shape.left = box.left + (box.width - g.width) / 2;
          break;
        case "right":
          shape.left = box.left + box.width - g.width;
          break;
        case "top":
          shape.top = box.top;
          break;
        case "middle":
          shape.top = box.top + (box.height - g.height) / 2;
          break;
        case "bottom":
          shape.top = box.top + box.height - g.height;
          break;
      }
    });
  });
}

/** Distribute selected shapes so the gaps between them are equal. */
export async function distribute(axis: DistributeAxis): Promise<void> {
  await withSelectedShapes(3, (shapes, geometry) => {
    const order = geometry
      .map((g, i) => ({ g, i }))
      .sort((a, b) =>
        axis === "horizontal" ? a.g.left - b.g.left : a.g.top - b.g.top,
      );

    const first = order[0].g;
    const last = order[order.length - 1].g;

    const span =
      axis === "horizontal"
        ? last.left + last.width - first.left
        : last.top + last.height - first.top;
    const usedBySizes = order.reduce(
      (sum, o) => sum + (axis === "horizontal" ? o.g.width : o.g.height),
      0,
    );
    const gap = (span - usedBySizes) / (order.length - 1);

    let cursor = axis === "horizontal" ? first.left : first.top;
    order.forEach((o) => {
      const shape = shapes[o.i];
      if (axis === "horizontal") {
        shape.left = cursor;
        cursor += o.g.width + gap;
      } else {
        shape.top = cursor;
        cursor += o.g.height + gap;
      }
    });
  });
}
