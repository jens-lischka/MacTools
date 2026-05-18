import type { Category, Feature } from "../../lib/types";
import { AlignmentPanel } from "../components/AlignmentPanel";
import { SizePositionPanel } from "../components/SizePositionPanel";
import { SelectSamePanel } from "../components/SelectSamePanel";
import { TextPanel } from "../components/TextPanel";
import { TransformPanel } from "../components/TransformPanel";
import { SlidesPanel } from "../components/SlidesPanel";
import { SpecialShapesPanel } from "../components/SpecialShapesPanel";
import { TablesPanel } from "../components/TablesPanel";
import { UtilitiesPanel } from "../components/UtilitiesPanel";
import { SettingsPanel } from "../components/SettingsPanel";

/**
 * Category registry driving the task pane. Each category renders a working
 * panel, optionally followed by a list of features still in progress.
 */

const f = (id: string, label: string, rating: Feature["rating"], description?: string): Feature => ({
  id,
  label,
  rating,
  description,
});

export const categories: Category[] = [
  { id: "alignment", label: "Alignment & Distribution", panel: AlignmentPanel },
  { id: "size", label: "Size & Position", panel: SizePositionPanel },
  { id: "selection", label: "Select Same", panel: SelectSamePanel },
  { id: "text", label: "Text & Paragraph", panel: TextPanel },
  {
    id: "properties",
    label: "Swap, Pick up & Apply",
    panel: TransformPanel,
    features: [
      f("tableProperties", "Pick up / Apply Table Formatting", "partial", "Planned for a later phase."),
    ],
  },
  {
    id: "tables",
    label: "Tables",
    panel: TablesPanel,
    features: [
      f("formatTable", "Table styling", "partial", "Planned for a later phase."),
      f("optimizeWidth", "Optimize Table Width", "partial", "Planned for a later phase."),
    ],
  },
  {
    id: "slides",
    label: "Slides & Presentation",
    panel: SlidesPanel,
    features: [
      f("pasteOnSlides", "Paste on Slides", "partial", "Clipboard access is constrained."),
    ],
  },
  { id: "special", label: "Special Shapes", panel: SpecialShapesPanel },
  {
    id: "utilities",
    label: "Utilities",
    panel: UtilitiesPanel,
    features: [
      f("optimizeFontSize", "Optimize Font Size", "partial", "Planned for a later phase."),
      f("conversionColours", "Conversion Assistant — Colours / Fonts", "partial", "Planned for a later phase."),
    ],
  },
  { id: "settings", label: "Settings", panel: SettingsPanel },
];
