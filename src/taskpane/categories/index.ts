import type { Category, Feature } from "../../lib/types";
import { AlignmentPanel } from "../components/AlignmentPanel";
import { SizePositionPanel } from "../components/SizePositionPanel";
import { SelectSamePanel } from "../components/SelectSamePanel";
import { TextPanel } from "../components/TextPanel";
import { ShapeCreationPanel } from "../components/ShapeCreationPanel";
import { TransformPanel } from "../components/TransformPanel";
import { SlidesPanel } from "../components/SlidesPanel";
import { SpecialShapesPanel } from "../components/SpecialShapesPanel";
import { UtilitiesPanel } from "../components/UtilitiesPanel";
import { SettingsPanel } from "../components/SettingsPanel";

/**
 * Category registry driving the task pane.
 *
 * Each category renders a working panel and/or a list of features. A feature
 * with no `run` handler renders disabled — it carries a feasibility rating so
 * not-yet-built and API-blocked items stay visible in-product.
 *
 * Ratings mirror docs/powerpoint-cross-platform-overview.md.
 */

const f = (id: string, label: string, rating: Feature["rating"], description?: string): Feature => ({
  id,
  label,
  rating,
  description,
});

export const categories: Category[] = [
  {
    id: "alignment",
    label: "Alignment & Distribution",
    panel: AlignmentPanel,
  },
  {
    id: "size",
    label: "Size & Position",
    panel: SizePositionPanel,
    features: [
      f("fillGap", "Fill to gap", "partial", "Planned for a later phase."),
      f("unifyShapes", "Unify Shapes / Arrows", "partial", "Planned for a later phase."),
      f("cropToCircle", "Crop to Circle", "blocked", "Picture crop API not exposed."),
      f("aspectRatioLock", "Lock Aspect Ratio", "blocked", "Property not exposed."),
    ],
  },
  {
    id: "selection",
    label: "Select Same",
    panel: SelectSamePanel,
    features: [
      f("showHide", "Show All / Hide Objects", "blocked", "Shape visibility not exposed."),
    ],
  },
  {
    id: "text",
    label: "Text & Paragraph",
    panel: TextPanel,
    features: [
      f("paragraphSpacing", "Paragraph Spacing", "blocked", "Space before/after has no JS API."),
    ],
  },
  {
    id: "insert",
    label: "Shape Creation",
    panel: ShapeCreationPanel,
    features: [
      f("multiply", "Multiply Shape", "blocked", "No shape-duplicate API."),
      f("connectors", "Connectors / Auto-Connect", "partial", "True connection sites not exposed."),
    ],
  },
  {
    id: "properties",
    label: "Swap, Pick up & Apply",
    panel: TransformPanel,
    features: [
      f("applyMatching", "Apply to Matching Objects", "ok", "Planned for a later phase."),
      f("tableProperties", "Pick up / Apply Table Formatting", "partial", "Needs the table API; planned."),
    ],
  },
  {
    id: "effects",
    label: "Effects",
    features: [
      f("dropShadow", "Drop Shadow / Remove Effects", "blocked", "Effect formatting not exposed."),
      f("grayscale", "Fix Grayscale", "blocked", "Picture recolour not exposed."),
    ],
  },
  {
    id: "tables",
    label: "Tables",
    features: [
      f("formatTable", "Format Table / Heading / Text", "partial", "Needs PowerPoint API 1.8; planned."),
      f("addRowColumn", "Add Row / Column", "partial", "Needs PowerPoint API 1.8; planned."),
      f("moveRemove", "Move / Remove Row / Column", "partial", "Needs PowerPoint API 1.8; planned."),
      f("tableToText", "Convert Table to Text", "partial", "Needs PowerPoint API 1.8; planned."),
      f("optimizeWidth", "Optimize Table Width", "partial", "Needs PowerPoint API 1.8; planned."),
      f("transpose", "Transpose Table", "blocked", "No structural transpose API."),
      f("splitTable", "Split Table", "blocked", "No split API."),
    ],
  },
  {
    id: "slides",
    label: "Slides & Presentation",
    panel: SlidesPanel,
    features: [
      f("stickyNoteManage", "Sticky Notes — show / hide / remove", "blocked", "Shape visibility not exposed."),
      f("pasteOnSlides", "Paste on Slides", "partial", "Clipboard access is constrained."),
      f("exportPictures", "Export as Pictures", "blocked", "No slide-render/export API."),
      f("sections", "Section management", "blocked", "No sections API."),
      f("slideGuides", "Slide Guides", "blocked", "No slide-guide API."),
      f("closeAll", "Close All", "blocked", "No multi-document control."),
    ],
  },
  {
    id: "special",
    label: "Special Shapes",
    panel: SpecialShapesPanel,
    features: [
      f("harvey", "Harvey Balls", "blocked", "Pie-segment angles are not settable via the API."),
    ],
  },
  {
    id: "utilities",
    label: "Utilities",
    panel: UtilitiesPanel,
    features: [
      f("optimizeFontSize", "Optimize Font Size", "partial", "Planned for a later phase."),
      f("conversionColours", "Conversion Assistant — Colours / Fonts", "partial", "Planned for a later phase."),
      f("airplaneMode", "Airplane Mode", "blocked", "Picture compression not exposed."),
      f("applyTemplate", "Apply Custom Template", "blocked", "No template-swap API."),
    ],
  },
  {
    id: "settings",
    label: "Settings",
    panel: SettingsPanel,
    features: [
      f("shortcuts", "Shortcut Manager", "blocked", "Global shortcut binding not available to add-ins."),
    ],
  },
];
