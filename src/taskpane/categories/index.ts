import type { Category, Feature } from "../../lib/types";
import { AlignmentPanel } from "../components/AlignmentPanel";
import { SizePositionPanel } from "../components/SizePositionPanel";
import { SelectSamePanel } from "../components/SelectSamePanel";
import { TextPanel } from "../components/TextPanel";
import { ShapeCreationPanel } from "../components/ShapeCreationPanel";
import { TransformPanel } from "../components/TransformPanel";
import { SlidesPanel } from "../components/SlidesPanel";
import { UtilitiesPanel } from "../components/UtilitiesPanel";

/**
 * Category registry driving the task pane.
 *
 * The Alignment category is fully implemented (custom panel). Every other
 * category is scaffolded: features are listed with their feasibility rating
 * so the roadmap is visible in-product, but handlers are wired in later
 * phases. Features without a `run` handler render disabled.
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
      f("selectFont", "Select by Font / Text", "ok", "Planned for a later phase."),
      f("selectPosition", "Select by Position", "ok", "Planned for a later phase."),
      f("showHide", "Show All / Hide Objects", "blocked", "Shape visibility not exposed."),
    ],
  },
  {
    id: "text",
    label: "Text & Paragraph",
    panel: TextPanel,
    features: [
      f("textStyles", "Text Styles (Body / Heading / …)", "ok", "Planned for a later phase."),
      f("mergeSplit", "Merge / Split / Copy Text", "ok", "Planned for a later phase."),
      f("specialChars", "Special Characters", "ok", "Planned for a later phase."),
      f("bullets", "Fix Bullets", "partial", "Bullet formatting API is limited; planned."),
      f("paragraphSpacing", "Paragraph Spacing", "partial", "Space before/after API is limited; planned."),
    ],
  },
  {
    id: "insert",
    label: "Shape Creation",
    panel: ShapeCreationPanel,
    features: [
      f("multiply", "Multiply Shape", "partial", "No shape-duplicate API; planned via recreate."),
      f("groupAs", "Group as Rows / Columns", "ok", "Planned for a later phase."),
      f("numberedCircle", "Numbered Circle", "ok", "Planned for a later phase."),
      f("connectors", "Connectors / Auto-Connect", "partial", "True connection sites not exposed."),
    ],
  },
  {
    id: "properties",
    label: "Swap, Pick up & Apply",
    panel: TransformPanel,
    features: [
      f("swapStyle", "Swap Fill / Outline / Line styles", "ok", "Planned for a later phase."),
      f("applyMatching", "Apply to Matching Objects", "ok", "Planned for a later phase."),
      f("tableProperties", "Pick up / Apply Table Formatting", "partial", "Table API is limited; planned."),
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
      f("formatTable", "Format Table / Heading / Text", "partial"),
      f("addRowColumn", "Add Row / Column", "partial"),
      f("moveRemove", "Move / Remove Row / Column", "partial"),
      f("tableToText", "Convert Table to Text", "partial"),
      f("optimizeWidth", "Optimize Table Width", "partial"),
      f("transpose", "Transpose Table", "blocked", "No structural transpose API."),
      f("splitTable", "Split Table", "blocked", "No split API."),
    ],
  },
  {
    id: "slides",
    label: "Slides & Presentation",
    panel: SlidesPanel,
    features: [
      f("stickyNotes", "Sticky Notes", "partial", "Planned for a later phase."),
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
    features: [
      f("title", "Slide Title", "ok"),
      f("conclusion", "Conclusion", "ok"),
      f("footnote", "Footnote", "ok"),
      f("ghost", "Ghost", "ok"),
      f("labels", "Labels", "ok"),
      f("harvey", "Harvey Balls", "partial", "Pie-segment angles are not settable via the API."),
    ],
  },
  {
    id: "utilities",
    label: "Utilities",
    panel: UtilitiesPanel,
    features: [
      f("optimizeFontSize", "Optimize Font Size", "partial", "Planned for a later phase."),
      f("fileSize", "File Size", "partial", "Planned for a later phase."),
      f("conversionColours", "Conversion Assistant — Colours / Fonts", "partial", "Planned for a later phase."),
      f("airplaneMode", "Airplane Mode", "blocked", "Picture compression not exposed."),
      f("applyTemplate", "Apply Custom Template", "blocked", "No template-swap API."),
    ],
  },
  {
    id: "settings",
    label: "Settings",
    features: [
      f("customizeSettings", "Customize Settings", "ok"),
      f("importSettings", "Import / Generate / Clear Settings", "ok"),
      f("onlineHelp", "Online Help", "ok"),
      f("shortcuts", "Shortcut Manager", "blocked", "Global shortcut binding not available to add-ins."),
    ],
  },
];
