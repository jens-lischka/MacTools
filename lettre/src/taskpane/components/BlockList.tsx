/* The "Add block" palette and the ordered list of blocks in the newsletter. */

import { Button, Text, makeStyles, tokens } from "@fluentui/react-components";
import type { Block, BlockType } from "../../newsletter/types";
import { CATALOG } from "../../newsletter/catalog";

const useStyles = makeStyles({
  section: { marginBottom: "14px" },
  sectionLabel: {
    display: "block",
    fontSize: "10px",
    textTransform: "uppercase",
    letterSpacing: "0.08em",
    color: tokens.colorNeutralForeground3,
    marginBottom: "6px",
  },
  palette: { display: "flex", flexWrap: "wrap", gap: "4px" },
  row: {
    display: "flex",
    alignItems: "center",
    gap: "6px",
    padding: "6px 8px",
    borderRadius: tokens.borderRadiusMedium,
    cursor: "pointer",
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    marginBottom: "4px",
  },
  rowSelected: {
    border: `1px solid ${tokens.colorBrandStroke1}`,
    backgroundColor: tokens.colorNeutralBackground1Selected,
  },
  rowMain: { flexGrow: 1, minWidth: 0 },
  rowType: {
    fontSize: "10px",
    textTransform: "uppercase",
    letterSpacing: "0.05em",
    color: tokens.colorNeutralForeground3,
  },
  rowPreview: {
    display: "block",
    fontSize: "12px",
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  controls: { display: "flex", flexShrink: 0 },
  empty: {
    display: "block",
    textAlign: "center",
    padding: "16px",
    color: tokens.colorNeutralForeground3,
    fontSize: "12px",
  },
});

const LABELS: Record<BlockType, string> = Object.fromEntries(
  CATALOG.map((c) => [c.type, c.label]),
) as Record<BlockType, string>;

/** A short, human-readable summary of a block for the list row. */
function previewText(block: Block): string {
  switch (block.type) {
    case "header":
      return block.data.title;
    case "heading":
      return block.data.text;
    case "text":
      return block.data.content;
    case "image":
      return block.data.alt || "Image";
    case "imageText":
      return block.data.heading;
    case "button":
      return block.data.text;
    case "twoColumn":
      return "Two columns";
    case "carousel":
      return `${block.data.items.length} slides`;
    case "accordion":
      return `${block.data.items.length} items`;
    case "divider":
      return "Divider";
    case "spacer":
      return `${block.data.height}px`;
    case "footer":
      return block.data.company;
  }
}

export function BlockList({
  blocks,
  selectedId,
  onAdd,
  onSelect,
  onMove,
  onRemove,
}: {
  blocks: Block[];
  selectedId: string | null;
  onAdd: (type: BlockType) => void;
  onSelect: (id: string) => void;
  onMove: (id: string, dir: -1 | 1) => void;
  onRemove: (id: string) => void;
}) {
  const styles = useStyles();
  return (
    <>
      <div className={styles.section}>
        <Text className={styles.sectionLabel}>Add block</Text>
        <div className={styles.palette}>
          {CATALOG.map((entry) => (
            <Button key={entry.type} size="small" onClick={() => onAdd(entry.type)}>
              {entry.label}
            </Button>
          ))}
        </div>
      </div>

      <Text className={styles.sectionLabel}>Blocks ({blocks.length})</Text>
      {blocks.map((block, i) => (
        <div
          key={block.id}
          className={`${styles.row} ${block.id === selectedId ? styles.rowSelected : ""}`}
          onClick={() => onSelect(block.id)}
        >
          <div className={styles.rowMain}>
            <span className={styles.rowType}>{LABELS[block.type]}</span>
            <span className={styles.rowPreview}>{previewText(block)}</span>
          </div>
          <div className={styles.controls} onClick={(e) => e.stopPropagation()}>
            <Button
              size="small"
              appearance="subtle"
              disabled={i === 0}
              aria-label="Move up"
              onClick={() => onMove(block.id, -1)}
            >
              ↑
            </Button>
            <Button
              size="small"
              appearance="subtle"
              disabled={i === blocks.length - 1}
              aria-label="Move down"
              onClick={() => onMove(block.id, 1)}
            >
              ↓
            </Button>
            <Button
              size="small"
              appearance="subtle"
              aria-label="Delete"
              onClick={() => onRemove(block.id)}
            >
              ✕
            </Button>
          </div>
        </div>
      ))}
      {blocks.length === 0 && (
        <Text className={styles.empty}>No blocks yet. Add one above.</Text>
      )}
    </>
  );
}
