import { useMemo, useState } from "react";
import {
  FluentProvider,
  Button,
  TabList,
  Tab,
  MessageBar,
  MessageBarBody,
  Text,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import type { TabValue } from "@fluentui/react-components";
import { lettreTheme } from "../theme";
import type { Block, BlockType } from "../../newsletter/types";
import { CATALOG, createBlock, defaultNewsletter } from "../../newsletter/catalog";
import { generateHtml } from "../../newsletter/render";
import { insertNewsletter } from "../../office/mailbox";
import { BlockList } from "./BlockList";
import { BlockEditor } from "./BlockEditor";
import { Preview } from "./Preview";

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", height: "100vh", overflow: "hidden" },
  header: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "8px",
    padding: "10px 12px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  brand: { fontFamily: "Georgia, serif", fontSize: "18px", fontWeight: 700 },
  dot: { color: tokens.colorBrandForeground1 },
  tabs: { padding: "0 8px", borderBottom: `1px solid ${tokens.colorNeutralStroke2}` },
  body: {
    display: "flex",
    flexDirection: "column",
    flexGrow: 1,
    minHeight: 0,
    padding: "12px",
  },
  scroll: { flexGrow: 1, minHeight: 0, overflowY: "auto" },
  editor: {
    marginTop: "12px",
    paddingTop: "10px",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  editorHead: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    marginBottom: "8px",
  },
  editorTitle: {
    fontSize: "11px",
    textTransform: "uppercase",
    letterSpacing: "0.06em",
    color: tokens.colorNeutralForeground3,
  },
  code: {
    margin: 0,
    flexGrow: 1,
    minHeight: 0,
    overflow: "auto",
    fontFamily: "ui-monospace, Menlo, Consolas, monospace",
    fontSize: "11px",
    lineHeight: 1.5,
    whiteSpace: "pre-wrap",
    wordBreak: "break-all",
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
    padding: "10px",
  },
  htmlBar: { display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "8px" },
});

type Status = { intent: "success" | "error"; message: string };

const LABELS = Object.fromEntries(CATALOG.map((c) => [c.type, c.label])) as Record<
  BlockType,
  string
>;

export default function App() {
  const styles = useStyles();
  const [blocks, setBlocks] = useState<Block[]>(defaultNewsletter);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [view, setView] = useState<TabValue>("build");
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState<Status | null>(null);

  const html = useMemo(() => generateHtml(blocks), [blocks]);
  const selected = blocks.find((b) => b.id === selectedId) ?? null;

  const addBlock = (type: BlockType) => {
    const block = createBlock(type);
    setBlocks((prev) => [...prev, block]);
    setSelectedId(block.id);
  };

  const updateBlock = (id: string, patch: object) =>
    setBlocks((prev) =>
      prev.map((b) => (b.id === id ? ({ ...b, data: { ...b.data, ...patch } } as Block) : b)),
    );

  const removeBlock = (id: string) => {
    setBlocks((prev) => prev.filter((b) => b.id !== id));
    if (selectedId === id) setSelectedId(null);
  };

  const moveBlock = (id: string, dir: -1 | 1) =>
    setBlocks((prev) => {
      const i = prev.findIndex((b) => b.id === id);
      const j = i + dir;
      if (i < 0 || j < 0 || j >= prev.length) return prev;
      const next = [...prev];
      [next[i], next[j]] = [next[j], next[i]];
      return next;
    });

  const insert = async () => {
    setBusy(true);
    setStatus(null);
    try {
      await insertNewsletter(html);
      setStatus({ intent: "success", message: "Newsletter inserted into the email." });
    } catch (error) {
      setStatus({
        intent: "error",
        message: error instanceof Error ? error.message : "Insert failed.",
      });
    } finally {
      setBusy(false);
    }
  };

  const copyHtml = async () => {
    try {
      await navigator.clipboard.writeText(html);
      setStatus({ intent: "success", message: "HTML copied to the clipboard." });
    } catch {
      setStatus({ intent: "error", message: "The clipboard is not available here." });
    }
  };

  return (
    <FluentProvider theme={lettreTheme} className={styles.root}>
      <div className={styles.header}>
        <span className={styles.brand}>
          Lettre<span className={styles.dot}>.</span>
        </span>
        <Button appearance="primary" disabled={busy || blocks.length === 0} onClick={insert}>
          {busy ? "Inserting…" : "Insert into email"}
        </Button>
      </div>

      {status && (
        <MessageBar intent={status.intent}>
          <MessageBarBody>{status.message}</MessageBarBody>
        </MessageBar>
      )}

      <div className={styles.tabs}>
        <TabList selectedValue={view} onTabSelect={(_, data) => setView(data.value)}>
          <Tab value="build">Build</Tab>
          <Tab value="preview">Preview</Tab>
          <Tab value="html">HTML</Tab>
        </TabList>
      </div>

      {view === "build" && (
        <div className={styles.body}>
          <div className={styles.scroll}>
            <BlockList
              blocks={blocks}
              selectedId={selectedId}
              onAdd={addBlock}
              onSelect={setSelectedId}
              onMove={moveBlock}
              onRemove={removeBlock}
            />
            {selected && (
              <div className={styles.editor}>
                <div className={styles.editorHead}>
                  <span className={styles.editorTitle}>Edit · {LABELS[selected.type]}</span>
                  <Button size="small" appearance="subtle" onClick={() => setSelectedId(null)}>
                    Close
                  </Button>
                </div>
                <BlockEditor block={selected} update={updateBlock} />
              </div>
            )}
          </div>
        </div>
      )}

      {view === "preview" && (
        <div className={styles.body}>
          <Preview html={html} />
        </div>
      )}

      {view === "html" && (
        <div className={styles.body}>
          <div className={styles.htmlBar}>
            <Text size={200}>{html.length.toLocaleString()} characters</Text>
            <Button size="small" onClick={copyHtml}>
              Copy HTML
            </Button>
          </div>
          <pre className={styles.code}>{html}</pre>
        </div>
      )}
    </FluentProvider>
  );
}
