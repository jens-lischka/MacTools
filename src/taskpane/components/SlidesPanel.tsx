import * as React from "react";
import { Textarea, Input, Field, Divider } from "@fluentui/react-components";
import {
  getSlideTitles,
  insertTableOfContents,
  addStickyNote,
} from "../../lib/slides";
import { getSetting, setSetting } from "../../lib/settings";
import { useActions } from "./ActionContext";
import { ToolButton } from "./ToolButton";
import { usePanelStyles } from "./panelStyles";

const INITIALS_KEY = "stickyNote:initials";

export const SlidesPanel: React.FC = () => {
  const styles = usePanelStyles();
  const { run, busy } = useActions();
  const [titles, setTitles] = React.useState("");
  const [exporting, setExporting] = React.useState(false);
  const [initials, setInitials] = React.useState("");

  React.useEffect(() => {
    let cancelled = false;
    void getSetting<string>(INITIALS_KEY, "").then((value) => {
      if (!cancelled) setInitials(value);
    });
    return () => {
      cancelled = true;
    };
  }, []);

  const exportTitles = async () => {
    setExporting(true);
    try {
      const list = await getSlideTitles();
      setTitles(list.map((t, i) => `${i + 1}. ${t}`).join("\n"));
    } catch (err) {
      setTitles(err instanceof Error ? err.message : String(err));
    } finally {
      setExporting(false);
    }
  };

  return (
    <div className={styles.section}>
      <Field label="Slides">
        <div className={styles.toolbar}>
          <ToolButton
            icon="ExportSlideTitles"
            label="Export slide titles"
            disabled={busy || exporting}
            onClick={() => void exportTitles()}
          />
          <ToolButton
            icon="TableOfContents"
            label="Insert table of contents"
            disabled={busy}
            onClick={() =>
              run("Insert table of contents", () => insertTableOfContents())
            }
          />
        </div>
      </Field>
      {titles && (
        <Field label="Slide titles (select and copy)">
          <Textarea value={titles} rows={8} readOnly />
        </Field>
      )}

      <Divider />

      <Field label="Sticky note — your initials">
        <div className={styles.toolbar}>
          <Input
            className={styles.grow}
            value={initials}
            placeholder="e.g. JL"
            onChange={(_, d) => setInitials(d.value)}
          />
          <ToolButton
            icon="AddNote"
            label="Add sticky note"
            disabled={busy}
            onClick={() =>
              run("Add sticky note", async () => {
                await setSetting(INITIALS_KEY, initials, "roaming");
                await addStickyNote(initials);
              })
            }
          />
        </div>
      </Field>
    </div>
  );
};
