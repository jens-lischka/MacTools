import * as React from "react";
import {
  Button,
  Textarea,
  Input,
  Field,
  Divider,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import {
  getSlideTitles,
  insertTableOfContents,
  addStickyNote,
} from "../../lib/slides";
import { getSetting, setSetting } from "../../lib/settings";
import { useActions } from "./ActionContext";
import { ToolIcon } from "./ToolIcon";

const useStyles = makeStyles({
  section: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
  row: { display: "flex", gap: tokens.spacingHorizontalS, flexWrap: "wrap" },
  grow: { flexGrow: 1 },
});

const INITIALS_KEY = "stickyNote:initials";

export const SlidesPanel: React.FC = () => {
  const styles = useStyles();
  const { run, busy } = useActions();
  const [titles, setTitles] = React.useState("");
  const [exporting, setExporting] = React.useState(false);
  const [initials, setInitials] = React.useState("");

  React.useEffect(() => {
    void getSetting<string>(INITIALS_KEY, "").then(setInitials);
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
      <Button
        disabled={busy || exporting}
        icon={<ToolIcon name="ExportSlideTitles" />}
        onClick={() => void exportTitles()}
      >
        Export Slide Titles
      </Button>
      {titles && (
        <Field label="Slide titles (select and copy)">
          <Textarea value={titles} rows={8} readOnly />
        </Field>
      )}

      <Divider />

      <Button
        disabled={busy}
        icon={<ToolIcon name="TableOfContents" />}
        onClick={() =>
          run("Insert table of contents", () => insertTableOfContents())
        }
      >
        Insert Table of Contents
      </Button>

      <Divider />

      <Field label="Sticky note — your initials">
        <div className={styles.row}>
          <Input
            className={styles.grow}
            value={initials}
            placeholder="e.g. JL"
            onChange={(_, d) => setInitials(d.value)}
          />
          <Button
            disabled={busy}
            icon={<ToolIcon name="AddNote" />}
            onClick={() =>
              run("Add sticky note", async () => {
                await setSetting(INITIALS_KEY, initials, "roaming");
                await addStickyNote(initials);
              })
            }
          >
            Add Note
          </Button>
        </div>
      </Field>
    </div>
  );
};
