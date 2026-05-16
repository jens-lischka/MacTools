import * as React from "react";
import {
  Button,
  Textarea,
  Field,
  Divider,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { getSlideTitles, insertTableOfContents } from "../../lib/slides";
import { useActions } from "./ActionContext";

const useStyles = makeStyles({
  section: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
});

export const SlidesPanel: React.FC = () => {
  const styles = useStyles();
  const { run, busy } = useActions();
  const [titles, setTitles] = React.useState("");
  const [exporting, setExporting] = React.useState(false);

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
        onClick={() =>
          run("Insert table of contents", () => insertTableOfContents())
        }
      >
        Insert Table of Contents
      </Button>
    </div>
  );
};
