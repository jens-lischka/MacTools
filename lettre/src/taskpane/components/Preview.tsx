/* Renders the generated newsletter HTML in an isolated iframe.
 *
 * The HTML is written imperatively with document.write — the `srcDoc` prop
 * can be skipped by the browser when React diffs the same iframe element
 * across renders, leaving a stale preview. */

import { useEffect, useRef } from "react";
import { makeStyles, tokens } from "@fluentui/react-components";

const useStyles = makeStyles({
  frame: {
    width: "100%",
    height: "100%",
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: "#ffffff",
  },
});

export function Preview({ html }: { html: string }) {
  const styles = useStyles();
  const ref = useRef<HTMLIFrameElement>(null);

  useEffect(() => {
    const doc = ref.current?.contentDocument;
    if (!doc) return;
    doc.open();
    doc.write(html);
    doc.close();
  }, [html]);

  return <iframe ref={ref} title="Newsletter preview" className={styles.frame} />;
}
