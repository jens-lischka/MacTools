import * as React from "react";
import {
  Accordion,
  AccordionHeader,
  AccordionItem,
  AccordionPanel,
  Divider,
  FluentProvider,
  MessageBar,
  MessageBarBody,
  Spinner,
  Subtitle2,
  Caption1,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { categories } from "../categories";
import { FeatureList } from "./FeatureList";
import { ActionContext, type ActionRunner } from "./ActionContext";
import { HelpDialog } from "./HelpDialog";
import { macToolsTheme } from "../theme";

const useStyles = makeStyles({
  root: {
    height: "100vh",
    display: "flex",
    flexDirection: "column",
    backgroundColor: tokens.colorNeutralBackground2,
  },
  header: {
    padding: tokens.spacingVerticalM,
    paddingBottom: tokens.spacingVerticalS,
    display: "flex",
    justifyContent: "space-between",
    alignItems: "flex-start",
    gap: tokens.spacingHorizontalS,
  },
  scroll: { flexGrow: 1, overflowY: "auto" },
  status: { margin: tokens.spacingVerticalS },
  panelInner: {
    padding: tokens.spacingVerticalS,
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
});

type Status = { kind: "info" | "success" | "error"; text: string } | null;

const App: React.FC = () => {
  const styles = useStyles();
  const [busy, setBusy] = React.useState(false);
  const [status, setStatus] = React.useState<Status>(null);

  const runner = React.useMemo<ActionRunner>(
    () => ({
      busy,
      run: (label, fn) => {
        setBusy(true);
        setStatus({ kind: "info", text: `${label}…` });
        fn()
          .then(() => setStatus({ kind: "success", text: `${label} — done.` }))
          .catch((err: unknown) =>
            setStatus({
              kind: "error",
              text: `${label} — ${err instanceof Error ? err.message : String(err)}`,
            }),
          )
          .finally(() => setBusy(false));
      },
    }),
    [busy],
  );

  return (
    <FluentProvider theme={macToolsTheme}>
      <ActionContext.Provider value={runner}>
        <div className={styles.root}>
          <div className={styles.header}>
            <div>
              <Subtitle2>StudioTools</Subtitle2>
              <br />
              <Caption1>A Creative Studio Solution</Caption1>
              <br />
              <Caption1>
                Enhanced features for macOS and web — built for streamlined
                presentation workflows.
              </Caption1>
            </div>
            <HelpDialog />
          </div>

          {status && (
            <MessageBar
              key={status.text}
              className={styles.status}
              intent={status.kind === "error" ? "error" : status.kind}
            >
              <MessageBarBody>
                {busy && <Spinner size="tiny" />} {status.text}
              </MessageBarBody>
            </MessageBar>
          )}

          <div className={styles.scroll}>
            <Accordion collapsible defaultOpenItems="alignment">
              {categories.map((category) => {
                const Panel = category.panel;
                const features = category.features ?? [];
                return (
                  <AccordionItem key={category.id} value={category.id}>
                    <AccordionHeader>{category.label}</AccordionHeader>
                    <AccordionPanel>
                      <div className={styles.panelInner}>
                        {Panel && <Panel />}
                        {Panel && features.length > 0 && <Divider />}
                        {features.length > 0 && <FeatureList features={features} />}
                      </div>
                    </AccordionPanel>
                  </AccordionItem>
                );
              })}
            </Accordion>
          </div>
        </div>
      </ActionContext.Provider>
    </FluentProvider>
  );
};

export default App;
