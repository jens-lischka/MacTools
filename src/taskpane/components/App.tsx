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
import { HelpButton } from "./HelpButton";
import { highestSupportedApi } from "../../lib/capabilities";
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
  scroll: { flexGrow: 1, overflowY: "scroll" },
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
              <Subtitle2>A Creative Studio Solution</Subtitle2>
              <br />
              <br />
              <Caption1>
                Enhanced features for macOS and web — built for streamlined
                presentation workflows.
              </Caption1>
              <br />
              <br />
              <Caption1>
                Version 0.1.2 · PowerPoint API {highestSupportedApi()}
              </Caption1>
            </div>
            <HelpButton />
          </div>

          {status && (
            <MessageBar
              layout="multiline"
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
