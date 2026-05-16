import * as React from "react";
import {
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  Button,
  Tooltip,
  makeStyles,
} from "@fluentui/react-components";
import { Dismiss24Regular } from "@fluentui/react-icons";
import { ToolIcon } from "./ToolIcon";

const useStyles = makeStyles({
  surface: { maxWidth: "min(960px, 94vw)", width: "min(960px, 94vw)" },
  frame: { width: "100%", height: "76vh", border: "0" },
});

/** Info button in the header that opens the manual inside a modal. */
export const HelpDialog: React.FC = () => {
  const styles = useStyles();
  return (
    <Dialog>
      <DialogTrigger disableButtonEnhancement>
        <Tooltip content="Manual & about" relationship="label">
          <Button
            appearance="subtle"
            icon={<ToolIcon name="Info" />}
            aria-label="Open the manual"
          />
        </Tooltip>
      </DialogTrigger>
      <DialogSurface className={styles.surface}>
        <DialogBody>
          <DialogTitle
            action={
              <DialogTrigger action="close" disableButtonEnhancement>
                <Button
                  appearance="subtle"
                  aria-label="Close"
                  icon={<Dismiss24Regular />}
                />
              </DialogTrigger>
            }
          >
            StudioTools Manual
          </DialogTitle>
          <DialogContent>
            <iframe
              className={styles.frame}
              src="manual.html"
              title="StudioTools manual"
            />
          </DialogContent>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};
