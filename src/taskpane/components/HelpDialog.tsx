import * as React from "react";
import {
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Button,
  Tooltip,
} from "@fluentui/react-components";
import { ToolIcon } from "./ToolIcon";
import { highestSupportedApi } from "../../lib/capabilities";

/** Info button in the header that opens an About / help modal. */
export const HelpDialog: React.FC = () => (
  <Dialog>
    <DialogTrigger disableButtonEnhancement>
      <Tooltip content="About MacTools" relationship="label">
        <Button
          appearance="subtle"
          icon={<ToolIcon name="Info" />}
          aria-label="About MacTools"
        />
      </Tooltip>
    </DialogTrigger>
    <DialogSurface>
      <DialogBody>
        <DialogTitle>MacTools for PowerPoint</DialogTitle>
        <DialogContent>
          <p>
            Cross-platform PowerPoint tools — alignment, sizing, text, shapes,
            tables and more, all in one task pane. Works in PowerPoint on the
            web, macOS and Windows.
          </p>
          <p>
            Most tools act on the shapes you have selected on the current
            slide. Tools that need a newer PowerPoint API than your Office
            build provides disable themselves automatically.
          </p>
          <p>
            Detected PowerPoint API set: <strong>{highestSupportedApi()}</strong>
          </p>
        </DialogContent>
        <DialogActions>
          <Button
            appearance="primary"
            onClick={() => window.open("manual.html", "_blank", "noopener")}
          >
            Open full manual
          </Button>
          <DialogTrigger disableButtonEnhancement>
            <Button appearance="secondary">Close</Button>
          </DialogTrigger>
        </DialogActions>
      </DialogBody>
    </DialogSurface>
  </Dialog>
);
