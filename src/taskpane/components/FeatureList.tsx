import * as React from "react";
import {
  Button,
  Badge,
  Tooltip,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import type { Feature, Rating } from "../../lib/types";
import { isApiSupported } from "../../lib/capabilities";
import { useActions } from "./ActionContext";

const useStyles = makeStyles({
  list: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
  },
  row: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
  },
  button: { flexGrow: 1, justifyContent: "flex-start" },
});

const RATING_META: Record<Rating, { color: "success" | "warning" | "danger"; text: string }> = {
  ok: { color: "success", text: "Ready" },
  partial: { color: "warning", text: "Partial" },
  blocked: { color: "danger", text: "Blocked" },
};

const FeatureRow: React.FC<{ feature: Feature }> = ({ feature }) => {
  const styles = useStyles();
  const { run, busy } = useActions();
  const meta = RATING_META[feature.rating];
  const apiMissing =
    feature.requiresApi !== undefined && !isApiSupported(feature.requiresApi);
  const handler = feature.run;
  const tooltip = apiMissing
    ? `Needs PowerPoint API ${feature.requiresApi}, which this client does not support.`
    : feature.description;

  const button = (
    <Button
      appearance="subtle"
      className={styles.button}
      disabled={busy || !handler || apiMissing}
      onClick={handler && !apiMissing ? () => run(feature.label, handler) : undefined}
    >
      {feature.label}
    </Button>
  );

  return (
    <div className={styles.row}>
      {tooltip ? (
        <Tooltip content={tooltip} relationship="description">
          {button}
        </Tooltip>
      ) : (
        button
      )}
      <Badge appearance="tint" color={apiMissing ? "warning" : meta.color} size="small">
        {apiMissing ? `API ${feature.requiresApi}` : meta.text}
      </Badge>
    </div>
  );
};

export const FeatureList: React.FC<{ features: Feature[] }> = ({ features }) => {
  const styles = useStyles();
  return (
    <div className={styles.list}>
      {features.map((feature) => (
        <FeatureRow key={feature.id} feature={feature} />
      ))}
    </div>
  );
};
