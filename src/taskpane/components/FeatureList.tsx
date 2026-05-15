import * as React from "react";
import {
  Button,
  Badge,
  Tooltip,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import type { Feature, Rating } from "../../lib/types";
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
  const handler = feature.run;

  const button = (
    <Button
      appearance="subtle"
      className={styles.button}
      disabled={busy || !handler}
      onClick={handler ? () => run(feature.label, handler) : undefined}
    >
      {feature.label}
    </Button>
  );

  return (
    <div className={styles.row}>
      {feature.description ? (
        <Tooltip content={feature.description} relationship="description">
          {button}
        </Tooltip>
      ) : (
        button
      )}
      <Badge appearance="tint" color={meta.color} size="small">
        {meta.text}
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
