/* Thin, typed wrappers over Fluent inputs so the block editor stays terse. */

import type { ReactNode } from "react";
import {
  Input,
  Textarea,
  Select,
  Label,
  makeStyles,
  tokens,
} from "@fluentui/react-components";

const useStyles = makeStyles({
  field: { display: "flex", flexDirection: "column", gap: "4px", marginBottom: "10px" },
  label: {
    fontSize: "11px",
    textTransform: "uppercase",
    letterSpacing: "0.04em",
    color: tokens.colorNeutralForeground3,
  },
  color: { display: "flex", gap: "6px", alignItems: "center" },
  swatch: {
    width: "36px",
    height: "32px",
    padding: 0,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusMedium,
    cursor: "pointer",
    flexShrink: 0,
  },
  grow: { flexGrow: 1 },
});

export function Field({ label, children }: { label: string; children: ReactNode }) {
  const styles = useStyles();
  return (
    <div className={styles.field}>
      <Label className={styles.label}>{label}</Label>
      {children}
    </div>
  );
}

export function TextInput({
  value,
  onChange,
  placeholder,
}: {
  value: string;
  onChange: (value: string) => void;
  placeholder?: string;
}) {
  return (
    <Input
      value={value}
      placeholder={placeholder}
      onChange={(_, data) => onChange(data.value)}
    />
  );
}

export function TextAreaInput({
  value,
  onChange,
  rows = 3,
}: {
  value: string;
  onChange: (value: string) => void;
  rows?: number;
}) {
  return (
    <Textarea
      value={value}
      rows={rows}
      resize="vertical"
      onChange={(_, data) => onChange(data.value)}
    />
  );
}

export function SelectInput({
  value,
  onChange,
  options,
}: {
  value: string;
  onChange: (value: string) => void;
  options: { value: string; label: string }[];
}) {
  return (
    <Select value={value} onChange={(_, data) => onChange(data.value)}>
      {options.map((opt) => (
        <option key={opt.value} value={opt.value}>
          {opt.label}
        </option>
      ))}
    </Select>
  );
}

export function ColorInput({
  value,
  onChange,
}: {
  value: string;
  onChange: (value: string) => void;
}) {
  const styles = useStyles();
  return (
    <div className={styles.color}>
      <input
        type="color"
        className={styles.swatch}
        value={value}
        onChange={(e) => onChange(e.target.value)}
      />
      <Input
        className={styles.grow}
        value={value}
        onChange={(_, data) => onChange(data.value)}
      />
    </div>
  );
}
