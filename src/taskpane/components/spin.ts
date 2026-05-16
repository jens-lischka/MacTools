import type { SpinButtonProps } from "@fluentui/react-components";

/** A SpinButton `onChange` handler that writes valid numbers to `setter`. */
export function spinHandler(
  setter: (value: number) => void,
): SpinButtonProps["onChange"] {
  return (_event, data) => {
    const next = data.value ?? Number(data.displayValue);
    if (Number.isFinite(next)) setter(next as number);
  };
}
