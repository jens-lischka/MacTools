/* Per-block property editor. `block` is a discriminated union, so each
 * `case` narrows `block.data` to the exact shape with no casts. */

import {
  Button,
  Slider,
  Text,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import type { Block } from "../../newsletter/types";
import { newAccordionItem, newCarouselItem } from "../../newsletter/catalog";
import { Field, TextInput, TextAreaInput, SelectInput, ColorInput } from "./fields";

const useStyles = makeStyles({
  item: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    padding: "10px",
    marginBottom: "8px",
    backgroundColor: tokens.colorNeutralBackground2,
  },
  itemHead: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    marginBottom: "6px",
  },
  itemLabel: {
    fontSize: "10px",
    textTransform: "uppercase",
    letterSpacing: "0.06em",
    color: tokens.colorNeutralForeground3,
  },
  note: {
    display: "block",
    padding: "8px 10px",
    marginBottom: "10px",
    fontSize: "11px",
    lineHeight: 1.5,
    color: tokens.colorNeutralForeground2,
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
  },
  addBtn: { width: "100%" },
});

const ALIGN = [
  { value: "left", label: "Left" },
  { value: "center", label: "Center" },
  { value: "right", label: "Right" },
];

export function BlockEditor({
  block,
  update,
}: {
  block: Block;
  update: (id: string, patch: object) => void;
}) {
  const styles = useStyles();
  const set = (patch: object) => update(block.id, patch);

  switch (block.type) {
    case "header": {
      const d = block.data;
      return (
        <>
          <Field label="Title">
            <TextInput value={d.title} onChange={(v) => set({ title: v })} />
          </Field>
          <Field label="Tagline">
            <TextInput value={d.tagline} onChange={(v) => set({ tagline: v })} />
          </Field>
          <Field label="Background">
            <ColorInput value={d.bg} onChange={(v) => set({ bg: v })} />
          </Field>
          <Field label="Text color">
            <ColorInput value={d.fg} onChange={(v) => set({ fg: v })} />
          </Field>
        </>
      );
    }

    case "heading": {
      const d = block.data;
      return (
        <>
          <Field label="Text">
            <TextInput value={d.text} onChange={(v) => set({ text: v })} />
          </Field>
          <Field label="Level">
            <SelectInput
              value={d.level}
              onChange={(v) => set({ level: v })}
              options={[
                { value: "h1", label: "H1 — large" },
                { value: "h2", label: "H2 — medium" },
                { value: "h3", label: "H3 — small" },
              ]}
            />
          </Field>
          <Field label="Alignment">
            <SelectInput value={d.align} onChange={(v) => set({ align: v })} options={ALIGN} />
          </Field>
          <Field label="Color">
            <ColorInput value={d.color} onChange={(v) => set({ color: v })} />
          </Field>
        </>
      );
    }

    case "text": {
      const d = block.data;
      return (
        <>
          <Field label="Content">
            <TextAreaInput rows={5} value={d.content} onChange={(v) => set({ content: v })} />
          </Field>
          <Field label="Alignment">
            <SelectInput
              value={d.align}
              onChange={(v) => set({ align: v })}
              options={[...ALIGN, { value: "justify", label: "Justify" }]}
            />
          </Field>
          <Field label="Color">
            <ColorInput value={d.color} onChange={(v) => set({ color: v })} />
          </Field>
        </>
      );
    }

    case "image": {
      const d = block.data;
      return (
        <>
          <Field label="Image URL">
            <TextInput value={d.src} onChange={(v) => set({ src: v })} placeholder="https://..." />
          </Field>
          <Field label="Alt text">
            <TextInput value={d.alt} onChange={(v) => set({ alt: v })} />
          </Field>
          <Field label="Link (optional)">
            <TextInput value={d.href} onChange={(v) => set({ href: v })} placeholder="https://..." />
          </Field>
        </>
      );
    }

    case "imageText": {
      const d = block.data;
      return (
        <>
          <Field label="Layout">
            <SelectInput
              value={d.layout}
              onChange={(v) => set({ layout: v })}
              options={[
                { value: "image-left", label: "Image left" },
                { value: "image-right", label: "Image right" },
              ]}
            />
          </Field>
          <Field label="Image URL">
            <TextInput value={d.src} onChange={(v) => set({ src: v })} />
          </Field>
          <Field label="Image alt">
            <TextInput value={d.alt} onChange={(v) => set({ alt: v })} />
          </Field>
          <Field label={`Image width: ${d.imageWidth}px`}>
            <Slider
              min={140}
              max={320}
              step={10}
              value={d.imageWidth}
              onChange={(_, data) => set({ imageWidth: data.value })}
            />
          </Field>
          <Field label="Heading">
            <TextInput value={d.heading} onChange={(v) => set({ heading: v })} />
          </Field>
          <Field label="Body">
            <TextAreaInput value={d.body} onChange={(v) => set({ body: v })} />
          </Field>
          <Field label="Link text (optional)">
            <TextInput value={d.linkText} onChange={(v) => set({ linkText: v })} />
          </Field>
          <Field label="Link URL (optional)">
            <TextInput value={d.linkUrl} onChange={(v) => set({ linkUrl: v })} />
          </Field>
        </>
      );
    }

    case "button": {
      const d = block.data;
      return (
        <>
          <Field label="Label">
            <TextInput value={d.text} onChange={(v) => set({ text: v })} />
          </Field>
          <Field label="Link URL">
            <TextInput value={d.href} onChange={(v) => set({ href: v })} />
          </Field>
          <Field label="Alignment">
            <SelectInput value={d.align} onChange={(v) => set({ align: v })} options={ALIGN} />
          </Field>
          <Field label="Background">
            <ColorInput value={d.bg} onChange={(v) => set({ bg: v })} />
          </Field>
          <Field label="Text color">
            <ColorInput value={d.fg} onChange={(v) => set({ fg: v })} />
          </Field>
        </>
      );
    }

    case "twoColumn": {
      const d = block.data;
      return (
        <>
          <Field label="Left column">
            <TextAreaInput value={d.left} onChange={(v) => set({ left: v })} />
          </Field>
          <Field label="Right column">
            <TextAreaInput value={d.right} onChange={(v) => set({ right: v })} />
          </Field>
        </>
      );
    }

    case "carousel": {
      const d = block.data;
      const setItem = (i: number, patch: object) =>
        set({ items: d.items.map((it, j) => (j === i ? { ...it, ...patch } : it)) });
      return (
        <>
          <Text className={styles.note}>
            Outlook can&apos;t run interactive carousels. This renders as a static row of{" "}
            {d.items.length} images — the standard email-safe equivalent.
          </Text>
          {d.items.map((item, i) => (
            <div key={i} className={styles.item}>
              <div className={styles.itemHead}>
                <span className={styles.itemLabel}>Slide {i + 1}</span>
                {d.items.length > 2 && (
                  <Button
                    size="small"
                    appearance="subtle"
                    onClick={() => set({ items: d.items.filter((_, j) => j !== i) })}
                  >
                    Remove
                  </Button>
                )}
              </div>
              <Field label="Image URL">
                <TextInput value={item.src} onChange={(v) => setItem(i, { src: v })} />
              </Field>
              <Field label="Alt">
                <TextInput value={item.alt} onChange={(v) => setItem(i, { alt: v })} />
              </Field>
              <Field label="Caption">
                <TextInput value={item.caption} onChange={(v) => setItem(i, { caption: v })} />
              </Field>
            </div>
          ))}
          {d.items.length < 4 && (
            <Button
              className={styles.addBtn}
              appearance="outline"
              onClick={() => set({ items: [...d.items, newCarouselItem()] })}
            >
              Add slide
            </Button>
          )}
        </>
      );
    }

    case "accordion": {
      const d = block.data;
      const setItem = (i: number, patch: object) =>
        set({ items: d.items.map((it, j) => (j === i ? { ...it, ...patch } : it)) });
      return (
        <>
          <Text className={styles.note}>
            Outlook can&apos;t run collapsible accordions. This renders as numbered, styled Q&amp;A
            cards — always expanded, but visually structured.
          </Text>
          {d.items.map((item, i) => (
            <div key={i} className={styles.item}>
              <div className={styles.itemHead}>
                <span className={styles.itemLabel}>
                  Item {String(i + 1).padStart(2, "0")}
                </span>
                {d.items.length > 1 && (
                  <Button
                    size="small"
                    appearance="subtle"
                    onClick={() => set({ items: d.items.filter((_, j) => j !== i) })}
                  >
                    Remove
                  </Button>
                )}
              </div>
              <Field label="Question">
                <TextInput value={item.question} onChange={(v) => setItem(i, { question: v })} />
              </Field>
              <Field label="Answer">
                <TextAreaInput
                  rows={2}
                  value={item.answer}
                  onChange={(v) => setItem(i, { answer: v })}
                />
              </Field>
            </div>
          ))}
          {d.items.length < 8 && (
            <Button
              className={styles.addBtn}
              appearance="outline"
              onClick={() => set({ items: [...d.items, newAccordionItem()] })}
            >
              Add item
            </Button>
          )}
        </>
      );
    }

    case "divider": {
      const d = block.data;
      return (
        <Field label="Line color">
          <ColorInput value={d.color} onChange={(v) => set({ color: v })} />
        </Field>
      );
    }

    case "spacer": {
      const d = block.data;
      return (
        <Field label={`Height: ${d.height}px`}>
          <Slider
            min={8}
            max={80}
            value={d.height}
            onChange={(_, data) => set({ height: data.value })}
          />
        </Field>
      );
    }

    case "footer": {
      const d = block.data;
      return (
        <>
          <Field label="Company">
            <TextInput value={d.company} onChange={(v) => set({ company: v })} />
          </Field>
          <Field label="Address">
            <TextInput value={d.address} onChange={(v) => set({ address: v })} />
          </Field>
          <Field label="Unsubscribe URL">
            <TextInput value={d.unsubUrl} onChange={(v) => set({ unsubUrl: v })} />
          </Field>
        </>
      );
    }
  }
}
