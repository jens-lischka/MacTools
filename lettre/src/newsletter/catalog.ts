/* Block catalogue, defaults, and factory helpers. */

import type { Block, BlockType, DataOf } from "./types";

export const uid = (): string => Math.random().toString(36).slice(2, 10);

const sampleImg = (seed: string): string =>
  `https://images.unsplash.com/photo-${seed}?w=800&q=80`;

/** Block types offered in the "Add block" palette, in display order. */
export const CATALOG: { type: BlockType; label: string }[] = [
  { type: "header", label: "Header" },
  { type: "heading", label: "Heading" },
  { type: "text", label: "Text" },
  { type: "image", label: "Image" },
  { type: "imageText", label: "Img + Text" },
  { type: "button", label: "Button" },
  { type: "twoColumn", label: "2 Columns" },
  { type: "carousel", label: "Carousel" },
  { type: "accordion", label: "Accordion" },
  { type: "divider", label: "Divider" },
  { type: "spacer", label: "Spacer" },
  { type: "footer", label: "Footer" },
];

/** Starting data for each block type. */
export const DEFAULTS: { [K in BlockType]: DataOf<K> } = {
  header: {
    title: "The Quarterly",
    tagline: "Issue 04 — November 2025",
    bg: "#1f2421",
    fg: "#f5f4ef",
  },
  heading: { text: "A new headline", level: "h2", align: "left", color: "#1a1a1a" },
  text: {
    content:
      "Write your paragraph here. Newsletters get scanned, not read — keep it tight, lead with the point, and let the rest fall away.",
    color: "#333333",
    align: "left",
  },
  image: {
    src: sampleImg("1497366216548-37526070297c"),
    alt: "Decorative image",
    href: "",
  },
  imageText: {
    src: sampleImg("1551434678-e076c223a692"),
    alt: "Article image",
    heading: "A short, scannable headline",
    body: "Two or three lines of supporting copy. Lead with the point — most readers stop after the first sentence anyway.",
    layout: "image-left",
    imageWidth: 220,
    linkText: "Read more",
    linkUrl: "https://example.com",
  },
  button: {
    text: "Read the full update",
    href: "https://example.com",
    bg: "#c2410c",
    fg: "#ffffff",
    align: "left",
  },
  twoColumn: {
    left: "Left column. Useful for pairing two updates side-by-side — product news on one side, team news on the other.",
    right:
      "Right column. Keep both sides roughly equal in length so the layout reads cleanly across clients.",
  },
  carousel: {
    items: [
      { src: sampleImg("1517245386807-bb43f82c33c4"), alt: "Slide 1", caption: "Munich offsite" },
      { src: sampleImg("1522071820081-009f0129c71c"), alt: "Slide 2", caption: "Engineering retreat" },
      { src: sampleImg("1531973576160-7125cd663d86"), alt: "Slide 3", caption: "Product launch" },
    ],
  },
  accordion: {
    items: [
      {
        question: "When does the new policy take effect?",
        answer:
          "The updated travel policy applies to all bookings made on or after 1 December 2025. Existing reservations are not affected.",
      },
      {
        question: "Who do I contact with questions?",
        answer:
          "Reach out to people-ops@example.com — the team aims to respond within one business day.",
      },
      {
        question: "Where can I find the full document?",
        answer:
          "The complete policy is on the intranet under People → Policies → Travel & Expenses.",
      },
    ],
  },
  divider: { color: "#e5e3dc" },
  spacer: { height: 24 },
  footer: {
    company: "Your Company GmbH",
    address: "Beispielstraße 1, 60311 Frankfurt am Main",
    unsubUrl: "https://example.com/unsubscribe",
  },
};

const clone = <T>(value: T): T => JSON.parse(JSON.stringify(value)) as T;

/** Create a fresh block with default data. */
export function createBlock(type: BlockType): Block {
  return { id: uid(), type, data: clone(DEFAULTS[type]) } as Block;
}

/** Template for the "New slide" / "New item" buttons. */
export const newCarouselItem = () => ({
  src: sampleImg("1500530855697-b586d89ba3ee"),
  alt: "New slide",
  caption: "",
});
export const newAccordionItem = () => ({
  question: "New question?",
  answer: "Answer text.",
});

/** A small starter newsletter shown when the task pane first opens. */
export function defaultNewsletter(): Block[] {
  const make = (type: BlockType, patch: object = {}): Block => {
    const block = createBlock(type);
    Object.assign(block.data, patch);
    return block;
  };
  return [
    make("header"),
    make("heading", { text: "A note from the editor" }),
    make("text"),
    make("imageText", { heading: "New travel policy in effect Dec 1" }),
    make("heading", { text: "Frequently asked", level: "h3" }),
    make("accordion"),
    make("button", { align: "center" }),
    make("divider"),
    make("footer"),
  ];
}
