/* Block model for the newsletter document.
 *
 * A newsletter is an ordered list of `Block`s. `Block` is a discriminated
 * union keyed on `type`, so the editor and renderers get the exact data
 * shape for each block without casts. */

export type Align = "left" | "center" | "right";
export type TextAlign = Align | "justify";

export interface CarouselItem {
  src: string;
  alt: string;
  caption: string;
}

export interface AccordionItem {
  question: string;
  answer: string;
}

export interface HeaderData {
  title: string;
  tagline: string;
  bg: string;
  fg: string;
}

export interface HeadingData {
  text: string;
  level: "h1" | "h2" | "h3";
  align: Align;
  color: string;
}

export interface TextData {
  content: string;
  color: string;
  align: TextAlign;
}

export interface ImageData {
  src: string;
  alt: string;
  href: string;
}

export interface ImageTextData {
  src: string;
  alt: string;
  heading: string;
  body: string;
  layout: "image-left" | "image-right";
  imageWidth: number;
  linkText: string;
  linkUrl: string;
}

export interface ButtonData {
  text: string;
  href: string;
  bg: string;
  fg: string;
  align: Align;
}

export interface TwoColumnData {
  left: string;
  right: string;
}

export interface CarouselData {
  items: CarouselItem[];
}

export interface AccordionData {
  items: AccordionItem[];
}

export interface DividerData {
  color: string;
}

export interface SpacerData {
  height: number;
}

export interface FooterData {
  company: string;
  address: string;
  unsubUrl: string;
}

interface BlockOf<T extends string, D> {
  id: string;
  type: T;
  data: D;
}

export type Block =
  | BlockOf<"header", HeaderData>
  | BlockOf<"heading", HeadingData>
  | BlockOf<"text", TextData>
  | BlockOf<"image", ImageData>
  | BlockOf<"imageText", ImageTextData>
  | BlockOf<"button", ButtonData>
  | BlockOf<"twoColumn", TwoColumnData>
  | BlockOf<"carousel", CarouselData>
  | BlockOf<"accordion", AccordionData>
  | BlockOf<"divider", DividerData>
  | BlockOf<"spacer", SpacerData>
  | BlockOf<"footer", FooterData>;

export type BlockType = Block["type"];

/** The data shape that belongs to a given block type. */
export type DataOf<T extends BlockType> = Extract<Block, { type: T }>["data"];
