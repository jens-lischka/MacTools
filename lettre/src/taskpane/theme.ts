import { webLightTheme, type Theme } from "@fluentui/react-components";

/* Lettre brand theme — a warm terracotta accent (#c2410c) layered onto
 * Fluent's light theme by overriding only the brand tokens. */

const BRAND = "#c2410c";
const BRAND_HOVER = "#9a3412";
const BRAND_PRESSED = "#7c2d12";

export const lettreTheme: Theme = {
  ...webLightTheme,

  colorBrandBackground: BRAND,
  colorBrandBackgroundHover: BRAND_HOVER,
  colorBrandBackgroundPressed: BRAND_PRESSED,
  colorBrandBackgroundSelected: BRAND_HOVER,

  colorCompoundBrandBackground: BRAND,
  colorCompoundBrandBackgroundHover: BRAND_HOVER,
  colorCompoundBrandBackgroundPressed: BRAND_PRESSED,

  colorBrandStroke1: BRAND,
  colorCompoundBrandStroke: BRAND,
  colorCompoundBrandStrokeHover: BRAND_HOVER,
  colorCompoundBrandStrokePressed: BRAND_PRESSED,

  colorBrandForeground1: BRAND,
  colorBrandForeground2: BRAND,
  colorBrandForegroundLink: BRAND,
  colorBrandForegroundLinkHover: BRAND_HOVER,

  colorNeutralForegroundOnBrand: "#ffffff",
};
