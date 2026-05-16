import { webLightTheme, type Theme } from "@fluentui/react-components";

/**
 * MacTools brand theme — primary colour #000F47 (deep navy) and Noto Sans.
 *
 * Built by overriding the brand tokens of Fluent's light theme rather than
 * generating a full brand ramp, which keeps the palette predictable.
 */

const BRAND = "#000F47";
const BRAND_HOVER = "#1A2A6B";
const BRAND_PRESSED = "#00072E";

export const macToolsTheme: Theme = {
  ...webLightTheme,
  fontFamilyBase: "'Noto Sans', 'Segoe UI', sans-serif",

  colorBrandBackground: BRAND,
  colorBrandBackgroundHover: BRAND_HOVER,
  colorBrandBackgroundPressed: BRAND_PRESSED,
  colorBrandBackgroundSelected: BRAND_HOVER,

  colorCompoundBrandBackground: BRAND,
  colorCompoundBrandBackgroundHover: BRAND_HOVER,
  colorCompoundBrandBackgroundPressed: BRAND_PRESSED,

  colorBrandStroke1: BRAND,
  colorBrandStroke2: "#A6AFCB",
  colorCompoundBrandStroke: BRAND,
  colorCompoundBrandStrokeHover: BRAND_HOVER,
  colorCompoundBrandStrokePressed: BRAND_PRESSED,

  colorBrandForeground1: BRAND,
  colorBrandForeground2: BRAND,
  colorBrandForegroundLink: BRAND,
  colorBrandForegroundLinkHover: BRAND_HOVER,

  colorNeutralForegroundOnBrand: "#FFFFFF",
};
