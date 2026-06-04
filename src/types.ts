/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export enum ThemeId {
  COSMIC_SLATE = "cosmic-slate",
  MINIMAL_SWISS = "minimal-swiss",
  WARM_EDITORIAL = "warm-editorial",
  BRUTALIST_MONO = "brutalist-mono",
  CREATIVE_AVANTGARDE = "creative-avantgarde"
}

export interface ColorPalette {
  primary: string;     // The brand/primary color
  secondary: string;   // Secondary brand color
  accent: string;      // Accent/highlight color
  background: string;  // Main canvas background
  text: string;        // Text body color
  textLight?: string;  // Muted/light text color for subtexts
  cardBg?: string;     // Color for cards inside slides
  border?: string;     // Border color for styling elements
}

export type SlideLayout =
  | "title-slide"
  | "two-column"
  | "headline-bullet"
  | "quote-slide"
  | "image-left"
  | "image-right"
  | "minimal-split"
  | "stats-bento";

export interface Slide {
  id: string;
  title: string;
  content: string[];
  layout: SlideLayout;
  imageUrl?: string;
  imageCaption?: string;
  notes?: string;
  imageSearchQuery?: string;
  badge?: string;
  bgColor?: string;
  textColor?: string;
  accentColor?: string;
  primaryColor?: string;
  cardBgColor?: string;
  borderColor?: string;
}

export interface Presentation {
  title: string;
  themeId: ThemeId;
  brandColor: string; // Dynamic hex brand color
  palette: ColorPalette;
  slides: Slide[];
}

export interface ThemeConfig {
  id: ThemeId;
  name: string;
  description: string;
  fontHeading: string; // Tailwind font class (sans/serif/mono) or system font
  fontBody: string;
  defaultPalette: (brandColor: string) => ColorPalette;
}
