/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { ColorPalette } from "../types";
import { normalizeHex } from "./themePresets";

/**
 * Calculates the relative luminance of a RGB color according to WCAG 2.1 guidelines
 * https://www.w3.org/WAI/GL/wiki/Relative_luminance
 */
export function getLuminance(hexColor: string): number {
  try {
    const cleanHex = normalizeHex(hexColor, "#000000").replace("#", "");
    const r = parseInt(cleanHex.substring(0, 2), 16) || 0;
    const g = parseInt(cleanHex.substring(2, 4), 16) || 0;
    const b = parseInt(cleanHex.substring(4, 6), 16) || 0;

    const [rs, gs, bs] = [r, g, b].map((val) => {
      const sRGB = val / 255;
      return sRGB <= 0.03928 ? sRGB / 12.92 : Math.pow((sRGB + 0.055) / 1.055, 2.4);
    });

    return 0.2126 * rs + 0.7152 * gs + 0.0722 * bs;
  } catch (e) {
    return 0;
  }
}

/**
 * Calculates WCAG contrast ratio between two colors (range: 1.0 to 21.0)
 */
export function getContrastRatio(colorA: string, colorB: string): number {
  const lumA = getLuminance(colorA);
  const lumB = getLuminance(colorB);
  const lighter = Math.max(lumA, lumB);
  const darker = Math.min(lumA, lumB);
  const ratio = (lighter + 0.05) / (darker + 0.05);
  return Math.round(ratio * 100) / 100;
}

export type WCAGLevel = "AAA" | "AA" | "AA-Large" | "Fail";

export interface AccessibilityTestResult {
  ratio: number;
  level: WCAGLevel;
  isAccessible: boolean; // meets at least AA for body or large text
  isAAPass: boolean; // meets 4.5:1
  isAAAPass: boolean; // meets 7.0:1
  label: string;
  badgeColor: string;
  suggestedTextColor: string;
  recommendation: string;
}

/**
 * Tests color accessibility between background and foreground colors
 */
export function testColorAccessibility(
  backgroundHex: string,
  foregroundHex: string,
  minRatio: number = 4.5
): AccessibilityTestResult {
  const bg = normalizeHex(backgroundHex, "#0B0F19");
  const fg = normalizeHex(foregroundHex, "#FFFFFF");
  const ratio = getContrastRatio(bg, fg);

  const isAAAPass = ratio >= 7.0;
  const isAAPass = ratio >= 4.5;
  const isAALargePass = ratio >= 3.0;

  let level: WCAGLevel = "Fail";
  let label = "Low Contrast (Fail)";
  let badgeColor = "bg-rose-500/20 text-rose-300 border-rose-500/40";
  let recommendation = "Color contrast is below WCAG AA guidelines. Text may be hard to read.";

  if (isAAAPass) {
    level = "AAA";
    label = `WCAG AAA (${ratio}:1)`;
    badgeColor = "bg-emerald-500/20 text-emerald-300 border-emerald-500/40";
    recommendation = "Excellent contrast! Exceeds WCAG AAA requirements for maximum readability.";
  } else if (isAAPass) {
    level = "AA";
    label = `WCAG AA (${ratio}:1)`;
    badgeColor = "bg-teal-500/20 text-teal-300 border-teal-500/40";
    recommendation = "Good contrast! Meets WCAG AA standard (4.5:1+) for normal body text.";
  } else if (isAALargePass) {
    level = "AA-Large";
    label = `WCAG AA Large (${ratio}:1)`;
    badgeColor = "bg-amber-500/20 text-amber-300 border-amber-500/40";
    recommendation = "Acceptable for large headings (3.0:1+), but body text should be higher contrast.";
  }

  // Suggest optimal high-contrast alternative if it fails
  const bgLum = getLuminance(bg);
  const suggestedTextColor = bgLum > 0.45 ? "#0F172A" : "#FFFFFF";

  return {
    ratio,
    level,
    isAccessible: ratio >= minRatio || isAALargePass,
    isAAPass,
    isAAAPass,
    label,
    badgeColor,
    suggestedTextColor,
    recommendation
  };
}

/**
 * Returns an accessible foreground color for a given background, ensuring at least minRatio contrast
 */
export function ensureAccessibleTextColor(
  backgroundHex: string,
  preferredColorHex: string = "#FFFFFF",
  minRatio: number = 4.5
): string {
  const bg = normalizeHex(backgroundHex, "#0B0F19");
  const preferred = normalizeHex(preferredColorHex, "#FFFFFF");
  const currentRatio = getContrastRatio(bg, preferred);

  if (currentRatio >= minRatio) {
    return preferred;
  }

  // If preferred color doesn't pass, choose between high-contrast white or dark charcoal
  const whiteRatio = getContrastRatio(bg, "#FFFFFF");
  const blackRatio = getContrastRatio(bg, "#0F172A");

  if (whiteRatio >= blackRatio) {
    return "#FFFFFF";
  } else {
    return "#0F172A";
  }
}

/**
 * Validates and adjusts an entire color palette to guarantee WCAG AA accessibility
 */
export function ensureAccessiblePalette(palette: ColorPalette): {
  palette: ColorPalette;
  testedResults: {
    textOnBg: AccessibilityTestResult;
    textOnCard: AccessibilityTestResult;
    primaryOnBg: AccessibilityTestResult;
    accentOnBg: AccessibilityTestResult;
  };
} {
  const bg = normalizeHex(palette.background, "#0B0F19");
  const cardBg = normalizeHex(palette.cardBg || bg, "#161E2E");

  const textOnBg = testColorAccessibility(bg, palette.text);
  const textOnCard = testColorAccessibility(cardBg, palette.text);
  const primaryOnBg = testColorAccessibility(bg, palette.primary, 3.0);
  const accentOnBg = testColorAccessibility(bg, palette.accent, 3.0);

  // Auto-adjust text if contrast is below 4.5:1 on background or card
  let adjustedText = palette.text;
  if (!textOnBg.isAAPass) {
    adjustedText = textOnBg.suggestedTextColor;
  }

  let adjustedTextLight = palette.textLight;
  if (adjustedTextLight) {
    const lightRatio = getContrastRatio(bg, adjustedTextLight);
    if (lightRatio < 3.0) {
      adjustedTextLight = getLuminance(bg) > 0.45 ? "#64748B" : "#94A3B8";
    }
  }

  const adjustedPalette: ColorPalette = {
    ...palette,
    text: adjustedText,
    textLight: adjustedTextLight
  };

  return {
    palette: adjustedPalette,
    testedResults: {
      textOnBg,
      textOnCard,
      primaryOnBg,
      accentOnBg
    }
  };
}
