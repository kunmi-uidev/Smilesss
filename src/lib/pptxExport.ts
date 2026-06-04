/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import pptxgen from "pptxgenjs";
import { Presentation, Slide, ColorPalette } from "../types";
import { THEME_PRESETS } from "./themePresets";

/**
 * Clean hex for PPTXGenJS (must be uppercase and without '#' prefix)
 */
function cleanHex(color: string): string {
  if (!color) return "000000";
  return color.trim().replace("#", "").toUpperCase();
}

/**
 * Generate and download a PowerPoint presentation from a Slide Presentation Model
 */
export async function exportToPPTX(presentation: Presentation): Promise<void> {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_16x9";
  pptx.title = presentation.title || "AI Generated Presentation";

  const theme = THEME_PRESETS[presentation.themeId];
  const palette = presentation.palette;

  // Determine standard font mappings
  let fontHeading = "Arial Black";
  let fontBody = "Arial";
  if (presentation.themeId === "warm-editorial") {
    fontHeading = "Georgia";
    fontBody = "Calibri";
  } else if (presentation.themeId === "brutalist-mono") {
    fontHeading = "Courier New";
    fontBody = "Courier New";
  } else if (presentation.themeId === "cosmic-slate") {
    fontHeading = "Trebuchet MS";
    fontBody = "Calibri";
  }

  // Add all slides
  presentation.slides.forEach((slideData: Slide, index: number) => {
    const slide = pptx.addSlide();
    
    // Resolve slide-specific colors (falling back to theme bounds)
    const currentBgHex = cleanHex(slideData.bgColor || palette.background);
    const currentTextHex = cleanHex(slideData.textColor || palette.text);
    const currentPrimaryHex = cleanHex(slideData.primaryColor || palette.primary);
    const currentAccentHex = cleanHex(slideData.accentColor || palette.accent);
    const currentCardBgHex = cleanHex(slideData.cardBgColor || palette.cardBg || palette.background);
    const currentBorderHex = cleanHex(slideData.borderColor || palette.border);

    slide.background = { fill: currentBgHex };

    // Format slide title depending on whether it's a title slide
    const isTitle = slideData.layout === "title-slide";

    // Optional top corner badge
    if (slideData.badge) {
      slide.addText(slideData.badge.toUpperCase(), {
        x: 0.6,
        y: 0.15,
        w: 5.0,
        h: 0.3,
        fontSize: 10,
        fontFace: fontBody,
        color: currentAccentHex,
        bold: true,
      });
    }

    if (isTitle) {
      // 1. Title Slide Layout
      slide.addText(presentation.title || "BUSINESS PRESENTATION", {
        x: 1.0,
        y: 1.8,
        w: 8.0,
        h: 1.5,
        fontSize: 36,
        fontFace: fontHeading,
        color: currentPrimaryHex,
        bold: true,
        align: "center",
      });

      slide.addText(slideData.title || "AI Powered Pitch Deck", {
        x: 1.0,
        y: 3.2,
        w: 8.0,
        h: 1.0,
        fontSize: 20,
        fontFace: fontBody,
        color: currentTextHex,
        align: "center",
      });

      if (slideData.content && slideData.content.length > 0) {
        slide.addText(slideData.content.join("  |  "), {
          x: 1.0,
          y: 4.4,
          w: 8.0,
          h: 0.6,
          fontSize: 12,
          fontFace: fontBody,
          color: currentAccentHex,
          align: "center",
          italic: true,
        });
      }

    } else if (slideData.layout === "two-column") {
      // 2. Two-Column Layout
      // Slide Title
      slide.addText(slideData.title, {
        x: 0.6,
        y: 0.5,
        w: 8.8,
        h: 0.8,
        fontSize: 26,
        fontFace: fontHeading,
        color: currentPrimaryHex,
        bold: true,
      });

      const half = Math.ceil(slideData.content.length / 2);
      const col1Bullets = slideData.content.slice(0, half);
      const col2Bullets = slideData.content.slice(half);

      // Card Column 1 Background & Text
      slide.addShape(pptx.shapes.RECTANGLE, {
        x: 0.6,
        y: 1.5,
        w: 4.2,
        h: 3.4,
        fill: { color: currentCardBgHex },
        line: { color: currentBorderHex, width: 1 },
      });

      slide.addText(
        col1Bullets.map(bullet => ({ text: "• " + bullet + "\n\n", options: { fontSize: 13, color: currentTextHex } })),
        {
          x: 0.8,
          y: 1.7,
          w: 3.8,
          h: 3.0,
          fontFace: fontBody,
          align: "left",
        }
      );

      // Card Column 2 Background & Text
      slide.addShape(pptx.shapes.RECTANGLE, {
        x: 5.2,
        y: 1.5,
        w: 4.2,
        h: 3.4,
        fill: { color: currentCardBgHex },
        line: { color: currentBorderHex, width: 1 },
      });

      slide.addText(
        col2Bullets.map(bullet => ({ text: "• " + bullet + "\n\n", options: { fontSize: 13, color: currentTextHex } })),
        {
          x: 5.4,
          y: 1.7,
          w: 3.8,
          h: 3.0,
          fontFace: fontBody,
          align: "left",
        }
      );

    } else if (slideData.layout === "quote-slide") {
      // 3. Quote Slide Block
      slide.addShape(pptx.shapes.RECTANGLE, {
        x: 1.0,
        y: 1.2,
        w: 8.0,
        h: 3.2,
        fill: { color: currentCardBgHex },
        line: { color: currentPrimaryHex, width: 2 },
      });

      slide.addText("“", {
        x: 1.2,
        y: 1.3,
        w: 1.0,
        h: 0.6,
        fontSize: 44,
        fontFace: fontHeading,
        color: currentAccentHex,
        bold: true,
      });

      slide.addText(slideData.title, {
        x: 1.5,
        y: 1.7,
        w: 7.0,
        h: 1.2,
        fontSize: 22,
        fontFace: fontHeading,
        color: currentTextHex,
        bold: true,
        italic: true,
        align: "center",
      });

      if (slideData.content && slideData.content.length > 0) {
        slide.addText("- " + slideData.content.join(", "), {
          x: 1.5,
          y: 3.1,
          w: 7.0,
          h: 0.8,
          fontSize: 14,
          fontFace: fontBody,
          color: currentPrimaryHex,
          align: "center",
        });
      }

    } else if (slideData.layout === "image-left" && slideData.imageUrl) {
      // 4. Image Left Layout
      // Left Image block
      slide.addImage({
        path: slideData.imageUrl,
        x: 0.6,
        y: 1.0,
        w: 4.2,
        h: 3.9,
      });

      // Right Text block
      slide.addText(slideData.title, {
        x: 5.2,
        y: 1.0,
        w: 4.2,
        h: 0.8,
        fontSize: 24,
        fontFace: fontHeading,
        color: currentPrimaryHex,
        bold: true,
      });

      slide.addText(
        slideData.content.map(bullet => ({ text: "• " + bullet + "\n\n", options: { fontSize: 13, color: currentTextHex } })),
        {
          x: 5.2,
          y: 1.9,
          w: 4.2,
          h: 3.0,
          fontFace: fontBody,
        }
      );

    } else if (slideData.layout === "image-right" && slideData.imageUrl) {
      // 5. Image Right Layout
      // Left Text block
      slide.addText(slideData.title, {
        x: 0.6,
        y: 1.0,
        w: 4.2,
        h: 0.8,
        fontSize: 24,
        fontFace: fontHeading,
        color: currentPrimaryHex,
        bold: true,
      });

      slide.addText(
        slideData.content.map(bullet => ({ text: "• " + bullet + "\n\n", options: { fontSize: 13, color: currentTextHex } })),
        {
          x: 0.6,
          y: 1.9,
          w: 4.2,
          h: 3.0,
          fontFace: fontBody,
        }
      );

      // Right Image block
      slide.addImage({
        path: slideData.imageUrl,
        x: 5.2,
        y: 1.0,
        w: 4.2,
        h: 3.9,
      });

    } else if (slideData.layout === "stats-bento") {
      // 6. Stats Bento Grid
      slide.addText(slideData.title, {
        x: 0.6,
        y: 0.4,
        w: 8.8,
        h: 0.7,
        fontSize: 26,
        fontFace: fontHeading,
        color: currentPrimaryHex,
        bold: true,
      });

      const grids = [
        { x: 0.6, y: 1.3, w: 4.2, h: 1.7 },
        { x: 5.2, y: 1.3, w: 4.2, h: 1.7 },
        { x: 0.6, y: 3.2, w: 4.2, h: 1.7 },
        { x: 5.2, y: 3.2, w: 4.2, h: 1.7 },
      ];

      for (let i = 0; i < 4; i++) {
        const item = slideData.content[i] || "";
        const grid = grids[i];

        // Bento card backing
        slide.addShape(pptx.shapes.RECTANGLE, {
          x: grid.x,
          y: grid.y,
          w: grid.w,
          h: grid.h,
          fill: { color: currentCardBgHex },
          line: { color: currentBorderHex, width: 1 },
        });

        // Split standard stat formats like "99% Growth rate"
        const doubleSpaceIndex = item.indexOf(" ");
        let bigMetric = item;
        let subText = "";

        if (doubleSpaceIndex > 0) {
          bigMetric = item.substring(0, doubleSpaceIndex);
          subText = item.substring(doubleSpaceIndex + 1);
        }

        slide.addText(bigMetric, {
          x: grid.x + 0.2,
          y: grid.y + 0.2,
          w: grid.w - 0.4,
          h: 0.6,
          fontSize: 24,
          fontFace: fontHeading,
          color: currentAccentHex,
          bold: true,
        });

        if (subText) {
          slide.addText(subText, {
            x: grid.x + 0.2,
            y: grid.y + 0.8,
            w: grid.w - 0.4,
            h: 0.7,
            fontSize: 11,
            fontFace: fontBody,
            color: currentTextHex,
          });
        }
      }

    } else {
      // 7. Standard / Headline Bullet & Minimal Split
      // Slide Title
      slide.addText(slideData.title, {
        x: 0.6,
        y: 0.5,
        w: 8.8,
        h: 0.8,
        fontSize: 28,
        fontFace: fontHeading,
        color: currentPrimaryHex,
        bold: true,
      });

      if (slideData.imageUrl) {
        // Splitting into background side image if minimal split setup
        slide.addImage({
          path: slideData.imageUrl,
          x: 5.4,
          y: 1.5,
          w: 4.0,
          h: 3.4,
        });

        slide.addText(
          slideData.content.map(bullet => ({ text: "• " + bullet + "\n\n", options: { fontSize: 13, color: currentTextHex } })),
          {
            x: 0.6,
            y: 1.5,
            w: 4.5,
            h: 3.4,
            fontFace: fontBody,
          }
        );
      } else {
        // Flat styled list
        slide.addText(
          slideData.content.map(bullet => ({ text: "• " + bullet + "\n\n", options: { fontSize: 14, color: currentTextHex } })),
          {
            x: 0.6,
            y: 1.6,
            w: 8.8,
            h: 3.4,
            fontFace: fontBody,
          }
        );
      }
    }
  });

  // Save/Download presentation file triggered automatically
  await pptx.writeFile({ fileName: `${presentation.title.toLowerCase().replace(/\s+/g, "_")}.pptx` });
}
