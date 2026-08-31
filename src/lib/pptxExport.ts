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
 * Fixed at 1920x1080 Full HD (16:9 Standard widescreen: 13.333 in x 7.5 in)
 */
export async function exportToPPTX(presentation: Presentation): Promise<void> {
  const pptx = new pptxgen();
  // Set 1920x1080 Full HD Layout (13.333 x 7.5 inches standard widescreen)
  pptx.defineLayout({ name: "FULL_HD_1920x1080", width: 13.333, height: 7.5 });
  pptx.layout = "FULL_HD_1920x1080";
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

  const totalSlides = presentation.slides.length;

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

    const isTitle = slideData.layout === "title-slide";
    const isThankYou = slideData.title.toLowerCase().includes("thank you") || (slideData.badge && slideData.badge.toLowerCase().includes("thank"));
    const scaleFactor = slideData.fontSize === "small" ? 0.85 : slideData.fontSize === "large" ? 1.15 : 1.0;

    // Optional top corner badge (1920x1080 scaled)
    if (slideData.badge && !isTitle) {
      slide.addText(slideData.badge.toUpperCase(), {
        x: 0.8,
        y: 0.45,
        w: 6.0,
        h: 0.35,
        fontSize: Math.round(11 * scaleFactor),
        fontFace: fontHeading,
        color: currentAccentHex,
        bold: true,
      });
    }

    if (isTitle) {
      // 1. Title Slide Layout (1920x1080)
      slide.addShape(pptx.shapes.RECTANGLE, {
        x: 0.8,
        y: 1.2,
        w: 11.733,
        h: 5.2,
        fill: { color: currentCardBgHex },
        line: { color: currentBorderHex, width: 1.5 },
      });

      if (slideData.badge) {
        slide.addText(slideData.badge.toUpperCase(), {
          x: 1.2,
          y: 1.6,
          w: 10.933,
          h: 0.4,
          fontSize: Math.round(12 * scaleFactor),
          fontFace: fontHeading,
          color: currentAccentHex,
          bold: true,
          align: "center",
        });
      }

      slide.addText(presentation.title || "BUSINESS PRESENTATION", {
        x: 1.2,
        y: 2.2,
        w: 10.933,
        h: 1.8,
        fontSize: Math.round(42 * scaleFactor),
        fontFace: fontHeading,
        color: currentPrimaryHex,
        bold: true,
        align: "center",
      });

      slide.addText(slideData.title || "AI Powered Presentation Deck", {
        x: 1.2,
        y: 4.1,
        w: 10.933,
        h: 1.0,
        fontSize: Math.round(24 * scaleFactor),
        fontFace: fontBody,
        color: currentTextHex,
        align: "center",
      });

      if (slideData.content && slideData.content.length > 0) {
        slide.addText(slideData.content.join("   •   "), {
          x: 1.2,
          y: 5.3,
          w: 10.933,
          h: 0.6,
          fontSize: Math.round(14 * scaleFactor),
          fontFace: fontBody,
          color: currentAccentHex,
          align: "center",
          italic: true,
        });
      }

    } else if (isThankYou) {
      // Special: Thank you for listening closing slide
      slide.addShape(pptx.shapes.RECTANGLE, {
        x: 1.2,
        y: 1.2,
        w: 10.933,
        h: 5.2,
        fill: { color: currentCardBgHex },
        line: { color: currentPrimaryHex, width: 2 },
      });

      slide.addText("“", {
        x: 1.5,
        y: 1.5,
        w: 1.0,
        h: 0.8,
        fontSize: Math.round(60 * scaleFactor),
        fontFace: fontHeading,
        color: currentAccentHex,
        bold: true,
      });

      slide.addText(slideData.title, {
        x: 1.5,
        y: 2.0,
        w: 10.333,
        h: 1.4,
        fontSize: Math.round(38 * scaleFactor),
        fontFace: fontHeading,
        color: currentPrimaryHex,
        bold: true,
        align: "center",
      });

      if (slideData.content && slideData.content.length > 0) {
        slide.addText(
          slideData.content.map(bullet => ({
            text: "• " + bullet.replace(/^[•\-\*\s]+/, '') + "\n",
            options: { fontSize: Math.round(16 * scaleFactor), color: currentTextHex }
          })),
          {
            x: 2.0,
            y: 3.6,
            w: 9.333,
            h: 2.2,
            fontFace: fontBody,
            align: "center",
          }
        );
      }

    } else if (slideData.layout === "two-column") {
      // 2. Two-Column Layout (1920x1080)
      slide.addText(slideData.title, {
        x: 0.8,
        y: 0.8,
        w: 11.733,
        h: 0.9,
        fontSize: Math.round(30 * scaleFactor),
        fontFace: fontHeading,
        color: currentPrimaryHex,
        bold: true,
      });

      const half = Math.ceil(slideData.content.length / 2);
      const col1Bullets = slideData.content.slice(0, half);
      const col2Bullets = slideData.content.slice(half);

      // Card Column 1
      slide.addShape(pptx.shapes.RECTANGLE, {
        x: 0.8,
        y: 1.8,
        w: 5.65,
        h: 4.8,
        fill: { color: currentCardBgHex },
        line: { color: currentBorderHex, width: 1.5 },
      });

      slide.addText(
        col1Bullets.map(bullet => ({
          text: (bullet.startsWith("•") || bullet.startsWith("-") ? "" : "• ") + bullet + "\n\n",
          options: { fontSize: Math.round(15 * scaleFactor), color: currentTextHex }
        })),
        {
          x: 1.1,
          y: 2.1,
          w: 5.05,
          h: 4.2,
          fontFace: fontBody,
          align: "left",
        }
      );

      // Card Column 2
      slide.addShape(pptx.shapes.RECTANGLE, {
        x: 6.88,
        y: 1.8,
        w: 5.65,
        h: 4.8,
        fill: { color: currentCardBgHex },
        line: { color: currentBorderHex, width: 1.5 },
      });

      slide.addText(
        col2Bullets.map(bullet => ({
          text: (bullet.startsWith("•") || bullet.startsWith("-") ? "" : "• ") + bullet + "\n\n",
          options: { fontSize: Math.round(15 * scaleFactor), color: currentTextHex }
        })),
        {
          x: 7.18,
          y: 2.1,
          w: 5.05,
          h: 4.2,
          fontFace: fontBody,
          align: "left",
        }
      );

    } else if (slideData.layout === "quote-slide") {
      // 3. Quote Slide Block (1920x1080)
      slide.addShape(pptx.shapes.RECTANGLE, {
        x: 1.2,
        y: 1.4,
        w: 10.933,
        h: 4.8,
        fill: { color: currentCardBgHex },
        line: { color: currentPrimaryHex, width: 2 },
      });

      slide.addText("“", {
        x: 1.5,
        y: 1.6,
        w: 1.2,
        h: 0.8,
        fontSize: Math.round(54 * scaleFactor),
        fontFace: fontHeading,
        color: currentAccentHex,
        bold: true,
      });

      slide.addText(slideData.title, {
        x: 1.8,
        y: 2.4,
        w: 9.733,
        h: 2.2,
        fontSize: Math.round(30 * scaleFactor),
        fontFace: fontHeading,
        color: currentTextHex,
        bold: true,
        italic: true,
        align: "center",
      });

      if (slideData.content && slideData.content.length > 0) {
        slide.addText("- " + slideData.content.join(" & "), {
          x: 1.8,
          y: 4.8,
          w: 9.733,
          h: 0.8,
          fontSize: Math.round(17 * scaleFactor),
          fontFace: fontBody,
          color: currentPrimaryHex,
          align: "center",
          bold: true,
        });
      }

    } else if (slideData.layout === "image-left" && slideData.imageUrl) {
      // 4. Image Left Layout (1920x1080)
      slide.addImage({
        path: slideData.imageUrl,
        x: 0.8,
        y: 1.5,
        w: 5.65,
        h: 5.1,
      });

      slide.addText(slideData.title, {
        x: 6.88,
        y: 1.5,
        w: 5.65,
        h: 1.0,
        fontSize: Math.round(28 * scaleFactor),
        fontFace: fontHeading,
        color: currentPrimaryHex,
        bold: true,
      });

      slide.addText(
        slideData.content.map(bullet => ({
          text: (bullet.startsWith("•") || bullet.startsWith("-") ? "" : "• ") + bullet + "\n\n",
          options: { fontSize: Math.round(15 * scaleFactor), color: currentTextHex }
        })),
        {
          x: 6.88,
          y: 2.6,
          w: 5.65,
          h: 4.0,
          fontFace: fontBody,
        }
      );

    } else if (slideData.layout === "image-right" && slideData.imageUrl) {
      // 5. Image Right Layout (1920x1080)
      slide.addText(slideData.title, {
        x: 0.8,
        y: 1.5,
        w: 5.65,
        h: 1.0,
        fontSize: Math.round(28 * scaleFactor),
        fontFace: fontHeading,
        color: currentPrimaryHex,
        bold: true,
      });

      slide.addText(
        slideData.content.map(bullet => ({
          text: (bullet.startsWith("•") || bullet.startsWith("-") ? "" : "• ") + bullet + "\n\n",
          options: { fontSize: Math.round(15 * scaleFactor), color: currentTextHex }
        })),
        {
          x: 0.8,
          y: 2.6,
          w: 5.65,
          h: 4.0,
          fontFace: fontBody,
        }
      );

      slide.addImage({
        path: slideData.imageUrl,
        x: 6.88,
        y: 1.5,
        w: 5.65,
        h: 5.1,
      });

    } else if (slideData.layout === "stats-bento") {
      // 6. Stats Bento Grid (1920x1080)
      slide.addText(slideData.title, {
        x: 0.8,
        y: 0.8,
        w: 11.733,
        h: 0.8,
        fontSize: Math.round(30 * scaleFactor),
        fontFace: fontHeading,
        color: currentPrimaryHex,
        bold: true,
      });

      const grids = [
        { x: 0.8, y: 1.8, w: 5.65, h: 2.3 },
        { x: 6.88, y: 1.8, w: 5.65, h: 2.3 },
        { x: 0.8, y: 4.3, w: 5.65, h: 2.3 },
        { x: 6.88, y: 4.3, w: 5.65, h: 2.3 },
      ];

      for (let i = 0; i < 4; i++) {
        const item = slideData.content[i] || "";
        const grid = grids[i];

        slide.addShape(pptx.shapes.RECTANGLE, {
          x: grid.x,
          y: grid.y,
          w: grid.w,
          h: grid.h,
          fill: { color: currentCardBgHex },
          line: { color: currentBorderHex, width: 1.5 },
        });

        const doubleSpaceIndex = item.indexOf(" ");
        let bigMetric = item;
        let subText = "";

        if (doubleSpaceIndex > 0) {
          bigMetric = item.substring(0, doubleSpaceIndex);
          subText = item.substring(doubleSpaceIndex + 1);
        }

        slide.addText(bigMetric, {
          x: grid.x + 0.3,
          y: grid.y + 0.3,
          w: grid.w - 0.6,
          h: 0.9,
          fontSize: Math.round(32 * scaleFactor),
          fontFace: fontHeading,
          color: currentAccentHex,
          bold: true,
        });

        if (subText) {
          slide.addText(subText, {
            x: grid.x + 0.3,
            y: grid.y + 1.2,
            w: grid.w - 0.6,
            h: 0.8,
            fontSize: Math.round(14 * scaleFactor),
            fontFace: fontBody,
            color: currentTextHex,
          });
        }
      }

    } else if (slideData.layout === "comparison-table") {
      // 7. Comparison Table Layout (1920x1080)
      slide.addText(slideData.title, {
        x: 0.8,
        y: 0.8,
        w: 11.733,
        h: 0.8,
        fontSize: Math.round(30 * scaleFactor),
        fontFace: fontHeading,
        color: currentPrimaryHex,
        bold: true,
      });

      // Parse rows
      const tableRows: pptxgen.TableRow[] = [];
      tableRows.push([
        { text: "EVALUATION DIMENSION", options: { bold: true, color: currentPrimaryHex, fill: { color: currentCardBgHex }, fontSize: Math.round(13 * scaleFactor) } },
        { text: "STANDARD PRACTICES", options: { bold: true, color: currentTextHex, fill: { color: currentCardBgHex }, fontSize: Math.round(13 * scaleFactor) } },
        { text: "SLIDESSS ADVANTAGE", options: { bold: true, color: currentAccentHex, fill: { color: currentCardBgHex }, fontSize: Math.round(13 * scaleFactor) } },
      ]);

      const contentLen = slideData.content.length;
      let hasSep = false;
      slideData.content.forEach(bullet => {
        if (bullet.includes(" vs ") || bullet.includes(" | ") || bullet.includes(" - ")) {
          hasSep = true;
        }
      });

      if (hasSep) {
        slideData.content.forEach((bullet, rIdx) => {
          let aspect = `Dimension ${rIdx + 1}`;
          let left = bullet;
          let right = "";
          for (const sep of [" vs ", " | ", " - "]) {
            if (bullet.includes(sep)) {
              const parts = bullet.split(sep);
              left = parts[0].trim();
              right = parts.slice(1).join(sep).trim();
              if (left.includes(": ")) {
                const p = left.split(": ");
                aspect = p[0].trim();
                left = p.slice(1).join(": ").trim();
              }
              break;
            }
          }
          tableRows.push([
            { text: aspect, options: { bold: true, color: currentPrimaryHex, fontSize: Math.round(12 * scaleFactor) } },
            { text: "✗ " + left, options: { color: currentTextHex, fontSize: Math.round(12 * scaleFactor) } },
            { text: "✓ " + right, options: { bold: true, color: currentAccentHex, fontSize: Math.round(12 * scaleFactor) } },
          ]);
        });
      } else {
        const half = Math.ceil(contentLen / 2);
        for (let i = 0; i < half; i++) {
          const l = slideData.content[i] || "—";
          const r = slideData.content[i + half] || "—";
          tableRows.push([
            { text: `Feature Point ${i + 1}`, options: { bold: true, color: currentPrimaryHex, fontSize: Math.round(12 * scaleFactor) } },
            { text: "✗ " + l, options: { color: currentTextHex, fontSize: Math.round(12 * scaleFactor) } },
            { text: "✓ " + r, options: { bold: true, color: currentAccentHex, fontSize: Math.round(12 * scaleFactor) } },
          ]);
        }
      }

      slide.addTable(tableRows, {
        x: 0.8,
        y: 1.8,
        w: 11.733,
        rowH: 0.7,
        fill: { color: currentCardBgHex },
        border: { type: "solid", pt: 1, color: currentBorderHex },
        fontFace: fontBody,
      });

    } else {
      // 8. Standard Headline / Bullet Layout (1920x1080)
      slide.addText(slideData.title, {
        x: 0.8,
        y: 0.8,
        w: 11.733,
        h: 0.9,
        fontSize: Math.round(32 * scaleFactor),
        fontFace: fontHeading,
        color: currentPrimaryHex,
        bold: true,
      });

      slide.addText(
        slideData.content.map(bullet => ({
          text: (bullet.startsWith("•") || bullet.startsWith("-") ? "" : "• ") + bullet + "\n\n",
          options: { fontSize: Math.round(16 * scaleFactor), color: currentTextHex }
        })),
        {
          x: 0.8,
          y: 2.0,
          w: 11.733,
          h: 4.6,
          fontFace: fontBody,
        }
      );
    }

    // Slide footer (1920x1080 fixed bounds)
    slide.addText(`Slidesss Presentation  |  Slide ${index + 1} of ${totalSlides}`, {
      x: 0.8,
      y: 6.85,
      w: 11.733,
      h: 0.35,
      fontSize: 10,
      fontFace: fontBody,
      color: currentTextHex,
      opacity: 60,
      align: "right",
    });
  });

  // Save/Download presentation file triggered automatically
  await pptx.writeFile({ fileName: `${presentation.title.toLowerCase().replace(/\s+/g, "_")}.pptx` });
}
