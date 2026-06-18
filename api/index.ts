/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import express from "express";
import dotenv from "dotenv";
import { GoogleGenAI, Type } from "@google/genai";

dotenv.config();

const app = express();

app.use(express.json({ limit: "20mb" }));

// Initialize Google GenAI Client with proper telemetry headers
const apiKey = process.env.GEMINI_API_KEY;
const ai = apiKey
  ? new GoogleGenAI({
      apiKey,
      httpOptions: {
        headers: {
          "User-Agent": "aistudio-build",
        },
      },
    })
  : null;

/// Reusable retry utility with exponential backoff for high-impact model endpoints
async function retryWithBackoff<T>(
  fn: () => Promise<T>,
  retries = 4,
  delay = 1000,
  factor = 2
): Promise<T> {
  try {
    return await fn();
  } catch (error: any) {
    let errorString = "";
    try {
      errorString = [
        error?.message,
        error?.toString(),
        JSON.stringify(error)
      ].filter(Boolean).join(" ").toLowerCase();
    } catch (_) {
      errorString = (error?.message || "").toString().toLowerCase();
    }

    const isRateLimitOrUnavailable =
      errorString.includes("503") ||
      errorString.includes("high demand") ||
      errorString.includes("unavailable") ||
      errorString.includes("429") ||
      errorString.includes("quota") ||
      errorString.includes("overburdened") ||
      error?.status === 503 ||
      error?.status === "UNAVAILABLE" ||
      error?.status === 429 ||
      error?.code === 503 ||
      error?.statusCode === 503;

    if (retries > 0 && isRateLimitOrUnavailable) {
      console.warn(`[GEMINI STATUS] Service bottleneck (503/429) detected. Retrying in ${delay}ms... Remaining: ${retries}. Error: ${errorString.slice(0, 150)}`);
      await new Promise((resolve) => setTimeout(resolve, delay));
      return retryWithBackoff(fn, retries - 1, delay * factor, factor);
    }
    throw error;
  }
}

// Endpoint 1: Convert parsed Word Document text into professional PowerPoint presentation slides JSON
app.post("/api/gemini/generate-presentation", async (req, res) => {
  try {
    if (!ai) {
      return res.status(500).json({
        error: "GEMINI_API_KEY is missing on the server. Please add your key in Settings > Secrets.",
      });
    }

    const { docText, brandColor, themeId, docName } = req.body;

    if (!docText || docText.trim().length === 0) {
      return res.status(400).json({ error: "Document text content is empty or invalid." });
    }

    const prompt = `
      You are a world-class Elite Creative Director and Senior Presentation Art Director.
      Your task is to convert the following parsed document into a highly-designed, extremely polished presentation deck.
      
      Document Title/Context/Source: ${docName || "Uploaded Business Document"}
      Desired Brand Color Hex Accent: ${brandColor || "#6366F1"}
      Proposed Layout Aesthetic: ${themeId || "cosmic-slate"}

      CRITICAL OBJECTIVE: Break completely away from the "generic AI look". Every single slide must feel individually directed.
      Use visual breaks, editorial pacing, custom content badges, and high-impact color shifts.

      DESIGN & PLOT ARCHITECTURE SYSTEM:
      1. PRESENTATION LENGTH:
         Adapt the length dynamically to give a thorough, comprehensive overview.
         - DO NOT limit or truncate how many pages/slides are generated. The total page count must be determined organically solely by the content in the uploaded document. If there is a lot of content, feel free to generate 15, 20, or even more slides to fully cover all details without any arbitrary limit or ceiling.
         - Maximize outline coverage while keeping each individual slide layout uncluttered.

      2. THEME-CONSTRUCTED SLIDE-SPECIFIC COSMETIC VARIATION:
         To make the deck look bespoke, you must define slide-level color configurations and content badges.
         Introduce 'visual breaks' by occasionally inverting or emphasizing slides!
         
         Here is how to design slide-specific colors based on the requested Layout Aesthetic "${themeId || "cosmic-slate"}":
         
         - Theme "cosmic-slate" (Darks & Space Neons):
           * Normal Slides: bgColor: "#0B0F19", textColor: "#F1F5F9", primaryColor: "${brandColor || "#6366F1"}", accentColor: "#EC4899", cardBgColor: "#161E2E", borderColor: "#1F2937"
           * Visual Break (e.g., Title Slide, Quote Slide, key stat, or Conclusion Slide):
             Invert to solid brand/neon color blocks! E.g.: bgColor: "${brandColor || "#6366F1"}", textColor: "#FFFFFF", primaryColor: "#FFFFFF", accentColor: "#EC4899", cardBgColor: "#0000002A", borderColor: "#FFFFFF33" OR use high-contrast deep purple-black ("#110E1C").

         - Theme "minimal-swiss" (High Structured Stark Minimalism):
           * Normal Slides: bgColor: "#FFFFFF", textColor: "#0F172A", primaryColor: "#000000", accentColor: "${brandColor || "#EF4444"}", cardBgColor: "#F8FAFC", borderColor: "#E2E8F0"
           * Visual Break slides:
             Invert to striking pitch black or intense corporate blue/red! E.g.: bgColor: "#090D16", textColor: "#FFFFFF", primaryColor: "#FFFFFF", accentColor: "${brandColor || "#EF4444"}", cardBgColor: "#1E293B", borderColor: "#334155"

         - Theme "warm-editorial" (Graceful Premium Bookish Serifs):
           * Normal Slides: bgColor: "#FCF9F2", textColor: "#1C1917", primaryColor: "#78350F", accentColor: "${brandColor || "#C2410C"}", cardBgColor: "#F5EFEB", borderColor: "#E7E0D4"
           * Visual Break slides (Warm clay color block partition):
             Invert to solid warm terracotta, deep amber, or dark roast forest clay! E.g.: bgColor: "${brandColor || "#78350F"}", textColor: "#FCF9F2", primaryColor: "#FCF9F2", accentColor: "#FCD34D", cardBgColor: "#0000001E", borderColor: "#FFFFFF1E"

         - Theme "brutalist-mono" (Stark Technical Code Grids):
           * Normal Slides: bgColor: "#F3F4F6", textColor: "#000000", primaryColor: "#000000", accentColor: "${brandColor || "#F59E0B"}", cardBgColor: "#FFFFFF", borderColor: "#000000"
           * Visual Break slides (Hazard high intensity break):
             Invert to heavy solid black background or pure vibrant warning hazards! E.g.: bgColor: "#000000", textColor: "#FFFFFF", primaryColor: "${brandColor || "#F59E0B"}", accentColor: "#10B981", cardBgColor: "#111827", borderColor: "#FFFFFF"

         - Theme "creative-avantgarde" (Playful Organic Pastel Gradients):
           * Normal Slides: bgColor: "#FAF5FF", textColor: "#1E1B4B", primaryColor: "${brandColor || "#EC4899"}", accentColor: "#3B82F6", cardBgColor: "#FFFFFF", borderColor: "#F3E8FF"
           * Visual Break slides:
             Delightful pastel peach or rich mauve backdrops! E.g.: bgColor: "#FEF3C7", textColor: "#451A03", primaryColor: "#92400E", accentColor: "#3B82F6", cardBgColor: "#FFFBEB", borderColor: "#FDE68A"

      3. BESPOKE METADATA BADGES:
         Every slide must receive a custom "badge" representing its visual purpose or strategic emphasis.
         - Do not repeat the same badge across more than 2 slides.
         - Examples of badges: "EXECUTIVE SUMMARY", "MARKET CONFORMANCE", "OPPORTUNITY", "KEY METRIC", "THE PROBLEM", "PROPOSED SOLUTION", "FINANCIAL DEVIATION", "NEXT STEPS", "CUSTOMER VOICE", "MILESTONES".

      4. CHOOSE EXPRESSIVE LAYOUTS:
         Tailor layouts specifically to fit the underlying slide copy:
         - "title-slide": Dedicated introduction.
         - "two-column": Ideal for side-by-side structures or dual-category bullets.
        - "comparison-table": CRITICAL layout choice for differences, versus comparisons (e.g. Legacy vs Slidesss), opposing viewpoints, pros/cons, or alternative features. Always select "comparison-table" when the content lists features, before vs after, or pros vs cons.
         - "headline-bullet": Elegant large subtitle statement with pristine items.
         - "quote-slide": Perfect for a strong single takeaway quotation or bold highlighted realization.
         - "image-left" / "image-right": Bold asymmetrical splits highlighting Unsplash visuals.
         - "stats-bento": 4 metric boxes. Perfect ONLY if the bullets express clean growth percentages, finances, or dynamic numbers (the bullets should contain actual numbers!).

      5. RESPECT TEXT FORMAT OF SOURCE DOCUMENT (CRITICAL):
         - Do NOT reduce text into list items or bullet points unless it is explicitly indicated as a bullet point, numbered list, or dash item in the source document.
         - If the source document contains a paragraph, explanation, or standard prose, keep it exactly as a regular paragraph/sentence block of text.
         - Only prepend items in your 'content' array with a bullet symbol style (like '• ', '- ' or numbered indicators) if they were bulleted or numbered list items in the uploaded document.

      Here is the source document content to rewrite and design:
      ---------
      ${docText}
      ---------
    `;

    const response = await retryWithBackoff(() =>
      ai.models.generateContent({
        model: "gemini-3.5-flash",
        contents: prompt,
        config: {
          systemInstruction: "You are an elite, senior visual art director and presentation systems scientist. Rearrange any given business texts into a multi-slide premium JSON form. Always select distinct badges, dynamic visual layout flow, and tailored aesthetic HEX overrides for each slide to break standard generic AI styling.",
          responseMimeType: "application/json",
          temperature: 0.35,
          responseSchema: {
            type: Type.OBJECT,
            properties: {
              title: {
                type: Type.STRING,
                description: "The primary title of the presentation.",
              },
              themeId: {
                type: Type.STRING,
                description: "The recommended theme ID matching cosmic-slate, minimal-swiss, warm-editorial, brutalist-mono, or creative-avantgarde.",
              },
              slides: {
                type: Type.ARRAY,
                description: "List of carefully structured presentation slides covering all content of the document. Let the quantity of slides be determined organically by the input document text complexity without any arbitrary maximum limit.",
                items: {
                  type: Type.OBJECT,
                  properties: {
                    title: {
                      type: Type.STRING,
                      description: "High-impact short slide heading.",
                    },
                    content: {
                      type: Type.ARRAY,
                      items: { type: Type.STRING },
                      description: "An array of text elements. If the document has bullet points, keep them as list items (preserving prefixes like '• ' or leading bullet dashes). If the document has standard paragraphs/sentences instead of bullets, each paragraph or coherent block of standard text must be returned as a plain text string here without converting or reducing it into bullets.",
                    },
                    layout: {
                      type: Type.STRING,
                      description: "The recommended visual layout: title-slide, two-column, headline-bullet, quote-slide, image-left, image-right, stats-bento, comparison-table.",
                    },
                    notes: {
                      type: Type.STRING,
                      description: "One single 1-sentence coaching note (under 12 words).",
                    },
                    imageSearchQuery: {
                      type: Type.STRING,
                      description: "An exact 2-3 word keyword search query for Unsplash (e.g. 'analytics growth').",
                    },
                    badge: {
                      type: Type.STRING,
                      description: "Short premium uppercase 1-2 word metadata label, e.g. 'FORECAST', 'STRATEGY', 'MARKET TREND'. Can be up to 15 characters.",
                    },
                    bgColor: {
                      type: Type.STRING,
                      description: "Strict 7-character HEX code override for background (e.g. #0B0F19). Follow theme-specific inverting directions rules to create high-impact premium rhythm.",
                    },
                    textColor: {
                      type: Type.STRING,
                      description: "Strict 7-character HEX code override for main body and general text.",
                    },
                    accentColor: {
                      type: Type.STRING,
                      description: "Strict 7-character HEX code override for aesthetic badges, accents, and visual highlights.",
                    },
                    primaryColor: {
                      type: Type.STRING,
                      description: "Strict 7-character HEX code override for slide headings.",
                    },
                    cardBgColor: {
                      type: Type.STRING,
                      description: "Strict 7-character HEX code override for card or bento layouts.",
                    },
                    borderColor: {
                      type: Type.STRING,
                      description: "Strict 7-character HEX code override for borders.",
                    },
                  },
                  required: ["title", "content", "layout", "imageSearchQuery", "badge", "bgColor", "textColor", "accentColor", "primaryColor", "cardBgColor", "borderColor"],
                },
              },
            },
            required: ["title", "themeId", "slides"],
          },
        },
      })
    );

    const textResult = response.text;
    if (!textResult) {
      throw new Error("Empty response received from presentation generation AI model.");
    }

    const payload = JSON.parse(textResult.trim());
    return res.json(payload);
  } catch (error: any) {
    console.error("Presentation Generation Error:", error);
    let errorString = "";
    try {
      errorString = [
        error?.message,
        error?.toString(),
        JSON.stringify(error)
      ].filter(Boolean).join(" ").toLowerCase();
    } catch (_) {
      errorString = (error?.message || "").toString().toLowerCase();
    }

    const isRateLimitOrUnavailable =
      errorString.includes("503") ||
      errorString.includes("high demand") ||
      errorString.includes("unavailable") ||
      errorString.includes("429") ||
      errorString.includes("quota") ||
      errorString.includes("overburdened") ||
      error?.status === 503 ||
      error?.status === "UNAVAILABLE" ||
      error?.status === 429 ||
      error?.code === 503 ||
      error?.statusCode === 503;

    if (isRateLimitOrUnavailable) {
      return res.status(400).json({
        isTransient: true,
        error: "The Gemini AI service is currently experiencing high demand. Click 'Convert' to trigger an auto-retry, or click the friendly 'Preview Offline Demo' button to load a gorgeous pre-designed slide deck instantly!"
      });
    }
    return res.status(400).json({ error: error?.message || "Internal generation failure." });
  }
});

// Endpoint 2: Provide Pro-Design / AI layout suggestions for a single slide to redesign it dynamically
app.post("/api/gemini/suggest-design", async (req, res) => {
  try {
    if (!ai) {
      return res.status(500).json({
        error: "GEMINI_API_KEY is missing on the server. Please check your system secrets configuration.",
      });
    }

    const { currentSlide, brandColor, themeId } = req.body;

    if (!currentSlide) {
      return res.status(400).json({ error: "Slide context missing." });
    }

    const prompt = `
      You are an elite slide visual layout strategist and presentation art director.
      Analyze the current slide's context and contents and provide a highly-designed alternative recommendation.
      
      Current Slide:
      - Title: "${currentSlide.title}"
      - Current Layout: "${currentSlide.layout}"
      - Content Lines: ${JSON.stringify(currentSlide.content)}
      - Theme Style: "${themeId || "cosmic-slate"}"
      - Brand Color: "${brandColor || "#6366F1"}"

      Your output must:
      1. Propose an optimized layout that better represents this specific data.
         Example: If the content is filled with metrics or stats, change layout to "stats-bento". If it's a quote, suggest "quote-slide". If it compares features, pros/cons, or legacy vs future, choose "comparison-table".
      2. Rewrite the title and content bullets to be even more professional, punchy, and design-forward.
      3. Suggest a layout type ('title-slide' | 'two-column' | 'headline-bullet' | 'quote-slide' | 'image-left' | 'image-right' | 'stats-bento' | 'comparison-table').
      4. Assign a custom, premium metadata "badge" appropriate for the content (e.g. "HIGHLIGHT", "METRIC", "SURVEY INDICATOR").
      5. Provide an elegant set of slide-specific background, text, accent, heading, card, and border HEX colors that perfectly fit the requested Theme Style "${themeId || "cosmic-slate"}" and create beautiful custom-brand contrast. If inverting the slide, make the contrast gorgeous.
    `;

    const response = await retryWithBackoff(() =>
      ai.models.generateContent({
        model: "gemini-3.5-flash",
        contents: prompt,
        config: {
          systemInstruction: "You are a professional slide designer. Output an enhanced slide representation exactly adhering to the JSON schema specified. Keep all text extremely concise to guarantee ultra-low latency.",
          responseMimeType: "application/json",
          temperature: 0.3,
          responseSchema: {
            type: Type.OBJECT,
            properties: {
              title: { type: Type.STRING, description: "Enhanced punchy presentation slide title." },
              content: { type: Type.ARRAY, items: { type: Type.STRING }, description: "Polished bullet points, revised concisely." },
              layout: { type: Type.STRING, description: "Optimized slide layout choice." },
              imageSearchQuery: { type: Type.STRING, description: "Refined 2-3 word search query." },
              notes: { type: Type.STRING, description: "Enhanced single-sentence delivery speaker coaching tips." },
              badge: { type: Type.STRING, description: "Custom professional slide-specific upper-case badge." },
              bgColor: { type: Type.STRING, description: "7-character HEX override for slide background." },
              textColor: { type: Type.STRING, description: "7-character HEX override for general body text." },
              accentColor: { type: Type.STRING, description: "7-character HEX override for badges and highlights." },
              primaryColor: { type: Type.STRING, description: "7-character HEX override for heading text." },
              cardBgColor: { type: Type.STRING, description: "7-character HEX override for bento cards." },
              borderColor: { type: Type.STRING, description: "7-character HEX override for bento and card borders." }
            },
            required: ["title", "content", "layout", "imageSearchQuery", "notes", "badge", "bgColor", "textColor", "accentColor", "primaryColor", "cardBgColor", "borderColor"]
          }
        }
      })
    );

    const result = response.text;
    if (!result) {
      throw new Error("No suggestion output received from the model.");
    }

    return res.json(JSON.parse(result.trim()));
  } catch (err: any) {
    console.error("AI Slide suggestion error:", err);
    let errorString = "";
    try {
      errorString = [
        err?.message,
        err?.toString(),
        JSON.stringify(err)
      ].filter(Boolean).join(" ").toLowerCase();
    } catch (_) {
      errorString = (err?.message || "").toString().toLowerCase();
    }

    const isRateLimitOrUnavailable =
      errorString.includes("503") ||
      errorString.includes("high demand") ||
      errorString.includes("unavailable") ||
      errorString.includes("429") ||
      errorString.includes("quota") ||
      errorString.includes("overburdened") ||
      err?.status === 503 ||
      err?.status === "UNAVAILABLE" ||
      err?.status === 429 ||
      err?.code === 503 ||
      err?.statusCode === 503;

    if (isRateLimitOrUnavailable) {
      return res.status(400).json({
        isTransient: true,
        error: "The Gemini AI service is currently experiencing temporary high demand blocks. Please retry this action in a few moments."
      });
    }
    return res.status(400).json({ error: err?.message || "AI suggestion failed." });
  }
});

// Endpoint 3: Unrestricted Live Graphics Database search 
app.get("/api/images/search", async (req, res) => {
  try {
    const query = (req.query.q as string || "").trim();
    if (!query) {
      return res.status(400).json({ error: "Missing search keyword query." });
    }
    
    // Fetch live photos using public secure endpoints
    const response = await fetch(`https://unsplash.com/napi/search/photos?query=${encodeURIComponent(query)}&per_page=24`, {
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
      }
    });
    
    if (!response.ok) {
      throw new Error(`Unsplash search query failed with status: ${response.status}`);
    }
    
    const data: any = await response.json();
    const results = (data.results || []).map((img: any) => ({
      id: img.id,
      url: img.urls?.regular || img.urls?.full || img.urls?.small,
      thumbnail: img.urls?.small || img.urls?.thumb,
      title: img.description || img.alt_description || "Bespoke digital photo artwork",
      author: img.user?.name || "Unsplash Creative Community"
    }));
    
    return res.json({ results });
  } catch (error: any) {
    console.error("Unsplash Search Bridge Error:", error);
    // Dynamic ultra-reliable fallback to Picsum Photos with deterministic seeds for seamless preview rendering
    const fallbacks = Array.from({ length: 12 }).map((_, index) => {
      const seedVal = Math.abs((query.split("").reduce((acc, char) => acc + char.charCodeAt(0), 0) + index * 19) % 300) + 1;
      return {
        id: `picsumfallback-${seedVal}`,
        url: `https://picsum.photos/id/${seedVal}/1200/800`,
        thumbnail: `https://picsum.photos/id/${seedVal}/300/200`,
        title: `Professional representation metadata visual for ${query}`,
        author: "Open Source Community (Picsum Photos)"
      };
    });
    return res.json({ results: fallbacks });
  }
});

export default app;
