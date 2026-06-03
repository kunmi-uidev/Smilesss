/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import express from "express";
import path from "path";
import dotenv from "dotenv";
import { GoogleGenAI, Type } from "@google/genai";

dotenv.config();

const app = express();
const PORT = 3000;

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
  retries = 3,
  delay = 1500,
  factor = 2
): Promise<T> {
  try {
    return await fn();
  } catch (error: any) {
    const errorMsg = (error?.message || error?.status || "").toString();
    const isRateLimitOrUnavailable =
      errorMsg.includes("503") ||
      errorMsg.includes("high demand") ||
      errorMsg.includes("UNAVAILABLE") ||
      errorMsg.includes("429") ||
      errorMsg.includes("quota") ||
      (error?.status === 503 || error?.status === 429);

    if (retries > 0 && isRateLimitOrUnavailable) {
      console.warn(`[GEMINI STATUS] Service bottleneck (503/429) detected. Retrying in ${delay}ms... Remaining: ${retries}`);
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
      You are a professional Creative Director and Presentation Designer at a top-tier design consultancy.
      Your task is to convert the following parsed document content into a highly designed, cohesive, and professional presentation slide deck.

      Document Title/Context/Source: ${docName || "Uploaded Business Document"}
      Desired Brand Color Hex: ${brandColor || "#6366F1"}
      Proposed Layout Aesthetic: ${themeId || "cosmic-slate"}

      Guidelines to Design like a Pro:
      1. Structure the deck with a logical flow: Title slide, context setting/problem, core solution/pillars, detailed analysis (using various layout types), and a strong conclusion.
      2. Choose expressive, slide-specific layouts based on the content of each slide:
         - "title-slide": Only for the initial introducing slide. Includes a main title and subtitle/hook.
         - "two-column": Ideal for comparing two ideas, features, props vs cons, before/after, or splitting content.
         - "headline-bullet": Elegant large headline with 2-4 clean scannable bullet points underneath.
         - "quote-slide": Perfect for a strong single takeaway sentence, business metric highlight, or a client quote.
         - "image-left": Bold visual on left side, 2-3 structured points on the right side.
         - "image-right": Structured points on the left, high-res visual on the right.
         - "minimal-split": A minimalistic divided column slide for modern professional impact.
         - "stats-bento": Perfect for slides containing numeric growth metrics, statistics, launch indicators, or dashboard-style summaries (the content bullets should contain metric names and high-contrast figures!).
      3. For bullets, rewrite copy into concise, professional, scannable statements (under 12 words per line). Never output wordy paragraphs in bullet text.
      4. For each slide, write strategic 'notes' (speaker details instructions) to guide a presenter's pacing.
      5. Provide an relevant imageSearchQuery for each slide (e.g. "modern office strategy", "abstract coding waves", "growth dashboard charts") so we can pull beautiful Unsplash assets.
      
      Here is the source document content to rewrite:
      ---------
      ${docText}
      ---------
    `;

    const response = await retryWithBackoff(() =>
      ai.models.generateContent({
        model: "gemini-3.5-flash",
        contents: prompt,
        config: {
          systemInstruction: "You are an elite presentation designer. Reorganize copy into pristine visual slide content structure. Ensure response complies exactly with the requested JSON schema. Avoid adding empty content. Do not use markdown backticks inside JSON strings.",
          responseMimeType: "application/json",
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
                description: "List of carefully structured presentation slides (6-12 slides).",
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
                      description: "Concise scannable bullet items, main takeaways, or structural columns.",
                    },
                    layout: {
                      type: Type.STRING,
                      description: "The recommended visual layout: title-slide, two-column, headline-bullet, quote-slide, image-left, image-right, minimal-split, stats-bento.",
                    },
                    notes: {
                      type: Type.STRING,
                      description: "Strategic speaker notes for verbal delivery during the presentation.",
                    },
                    imageSearchQuery: {
                      type: Type.STRING,
                      description: "An exact 2-3 word keyword search query for Unsplash to load a background or contextual photo (e.g. 'analytics growth', 'finance team').",
                    },
                  },
                  required: ["title", "content", "layout", "imageSearchQuery"],
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
    const errorMsg = (error?.message || error?.status || "").toString();
    if (
      errorMsg.includes("503") ||
      errorMsg.includes("high demand") ||
      errorMsg.includes("UNAVAILABLE") ||
      errorMsg.includes("429") ||
      errorMsg.includes("quota") ||
      error?.status === 503 ||
      error?.status === 429
    ) {
      return res.status(503).json({
        isTransient: true,
        error: "The Gemini AI service is currently experiencing high demand. Click 'Convert' to trigger an auto-retry, or click the friendly 'Preview Offline Demo' button to load a gorgeous pre-designed slide deck instantly!"
      });
    }
    return res.status(500).json({ error: error?.message || "Internal generation failure." });
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
      You are an elite slide visual layout strategist.
      Analyze the current slide's context and contents and provide a highly-designed alternative recommendation.
      
      Current Slide:
      - Title: "${currentSlide.title}"
      - Current Layout: "${currentSlide.layout}"
      - Content Lines: ${JSON.stringify(currentSlide.content)}
      - Theme Style: "${themeId || "cosmic-slate"}"
      - Brand Color: "${brandColor || "#6366F1"}"

      Your output must:
      1. Propose an optimized layout that better represents this specific data.
         Example: If the content is filled with metrics or stats, change layout to "stats-bento". If it's a quote, suggest "quote-slide".
      2. Rewrite the title and content bullets to be even more professional, punchy, and design-forward.
      3. Suggest a layout type ('title-slide' | 'two-column' | 'headline-bullet' | 'quote-slide' | 'image-left' | 'image-right' | 'minimal-split' | 'stats-bento').
      4. Make sure formatting is clean.
    `;

    const response = await retryWithBackoff(() =>
      ai.models.generateContent({
        model: "gemini-3.5-flash",
        contents: prompt,
        config: {
          systemInstruction: "You are a professional slide designer. Output an enhanced slide representation exactly adhering to the JSON schema specified.",
          responseMimeType: "application/json",
          responseSchema: {
            type: Type.OBJECT,
            properties: {
              title: { type: Type.STRING, description: "Enhanced punchy presentation slide title." },
              content: { type: Type.ARRAY, items: { type: Type.STRING }, description: "Polished bullet points, revised concisely." },
              layout: { type: Type.STRING, description: "Optimized slide layout choice." },
              imageSearchQuery: { type: Type.STRING, description: "Refined 2-3 word search query." },
              notes: { type: Type.STRING, description: "Enhanced delivery speaker coaching tips." }
            },
            required: ["title", "content", "layout", "imageSearchQuery", "notes"]
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
    const errorMsg = (err?.message || err?.status || "").toString();
    if (
      errorMsg.includes("503") ||
      errorMsg.includes("high demand") ||
      errorMsg.includes("UNAVAILABLE") ||
      errorMsg.includes("429") ||
      errorMsg.includes("quota") ||
      err?.status === 503 ||
      err?.status === 429
    ) {
      return res.status(503).json({
        isTransient: true,
        error: "The Gemini AI service is currently experiencing temporary high demand blocks. Please retry this action in a few moments."
      });
    }
    return res.status(500).json({ error: err?.message || "AI suggestion failed." });
  }
});

// Serve frontend assets in development and production
async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const { createServer: createViteServer } = await import("vite");
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
    console.log("Vite Development Middleware bound successfully.");
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
    console.log("Production static distribution bound successfully.");
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Presentation master server online on port ${PORT}`);
  });
}

// Only start the listening server if we are running in a standalone Node process
// (On Vercel, Vercel loads the exported app and handles routing natively, process.env.VERCEL is true)
if (!process.env.VERCEL) {
  startServer();
}

export default app;
