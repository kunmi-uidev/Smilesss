/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef } from "react";
import mammoth from "mammoth";
import {
  Upload, Sparkles, Wand2, Download, RefreshCw, Plus, Trash2, ArrowUp, ArrowDown,
  Image as ImageIcon, Layout, Palette, Search, FileText, Check, HelpCircle, ChevronRight,
  Monitor, Play, Printer, X, Sliders, ExternalLink, WifiOff, AlertCircle
} from "lucide-react";
import { ThemeId, ColorPalette, SlideLayout, Slide, Presentation } from "./types";
import { THEME_PRESETS, CURATED_UNSPLASH_IMAGES, getKeywordImage, normalizeHex, hexToRgbA } from "./lib/themePresets";
import { exportToPPTX } from "./lib/pptxExport";
import { testColorAccessibility, ensureAccessiblePalette, ensureAccessibleTextColor, getContrastRatio, getLuminance } from "./lib/accessibility";

const SAMPLE_DOCX_TEXT = `# EXCELSIOR GLOBAL SYSTEMS Q3 MARKETING STRATEGY
## Brand Alignment & Expansion Outline
### Confined & Internal Strategy Document

"Authentic visual intelligence is no longer a corporate luxury—it is the single highest leverage vector for corporate clarity and authority." 
Our primary focus is removing creative barriers for remote partners.

## RECENT PERFORMANCE STATS
Our growth is compounding across four core KPIs:
- 142% Compound expansion rate year-over-year
- 4.9 Star customer satisfaction metric
- 99.98% Core cloud server availability
- $12M Reinvested in creative design automation

## CORE EXPANSION PILLARS
We plan on directing investment in three key milestones:
1. DESIGN DEMOCRATIZATION: Providing automated interfaces for business non-designers to draft stunning marketing materials.
2. BRAND ENGINE INTEGRATION: Forcing deep mathematical palette consistency across global regions with one-click theme styling.
3. VISUAL ASSETS DEPTH: Partnering with premium royalty-free graphics registries to render high-contrast visual cues on slides.

## COMPETITIVE POSTURE SUMMARY
Comparison reveals severe gaps in old-world tech:
- Excelsior Systems: Expressive bento grids, real-time color customizers, instant PPTX delivery.
- Legacy Software: Static drab tables, rigid layouts, expensive login loops.

## EXECUTION TIMELINE
- PHASE I (JUNE): Finalize full-stack Gemini API JSON validation middleware.
- PHASE II (JULY): Launch public PPTX client-side rendering pipeline.
- PHASE III (AUGUST): Open Unsplash visual search databases to users.
- PHASE IV (SEPTEMBER): General market announcement.`;

// Diagnostic Helper to structure raw error inputs into clean, non-technical category templates
interface ParsedError {
  type: "network" | "technical" | "input";
  title: string;
  explanation: string;
  cause: string;
  cta: string;
}

function getDetailedError(message: string | null): ParsedError {
  if (!message) {
    return {
      type: "technical",
      title: "Technical System Error",
      explanation: "An unexpected internal error interrupted the process of designing your slide deck.",
      cause: "Internal compiler timeout.",
      cta: "Please try again later. If this persists, feel free to contact system support."
    };
  }

  const lower = message.toLowerCase();

  // 1. Check for empty template or prompt issue (input Validation notice)
  if (
    lower.includes("empty") || 
    lower.includes("please select a word") || 
    lower.includes("outline") ||
    lower.includes("no document text")
  ) {
    return {
      type: "input",
      title: "No Presentation Outline Found",
      explanation: "We couldn't find any presentation headlines or text content in the Manual Editor or the uploaded file.",
      cause: "The source text is empty or missing details.",
      cta: "Type out some high-level key points in the editor box, or load our predesigned company business template instantly."
    };
  }

  // 2. Check for network connection, server load, and service rate/quota blocks (Network Error)
  const isNetwork =
    lower.includes("fetch") ||
    lower.includes("network") ||
    lower.includes("connect") ||
    lower.includes("timeout") ||
    lower.includes("503") ||
    lower.includes("unavailable") ||
    lower.includes("429") ||
    lower.includes("quota") ||
    lower.includes("rate limit") ||
    lower.includes("overburdened") ||
    lower.includes("unreachable") ||
    lower.includes("key is missing") ||
    lower.includes("html template") ||
    lower.includes("doctype") ||
    lower.includes("unauthorized") ||
    lower.includes("offline") ||
    lower.includes("bottleneck");

  if (isNetwork) {
    return {
      type: "network",
      title: "Network Connection Issue",
      explanation: "Our system is experiencing a temporary network bottleneck or transient traffic overload. No worries, this is usually temporary!",
      cause: "Temporary AI model/network bottleneck.",
      cta: "Please verify your active internet connection or click 'Try Again' in a few seconds to retry instantly."
    };
  }

  // 3. Otherwise treat as a stable major internal Technical Error (Technical issue etc.)
  return {
    type: "technical",
    title: "Technical Error Details",
    explanation: "An unexpected technical error occurred while rendering or formulating your presentation slides.",
    cause: "Internal layout compiler exception.",
    cta: "Click below to contact our administrator or bypass with an offline demo presentation layout."
  };
}

function TypewriterText() {
  const text = "Instantly rewrite wordy files into designer slides.";
  const [displayedText, setDisplayedText] = useState("");
  const [index, setIndex] = useState(0);

  useEffect(() => {
    if (index < text.length) {
      const timeout = setTimeout(() => {
        setDisplayedText((prev) => prev + text[index]);
        setIndex((prev) => prev + 1);
      }, 45); // 45ms per character is highly legible and premium
      return () => clearTimeout(timeout);
    }
  }, [index, text]);

  const baseText = "Instantly rewrite wordy files into ";

  if (displayedText.length <= baseText.length) {
    return (
      <span className="relative inline">
        <span>{displayedText}</span>
        {index < text.length && (
          <span className="inline-block w-[3px] h-[0.85em] bg-indigo-400 ml-1 translate-y-[2px] animate-pulse" />
        )}
      </span>
    );
  }

  const part1 = baseText;
  const part2 = displayedText.substring(baseText.length);

  return (
    <span className="relative inline">
      <span>{part1}</span>
      <span className="text-transparent bg-gradient-to-r from-indigo-400 via-pink-400 to-amber-400 bg-clip-text">
        {part2}
      </span>
      {index < text.length && (
        <span className="inline-block w-[3px] h-[0.85em] bg-indigo-400 ml-1 translate-y-[2px] animate-pulse" />
      )}
    </span>
  );
}

function getContrastColor(bgColorHex: string): { text: string; mtext: string; card: string; border: string; contrastRatio: number; rating: string } {
  try {
    const hex = normalizeHex(bgColorHex, "#0B0F19").replace("#", "");
    const r = parseInt(hex.substring(0, 2), 16) || 0;
    const g = parseInt(hex.substring(2, 4), 16) || 0;
    const b = parseInt(hex.substring(4, 6), 16) || 0;
    
    const a = [r, g, b].map((v) => {
      v /= 255;
      return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4);
    });
    const luminance = 0.2126 * a[0] + 0.7152 * a[1] + 0.0722 * a[2];

    const ratioWithWhite = (1.0 + 0.05) / (luminance + 0.05);
    const ratioWithBlack = (luminance + 0.05) / 0.05;

    const useDark = luminance > 0.45;
    const contrastRatio = useDark ? ratioWithBlack : ratioWithWhite;
    
    const rating = contrastRatio >= 7.0 ? "AAA (Excellent)" : contrastRatio >= 4.5 ? "AA (Pass)" : "Low Contrast (Warning)";

    return {
      text: useDark ? "#0F172A" : "#F8FAFC",
      mtext: useDark ? "#475569" : "#94A3B8",
      card: useDark ? "rgba(15, 23, 42, 0.05)" : "rgba(255, 255, 255, 0.07)",
      border: useDark ? "rgba(15, 23, 42, 0.12)" : "rgba(255, 255, 255, 0.15)",
      contrastRatio,
      rating
    };
  } catch (err) {
    return {
      text: "#F8FAFC",
      mtext: "#94A3B8",
      card: "rgba(255, 255, 255, 0.07)",
      border: "rgba(255, 255, 255, 0.15)",
      contrastRatio: 21,
      rating: "AAA"
    };
  }
}

export default function App() {
  // Page state: 'upload' | 'editor'
  const [activeStep, setActiveStep] = useState<"upload" | "editor">("upload");

  // Input states
  const [brandColor, setBrandColor] = useState<string>("#4F46E5"); // Indigo default
  const [selectedThemeId, setSelectedThemeId] = useState<ThemeId>(ThemeId.COSMIC_SLATE);
  const [rawText, setRawText] = useState<string>("");
  const [docName, setDocName] = useState<string>("");
  const [isParsingDoc, setIsParsingDoc] = useState<boolean>(false);
  const [isGenerating, setIsGenerating] = useState<boolean>(false);
  const [generationProgress, setGenerationProgress] = useState<number>(0);
  const [statusMessage, setStatusMessage] = useState<string>("");

  // Slide Deck Page Count Configuration Modal state (Let user specify slide count with skip option)
  const [isDeckConfigOpen, setIsDeckConfigOpen] = useState<boolean>(false);
  const [targetSlideCount, setTargetSlideCount] = useState<number>(7);
  const [selectedSlideOption, setSelectedSlideOption] = useState<"auto" | 5 | 7 | 10 | 14 | "custom">("auto");

  // Presentation State
  const [presentation, setPresentation] = useState<Presentation | null>(null);
  const [activeSlideIndex, setActiveSlideIndex] = useState<number>(0);

  // Unsplash search pop-up states
  const [isUnsplashOpen, setIsUnsplashOpen] = useState<boolean>(false);
  const [unsplashSearch, setUnsplashSearch] = useState<string>("");
  const [unsplashCategory, setUnsplashCategory] = useState<string>("Business & Leadership");
  const [unsplashResults, setUnsplashResults] = useState<any[]>([]);
  const [isSearchingUnsplash, setIsSearchingUnsplash] = useState<boolean>(false);

  // AI Suggestion states
  const [isSuggestingAI, setIsSuggestingAI] = useState<boolean>(false);
  const [aiSuggestion, setAiSuggestion] = useState<any | null>(null);
  const [showSuggestionModal, setShowSuggestionModal] = useState<boolean>(false);

  // Error handling
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

  // Brand color alignment export modal states
  const [isExportSettingsOpen, setIsExportSettingsOpen] = useState<boolean>(false);
  const [pendingExportType, setPendingExportType] = useState<"pdf" | "pptx" | null>(null);

  // Real-time palette for preview
  const activePresetTheme = THEME_PRESETS[selectedThemeId];
  const activePalette = activePresetTheme.defaultPalette(brandColor);

  // Interval for fancy loading messages
  const loadingMessages = [
    "Analyzing document hierarchy...",
    "Extracting core statistics & metrics...",
    "Testing color contrast for WCAG accessibility...",
    "Selecting semantic professional slide layouts...",
    "Aligning color spaces to your brand guidelines...",
    "Formulating concise high-impact summaries...",
    "Synthesizing Unsplash metadata queries...",
    "Constructing designer presentation system..."
  ];

  // Auto-generate dynamic colors palette dynamically on brandColor change with accessibility checking
  useEffect(() => {
    if (presentation) {
      const rawPalette = THEME_PRESETS[presentation.themeId].defaultPalette(brandColor);
      const { palette: accessiblePalette } = ensureAccessiblePalette(rawPalette);
      setPresentation(prev => prev ? {
        ...prev,
        brandColor,
        palette: accessiblePalette
      } : null);
    }
  }, [brandColor]);

  // Handle uploaded docx file
  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setDocName(file.name);
    setIsParsingDoc(true);
    setErrorMessage(null);

    try {
      if (!file.name.endsWith(".docx")) {
        throw new Error("Only Microsoft Word File format (.docx) is accepted for extraction.");
      }
      
      const arrayBuffer = await file.arrayBuffer();
      const result = await mammoth.extractRawText({ arrayBuffer });
      
      if (!result.value || result.value.trim().length === 0) {
        throw new Error("Could not find readable text layers in uploaded Word document. Try checking if it is an empty file.");
      }

      setRawText(result.value);
    } catch (err: any) {
      console.error(err);
      setErrorMessage(err.message || "Failed parsing the Word document. Please make sure it's a valid non-corrupted .docx file.");
    } finally {
      setIsParsingDoc(false);
    }
  };

  // Load sample content for onboarding
  const handleLoadSample = () => {
    setDocName("Sample_Excelsior_Strategy_Report.docx");
    setRawText(SAMPLE_DOCX_TEXT);
  };

  // Instantly pre-populate presentation studio bypassing remote AI bottlenecks (e.g. 503 limits)
  const handleLoadDemoPresentation = () => {
    setDocName("Sample_Excelsior_Strategy_Report.docx");
    setRawText(SAMPLE_DOCX_TEXT);
    
    // Construct premium presentation slides with accessibility validated and mandatory Thank You slide
    const sampleSlides: Slide[] = [
      {
        id: `slide-demo-1-${Date.now()}`,
        title: "EXCELSIOR GLOBAL SYSTEMS Q3 MARKETING STRATEGY",
        badge: "EXECUTIVE BRIEF",
        content: [
          "Cohesive corporate brand strategy designed for active partners",
          "Presented by Excelsior Strategy Group",
          "Brand Alignment & Expansion Outline"
        ],
        layout: "title-slide",
        notes: "Welcome the leadership team and declare the priority: visual corporate authority and partner expansion."
      },
      {
        id: `slide-demo-2-${Date.now()}`,
        title: "RECENT PERFORMANCE STATS",
        badge: "GROWTH METRICS",
        content: [
          "142% year-over-year compound expansion rate",
          "4.9 customer satisfaction score representing extreme fidelity",
          "99.98% core cloud server availability",
          "$12M reinvested in our creative automation engine"
        ],
        layout: "stats-bento",
        notes: "Highlight our compounding metrics. The 142% CAGR is our highest leverage indicator for investor calls."
      },
      {
        id: `slide-demo-3-${Date.now()}`,
        title: "CORE EXPANSION PILLARS",
        badge: "STRATEGY PILLARS",
        content: [
          "DESIGN DEMOCRATIZATION: Providing simple, elegant web-based layout systems for business non-designers.",
          "BRAND ENGINE ALIGNMENT: One-click theme styles to maintain complete brand consistency.",
          "VISUAL ASSETS DENSITY: Deep Unsplash image registry integration for high-contrast context cues."
        ],
        layout: "two-column",
        notes: "These three strategic pillars hold our product vision together. We prioritize democratization first."
      },
      {
        id: `slide-demo-4-${Date.now()}`,
        title: "COMPETITIVE POSTURE SUMMARY",
        badge: "MARKET ADVANTAGE",
        content: [
          "Visual Architecture: Expressive bento grids vs Static drab tables",
          "Palette Calibration: Mathematical real-time contrast checking vs Inconsistent arbitrary colors",
          "Turnaround Velocity: Instant 1920x1080 PPTX/PDF delivery vs Rigid expensive design loops"
        ],
        layout: "comparison-table",
        notes: "Position Excelsior as the modern, high-velocity design intelligence platform."
      },
      {
        id: `slide-demo-5-${Date.now()}`,
        title: "CORPORATE FOCUS DIRECTIVE",
        badge: "CORE PRINCIPLE",
        content: [
          "\"Authentic visual intelligence is no longer a corporate luxury—it is the single highest leverage vector for corporate clarity and authority.\""
        ],
        layout: "quote-slide",
        notes: "Read this quote with high conviction. Visual intelligence is our brand's authority signature."
      },
      {
        id: `slide-demo-6-${Date.now()}`,
        title: "EXECUTION TIMELINE",
        badge: "MILESTONES",
        content: [
          "PHASE I (JUNE): Finalize full-stack Gemini JSON validation schemas with exponential backoff handlers.",
          "PHASE II (JULY): Ship public pptxgenjs design pipeline & 1920x1080 visual theme editors.",
          "PHASE III (AUGUST): Release premium curated visual searching tools.",
          "PHASE IV (SEPTEMBER): Announce general public market launch."
        ],
        layout: "headline-bullet",
        notes: "Walk the board through our rapid execution deliverables. All items are on track for Q3 delivery."
      },
      {
        id: `slide-demo-7-${Date.now()}`,
        title: "Thank you for listening",
        badge: "THANK YOU & Q&A",
        content: [
          "We appreciate your time, partnership, and strategic focus.",
          "Floor is now open for Q&A, executive discussion, and roadmap feedback.",
          "Connect with our project leadership team: leadership@excelsiorsystems.io"
        ],
        layout: "headline-bullet",
        notes: "Thank the audience warmly for their time and invite open questions and discussion."
      }
    ];

    // Inject Unsplash images based on standard queries
    const slidesWithImages = sampleSlides.map(slide => ({
      ...slide,
      imageUrl: getKeywordImage(slide.title.toLowerCase() || "business strategy")
    }));

    const rawPalette = THEME_PRESETS[selectedThemeId].defaultPalette(brandColor);
    const { palette: accessiblePalette } = ensureAccessiblePalette(rawPalette);

    const demoPresentation: Presentation = {
      title: docName ? docName.replace(/\.[^/.]+$/, "").replace(/_/g, " ") : "Sample Excelsior Strategy Report",
      themeId: selectedThemeId,
      brandColor: brandColor,
      palette: accessiblePalette,
      slides: slidesWithImages
    };

    setPresentation(demoPresentation);
    setActiveSlideIndex(0);
    setErrorMessage(null);
    setActiveStep("editor");
  };

  // Open the slide count configuration modal or trigger generation directly
  const handleOpenDeckConfigOrGenerate = () => {
    if (!rawText.trim()) {
      setErrorMessage("Please select a Word document to upload or load our professional sample text outline.");
      return;
    }
    setIsDeckConfigOpen(true);
  };

  // Submit parsed text to backend to build full presentation slides using Gemini
  const handleGeneratePresentation = async (requestedCount?: number | null) => {
    if (!rawText.trim()) {
      setErrorMessage("Please select a Word document to upload or load our professional sample text outline.");
      return;
    }

    setIsDeckConfigOpen(false);
    setIsGenerating(true);
    setErrorMessage(null);
    setGenerationProgress(10);
    setStatusMessage(loadingMessages[0]);

    // Fast loading indicator increments
    const progressTimer = setInterval(() => {
      setGenerationProgress(prev => {
        if (prev >= 90) return prev;
        const jump = Math.floor(Math.random() * 10) + 5;
        const next = prev + jump;
        // Cycle status messages
        const msgIndex = Math.floor(next / 12) % loadingMessages.length;
        setStatusMessage(loadingMessages[msgIndex]);
        return next;
      });
    }, 1100);

    try {
      const res = await fetch("/api/gemini/generate-presentation", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          docText: rawText,
          brandColor: brandColor,
          themeId: selectedThemeId,
          docName: docName,
          targetSlideCount: requestedCount ?? (selectedSlideOption === "auto" ? undefined : targetSlideCount)
        })
      });

      let responseText = "";
      try {
        responseText = await res.text();
      } catch (e) {
        responseText = "";
      }

      const contentType = res.headers.get("Content-Type") || "";
      if (!contentType.includes("application/json")) {
        const cleanText = responseText.replace(/<[^>]*>/g, '').trim();
        throw new Error(`The server returned an HTML template instead of data. This typically indicates a network bottleneck or that your Gemini API key is missing or expired. Server message: ${cleanText.slice(0, 150) || "None"}`);
      }

      if (!res.ok) {
        let errorMsg = "Generation endpoint refused to formulate slide layouts.";
        try {
          const errPayload = JSON.parse(responseText);
          errorMsg = errPayload.error || errorMsg;
        } catch (e) {
          const cleanText = responseText.replace(/<[^>]*>/g, '').trim();
          errorMsg = `Server returned (${res.status}): ${cleanText.slice(0, 250) || "No response body"}`;
        }
        throw new Error(errorMsg);
      }

      let deck;
      try {
        deck = JSON.parse(responseText);
      } catch (e) {
        throw new Error(`Failed to parse presentation JSON response from server: ${responseText.slice(0, 150)}`);
      }
      
      const activeThemeId = (deck.themeId as ThemeId) || selectedThemeId;
      const rawPalette = THEME_PRESETS[activeThemeId].defaultPalette(brandColor);
      
      // Accessibility test before applying colors
      const { palette: accessiblePalette } = ensureAccessiblePalette(rawPalette);

      // Inject unique slide IDs and build standard initial presentation state with accessibility testing
      const initialSlides = (deck.slides || []).map((slide: any, i: number) => {
        // Test color contrast on slide level
        const slideBg = slide.bgColor || accessiblePalette.background;
        const slideText = ensureAccessibleTextColor(slideBg, slide.textColor || accessiblePalette.text);
        const slidePrimary = ensureAccessibleTextColor(slideBg, slide.primaryColor || accessiblePalette.primary, 3.0);

        return {
          ...slide,
          id: `slide-${Date.now()}-${i}-${Math.random().toString(36).substr(2, 5)}`,
          imageUrl: getKeywordImage(slide.imageSearchQuery || "business teamwork outline"),
          bgColor: slideBg,
          textColor: slideText,
          primaryColor: slidePrimary
        };
      });

      // Ensure mandatory closing slide exists
      const lastSlide = initialSlides[initialSlides.length - 1];
      const hasThankYou = lastSlide && (
        lastSlide.title.toLowerCase().includes("thank you") ||
        lastSlide.title.toLowerCase().includes("thanks") ||
        (lastSlide.badge && lastSlide.badge.toLowerCase().includes("thank"))
      );

      if (!hasThankYou) {
        initialSlides.push({
          id: `slide-thank-you-${Date.now()}`,
          title: "Thank you for listening",
          badge: "THANK YOU & Q&A",
          layout: "headline-bullet",
          content: [
            "Thank you for your time, focus, and participation.",
            "We welcome any questions, discussion topics, or immediate feedback.",
            "Connect with our project leadership team for execution next steps."
          ],
          notes: "Thank the audience for their time and open the floor for Q&A.",
          imageUrl: getKeywordImage("celebration teamwork"),
          bgColor: accessiblePalette.background,
          textColor: accessiblePalette.text,
          primaryColor: accessiblePalette.primary
        });
      }

      const loadedPresentation: Presentation = {
        title: deck.title || (docName ? docName.replace(/\.[^/.]+$/, "").replace(/_/g, " ") : "Untitled Presentation"),
        themeId: activeThemeId,
        brandColor: brandColor,
        palette: accessiblePalette,
        slides: initialSlides
      };

      setPresentation(loadedPresentation);
      setSelectedThemeId(activeThemeId);
      setActiveSlideIndex(0);
      setGenerationProgress(100);
      setStatusMessage("Pristine 1920x1080 slide deck successfully crafted!");

      // Small pause for visual feedback
      setTimeout(() => {
        setActiveStep("editor");
        setIsGenerating(false);
      }, 700);

    } catch (err: any) {
      console.error(err);
      setErrorMessage(err.message || "PowerPoint deck creation failed. Please verify your internet connection or verify your Gemini credentials.");
      setIsGenerating(false);
    } finally {
      clearInterval(progressTimer);
    }
  };

  // Apply another global layout theme with accessibility enforcement
  const handleSwitchTheme = (themeId: ThemeId) => {
    if (presentation) {
      const rawPalette = THEME_PRESETS[themeId].defaultPalette(brandColor);
      const { palette: accessiblePalette } = ensureAccessiblePalette(rawPalette);
      setPresentation({
        ...presentation,
        themeId,
        palette: accessiblePalette
      });
    }
    setSelectedThemeId(themeId);
  };

  // Auto-Fix All Slides to meet WCAG AA contrast standard
  const handleAutoFixAllSlidesContrast = () => {
    if (!presentation) return;
    const { palette: accessiblePalette } = ensureAccessiblePalette(presentation.palette);
    const updatedSlides = presentation.slides.map(slide => {
      const bg = slide.bgColor || accessiblePalette.background;
      const text = ensureAccessibleTextColor(bg, slide.textColor || accessiblePalette.text);
      const primary = ensureAccessibleTextColor(bg, slide.primaryColor || accessiblePalette.primary, 3.0);
      return {
        ...slide,
        bgColor: bg,
        textColor: text,
        primaryColor: primary
      };
    });

    setPresentation({
      ...presentation,
      palette: accessiblePalette,
      slides: updatedSlides
    });
  };

  // Slide navigation actions
  const activeSlide = presentation?.slides[activeSlideIndex] || null;
  const slideBg = activeSlide?.bgColor || presentation?.palette.background || "#000000";
  const slideText = activeSlide?.textColor || presentation?.palette.text || "#FFFFFF";
  const slideAccent = activeSlide?.accentColor || presentation?.palette.accent || "#3B82F6";
  const slidePrimary = activeSlide?.primaryColor || presentation?.palette.primary || "#3B82F6";
  const slideCardBg = activeSlide?.cardBgColor || presentation?.palette.cardBg || "#111827";
  const slideBorder = activeSlide?.borderColor || presentation?.palette.border || "#1F2937";

  const slideScale = activeSlide?.fontSize === "small" ? 0.82 : activeSlide?.fontSize === "large" ? 1.18 : 1.0;

  const getFontSizeStyle = (baseSizeRem: number, scale: number = slideScale) => {
    return `${baseSizeRem * scale}rem`;
  };

  const updateActiveSlide = (fields: Partial<Slide>) => {
    if (!presentation || !activeSlide) return;
    const updatedSlides = [...presentation.slides];
    updatedSlides[activeSlideIndex] = {
      ...activeSlide,
      ...fields
    };
    setPresentation({
      ...presentation,
      slides: updatedSlides
    });
  };

  const handleSlideBgChange = (newColor: string) => {
    const norm = normalizeHex(newColor, "#0B0F19");
    const acc = getContrastColor(norm);
    updateActiveSlide({
      bgColor: norm,
      textColor: acc.text,
      cardBgColor: acc.card,
      borderColor: acc.border
    });
  };

  // Inline bullet items controllers
  const handleEditBullet = (bulletIndex: number, text: string) => {
    if (!activeSlide) return;
    const bullets = [...activeSlide.content];
    bullets[bulletIndex] = text;
    updateActiveSlide({ content: bullets });
  };

  const handleAddBullet = () => {
    if (!activeSlide) return;
    updateActiveSlide({ content: [...activeSlide.content, "New strategy bullet point"] });
  };

  const handleDeleteBullet = (bulletIndex: number) => {
    if (!activeSlide) return;
    const bullets = activeSlide.content.filter((_, idx) => idx !== bulletIndex);
    updateActiveSlide({ content: bullets });
  };

  // Rearranging slides
  const handleMoveSlide = (index: number, direction: "up" | "down") => {
    if (!presentation) return;
    const slides = [...presentation.slides];
    const newIndex = direction === "up" ? index - 1 : index + 1;
    if (newIndex < 0 || newIndex >= slides.length) return;

    // Swap position indices
    const temp = slides[index];
    slides[index] = slides[newIndex];
    slides[newIndex] = temp;

    setPresentation({
      ...presentation,
      slides
    });
    setActiveSlideIndex(newIndex);
  };

  const handleAddBlankSlide = () => {
    if (!presentation) return;
    const newSlide: Slide = {
      id: `slide-custom-${Date.now()}`,
      title: "Strategy Objective Keynote",
      content: ["Focus benchmark indicator target", "Cohesive brand deployment plan"],
      layout: "headline-bullet",
      imageUrl: getKeywordImage("corporate presentation strategy"),
      notes: "Speaker prompts for the updated strategy slide."
    };
    const updatedSlides = [...presentation.slides];
    updatedSlides.splice(activeSlideIndex + 1, 0, newSlide);
    setPresentation({
      ...presentation,
      slides: updatedSlides
    });
    setActiveSlideIndex(activeSlideIndex + 1);
  };

  const handleDeleteSlide = (index: number) => {
    if (!presentation || presentation.slides.length <= 1) return;
    const updatedSlides = presentation.slides.filter((_, idx) => idx !== index);
    setPresentation({
      ...presentation,
      slides: updatedSlides
    });
    setActiveSlideIndex(Math.max(0, index - 1));
  };

  // Unsplash search trigger
  const triggerUnsplashSearch = async () => {
    setIsSearchingUnsplash(true);
    try {
      const term = unsplashSearch.trim();
      const searchQuery = term || unsplashCategory;
      const res = await fetch(`/api/images/search?q=${encodeURIComponent(searchQuery)}`);
      if (!res.ok) {
        throw new Error("Failed to load images");
      }
      const data = await res.json();
      setUnsplashResults(data.results || []);
    } catch (e) {
      console.error("Failed to query live images:", e);
      setUnsplashResults(CURATED_UNSPLASH_IMAGES[unsplashCategory] || []);
    } finally {
      setIsSearchingUnsplash(false);
    }
  };

  useEffect(() => {
    if (isUnsplashOpen) {
      triggerUnsplashSearch();
    }
  }, [unsplashCategory, isUnsplashOpen]);

  const handleSelectImage = (url: string) => {
    updateActiveSlide({ imageUrl: url });
    setIsUnsplashOpen(false);
  };

  // AI Design Suggestion (Redesign current active slide with expert prompt)
  const handleRequestDesignSuggestion = async () => {
    if (!activeSlide) return;

    setIsSuggestingAI(true);
    setErrorMessage(null);

    try {
      const res = await fetch("/api/gemini/suggest-design", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          currentSlide: activeSlide,
          brandColor,
          themeId: presentation?.themeId
        })
      });

      let responseText = "";
      try {
        responseText = await res.text();
      } catch (e) {
        responseText = "";
      }

      const contentType = res.headers.get("Content-Type") || "";
      if (!contentType.includes("application/json")) {
        const cleanText = responseText.replace(/<[^>]*>/g, '').trim();
        throw new Error(`The server returned an HTML template instead of data. This typically indicates a network bottleneck or that your Gemini API key is missing or expired. Server message: ${cleanText.slice(0, 150) || "None"}`);
      }

      if (!res.ok) {
        let errorMsg = "Design Suggestion server endpoint refused proposal.";
        try {
          const errPayload = JSON.parse(responseText);
          errorMsg = errPayload.error || errorMsg;
        } catch (e) {
          const cleanText = responseText.replace(/<[^>]*>/g, '').trim();
          errorMsg = `Server returned (${res.status}): ${cleanText.slice(0, 250) || "No response body"}`;
        }
        throw new Error(errorMsg);
      }

      let suggestion;
      try {
        suggestion = JSON.parse(responseText);
      } catch (e) {
        throw new Error("Design Suggestion response was not valid JSON.");
      }

      setAiSuggestion(suggestion);
      setShowSuggestionModal(true);
    } catch (err: any) {
      console.error(err);
      setErrorMessage(err.message || "Could not generate design suggestion. Ensure server.ts API routes are responsive.");
    } finally {
      setIsSuggestingAI(false);
    }
  };

  const handleApplySuggestion = () => {
    if (!aiSuggestion || !activeSlide) return;
    updateActiveSlide({
      title: aiSuggestion.title,
      content: aiSuggestion.content,
      layout: aiSuggestion.layout as SlideLayout,
      notes: aiSuggestion.notes,
      imageUrl: getKeywordImage(aiSuggestion.imageSearchQuery),
      imageSearchQuery: aiSuggestion.imageSearchQuery,
      badge: aiSuggestion.badge,
      bgColor: aiSuggestion.bgColor,
      textColor: aiSuggestion.textColor,
      accentColor: aiSuggestion.accentColor,
      primaryColor: aiSuggestion.primaryColor,
      cardBgColor: aiSuggestion.cardBgColor,
      borderColor: aiSuggestion.borderColor
    });
    setShowSuggestionModal(false);
    setAiSuggestion(null);
  };

  // PPTX Export Download
  const handleExportPPTX = async () => {
    if (!presentation) return;
    try {
      await exportToPPTX(presentation);
    } catch (err) {
      console.error(err);
      alert("Error building download bundle. Ensure pptxgenjs has computed all margins.");
    }
  };

  // Graceful client-side image loading error fallback engine 
  const handleImageError = (e: React.SyntheticEvent<HTMLImageElement, Event>) => {
    const target = e.target as HTMLImageElement;
    if (target.dataset.fallbackTriggered) return;
    target.dataset.fallbackTriggered = "true";
    
    // Generate a deterministic integer seed index between 1 and 200 based on previous src code to keep seed stable
    const currentSrc = target.src || "";
    let hash = 0;
    for (let i = 0; i < currentSrc.length; i++) {
      hash = (hash << 5) - hash + currentSrc.charCodeAt(i);
      hash |= 0;
    }
    const seedIndex = (Math.abs(hash) % 200) + 1;
    target.src = `https://picsum.photos/seed/slidecraft-err-${seedIndex}/1200/800`;
  };

  // High-fidelity PDF Export Trigger (creates an isolated document outside iframe bottlenecks and resolves layouts)
  const handleExportPDF = () => {
    const printContent = document.getElementById("pdf-print-container")?.innerHTML;
    if (!printContent) {
      window.print();
      return;
    }

    try {
      const printWindow = window.open("", "_blank");
      if (!printWindow) {
        alert("Your browser blocked the presentation printing window. Standard printer mode will start now. For premium designed orientation templates, please allow pop-ups for this tab or expand the application.");
        window.print();
        return;
      }

      printWindow.document.write(`
        <!DOCTYPE html>
        <html>
          <head>
            <title>${presentation?.title || "Slidesss - Presentation Deck"}</title>
            <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500;700&family=Outfit:wght@400;600;800;900&family=Playfair+Display:ital,wght@0,400;0,600;0,700;1,400&display=swap">
            <script src="https://cdn.tailwindcss.com"></script>
            <script>
              window.tailwind = window.tailwind || {};
              tailwind.config = {
                theme: {
                  extend: {
                    fontFamily: {
                      sans: ["Inter", "sans-serif"],
                      heading: ["Outfit", "Inter", "sans-serif"],
                      serif: ["Playfair Display", "Georgia", "serif"],
                      mono: ["JetBrains Mono", "monospace"],
                    }
                  }
                }
              }
            </script>
            <style>
              body {
                margin: 0 !important;
                padding: 0 !important;
                background-color: #ffffff;
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
              }
              .print-page {
                width: 1920px !important;
                height: 1080px !important;
                min-width: 1920px !important;
                min-height: 1080px !important;
                max-width: 1920px !important;
                max-height: 1080px !important;
                page-break-after: always !important;
                break-after: page !important;
                box-sizing: border-box !important;
                position: relative !important;
                overflow: hidden !important;
                display: block !important;
              }
              @page {
                size: 1920px 1080px;
                margin: 0;
              }
              @media print {
                body {
                  margin: 0 !important;
                  background-color: #ffffff;
                }
                .print-page {
                  margin: 0 !important;
                  border: none !important;
                  box-shadow: none !important;
                  width: 1920px !important;
                  height: 1080px !important;
                }
              }
            </style>
          </head>
          <body>
            <div>
              ${printContent}
            </div>
            <script>
              window.onload = function() {
                setTimeout(function() {
                  window.print();
                }, 850);
              };
            </script>
          </body>
        </html>
      `);
      printWindow.document.close();
    } catch (err) {
      console.error("Popup window print redirect exception:", err);
      window.print();
    }
  };

  return (
    <div id="app-root" className="relative min-h-screen bg-neutral-950 text-neutral-100 font-sans selection:bg-indigo-600 selection:text-white flex flex-col">
      
      {/* COSMIC NEON INTERGALACTIC BACKGROUND */}
      <div className="absolute inset-0 -z-10 pointer-events-none overflow-hidden no-print">
        {/* Deep stellar gradient */}
        <div className="absolute inset-0 bg-radial-at-t from-indigo-950/20 via-neutral-950 to-neutral-950" />
        
        {/* Neon laser structural alignment grid */}
        <div className="absolute inset-0 bg-[linear-gradient(to_right,#1e1b4b_1px,transparent_1px),linear-gradient(to_bottom,#1e1b4b_1px,transparent_1px)] bg-[size:4.5rem_4.5rem] [mask-image:radial-gradient(ellipse_60%_50%_at_50%_40%,#000_70%,transparent_100%)] opacity-25" />
        
        {/* Glowing Nebula dust clouds */}
        <div className="absolute top-[10%] left-[5%] w-[450px] h-[450px] rounded-full bg-indigo-500/10 blur-[130px] mix-blend-screen" />
        <div className="absolute top-[30%] right-[-5%] w-[400px] h-[400px] rounded-full bg-pink-500/8 blur-[110px] mix-blend-screen" />
        <div className="absolute bottom-[10%] left-[25%] w-[600px] h-[600px] rounded-full bg-blue-600/5 blur-[160px] mix-blend-screen" />
        <div className="absolute top-[5%] left-[50%] -translate-x-1/2 w-80 h-80 rounded-full bg-amber-500/5 blur-[110px] mix-blend-screen" />

        {/* Dynamic shining star assets */}
        <div className="absolute top-[12%] left-[24%] w-1.5 h-1.5 bg-indigo-300 rounded-full animate-pulse opacity-60" />
        <div className="absolute top-[38%] left-[78%] w-1 h-1 bg-pink-300 rounded-full animate-ping opacity-40" />
        <div className="absolute top-[62%] left-[12%] w-1.5 h-1.5 bg-amber-200 rounded-full animate-pulse opacity-50" />
        <div className="absolute top-[82%] left-[58%] w-2 h-2 bg-blue-400 rounded-full animate-pulse opacity-75" />
        <div className="absolute top-[48%] left-[42%] w-0.5 h-0.5 bg-white rounded-full opacity-35" />
        <div className="absolute top-[24%] left-[88%] w-1 h-1 bg-white rounded-full opacity-50 animate-pulse" />
        <div className="absolute top-[70%] left-[80%] w-0.5 h-0.5 bg-white rounded-full opacity-30" />
      </div>

      {/* HEADER BAR */}
      <header id="app-header" className="no-print border-b border-neutral-900 bg-neutral-950/80 backdrop-blur-md px-4 sm:px-6 py-3.5 flex flex-col md:flex-row items-center justify-between gap-3 md:gap-4 sticky top-0 z-40">
        <div className="flex items-center space-x-3 w-full md:w-auto justify-between md:justify-start">
          <div className="flex items-center space-x-3">
            <div className="bg-indigo-600 p-2 rounded-xl text-white shadow-lg shadow-indigo-600/30 shrink-0">
              <Sparkles className="w-5 h-5 animate-pulse" />
            </div>
            <div>
              <h1 className="font-heading font-black text-base sm:text-lg tracking-tight bg-gradient-to-r from-indigo-400 via-pink-400 to-amber-400 bg-clip-text text-transparent">
                Slidesss
              </h1>
              <p className="text-[10px] sm:text-xs text-neutral-400 font-mono">Word Document to High-End Presentation</p>
            </div>
          </div>
        </div>

        {/* Global Action items depend on views */}
        {activeStep === "editor" && presentation && (
          <div className="flex flex-wrap items-center justify-center md:justify-end gap-2.5 sm:gap-3 w-full md:w-auto">
            <button
              id="btn-back-upload"
              onClick={() => {
                if (confirm("Are you sure you want to return to upload? Your currently custom-edited slides will be lost.")) {
                  setActiveStep("upload");
                }
              }}
              className="px-3 py-1.5 text-[11px] sm:text-xs font-semibold text-neutral-400 hover:text-white hover:bg-neutral-900 rounded-xl transition-all cursor-pointer"
            >
              Start Over
            </button>

            {/* REAL-TIME PREVIEW COLOR BARS */}
            <div className="hidden lg:flex items-center space-x-2 bg-neutral-900 p-1.5 rounded-xl border border-neutral-800">
              <span className="text-[10px] font-mono font-medium text-neutral-400 px-2 uppercase">Brand Code:</span>
              <input
                id="input-brand-color-editor"
                type="color"
                value={brandColor}
                onChange={(e) => setBrandColor(e.target.value)}
                className="w-8 h-8 rounded-lg cursor-pointer border-0 bg-transparent"
              />
              <div className="flex -space-x-1.5 px-2">
                <div className="w-4 h-4 rounded-full border border-neutral-800 shadow-sm" style={{ backgroundColor: presentation.palette.primary }} />
                <div className="w-4 h-4 rounded-full border border-neutral-800 shadow-sm" style={{ backgroundColor: presentation.palette.secondary }} />
                <div className="w-4 h-4 rounded-full border border-neutral-800 shadow-sm" style={{ backgroundColor: presentation.palette.accent }} />
                <div className="w-4 h-4 rounded-full border border-neutral-800 shadow-sm" style={{ backgroundColor: presentation.palette.background }} />
                <div className="w-4 h-4 rounded-full border border-neutral-800 shadow-sm" style={{ backgroundColor: presentation.palette.text }} />
              </div>
            </div>

            <div className="flex items-center bg-neutral-900 rounded-xl p-1 border border-neutral-800 max-w-full overflow-x-auto shrink-0">
              {Object.values(ThemeId).map((tid) => (
                <button
                  key={tid}
                  id={`btn-editor-theme-${tid}`}
                  onClick={() => handleSwitchTheme(tid)}
                  title={THEME_PRESETS[tid].name}
                  className={`px-2.5 sm:px-3 py-1.5 text-[10px] sm:text-xs font-medium rounded-lg transition-all ${
                    presentation.themeId === tid
                      ? "bg-indigo-600 text-white shadow"
                      : "text-neutral-400 hover:text-white"
                  }`}
                >
                  {THEME_PRESETS[tid].name.split(" ")[0]}
                </button>
              ))}
            </div>

            <div className="flex items-center space-x-2 shrink-0">
              <button
                id="btn-export-pdf"
                onClick={() => {
                  setPendingExportType("pdf");
                  setIsExportSettingsOpen(true);
                }}
                className="bg-neutral-800 hover:bg-neutral-700 text-neutral-100 text-[11px] sm:text-xs font-bold px-3 py-1.5 sm:py-2 rounded-xl flex items-center space-x-1 border border-neutral-750 transition cursor-pointer"
              >
                <Printer className="w-3.5 h-3.5" />
                <span>
                  <span className="hidden sm:inline">Export </span>PDF
                </span>
              </button>
              <button
                id="btn-export-pptx"
                onClick={() => {
                  setPendingExportType("pptx");
                  setIsExportSettingsOpen(true);
                }}
                className="bg-gradient-to-r from-orange-500 to-amber-500 hover:from-orange-600 hover:to-amber-600 text-white text-[11px] sm:text-xs font-bold px-3 py-1.5 sm:py-2 rounded-xl flex items-center space-x-1 shadow-lg shadow-amber-500/25 transition cursor-pointer"
              >
                <Download className="w-3.5 h-3.5" />
                <span>
                  <span className="hidden sm:inline">Download </span>PPTX
                </span>
              </button>
            </div>
          </div>
        )}
      </header>

      {/* COMPREHENSIVE PRO DIAGNOSTICS ERROR MODAL */}
      {errorMessage && (() => {
        const errDetails = getDetailedError(errorMessage);
        return (
          <div id="modal-error-diagnostics" className="no-print fixed inset-0 z-50 bg-neutral-950/80 backdrop-blur-md flex items-center justify-center p-4 animate-in fade-in duration-200">
            <div className="bg-neutral-900 border border-neutral-800 rounded-3xl max-w-md w-full p-6 sm:p-8 space-y-6 shadow-2xl shadow-red-950/10 animate-in zoom-in-95 duration-200 text-left">
              
              {/* Simplified Visual Icon and Header */}
              <div className="flex flex-col items-center text-center space-y-4">
                <div className="p-3 bg-neutral-800/60 border border-neutral-750 rounded-full shrink-0">
                  {errDetails.type === "network" ? (
                    <WifiOff className="w-10 h-10 text-amber-500 animate-pulse" />
                  ) : errDetails.type === "input" ? (
                    <FileText className="w-10 h-10 text-indigo-400" />
                  ) : (
                    <AlertCircle className="w-10 h-10 text-red-500 animate-bounce" />
                  )}
                </div>
                
                <div className="space-y-1.5 w-full">
                  <span className="text-[10px] font-mono uppercase tracking-widest text-neutral-400 block font-extrabold">
                    {errDetails.type === "network" ? "Connection Notice" : errDetails.type === "input" ? "Input Required" : "System Error"}
                  </span>
                  <h3 className="font-heading font-black text-lg sm:text-xl text-white leading-tight">
                    {errDetails.title}
                  </h3>
                </div>
              </div>

              {/* Simplified, friendly explanation of the error (no technical jargon) */}
              <div className="bg-neutral-950/60 border border-neutral-850 p-5 rounded-2xl text-center space-y-3">
                <p className="text-neutral-300 text-xs sm:text-sm leading-relaxed">
                  {errDetails.explanation}
                </p>
                <p className="text-neutral-500 text-[11px] font-medium leading-relaxed italic">
                  Suggested Action: {errDetails.cta}
                </p>
              </div>

              {/* Footer Controls based on Error Type */}
              <div className="flex flex-col gap-2.5 pt-2 border-t border-neutral-800/65">
                <div className="grid grid-cols-2 gap-2">
                  <button
                    id="btn-error-close-primary"
                    onClick={() => setErrorMessage(null)}
                    className="px-4 py-2.5 border border-neutral-800 hover:bg-neutral-800 text-neutral-400 hover:text-white text-xs font-bold rounded-xl transition cursor-pointer text-center"
                  >
                    Cancel
                  </button>

                  <button
                    id="btn-error-bypass-demo"
                    onClick={() => {
                      setErrorMessage(null);
                      handleLoadDemoPresentation();
                    }}
                    className="px-4 py-2.5 bg-neutral-800/80 hover:bg-neutral-750 text-indigo-400 hover:text-indigo-300 border border-neutral-750 text-xs font-bold rounded-xl transition cursor-pointer flex items-center justify-center space-x-1"
                  >
                    <Sparkles className="w-3 h-3" />
                    <span>Load Demo Deck</span>
                  </button>
                </div>

                {/* Main Action buttons customized per type */}
                {errDetails.type === "network" ? (
                  <button
                    id="btn-error-retry-action"
                    onClick={() => {
                      setErrorMessage(null);
                      if (!rawText.trim()) {
                        handleLoadSample();
                      } else {
                        handleGeneratePresentation();
                      }
                    }}
                    className="w-full px-5 py-3 bg-gradient-to-r from-amber-500 to-orange-600 hover:from-amber-400 hover:to-orange-500 text-white text-xs font-bold rounded-xl shadow-lg shadow-amber-500/10 flex items-center justify-center space-x-2 transition cursor-pointer"
                  >
                    <RefreshCw className="w-4 h-4" />
                    <span>Try Again Now</span>
                  </button>
                ) : errDetails.type === "technical" ? (
                  <a
                    id="btn-error-contact-admin"
                    href={`mailto:bukunmiogunneye0@gmail.com?subject=AI%20Presentation%20System%20-%20Technical%20Error&body=Hello%20Support%20Team%2C%0A%0AI%20encountered%20an%20unexpected%20technical%20system%20error%20while%20attempting%20to%20build%20my%20presentation.%0A%0AError%20Details%3A%20${encodeURIComponent(errorMessage || "Internal layout issue.")}%0A%0APlease%20assist%20with%20resolving%20this%20blocker.`}
                    className="w-full px-5 py-3 bg-gradient-to-r from-indigo-600 to-indigo-700 hover:from-indigo-500 hover:to-indigo-600 text-white text-xs font-bold rounded-xl shadow-lg shadow-indigo-600/10 flex items-center justify-center space-x-2 transition cursor-pointer text-center"
                  >
                    <ExternalLink className="w-4 h-4" />
                    <span>Contact Administrator</span>
                  </a>
                ) : (
                  <button
                    id="btn-error-load-outline"
                    onClick={() => {
                      setErrorMessage(null);
                      handleLoadSample();
                    }}
                    className="w-full px-5 py-3 bg-gradient-to-r from-indigo-600 to-indigo-700 hover:from-indigo-500 hover:to-indigo-600 text-white text-xs font-bold rounded-xl shadow-lg shadow-indigo-600/10 flex items-center justify-center space-x-2 transition cursor-pointer"
                  >
                    <FileText className="w-4 h-4" />
                    <span>Use Sample Outlines</span>
                  </button>
                )}
              </div>

            </div>
          </div>
        );
      })()}

      {/* MAIN SCREEN WRAP */}
      <main className="no-print flex-1 flex flex-col">
        {activeStep === "upload" ? (
          
          /* VIEW 1: LANDING & INITIAL CONVERSION STEP (CENTERED UX) */
          <div id="view-upload-sandbox" className="flex-1 w-full bg-gradient-to-r from-indigo-400 via-pink-400 to-amber-400 p-[16px] sm:p-[40px]">
            <div className="w-full bg-neutral-950 rounded-[24px] py-12 px-4 sm:px-6 md:px-8 flex flex-col items-center justify-start space-y-12 min-h-screen">
              
              {/* Max width container for keeping content elegant and readable on desktop */}
              <div className="max-w-3xl w-full flex flex-col items-center justify-start space-y-10">
                
                {/* HERO INTRODUCTION ZONE */}
                <div className="text-center space-y-4 max-w-2xl px-2 sm:px-0">
                  <span className="px-3.5 py-1 bg-indigo-500/10 text-indigo-400 border border-indigo-500/20 text-xs font-mono rounded-full font-bold inline-block mx-auto">
                    No sign-up required.
                  </span>
                  <h2 className="font-heading font-black text-[32px] sm:text-[48px] leading-tight tracking-tight text-white max-w-xl mx-auto line-clamp-3 sm:line-clamp-none min-h-[96px] sm:min-h-[144px]">
                    <TypewriterText />
                  </h2>
                  <p className="text-[14px] sm:text-sm text-neutral-400 leading-normal max-w-md mx-auto line-clamp-2 sm:line-clamp-none">
                    Upload any Microsoft Word document. Our advanced AI system constructs beautiful, custom-branded presentation slides instantly.
                  </p>
                </div>

                {/* INTEGRATED SEMANTIC PROCESSING CONTAINER (DRAG AND DROP BOX COMES NEXT) */}
                <div className="w-full bg-neutral-900/65 backdrop-blur-xl border border-neutral-800/90 rounded-3xl p-6 lg:p-8 space-y-6 shadow-xl shadow-indigo-950/10">
              <div className="flex items-center justify-between">
                <div className="flex items-center space-x-2 text-indigo-400">
                  <FileText className="w-4 h-4" />
                  <span className="font-bold text-xs uppercase tracking-widest font-mono">Word File Processing Zone</span>
                </div>
                <div className="hidden sm:flex items-center space-x-2">
                  <button
                    id="btn-load-sample"
                    onClick={handleLoadSample}
                    className="px-2.5 py-1.5 bg-neutral-850/70 hover:bg-neutral-800/90 border border-neutral-750 text-indigo-300 text-[10px] font-semibold rounded-lg flex items-center space-x-1.5 transition-all cursor-pointer"
                    title="Load document text from sample corporate strategy"
                  >
                    <Sliders className="w-3 h-3" />
                    <span>Load Text</span>
                  </button>
                  <button
                    id="btn-preview-demo"
                    onClick={handleLoadDemoPresentation}
                    className="px-2.5 py-1.5 bg-gradient-to-r from-teal-500/20 to-indigo-500/20 hover:from-teal-500/30 hover:to-indigo-500/30 border border-indigo-500/30 text-teal-300 text-[10px] font-bold rounded-lg flex items-center space-x-1.5 transition-all shadow-md shadow-indigo-950/20 cursor-pointer"
                    title="Instantly open presentation studio with a fully designed demo deck (bypasses remote AI generation limits)"
                  >
                    <Play className="w-3 h-3 text-teal-400" />
                    <span>Demo Studio</span>
                  </button>
                </div>
              </div>

              {/* FILE SELECTOR CONTAINER (DRAG AND DROP BOX COMES IMMEDIATELY NEXT) */}
              <div className="relative border-2 border-dashed border-neutral-800 hover:border-indigo-500 bg-neutral-950/35 hover:bg-neutral-950/50 rounded-2xl p-6 flex flex-col items-center justify-center text-center transition-all min-h-48 group">
                <input
                  id="file-upload-input"
                  type="file"
                  accept=".docx"
                  onChange={handleFileChange}
                  className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
                />
                
                {isParsingDoc ? (
                  <div className="space-y-3">
                    <RefreshCw className="w-10 h-10 text-indigo-400 animate-spin mx-auto" />
                    <p className="text-xs text-neutral-300 font-mono">Extracting semantic layout coordinates from Word file...</p>
                  </div>
                ) : docName ? (
                  <div className="space-y-2">
                    <div className="bg-emerald-500/10 text-emerald-400 p-3 rounded-full w-12 h-12 flex items-center justify-center mx-auto border border-emerald-500/20">
                      <Check className="w-6 h-6" />
                    </div>
                    <p className="text-xs font-bold text-white font-mono break-all">{docName}</p>
                    <p className="text-[10px] text-neutral-500 uppercase font-mono">Ready for presentation synthesis</p>
                  </div>
                ) : (
                  <div className="space-y-4">
                    <div className="bg-neutral-900/50 text-neutral-400 p-3.5 rounded-xl group-hover:bg-indigo-600 group-hover:text-white transition-all w-12 h-12 flex items-center justify-center mx-auto">
                      <Upload className="w-5 h-5" />
                    </div>
                    <div>
                      <p className="text-xs font-bold text-neutral-200">Drag & Drop Word Document Here</p>
                      <p className="text-[11px] text-neutral-400 mt-1 font-mono">Accepts Microsoft Word DOCX formats to parse content structures</p>
                    </div>
                  </div>
                )}
              </div>

              {/* REVIEW EXTRACTED TEXTBOX */}
              <div className="flex flex-col space-y-2">
                <div className="flex items-center justify-between text-xs text-neutral-400">
                  <span className="font-mono uppercase text-[10px] font-bold">Review Parsed Content Outlines</span>
                  {rawText && (
                    <span className="font-mono text-[9px] bg-neutral-800 text-neutral-300 px-1.5 py-0.5 rounded">
                      {rawText.length} characters parsed
                    </span>
                  )}
                </div>
                <textarea
                  id="text-outline-editor"
                  value={rawText}
                  onChange={(e) => setRawText(e.target.value)}
                  placeholder="Review parsed document texts, or type your slide notes directly here to construct layouts..."
                  className="bg-neutral-950/35 border border-neutral-800 rounded-xl p-4 text-xs text-neutral-300 font-mono leading-relaxed focus:outline-none focus:border-indigo-500 focus:bg-neutral-950/60 resize-none h-36"
                />
              </div>

              {/* GLOBAL DESIGN AESTHETIC STYLE CHOOSER */}
              <div className="space-y-2 pt-2 border-t border-neutral-800/40">
                <div className="flex items-center space-x-2 text-indigo-400">
                  <Layout className="w-4 h-4" />
                  <span className="font-mono uppercase text-[10px] font-bold tracking-wide">Select Slide Theme Style</span>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  {Object.values(THEME_PRESETS).map((t) => (
                    <button
                      key={t.id}
                      id={`btn-landing-theme-${t.id}`}
                      onClick={() => setSelectedThemeId(t.id)}
                      className={`text-left p-3 rounded-xl border transition-all flex flex-col justify-between h-20 ${
                        selectedThemeId === t.id
                          ? "bg-indigo-950/20 border-indigo-500/80 shadow-md shadow-indigo-500/10"
                          : "bg-neutral-950/35 border-neutral-800 hover:border-neutral-750"
                      }`}
                    >
                      <span className="font-bold text-xs text-white block">{t.name}</span>
                      <span className="text-[10px] text-neutral-400 leading-tight block line-clamp-2">{t.description}</span>
                    </button>
                  ))}
                </div>
              </div>

              {/* CONVERTER SYSTEM ACTION TRIGGER */}
              <button
                id="btn-generate-deck"
                disabled={!rawText.trim() || isGenerating}
                onClick={handleOpenDeckConfigOrGenerate}
                className={`w-full py-4 px-6 rounded-2xl font-bold text-sm tracking-wide flex items-center justify-center space-x-2 shadow-lg transition-all ${
                  isGenerating || !rawText.trim()
                    ? "bg-neutral-800 text-neutral-500 cursor-not-allowed"
                    : "bg-gradient-to-r from-indigo-600 via-pink-600 to-indigo-600 text-white hover:opacity-95 shadow-indigo-600/20 cursor-pointer"
                }`}
              >
                {isGenerating ? (
                  <>
                    <RefreshCw className="w-4 h-4 animate-spin" />
                    <span>{statusMessage} ({generationProgress}%)</span>
                  </>
                ) : (
                  <>
                    <Sparkles className="w-4 h-4 text-yellow-300 animate-pulse" />
                    <span>Generate PDF/PPTX (1920×1080)</span>
                  </>
                )}
              </button>

              {/* TRANSIENT SERVICE BOTTLENECK FRIENDLY FALLBACK */}
              <div className="text-center pt-1 border-t border-neutral-800/40">
                <p className="text-xs text-neutral-400">
                  Experiencing high demand or API delay?{" "}
                  <button
                    type="button"
                    onClick={handleLoadDemoPresentation}
                    className="text-teal-400 hover:text-teal-300 font-bold underline cursor-pointer transition-all bg-transparent border-0 p-0 inline-block focus:outline-none"
                  >
                    Load Offline Demo Deck
                  </button>
                </p>
              </div>
            </div>

          </div>

          {/* DEPLOYMENT LINE FOOTER */}
          <div className="text-neutral-500 text-[10px] font-mono pt-4 flex items-center space-x-2 justify-center pb-8">
            <span className="inline-block w-2 h-2 bg-emerald-500 rounded-full animate-ping" />
            <span>Express Server Engine Ready  |  Gemini Core v3.5-flash Online</span>
          </div>

        </div>
      </div>
        ) : (
          
          /* VIEW 2: SLIDE PRESENTATION MASTER STUDIO (PRO SLIDE EDITOR) */
          <div id="view-presentation-studio" className="flex-1 flex flex-col lg:flex-row items-stretch overflow-hidden">
            
            {/* COLUMN A: SLIDE NAVIGATION SIDEBAR */}
            <div id="editor-sidebar" className="no-print w-full lg:w-64 border-r border-neutral-900 bg-neutral-950 flex flex-col overflow-y-auto shrink-0 justify-between">
              
              <div className="p-4 space-y-3">
                <div className="flex items-center justify-between text-xs">
                  <span className="font-mono text-neutral-400 uppercase font-bold tracking-widest text-[9px]">Presentation Deck Structure</span>
                  <span className="px-2 py-0.5 bg-neutral-900 text-neutral-400 font-mono text-[9px] rounded-full">
                    {presentation?.slides.length || 0} Slides
                  </span>
                </div>

                {/* THUMBNAILS CONTAINER */}
                <div className="space-y-2 mt-2">
                  {presentation?.slides.map((s, index) => {
                    const isSelected = index === activeSlideIndex;
                    return (
                      <div
                        key={s.id}
                        id={`slide-thumb-${index}`}
                        onClick={() => setActiveSlideIndex(index)}
                        className={`group p-3 rounded-xl border transition-all text-left cursor-pointer flex items-start justify-between ${
                          isSelected
                            ? "bg-neutral-900 border-indigo-600 shadow-sm"
                            : "bg-neutral-950/40 border-neutral-900 hover:bg-neutral-900/60"
                        }`}
                      >
                        <div className="flex space-x-2.5 items-start overflow-hidden mr-1">
                          {/* Slide counter */}
                          <div className={`mt-0.5 w-5 h-5 rounded flex items-center justify-center text-[10px] font-mono shrink-0 font-bold ${
                            isSelected ? "bg-indigo-600 text-white" : "bg-neutral-900 text-neutral-500"
                          }`}>
                            {index + 1}
                          </div>
                          
                          {/* Title and layout info */}
                          <div className="overflow-hidden">
                            <h4 className="font-bold text-xs text-neutral-200 truncate group-hover:text-white">{s.title || "Untitled Slide"}</h4>
                            <p className="text-[9px] text-neutral-500 font-mono uppercase mt-0.5 tracking-lighter">{s.layout}</p>
                          </div>
                        </div>

                        {/* Move Actions overlays on thumbnails */}
                        <div className="flex flex-col space-y-1 opacity-0 group-hover:opacity-100 transition-all ml-1 shrink-0">
                          <button
                            id={`btn-move-up-${index}`}
                            title="Move slide representation up"
                            disabled={index === 0}
                            onClick={(e) => { e.stopPropagation(); handleMoveSlide(index, "up"); }}
                            className="bg-neutral-900 p-1 rounded text-neutral-400 hover:text-white disabled:opacity-30 disabled:hover:text-neutral-400"
                          >
                            <ArrowUp className="w-2.5 h-2.5" />
                          </button>
                          <button
                            id={`btn-move-down-${index}`}
                            title="Move slide representation down"
                            disabled={index === presentation.slides.length - 1}
                            onClick={(e) => { e.stopPropagation(); handleMoveSlide(index, "down"); }}
                            className="bg-neutral-900 p-1 rounded text-neutral-400 hover:text-white disabled:opacity-30 disabled:hover:text-neutral-400"
                          >
                            <ArrowDown className="w-2.5 h-2.5" />
                          </button>
                        </div>
                      </div>
                    );
                  })}
                </div>

                {/* ADD STRATEGY SLIDE KEYNOTES */}
                <button
                  id="btn-add-slide"
                  onClick={handleAddBlankSlide}
                  className="w-full py-2.5 bg-neutral-900 hover:bg-neutral-850 border border-neutral-800 text-indigo-300 text-xs font-semibold rounded-xl flex items-center justify-center space-x-1.5 transition-all mt-4 cursor-pointer"
                >
                  <Plus className="w-4 h-4" />
                  <span>Insert Strategy Slide</span>
                </button>
              </div>

              {/* CURRENT PALETTE INFO PANEL */}
              {presentation && (
                <div className="p-4 border-t border-neutral-900 bg-neutral-950/80 space-y-3 font-mono text-[9px]">
                  <span className="text-neutral-500 uppercase font-black tracking-widest block">Active Branding Schema</span>
                  
                  <div className="space-y-1 bg-neutral-900 p-2.5 rounded-lg border border-neutral-850">
                    <div className="flex justify-between items-center text-neutral-400">
                      <span>Brand HEX:</span>
                      <span className="text-white font-bold">{presentation.brandColor}</span>
                    </div>
                    <div className="flex justify-between items-center text-neutral-400">
                      <span>Active Theme:</span>
                      <span className="text-white font-bold">{THEME_PRESETS[presentation.themeId].name}</span>
                    </div>
                  </div>

                  <div className="text-neutral-500 leading-tight">
                    Double-click bullet points on the central preview slide card to edit copy context instantly.
                  </div>
                </div>
              )}
            </div>

            {/* COLUMN B: MAIN WORKSPACE CARD (SLIDE VIEWPORT AND FIELD CONTROLS) */}
            <div id="editor-workspace" className="flex-1 bg-neutral-900 p-6 flex flex-col justify-between items-stretch overflow-y-auto space-y-6">
              
              {activeSlide ? (
                <div className="flex-1 flex flex-col lg:flex-row gap-6 items-stretch">
                  
                  {/* LEFT CHUNKS: SLIDE PREVIEW BOX AND SPEAKER NOTES */}
                  <div className="w-full lg:w-[65%] flex flex-col justify-between space-y-4">
                    
                    {/* THE SLIDE CANVAS VIEWPORT */}
                    <div
                      id="slide-preview-viewport"
                      className="relative w-full rounded-2xl overflow-hidden shadow-2xl border border-neutral-800/80 slide-preview-container print-page"
                      style={{
                        backgroundColor: slideBg,
                        color: slideText,
                        "--bg-color": slideBg,
                        "--text-color": slideText
                      } as React.CSSProperties}
                    >
                      
                      {/* DECORATIVE LIGHT SHAPES FOR AVANTGARDE */}
                      {presentation?.themeId === "creative-avantgarde" && (
                        <div className="absolute -left-12 -top-12 w-48 h-48 bg-pink-300/10 rounded-full blur-2xl pointer-events-none" />
                      )}

                      {/* DECORATIVE RADIAL GLOWS FOR COSMIC */}
                      {presentation?.themeId === "cosmic-slate" && (
                        <div className="absolute right-0 bottom-0 w-80 h-80 bg-indigo-500/5 rounded-full blur-3xl pointer-events-none" />
                      )}

                      {/* SLIDES CONTENT LAYOUTS INNER BORDER BOX */}
                      <div className={`w-full h-full p-8 lg:p-12 flex flex-col justify-between select-none relative ${
                        presentation?.themeId === "brutalist-mono" ? "border-4 border-black" : ""
                      }`}>
                        
                        {/* RUNTIME SLIDE BRAND INDICATOR (Muted Header) */}
                        <div className="flex items-center justify-between no-print mb-2 border-b border-neutral-700/10 pb-2">
                          <span
                            className="text-[9px] font-mono tracking-widest font-black uppercase"
                            style={{ color: slideAccent }}
                          >
                            {presentation?.title || docName || "Untitled Presentation"}
                          </span>
                          <span className="text-[10px] font-mono text-neutral-500">
                            Slide {activeSlideIndex + 1} of {presentation?.slides.length || 0}
                          </span>
                        </div>

                        {/* RENDER ACTIVE PREVIEW BY LAYOUT TYPE SELECT */}
                        <div className="flex-1 flex flex-col justify-center">
                          {activeSlide.layout === "title-slide" ? (
                            
                            /* LAYOUT 1: TITLE SLIDE */
                            <div className="text-center space-y-6 py-6">
                              {activeSlide.badge && (
                                <div className="mb-1">
                                  <span
                                    className="text-[10px] font-mono font-bold uppercase tracking-wider px-2.5 py-1 rounded-full border shadow-sm inline-block"
                                    style={{
                                      backgroundColor: `${slideAccent}1d`,
                                      color: slideAccent,
                                      borderColor: `${slideAccent}3a`
                                    }}
                                  >
                                    {activeSlide.badge}
                                  </span>
                                </div>
                              )}
                              <h3
                                className={`tracking-tight leading-tight ${activePresetTheme.fontHeading}`}
                                style={{ color: slidePrimary, fontSize: getFontSizeStyle(2.7) }}
                              >
                                {presentation?.title || "BUSINESS PRESENTATION"}
                              </h3>
                              <div className="block mt-4">
                                <p
                                  className="font-mono font-medium opacity-85"
                                  style={{ color: slideText, fontSize: getFontSizeStyle(1.1) }}
                                >
                                  {activeSlide.title}
                                </p>
                              </div>
                              {activeSlide.content.length > 0 && (
                                <div className="flex flex-wrap items-center justify-center gap-3 pt-4">
                                  {activeSlide.content.map((point, idx) => (
                                    <span
                                      key={idx}
                                      className="bg-neutral-500/5 border border-neutral-750/30 rounded-full font-medium"
                                      style={{ color: slideAccent, fontSize: getFontSizeStyle(0.75) }}
                                    >
                                      {point}
                                    </span>
                                  ))}
                                </div>
                              )}
                            </div>

                          ) : activeSlide.layout === "two-column" ? (
                            
                            /* LAYOUT 2: TWO COLUMN SPLIT */
                            <div className="space-y-6">
                              <div>
                                {activeSlide.badge && (
                                  <div className="mb-2">
                                    <span
                                      className="text-[10px] font-mono font-bold uppercase tracking-wider px-2.5 py-1 rounded-full border shadow-sm inline-block"
                                      style={{
                                        backgroundColor: `${slideAccent}1d`,
                                        color: slideAccent,
                                        borderColor: `${slideAccent}3a`
                                      }}
                                    >
                                      {activeSlide.badge}
                                    </span>
                                  </div>
                                )}
                                <h3 className={`font-bold ${activePresetTheme.fontHeading} mb-2`} style={{ color: slidePrimary, fontSize: getFontSizeStyle(1.5) }}>
                                  {activeSlide.title}
                                </h3>
                              </div>
                              <div className="grid grid-cols-1 md:grid-cols-2 gap-6 pt-2">
                                <div className="p-4 rounded-xl space-y-3" style={{ backgroundColor: slideCardBg, border: `1px solid ${slideBorder}` }}>
                                  <span className="h-1.5 w-6 rounded-full block" style={{ backgroundColor: slideAccent }} />
                                  <ul className="space-y-2.5">
                                    {activeSlide.content.slice(0, Math.ceil(activeSlide.content.length / 2)).map((bullet, idx) => {
                                      const isListItem = bullet.trim().startsWith("•") || bullet.trim().startsWith("-") || bullet.trim().startsWith("*") || /^\d+\./.test(bullet.trim());
                                      let displayText = bullet;
                                      if (isListItem) {
                                        displayText = bullet.replace(/^[•\-\*\s]+/, '').replace(/^\d+\.\s*/, '');
                                      }
                                      return (
                                        <li key={idx} className={`flex items-start ${isListItem ? "space-x-2" : ""} ${activePresetTheme.fontBody}`} style={{ color: slideText, fontSize: getFontSizeStyle(0.85) }}>
                                          {isListItem && (
                                            <span className="mt-1.5 w-1.5 h-1.5 rounded-full shrink-0" style={{ backgroundColor: slidePrimary }} />
                                          )}
                                          <span>{displayText}</span>
                                        </li>
                                      );
                                    })}
                                  </ul>
                                </div>
                                <div className="p-4 rounded-xl space-y-3" style={{ backgroundColor: slideCardBg, border: `1px solid ${slideBorder}` }}>
                                  <span className="h-1.5 w-6 rounded-full block" style={{ backgroundColor: slidePrimary }} />
                                  <ul className="space-y-2.5">
                                    {activeSlide.content.slice(Math.ceil(activeSlide.content.length / 2)).map((bullet, idx) => {
                                      const isListItem = bullet.trim().startsWith("•") || bullet.trim().startsWith("-") || bullet.trim().startsWith("*") || /^\d+\./.test(bullet.trim());
                                      let displayText = bullet;
                                      if (isListItem) {
                                        displayText = bullet.replace(/^[•\-\*\s]+/, '').replace(/^\d+\.\s*/, '');
                                      }
                                      return (
                                        <li key={idx} className={`flex items-start ${isListItem ? "space-x-2" : ""} ${activePresetTheme.fontBody}`} style={{ color: slideText, fontSize: getFontSizeStyle(0.85) }}>
                                          {isListItem && (
                                            <span className="mt-1.5 w-1.5 h-1.5 rounded-full shrink-0" style={{ backgroundColor: slideAccent }} />
                                          )}
                                          <span>{displayText}</span>
                                        </li>
                                      );
                                    })}
                                  </ul>
                                </div>
                              </div>
                            </div>

                          ) : activeSlide.layout === "quote-slide" ? (
                            
                            /* LAYOUT 3: STYLIZED QUOTE */
                            <div className="text-center max-w-4xl mx-auto space-y-4 py-4">
                              {activeSlide.badge && (
                                <div className="mb-2">
                                  <span
                                    className="text-[10px] font-mono font-bold uppercase tracking-wider px-2.5 py-1 rounded-full border shadow-sm inline-block"
                                    style={{
                                      backgroundColor: `${slideAccent}1d`,
                                      color: slideAccent,
                                      borderColor: `${slideAccent}3a`
                                    }}
                                  >
                                    {activeSlide.badge}
                                  </span>
                                </div>
                              )}
                              <span className="text-5xl leading-none font-black opacity-30 block" style={{ color: slideAccent }}>“</span>
                              <h4
                                className="font-black italic tracking-wide text-center"
                                style={{ color: slideText, fontSize: getFontSizeStyle(1.4) }}
                              >
                                {activeSlide.title}
                              </h4>
                              {activeSlide.content.length > 0 && (
                                <div className="block pt-2">
                                  <p className="uppercase tracking-widest font-mono font-bold text-center opacity-85" style={{ color: slidePrimary, fontSize: getFontSizeStyle(0.75) }}>
                                    - {activeSlide.content.join(" & ")}
                                  </p>
                                </div>
                              )}
                            </div>

                          ) : activeSlide.layout === "image-left" && activeSlide.imageUrl ? (
                            
                            /* LAYOUT 4: VISUAL LEFT SPLIT */
                            <div className="grid grid-cols-1 md:grid-cols-2 gap-8 items-center h-full">
                              <div className="relative rounded-xl overflow-hidden shadow-md h-full min-h-48 group">
                                <img
                                  src={activeSlide.imageUrl}
                                  alt="Slide graphic"
                                  className="absolute inset-0 w-full h-full object-cover rounded-xl"
                                  referrerPolicy="no-referrer"
                                  onError={handleImageError}
                                />
                                <div className="absolute inset-0 bg-neutral-950/20" />
                                {activeSlide.imageCaption && (
                                  <div className="absolute bottom-2 left-2 bg-neutral-950/80 text-[9px] px-2 py-0.5 rounded text-neutral-300">
                                    {activeSlide.imageCaption}
                                  </div>
                                )}
                              </div>
                              <div className="space-y-4">
                                <div>
                                  {activeSlide.badge && (
                                    <div className="mb-2">
                                      <span
                                        className="text-[10px] font-mono font-bold uppercase tracking-wider px-2.5 py-1 rounded-full border shadow-sm inline-block"
                                        style={{
                                          backgroundColor: `${slideAccent}1d`,
                                          color: slideAccent,
                                          borderColor: `${slideAccent}3a`
                                        }}
                                      >
                                        {activeSlide.badge}
                                      </span>
                                    </div>
                                  )}
                                  <h3 className={`font-bold ${activePresetTheme.fontHeading} mb-2`} style={{ color: slidePrimary, fontSize: getFontSizeStyle(1.5) }}>
                                    {activeSlide.title}
                                  </h3>
                                </div>
                                <ul className="space-y-3">
                                  {activeSlide.content.map((bullet, idx) => {
                                    const isListItem = bullet.trim().startsWith("•") || bullet.trim().startsWith("-") || bullet.trim().startsWith("*") || /^\d+\./.test(bullet.trim());
                                    let displayText = bullet;
                                    if (isListItem) {
                                      displayText = bullet.replace(/^[•\-\*\s]+/, '').replace(/^\d+\.\s*/, '');
                                    }
                                    return (
                                      <li key={idx} className={`flex items-start ${isListItem ? "space-x-2.5" : ""} ${activePresetTheme.fontBody}`} style={{ color: slideText, fontSize: getFontSizeStyle(0.85) }}>
                                        {isListItem && (
                                          <span className="mt-1.5 w-1.5 h-1.5 rounded-full shrink-0" style={{ backgroundColor: slideAccent }} />
                                        )}
                                        <span>{displayText}</span>
                                      </li>
                                    );
                                  })}
                                </ul>
                              </div>
                            </div>

                          ) : activeSlide.layout === "image-right" && activeSlide.imageUrl ? (
                            
                            /* LAYOUT 5: VISUAL RIGHT SPLIT */
                            <div className="grid grid-cols-1 md:grid-cols-2 gap-8 items-center h-full">
                              <div className="space-y-4">
                                <div>
                                  {activeSlide.badge && (
                                    <div className="mb-2">
                                      <span
                                        className="text-[10px] font-mono font-bold uppercase tracking-wider px-2.5 py-1 rounded-full border shadow-sm inline-block"
                                        style={{
                                          backgroundColor: `${slideAccent}1d`,
                                          color: slideAccent,
                                          borderColor: `${slideAccent}3a`
                                        }}
                                      >
                                        {activeSlide.badge}
                                      </span>
                                    </div>
                                  )}
                                  <h3 className={`font-bold ${activePresetTheme.fontHeading} mb-2`} style={{ color: slidePrimary, fontSize: getFontSizeStyle(1.5) }}>
                                    {activeSlide.title}
                                  </h3>
                                </div>
                                <ul className="space-y-3">
                                  {activeSlide.content.map((bullet, idx) => {
                                    const isListItem = bullet.trim().startsWith("•") || bullet.trim().startsWith("-") || bullet.trim().startsWith("*") || /^\d+\./.test(bullet.trim());
                                    let displayText = bullet;
                                    if (isListItem) {
                                      displayText = bullet.replace(/^[•\-\*\s]+/, '').replace(/^\d+\.\s*/, '');
                                    }
                                    return (
                                      <li key={idx} className={`flex items-start ${isListItem ? "space-x-2.5" : ""} ${activePresetTheme.fontBody}`} style={{ color: slideText, fontSize: getFontSizeStyle(0.85) }}>
                                        {isListItem && (
                                          <span className="mt-1.5 w-1.5 h-1.5 rounded-full shrink-0" style={{ backgroundColor: slideAccent }} />
                                        )}
                                        <span>{displayText}</span>
                                      </li>
                                    );
                                  })}
                                </ul>
                              </div>
                              <div className="relative rounded-xl overflow-hidden shadow-md h-full min-h-48">
                                <img
                                  src={activeSlide.imageUrl}
                                  alt="Slide graphic"
                                  className="absolute inset-0 w-full h-full object-cover rounded-xl"
                                  referrerPolicy="no-referrer"
                                  onError={handleImageError}
                                />
                                <div className="absolute inset-0 bg-neutral-950/20" />
                              </div>
                            </div>

                          ) : activeSlide.layout === "stats-bento" ? (
                            
                            /* LAYOUT 6: METRICS BENTO GRID */
                            <div className="space-y-4">
                              <div>
                                {activeSlide.badge && (
                                  <div className="mb-2">
                                    <span
                                      className="text-[10px] font-mono font-bold uppercase tracking-wider px-2.5 py-1 rounded-full border shadow-sm inline-block"
                                      style={{
                                        backgroundColor: `${slideAccent}1d`,
                                        color: slideAccent,
                                        borderColor: `${slideAccent}3a`
                                      }}
                                    >
                                      {activeSlide.badge}
                                    </span>
                                  </div>
                                )}
                                <h3 className={`font-bold ${activePresetTheme.fontHeading} mb-2`} style={{ color: slidePrimary, fontSize: getFontSizeStyle(1.25) }}>
                                  {activeSlide.title}
                                </h3>
                              </div>
                              <div className="grid grid-cols-2 gap-4">
                                {[0, 1, 2, 3].map((val) => {
                                  const textVal = activeSlide.content[val] || "N/A metric value";
                                  const spaceIdx = textVal.indexOf(" ");
                                  let metricNum = textVal;
                                  let metricLabel = "";
                                  if (spaceIdx > 0) {
                                    metricNum = textVal.substring(0, spaceIdx);
                                    metricLabel = textVal.substring(spaceIdx + 1);
                                  }

                                  const numericStr = metricNum.replace(/[^0-9.]/g, '');
                                  const numericVal = numericStr ? parseFloat(numericStr) : null;
                                  const isPercent = metricNum.includes("%");

                                  return (
                                    <div
                                      key={val}
                                      className="p-4 rounded-xl space-y-2 flex flex-col justify-between transition-all"
                                      style={{
                                        backgroundColor: slideCardBg,
                                        border: `1px solid ${slideBorder}`
                                      }}
                                    >
                                      <div className="space-y-0.5">
                                        <span className="font-heading font-black block" style={{ color: slideAccent, fontSize: getFontSizeStyle(1.3) }}>
                                          {metricNum}
                                        </span>
                                        <span className="font-mono font-medium block leading-tight text-neutral-400 uppercase tracking-wide" style={{ fontSize: getFontSizeStyle(0.55) }}>
                                          {metricLabel || "Strategic target metrics"}
                                        </span>
                                      </div>

                                      {/* BAR CHART OR DIAGRAM WIDGET AS REQUESTED */}
                                      <div className="w-full pt-1">
                                        {isPercent && numericVal !== null ? (
                                          <div className="space-y-1">
                                            <div className="flex justify-between items-center text-[7px] font-mono opacity-50">
                                              <span>Progress Ratio</span>
                                              <span>{Math.min(100, Math.round(numericVal))}%</span>
                                            </div>
                                            <div className="w-full h-1 bg-neutral-900/60 rounded-full overflow-hidden p-[0.5px] border border-neutral-800/15">
                                              <div 
                                                className="h-full rounded-full transition-all duration-1000"
                                                style={{
                                                  width: `${Math.min(100, Math.max(8, numericVal))}%`,
                                                  background: `linear-gradient(to right, ${slidePrimary}, ${slideAccent})`,
                                                  boxShadow: `0 0 4px ${slideAccent}40`
                                                }}
                                              />
                                            </div>
                                          </div>
                                        ) : numericVal !== null ? (
                                          <div className="space-y-1">
                                            <div className="flex justify-between items-center text-[7px] font-mono opacity-50">
                                              <span>Comparative Bar Graph</span>
                                              <span>Val: {metricNum}</span>
                                            </div>
                                            <div className="h-4 flex items-end justify-between gap-[1.5px] w-full">
                                              {[20, 35, 50, 30, 60, 75, 45, 65, 85, 100].map((hValue, barIdx) => {
                                                const scaleFactor = Math.min(1.2, Math.max(0.3, numericVal / 100));
                                                const calcHeight = Math.min(100, Math.max(15, hValue * scaleFactor));
                                                return (
                                                  <div 
                                                    key={barIdx}
                                                    className="w-full rounded-t-[1px] transition-all"
                                                    style={{
                                                      height: `${calcHeight}%`,
                                                      backgroundColor: barIdx >= 8 ? slideAccent : `${slidePrimary}25`,
                                                    }}
                                                  />
                                                );
                                              })}
                                            </div>
                                          </div>
                                        ) : (
                                          <div className="space-y-1">
                                            <div className="flex justify-between items-center text-[7px] font-mono opacity-50">
                                              <span>Trajectory Trend</span>
                                              <span>Active Metrics</span>
                                            </div>
                                            <div className="h-4 w-full flex items-center justify-center">
                                              <svg className="w-full h-full" viewBox="0 0 100 20" preserveAspectRatio="none">
                                                <path 
                                                  d="M 0 15 Q 15 5, 30 14 T 60 4 T 90 12 L 100 8" 
                                                  fill="none" 
                                                  stroke={slideAccent} 
                                                  strokeWidth="1.5" 
                                                  strokeLinecap="round"
                                                  className="opacity-90"
                                                />
                                                <path 
                                                  d="M 0 15 Q 15 5, 30 14 T 60 4 T 90 12 L 100 8 L 100 20 L 0 20 Z" 
                                                  fill={`url(#ambient-grad-${val})`}
                                                  className="opacity-25"
                                                />
                                                <defs>
                                                  <linearGradient id={`ambient-grad-${val}`} x1="0" y1="0" x2="0" y2="1">
                                                    <stop offset="0%" stopColor={slideAccent} />
                                                    <stop offset="100%" stopColor="transparent" />
                                                  </linearGradient>
                                                </defs>
                                              </svg>
                                            </div>
                                          </div>
                                        )}
                                      </div>
                                    </div>
                                  );
                                })}
                              </div>
                            </div>

                          ) : activeSlide.layout === "comparison-table" ? (
                            
                            /* LAYOUT 8: COMPARISON TABLE */
                            <div className="space-y-4">
                              <div>
                                {activeSlide.badge && (
                                  <div className="mb-2">
                                    <span
                                      className="text-[10px] font-mono font-bold uppercase tracking-wider px-2.5 py-1 rounded-full border shadow-sm inline-block"
                                      style={{
                                        backgroundColor: `${slideAccent}1d`,
                                        color: slideAccent,
                                        borderColor: `${slideAccent}3a`
                                      }}
                                    >
                                      {activeSlide.badge}
                                    </span>
                                  </div>
                                )}
                                <h3 className={`font-bold ${activePresetTheme.fontHeading} mb-2`} style={{ color: slidePrimary, fontSize: getFontSizeStyle(1.1) }}>
                                  {activeSlide.title}
                                </h3>
                              </div>

                              <div className="overflow-hidden rounded-xl border shadow-lg" style={{ borderColor: slideBorder, backgroundColor: slideCardBg }}>
                                <table className="w-full text-left border-collapse table-fixed">
                                  <thead>
                                    <tr style={{ borderBottom: `2px solid ${slideBorder}`, backgroundColor: `${slidePrimary}0d` }}>
                                      <th className="px-3 py-2 w-1/4 text-[8px] uppercase tracking-wider font-mono opacity-70" style={{ color: slideText, fontSize: getFontSizeStyle(0.55) }}>Aspect</th>
                                      <th className="px-3 py-2 w-3/8 font-bold" style={{ color: slidePrimary, fontSize: getFontSizeStyle(0.7) }}>Traditional Approach</th>
                                      <th className="px-3 py-2 w-3/8 font-bold flex items-center gap-1" style={{ color: slideAccent, fontSize: getFontSizeStyle(0.7) }}>
                                        Slidesss Systems
                                        <span className="text-[6px] font-mono uppercase bg-emerald-500/10 text-emerald-400 border border-emerald-400/20 px-1 py-0.2 rounded animate-pulse">Optimal</span>
                                      </th>
                                    </tr>
                                  </thead>
                                  <tbody>
                                    {(() => {
                                      const rows: { aspect: string; left: string; right: string }[] = [];
                                      const contentLen = activeSlide.content.length;
                                      
                                      let hasSeparator = false;
                                      activeSlide.content.forEach(bullet => {
                                        if (bullet.includes(" vs ") || bullet.includes(" | ") || bullet.includes(" - ")) {
                                          hasSeparator = true;
                                        }
                                      });

                                      if (hasSeparator) {
                                        activeSlide.content.forEach((bullet, idx) => {
                                          let aspect = `Dimension ${idx + 1}`;
                                          let left = bullet;
                                          let right = "";
                                          const separators = [" vs ", " | ", " - "];
                                          for (const sep of separators) {
                                            if (bullet.includes(sep)) {
                                              const parts = bullet.split(sep);
                                              left = parts[0].trim();
                                              right = parts.slice(1).join(sep).trim();
                                              if (left.includes(": ")) {
                                                const aspParts = left.split(": ");
                                                aspect = aspParts[0].trim();
                                                left = aspParts.slice(1).join(": ").trim();
                                              }
                                              break;
                                            }
                                          }
                                          rows.push({ aspect, left, right });
                                        });
                                      } else {
                                        const half = Math.ceil(contentLen / 2);
                                        for (let i = 0; i < half; i++) {
                                          const leftBullet = activeSlide.content[i] || "";
                                          const rightBullet = activeSlide.content[i + half] || "";
                                          let aspect = `Dimension ${i + 1}`;
                                          let leftText = leftBullet;
                                          let rightText = rightBullet;
                                          if (leftBullet.includes(": ")) {
                                            const parts = leftBullet.split(": ");
                                            aspect = parts[0].trim();
                                            leftText = parts.slice(1).join(": ").trim();
                                          }
                                          if (rightBullet.includes(": ")) {
                                            const parts = rightBullet.split(": ");
                                            rightText = parts.slice(1).join(": ").trim();
                                          }
                                          rows.push({ aspect, left: leftText, right: rightText });
                                        }
                                      }

                                      return rows.map((row, rIdx) => (
                                        <tr key={rIdx} className="hover:bg-neutral-900/10 transition-colors" style={{ borderBottom: rIdx < rows.length - 1 ? `1px solid ${slideBorder}` : 'none' }}>
                                          <td className="px-3 py-2 font-mono opacity-80 break-words" style={{ color: slideText, fontSize: getFontSizeStyle(0.6) }}>{row.aspect}</td>
                                          <td className="px-3 py-2 opacity-85 break-words" style={{ color: slideText, fontSize: getFontSizeStyle(0.7) }}>
                                            <div className="flex items-start gap-1">
                                              <span className="text-neutral-550 shrink-0 mt-0.5">✗</span>
                                              <span>{row.left || "—"}</span>
                                            </div>
                                          </td>
                                          <td className="px-3 py-2 font-medium break-words" style={{ color: slidePrimary, fontSize: getFontSizeStyle(0.7) }}>
                                            <div className="flex items-start gap-1">
                                              <span className="text-emerald-500 shrink-0 mt-0.5">✓</span>
                                              <span>{row.right || "—"}</span>
                                            </div>
                                          </td>
                                        </tr>
                                      ));
                                    })()}
                                  </tbody>
                                </table>
                              </div>
                            </div>

                          ) : (
                            
                            /* LAYOUT 7: DEFAULT / HEADLINE BULLET */
                            <div className="space-y-4">
                              <div>
                                {activeSlide.badge && (
                                  <div className="mb-2">
                                    <span
                                      className="text-[10px] font-mono font-bold uppercase tracking-wider px-2.5 py-1 rounded-full border shadow-sm inline-block"
                                      style={{
                                        backgroundColor: `${slideAccent}1d`,
                                        color: slideAccent,
                                        borderColor: `${slideAccent}3a`
                                      }}
                                    >
                                      {activeSlide.badge}
                                    </span>
                                  </div>
                                )}
                                <h3 className={`font-bold ${activePresetTheme.fontHeading} mb-2`} style={{ color: slidePrimary, fontSize: getFontSizeStyle(1.5) }}>
                                  {activeSlide.title}
                                </h3>
                              </div>
                              <ul className="space-y-3.5">
                                {activeSlide.content.map((bullet, idx) => {
                                  const isListItem = bullet.trim().startsWith("•") || bullet.trim().startsWith("-") || bullet.trim().startsWith("*") || /^\d+\./.test(bullet.trim());
                                  let displayText = bullet;
                                  if (isListItem) {
                                    displayText = bullet.replace(/^[•\-\*\s]+/, '').replace(/^\d+\.\s*/, '');
                                  }
                                  return (
                                    <li key={idx} className={`flex items-start ${isListItem ? "space-x-2.5" : ""} ${activePresetTheme.fontBody}`} style={{ color: slideText, fontSize: getFontSizeStyle(0.85) }}>
                                      {isListItem && (
                                        <span className="mt-1.5 w-1.5 h-1.5 rounded-full shrink-0 animate-pulse" style={{ backgroundColor: slideAccent }} />
                                      )}
                                      <span>{displayText}</span>
                                    </li>
                                  );
                                })}
                              </ul>
                            </div>
                          )}
                        </div>

                        {/* SLIDE FOOTER BRAND LINE */}
                        <div className="flex items-center justify-between mt-4 text-[9px] border-t border-neutral-700/10 pt-2 opacity-60">
                          <span style={{ color: slideText }}>Excelsior Systems © 2026</span>
                          <span className="font-mono">{presentation?.themeId.toUpperCase()} EDITION</span>
                        </div>
                      </div>
                    </div>

                    {/* DYNAMIC SPEAKER NOTES CONTAINER */}
                    <div className="no-print bg-neutral-950 border border-neutral-850 p-4 rounded-xl space-y-2">
                      <div className="flex items-center space-x-1 text-xs text-neutral-400">
                        <Monitor className="w-3.5 h-3.5" />
                        <span className="font-mono font-bold uppercase text-[9px]">Speaker Coaching Guidelines</span>
                      </div>
                      <textarea
                        id="input-slide-notes"
                        value={activeSlide.notes || ""}
                        onChange={(e) => updateActiveSlide({ notes: e.target.value })}
                        placeholder="Add professional commentary notes here to guide your meeting presentation pacing."
                        className="w-full bg-transparent text-xs text-neutral-300 font-mono resize-none h-16 focus:outline-none"
                      />
                    </div>
                  </div>

                  {/* RIGHT CHUNKS: SLIDE COPY EDITORS & AI LAYOUT TOOLBAR */}
                  <div className="no-print w-full lg:w-[35%] flex flex-col justify-between space-y-4">
                    
                    {/* ACCORDION 1: SLIDE CONTENT FIELDS TEXT EDITOR */}
                    <div className="bg-neutral-950 border border-neutral-850 rounded-2xl p-5 space-y-4 flex-1 flex flex-col justify-between">
                      <div className="space-y-3 flex-1 overflow-y-auto">
                        <div className="flex items-center space-x-2 text-indigo-400 pb-1 border-b border-neutral-900">
                          <Sliders className="w-4 h-4" />
                          <h3 className="font-bold text-xs tracking-wider uppercase font-mono">Edit Slide Content</h3>
                        </div>

                        {/* EDITABLE TITLE */}
                        <div className="space-y-1">
                          <label className="text-[10px] font-mono text-neutral-500 uppercase font-bold block">Slide Heading</label>
                          <input
                            id="edit-slide-title"
                            type="text"
                            value={activeSlide.title}
                            onChange={(e) => updateActiveSlide({ title: e.target.value })}
                            className="w-full bg-neutral-900 border border-neutral-800 rounded-lg px-3 py-2 text-xs font-bold text-white focus:outline-none focus:border-indigo-500"
                          />
                        </div>

                        {/* EDITABLE BULLETS */}
                        <div className="space-y-2 pt-2">
                          <label className="text-[10px] font-mono text-neutral-500 uppercase font-bold block">Key content bullet points</label>
                          <div className="space-y-2 max-h-48 overflow-y-auto pr-1">
                            {activeSlide.content.map((point, index) => (
                              <div key={index} className="flex items-center space-x-1.5 group">
                                <input
                                  type="text"
                                  value={point}
                                  onChange={(e) => handleEditBullet(index, e.target.value)}
                                  className="flex-1 bg-neutral-900 border border-neutral-800 rounded-lg px-3 py-1.5 text-xs text-neutral-300 focus:outline-none focus:border-indigo-500"
                                />
                                <button
                                  id={`btn-delete-bullet-${index}`}
                                  onClick={() => handleDeleteBullet(index)}
                                  className="text-neutral-500 hover:text-red-400 p-1.5 rounded-lg hover:bg-neutral-900 transition-all opacity-0 group-hover:opacity-100"
                                  title="Remove bullet point"
                                >
                                  <Trash2 className="w-3.5 h-3.5" />
                                </button>
                              </div>
                            ))}
                          </div>
                          
                          <button
                            id="btn-add-bullet"
                            onClick={handleAddBullet}
                            className="text-[10px] font-mono text-indigo-400 hover:text-indigo-300 font-bold flex items-center space-x-1 py-1 px-1.5 rounded hover:bg-neutral-900 transition-all"
                          >
                            <Plus className="w-3 h-3" />
                            <span>Add Bullet Item</span>
                          </button>
                        </div>
                      </div>

                      {/* LAYOUT CHOICE PICKER PANEL */}
                      <div className="pt-4 border-t border-neutral-900 space-y-2">
                        <label className="text-[10px] font-mono text-neutral-500 uppercase font-bold block">Visual Slide Grid Layout</label>
                        <div className="grid grid-cols-2 gap-1.5">
                          {[
                            { id: "title-slide", label: "Title Cover" },
                            { id: "two-column", label: "2x Column Card" },
                            { id: "headline-bullet", label: "Standard Bullets" },
                            { id: "quote-slide", label: "Keynote Quote" },
                            { id: "image-left", label: "Visual Left" },
                            { id: "image-right", label: "Visual Right" },
                            { id: "stats-bento", label: "Stats Bento" },
                            { id: "comparison-table", label: "Comparison Table" }
                          ].map((gridLay) => (
                            <button
                              key={gridLay.id}
                              id={`btn-lay-select-${gridLay.id}`}
                              onClick={() => updateActiveSlide({ layout: gridLay.id as SlideLayout })}
                              className={`py-1.5 px-2 rounded text-[10px] font-medium text-left transition-all ${
                                activeSlide.layout === gridLay.id
                                  ? "bg-indigo-950/40 text-indigo-400 border border-indigo-500/40"
                                  : "bg-neutral-900 border border-neutral-850 text-neutral-400 hover:text-neutral-300"
                              }`}
                            >
                              {gridLay.label}
                            </button>
                          ))}
                        </div>
                      </div>

                      {/* TYPOGRAPHY DYNAMIC SCALER SCREEN */}
                      <div className="pt-4 border-t border-neutral-900 space-y-2">
                        <div className="flex items-center justify-between">
                          <label className="text-[10px] font-mono text-neutral-500 uppercase font-bold block">Typography Scale Tuning</label>
                          <span className="text-[9px] font-mono font-bold uppercase text-indigo-400 bg-indigo-950/30 px-1.5 py-0.5 rounded border border-indigo-500/10">
                            {activeSlide.fontSize || "medium"}
                          </span>
                        </div>
                        <div className="grid grid-cols-3 gap-2">
                          {[
                            { id: "small", label: "Small", marker: "A", desc: "Compact spacing" },
                            { id: "medium", label: "Medium", marker: "AA", desc: "Default density" },
                            { id: "large", label: "Large", marker: "AAA", desc: "High visibility" }
                          ].map((sz) => (
                            <button
                              key={sz.id}
                              id={`btn-font-size-${sz.id}`}
                              onClick={() => updateActiveSlide({ fontSize: sz.id as "small" | "medium" | "large" })}
                              className={`py-2 px-1 rounded-xl flex flex-col items-center justify-center transition-all ${
                                (activeSlide.fontSize || "medium") === sz.id
                                  ? "bg-indigo-950/40 text-indigo-400 border border-indigo-500/40 shadow-[0_0_12px_rgba(79,70,229,0.15)]"
                                  : "bg-neutral-900 border border-neutral-850 text-neutral-400 hover:text-neutral-300 hover:border-neutral-700"
                              }`}
                              title={sz.desc}
                            >
                              <span className={`font-black tracking-tighter ${
                                sz.id === "small" ? "text-xs" : sz.id === "medium" ? "text-sm" : "text-base"
                              }`}>
                                {sz.marker}
                              </span>
                              <span className="text-[9px] font-semibold tracking-wider uppercase mt-1 opacity-80">{sz.label}</span>
                            </button>
                          ))}
                        </div>
                      </div>

                      {/* SLIDE CANVAS BACKGROUND COLOR & CONTRAST ACCESSIBILITY CHECKER */}
                      <div className="pt-4 border-t border-neutral-900 space-y-2.5">
                        <div className="flex items-center justify-between">
                          <label className="text-[10px] font-mono text-neutral-500 uppercase font-bold block">Slide Background Color</label>
                          <span className="text-[9px] font-mono font-bold uppercase text-indigo-400 bg-indigo-950/30 px-1.5 py-0.5 rounded border border-indigo-500/10">
                            {activeSlide.bgColor || presentation?.palette.background || "Default"}
                          </span>
                        </div>
                        
                        {/* Quick Presets */}
                        <div className="flex flex-wrap gap-2 items-center">
                          {[
                            { label: "Midnight Space", color: "#0B0F19" },
                            { label: "Crisp Snow", color: "#FFFFFF" },
                            { label: "Book Cream", color: "#FCF9F2" },
                            { label: "Slate Grey", color: "#1E293B" },
                            { label: "Matte Black", color: "#111827" },
                            { label: "Silver Muted", color: "#F3F4F6" }
                          ].map((preset) => (
                            <button
                              key={preset.color}
                              type="button"
                              onClick={() => handleSlideBgChange(preset.color)}
                              className={`w-6 h-6 rounded-full border transition-all cursor-pointer relative flex items-center justify-center ${
                                (activeSlide.bgColor || presentation?.palette.background) === preset.color
                                  ? "ring-2 ring-indigo-500 scale-110 shadow-md"
                                  : "border-neutral-850 hover:scale-105"
                              }`}
                              style={{ backgroundColor: preset.color }}
                              title={`${preset.label} (${preset.color})`}
                            >
                              {(activeSlide.bgColor || presentation?.palette.background) === preset.color && (
                                <Check className={`w-3.5 h-3.5 ${preset.color === "#FFFFFF" || preset.color === "#FCF9F2" || preset.color === "#F3F4F6" ? "text-slate-900" : "text-white"}`} />
                              )}
                            </button>
                          ))}

                          {/* Color Input container */}
                          <div className="relative flex items-center gap-1 shrink-0 bg-neutral-900 border border-neutral-800 rounded-lg p-1">
                            <input
                              type="color"
                              value={activeSlide.bgColor || presentation?.palette.background || "#0B0F19"}
                              onChange={(e) => handleSlideBgChange(e.target.value)}
                              className="w-5 h-5 rounded cursor-pointer border-0 bg-transparent p-0"
                            />
                            <input
                              type="text"
                              maxLength={7}
                              value={activeSlide.bgColor || presentation?.palette.background || "#0B0F19"}
                              onChange={(e) => handleSlideBgChange(e.target.value)}
                              className="w-16 bg-transparent text-[10px] font-mono font-bold text-neutral-300 focus:outline-none uppercase text-center"
                            />
                          </div>
                        </div>

                        {/* Accessibility Audit Feedback Badge */}
                        {(() => {
                          const currentBg = activeSlide.bgColor || presentation?.palette.background || "#0B0F19";
                          const currentText = activeSlide.textColor || presentation?.palette.text || "#FFFFFF";
                          const contrastObj = testColorAccessibility(currentBg, currentText);
                          return (
                            <div className="bg-neutral-900/45 border border-neutral-850 p-2.5 rounded-xl space-y-2 text-[11px] font-mono leading-tight">
                              <div className="flex items-center justify-between">
                                <span className="text-neutral-400">Contrast: <strong className="text-white font-mono">{contrastObj.ratio.toFixed(1)}:1</strong></span>
                                <span className={`px-2 py-0.5 rounded-full font-bold text-[9px] uppercase tracking-wide flex items-center gap-1 ${
                                  contrastObj.level === "AAA"
                                    ? "bg-emerald-950/30 text-emerald-400 border border-emerald-500/20"
                                    : contrastObj.level === "AA"
                                    ? "bg-indigo-950/30 text-indigo-400 border border-indigo-500/20"
                                    : "bg-amber-950/30 text-amber-400 border border-amber-500/20"
                                }`}>
                                  <span className={`inline-block w-1.5 h-1.5 rounded-full ${
                                    contrastObj.level === "AAA" ? "bg-emerald-400" : contrastObj.level === "AA" ? "bg-indigo-400" : "bg-amber-400 animate-pulse"
                                  }`} />
                                  {contrastObj.label}
                                </span>
                              </div>
                              {!contrastObj.isAAPass && (
                                <div className="pt-1 flex items-center justify-between border-t border-neutral-800/60">
                                  <span className="text-[10px] text-amber-300">Below WCAG AA (4.5:1)</span>
                                  <button
                                    type="button"
                                    onClick={() => {
                                      const optimalText = ensureAccessibleTextColor(currentBg, currentText);
                                      updateActiveSlide({ textColor: optimalText });
                                    }}
                                    className="px-2 py-0.5 text-[9px] font-bold bg-amber-500/20 hover:bg-amber-500/30 text-amber-200 border border-amber-500/30 rounded transition cursor-pointer"
                                  >
                                    Auto-Fix Contrast
                                  </button>
                                </div>
                              )}
                            </div>
                          );
                        })()}
                      </div>
                    </div>

                    {/* ACCORDION 2: GRAPHICS AND UNSPLASH INTEGRATIONS */}
                    <div className="bg-neutral-950 border border-neutral-850 rounded-2xl p-5 space-y-4">
                      <div className="flex items-center justify-between pb-1 border-b border-neutral-900">
                        <div className="flex items-center space-x-2 text-indigo-400">
                          <ImageIcon className="w-4 h-4" />
                          <h3 className="font-bold text-xs tracking-wider uppercase font-mono">Integrated Visual Elements</h3>
                        </div>
                        <button
                          id="btn-trigger-unsplash"
                          onClick={() => setIsUnsplashOpen(true)}
                          className="px-2.5 py-1 bg-indigo-600/15 text-indigo-400 border border-indigo-500/20 hover:bg-indigo-600 hover:text-white text-[10px] font-mono font-bold rounded-lg transition-all"
                        >
                          Unsplash Database
                        </button>
                      </div>

                      {/* MINI IMAGE VIEW PREVIEW IN CONTROL COLUMN */}
                      {activeSlide.imageUrl ? (
                        <div className="flex items-center space-x-3.5 bg-neutral-900 p-2.5 rounded-xl border border-neutral-850">
                          <img
                            src={activeSlide.imageUrl}
                            alt="Miniature"
                            className="w-12 h-12 rounded-lg object-cover border border-neutral-800"
                            referrerPolicy="no-referrer"
                            onError={handleImageError}
                          />
                          <div className="flex-1 overflow-hidden">
                            <span className="text-[10px] font-mono text-neutral-400 block uppercase">Context Search Phrase:</span>
                            <span className="text-xs font-bold text-white truncate block">{activeSlide.imageSearchQuery || "No Search Tag"}</span>
                          </div>
                          <button
                            id="btn-remove-unsplash-image"
                            onClick={() => updateActiveSlide({ imageUrl: undefined })}
                            className="p-1 hover:bg-neutral-800 rounded text-neutral-500 hover:text-red-400 transition"
                            title="Remove layout visual helper"
                          >
                            ✕
                          </button>
                        </div>
                      ) : (
                        <p className="text-xs text-neutral-500 italic leading-relaxed">
                          No graphic asset mapped. Set slide layout to 'Visual Left' or 'Visual Right' and select an Unsplash image to skin.
                        </p>
                      )}
                    </div>

                    {/* ACCORDION 3: PRO-AI DESIGN REDESIGN SUGGESTION CHUNKS */}
                    <div className="bg-neutral-950 border border-neutral-850 rounded-2xl p-5 space-y-4">
                      <div className="flex items-center space-x-2 text-indigo-400 pb-1 border-b border-neutral-900">
                        <Wand2 className="w-4 h-4" />
                        <h3 className="font-bold text-xs tracking-wider uppercase font-mono">AI Creative Assistant</h3>
                      </div>
                      
                      <p className="text-[11px] text-neutral-400 leading-relaxed">
                        Stuck on structure? Let Gemini review this slide's copy context. It will redesign typography headings, condense lines, and select the optimal structural layout type.
                      </p>

                      <button
                        id="btn-suggest-redesign"
                        disabled={isSuggestingAI}
                        onClick={handleRequestDesignSuggestion}
                        className="w-full py-2.5 bg-neutral-900 hover:bg-indigo-950/20 border border-neutral-800 hover:border-indigo-500/50 text-neutral-100 hover:text-indigo-400 text-xs font-bold rounded-xl flex items-center justify-center space-x-1.5 transition-all cursor-pointer"
                      >
                        {isSuggestingAI ? (
                          <>
                            <RefreshCw className="w-4 h-4 animate-spin text-indigo-400" />
                            <span>Asking Creative Director...</span>
                          </>
                        ) : (
                          <>
                            <Sparkles className="w-4 h-4 text-amber-400" />
                            <span>Pro AI Design Redesign Suggestion</span>
                          </>
                        )}
                      </button>
                    </div>

                    {/* DELETE CONVENTION CONTROL FOOTER */}
                    <div className="pt-2">
                      <button
                        id="btn-delete-slide"
                        disabled={presentation.slides.length <= 1}
                        onClick={() => handleDeleteSlide(activeSlideIndex)}
                        className="w-full py-2 bg-neutral-950 hover:bg-red-950/20 border border-neutral-900 hover:border-red-900/40 text-neutral-500 hover:text-red-400 text-xs font-semibold rounded-xl flex items-center justify-center space-x-1 transition-all disabled:opacity-30 disabled:cursor-not-allowed"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                        <span>Delete Current Slide</span>
                      </button>
                    </div>

                  </div>

                </div>
              ) : (
                <div className="text-center py-12">
                  <p className="text-neutral-400">Empty presentation deck layout state found.</p>
                </div>
              )}

            </div>

          </div>
        )}
      </main>

      {/* MODAL 1: UNSPLASH GLOBAL SEARCH DIALOG OVERLAY */}
      {isUnsplashOpen && (
        <div id="modal-unsplash" className="no-print fixed inset-0 z-50 bg-neutral-950/80 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-neutral-900 border border-neutral-800 rounded-3xl max-w-2xl w-full p-6 space-y-5 animate-in fade-in zoom-in-95 duration-200">
            <div className="flex items-center justify-between">
              <div className="flex items-center space-x-2 text-indigo-400">
                <ImageIcon className="w-5 h-5" />
                <h3 className="font-heading font-extrabold text-lg text-white">Unsplash High-Res Graphics Database</h3>
              </div>
              <button
                id="btn-close-unsplash"
                onClick={() => setIsUnsplashOpen(false)}
                className="text-neutral-500 hover:text-white text-sm"
              >
                ✕
              </button>
            </div>

            {/* SEARCH FORM INPUT */}
            <form
              onSubmit={(e) => { e.preventDefault(); triggerUnsplashSearch(); }}
              className="flex items-center space-x-2 bg-neutral-950 p-1.5 rounded-xl border border-neutral-850"
            >
              <div className="flex-1 flex items-center space-x-2 px-2.5">
                <Search className="w-4 h-4 text-neutral-500" />
                <input
                  id="unsplash-search-input"
                  type="text"
                  value={unsplashSearch}
                  onChange={(e) => setUnsplashSearch(e.target.value)}
                  placeholder="Search over 3M business, marketing, code and structure photos..."
                  className="bg-transparent text-sm text-white w-full focus:outline-none"
                />
              </div>
              <button
                id="btn-submit-unsplash-search"
                type="submit"
                className="px-4 py-2 bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-bold rounded-lg cursor-pointer"
              >
                Search
              </button>
            </form>

            {/* CURATED SELECTOR TABS */}
            <div className="flex flex-wrap items-center gap-1.5 border-b border-neutral-800 pb-3">
              {Object.keys(CURATED_UNSPLASH_IMAGES).map((cat) => (
                <button
                  key={cat}
                  id={`btn-unsplash-cat-${cat.replace(/\s+/g, "_")}`}
                  onClick={() => { setUnsplashSearch(""); setUnsplashCategory(cat); }}
                  className={`px-3 py-1 text-[10px] font-mono font-bold uppercase rounded-lg transition-all ${
                    unsplashCategory === cat && !unsplashSearch
                      ? "bg-neutral-850 text-indigo-300"
                      : "text-neutral-400 hover:text-neutral-200"
                  }`}
                >
                  {cat.split(" & ")[0]}
                </button>
              ))}
            </div>

            {/* GRID RESULTS OF SEARCH OUTLINE */}
            <div className="h-64 overflow-y-auto pr-1">
              {isSearchingUnsplash ? (
                <div className="h-full flex flex-col items-center justify-center text-center space-y-2">
                  <RefreshCw className="w-6 h-6 text-indigo-400 animate-spin" />
                  <p className="text-xs text-neutral-400 font-mono">Pulling royalty-free media layers...</p>
                </div>
              ) : unsplashResults.length > 0 ? (
                <div className="grid grid-cols-2 sm:grid-cols-4 gap-3.5">
                  {unsplashResults.map((img) => (
                    <div
                      key={img.id}
                      onClick={() => handleSelectImage(img.url)}
                      className="group relative rounded-xl overflow-hidden cursor-pointer border border-neutral-800 hover:border-indigo-500 transition-all shadow-sm flex flex-col h-28"
                    >
                      <img
                        src={img.thumbnail}
                        alt={img.title}
                        className="w-full h-full object-cover group-hover:scale-105 transition"
                        referrerPolicy="no-referrer"
                        onError={handleImageError}
                      />
                      <div className="absolute inset-0 bg-gradient-to-t from-neutral-950 via-transparent opacity-0 group-hover:opacity-100 transition-all flex items-end p-2">
                        <span className="text-[8px] font-mono text-neutral-300 truncate w-full block">photo of {img.author}</span>
                      </div>
                    </div>
                  ))}
                </div>
              ) : (
                <div className="h-full flex items-center justify-center text-center text-xs text-neutral-500">
                  No images found. Try refining search keywords or browse other catalogs.
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* MODAL 2: SIDE-BY-SIDE AI PRO DESIGN LAYOUT SUGGESTION CARDS */}
      {showSuggestionModal && aiSuggestion && activeSlide && (
        <div id="modal-ai-suggestion" className="no-print fixed inset-0 z-50 bg-neutral-950/80 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-neutral-900 border border-neutral-800 rounded-3xl max-w-4xl w-full p-6 space-y-5 animate-in fade-in zoom-in-95 duration-200">
            <div className="flex items-center justify-between">
              <div className="flex items-center space-x-2 text-indigo-400">
                <Wand2 className="w-5 h-5 text-amber-400" />
                <h3 className="font-heading font-extrabold text-lg text-white">Gemini Design Pro Strategy Proposal</h3>
              </div>
              <button
                id="btn-close-ai-suggest"
                onClick={() => setShowSuggestionModal(false)}
                className="text-neutral-500 hover:text-white text-sm"
              >
                ✕
              </button>
            </div>

            <p className="text-xs text-neutral-400 leading-relaxed">
              We parsed your current slide copy against business templates and aligned them with your company colors. Here is the revised comparison deck:
            </p>

            {/* COMPARISON SLIDES DECK GRID */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              
              {/* CARD A: ORIGINAL SLIDE DESIGN */}
              <div className="bg-neutral-950 border border-neutral-850 p-4.5 rounded-2xl space-y-3">
                <span className="text-[9px] font-mono text-neutral-500 uppercase block font-bold">Standard Slide Structure (Original)</span>
                <h4 className="text-sm font-bold text-white leading-tight truncate">{activeSlide.title}</h4>
                <div className="p-3 bg-neutral-900 rounded-xl space-y-2 border border-neutral-850">
                  <span className="text-[10px] font-mono font-semibold text-neutral-400 block uppercase">LAYOUT: {activeSlide.layout.toUpperCase()}</span>
                  <ul className="space-y-1.5">
                    {activeSlide.content.map((b, i) => (
                      <li key={i} className="text-[10px] text-neutral-300 truncate">
                        • {b}
                      </li>
                    ))}
                  </ul>
                </div>
              </div>

              {/* CARD B: AI ENHANCED SLIDE DESIGN */}
              <div className="bg-gradient-to-b from-indigo-950/10 to-neutral-950 border border-indigo-500/35 p-4.5 rounded-2xl space-y-3 shadow-md shadow-indigo-500/5">
                <div className="flex items-center justify-between">
                  <span className="text-[9px] font-mono text-indigo-400 uppercase block font-bold">Gemini Redesigned Slide Layout</span>
                  <span className="bg-indigo-500/20 text-indigo-300 text-[8px] font-mono px-1 rounded uppercase">PRO DESIGN</span>
                </div>
                <h4 className="text-sm font-bold text-indigo-300 leading-tight truncate">{aiSuggestion.title}</h4>
                
                <div className="p-3 bg-neutral-900 rounded-xl space-y-2 border border-indigo-500/20">
                  <div className="flex justify-between items-center">
                    <span className="text-[10px] font-mono font-semibold text-amber-400 block uppercase">LAYOUT: {aiSuggestion.layout.toUpperCase()}</span>
                    <span className="text-[9px] text-neutral-400 font-mono italic">Revised & Simplified</span>
                  </div>
                  <ul className="space-y-1.5">
                    {aiSuggestion.content.map((b: string, i: number) => (
                      <li key={i} className="text-[10px] text-white">
                        • {b}
                      </li>
                    ))}
                  </ul>
                </div>
              </div>

            </div>

            {/* ACTION TRIGGERS IN FOOTER */}
            <div className="flex items-center justify-end space-x-3 pt-3 border-t border-neutral-850">
              <button
                id="btn-dismiss-suggestion"
                onClick={() => setShowSuggestionModal(false)}
                className="px-4 py-2 border border-neutral-850 hover:bg-neutral-850 text-neutral-400 hover:text-neutral-200 text-xs font-bold rounded-xl transition"
              >
                Dismiss Suggestion
              </button>
              <button
                id="btn-apply-suggestion"
                onClick={handleApplySuggestion}
                className="px-5 py-2 bg-gradient-to-r from-indigo-600 to-indigo-500 hover:from-indigo-500 hover:to-indigo-600 text-white text-xs font-bold rounded-xl shadow-lg shadow-indigo-500/20 flex items-center space-x-1 transition cursor-pointer"
              >
                <Check className="w-4 h-4" />
                <span>Apply AI Redesign Suggestions</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 3: EXPORT DESIGN PRO BRAND COLOR CONFIGURATION (SHOWS ON GENERATE PDF OR PPTX) */}
      {isExportSettingsOpen && presentation && (
        <div id="modal-brand-color-export" className="no-print fixed inset-0 z-50 bg-neutral-950/80 backdrop-blur-sm flex items-center justify-center p-4 animate-in fade-in duration-150">
          <div className="bg-neutral-900 border border-neutral-800 rounded-3xl max-w-md w-full p-6 space-y-5 animate-in fade-in zoom-in-95 duration-200">
            <div className="flex items-center justify-between">
              <div className="flex items-center space-x-2 text-indigo-400 font-heading">
                <Palette className="w-5 h-5 text-amber-400 font-bold" />
                <h3 className="font-heading font-extrabold text-lg text-white">Export Deck Settings</h3>
              </div>
              <button
                id="btn-close-export-settings"
                onClick={() => setIsExportSettingsOpen(false)}
                className="text-neutral-500 hover:text-white text-sm"
              >
                ✕
              </button>
            </div>

            {/* RESOLUTION SPECIFICATION BADGE */}
            <div className="p-3 bg-neutral-950 rounded-2xl border border-neutral-800 flex items-center justify-between">
              <div className="space-y-0.5">
                <span className="text-[10px] font-mono text-neutral-400 uppercase font-bold block">Output Resolution</span>
                <span className="text-xs font-bold text-white font-mono">1920 × 1080 Full HD (16:9)</span>
              </div>
              <span className="px-2.5 py-1 text-[10px] font-mono font-bold bg-indigo-500/20 text-indigo-300 border border-indigo-500/30 rounded-lg">
                Fixed 1080p
              </span>
            </div>

            {/* BRAND COLOR ALIGNMENT */}
            <div className="bg-neutral-950 border border-neutral-800 p-4 rounded-2xl space-y-3">
              <div className="flex items-center space-x-4">
                <div className="relative">
                  <input
                    id="input-hex-picker-modal"
                    type="color"
                    value={brandColor}
                    onChange={(e) => setBrandColor(e.target.value)}
                    className="w-12 h-12 rounded-xl cursor-pointer border border-neutral-750 bg-transparent p-1"
                  />
                </div>
                <div className="flex-1 space-y-1">
                  <label className="text-[10px] font-mono text-neutral-400 block uppercase font-bold">Brand HEX Color</label>
                  <input
                    id="input-hex-text-modal"
                    type="text"
                    value={brandColor}
                    onChange={(e) => setBrandColor(normalizeHex(e.target.value, brandColor))}
                    placeholder="#4F46E5"
                    className="bg-neutral-900 border border-neutral-800 px-3 py-1.5 text-xs rounded-lg text-white font-mono w-full focus:outline-none focus:border-indigo-500"
                  />
                </div>
              </div>

              {/* REAL TIME PALETTE PROJECTION PREVIEW */}
              <div className="space-y-1 pt-1">
                <div className="flex items-center justify-between">
                  <span className="text-[9px] font-mono font-bold text-neutral-400 block uppercase">Real-Time Palette</span>
                  {(() => {
                    const testAudit = testColorAccessibility(presentation.palette.background, presentation.palette.text);
                    return (
                      <span className={`text-[9px] font-mono font-bold px-1.5 py-0.5 rounded ${
                        testAudit.isAAPass ? "text-emerald-400 bg-emerald-950/30" : "text-amber-400 bg-amber-950/30"
                      }`}>
                        WCAG {testAudit.level} ({testAudit.ratio}:1)
                      </span>
                    );
                  })()}
                </div>
                <div className="grid grid-cols-5 gap-1.5 p-1 bg-neutral-950 rounded-xl border border-neutral-850">
                  <div className="text-center p-1 rounded-lg border border-neutral-900" style={{ backgroundColor: presentation.palette.primary }}>
                    <span className="text-[8px] font-mono font-bold px-1" style={{ color: presentation.palette.text }}>Brand</span>
                  </div>
                  <div className="text-center p-1 rounded-lg border border-neutral-900" style={{ backgroundColor: presentation.palette.secondary }}>
                    <span className="text-[8px] font-mono font-bold px-1" style={{ color: presentation.palette.background }}>Sec</span>
                  </div>
                  <div className="text-center p-1 rounded-lg border border-neutral-900" style={{ backgroundColor: presentation.palette.accent }}>
                    <span className="text-[8px] font-mono font-bold px-1" style={{ color: presentation.palette.text }}>Acc</span>
                  </div>
                  <div className="text-center p-1 rounded-lg border border-neutral-900" style={{ backgroundColor: presentation.palette.cardBg }}>
                    <span className="text-[8px] font-mono font-bold px-1" style={{ color: presentation.palette.text }}>Card</span>
                  </div>
                  <div className="text-center p-1 rounded-lg border border-neutral-900" style={{ backgroundColor: presentation.palette.background }}>
                    <span className="text-[8px] font-mono font-bold px-1" style={{ color: presentation.palette.text }}>Canvas</span>
                  </div>
                </div>
              </div>

              {/* WCAG Auto Fix Action Button */}
              <div className="pt-2 flex justify-end">
                <button
                  type="button"
                  id="btn-fix-all-contrast"
                  onClick={handleAutoFixAllSlidesContrast}
                  className="text-[10px] font-mono font-bold text-indigo-400 hover:text-indigo-300 flex items-center space-x-1"
                >
                  <Check className="w-3 h-3" />
                  <span>Enforce WCAG AA on All Slides</span>
                </button>
              </div>
            </div>

            {/* ACTION TRIGGERS IN FOOTER */}
            <div className="flex items-center justify-end space-x-3 pt-2 border-t border-neutral-800/60">
              <button
                id="btn-cancel-export-modal"
                onClick={() => setIsExportSettingsOpen(false)}
                className="px-4 py-2 border border-neutral-800 hover:bg-neutral-850 text-neutral-400 hover:text-neutral-200 text-xs font-bold rounded-xl transition cursor-pointer"
              >
                Cancel
              </button>
              <button
                id="btn-confirm-export-modal"
                onClick={() => {
                  setIsExportSettingsOpen(false);
                  if (pendingExportType === "pdf") {
                    handleExportPDF();
                  } else if (pendingExportType === "pptx") {
                    handleExportPPTX();
                  }
                }}
                className="px-5 py-2 bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-bold rounded-xl shadow-lg shadow-indigo-500/20 flex items-center space-x-1.5 transition cursor-pointer"
              >
                {pendingExportType === "pdf" ? <Printer className="w-4 h-4" /> : <Download className="w-4 h-4" />}
                <span>
                  Confirm & Export {pendingExportType === "pdf" ? "PDF (1080p)" : "PPTX (1080p)"}
                </span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 4: SLIDE DECK PAGE COUNT & STRUCTURE CONFIGURATION (PROMPTS BEFORE GENERATING SLIDES) */}
      {isDeckConfigOpen && (
        <div id="modal-deck-config" className="no-print fixed inset-0 z-50 bg-neutral-950/85 backdrop-blur-md flex items-center justify-center p-4 animate-in fade-in duration-200">
          <div className="bg-neutral-900 border border-neutral-800 rounded-3xl max-w-lg w-full p-6 sm:p-7 space-y-6 shadow-2xl animate-in zoom-in-95 duration-200 text-left">
            
            <div className="flex items-center justify-between border-b border-neutral-800/80 pb-4">
              <div className="flex items-center space-x-2.5 text-indigo-400">
                <div className="p-2 bg-indigo-500/10 border border-indigo-500/20 rounded-xl">
                  <Sliders className="w-5 h-5 text-indigo-400" />
                </div>
                <div>
                  <h3 className="font-heading font-extrabold text-base sm:text-lg text-white">Slide Deck Structure</h3>
                  <p className="text-[11px] text-neutral-400">How many pages would you like to generate?</p>
                </div>
              </div>
              <button
                id="btn-close-deck-config"
                onClick={() => setIsDeckConfigOpen(false)}
                className="text-neutral-500 hover:text-white p-1 rounded-lg transition"
              >
                ✕
              </button>
            </div>

            {/* PRESET CHIPS */}
            <div className="space-y-2.5">
              <label className="text-[10px] font-mono text-neutral-400 uppercase font-bold block">Page Count Presets</label>
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                {[
                  { count: 5, label: "5 Slides", desc: "Brief Briefing" },
                  { count: 7, label: "7 Slides", desc: "Pitch (Recommended)" },
                  { count: 10, label: "10 Slides", desc: "Comprehensive" },
                  { count: 14, label: "14 Slides", desc: "Deep Dive" }
                ].map(item => (
                  <button
                    key={item.count}
                    id={`btn-preset-count-${item.count}`}
                    type="button"
                    onClick={() => {
                      setSelectedSlideOption(item.count as 5 | 7 | 10 | 14);
                      setTargetSlideCount(item.count);
                    }}
                    className={`p-3 rounded-xl border text-left transition-all ${
                      selectedSlideOption === item.count
                        ? "bg-indigo-950/40 border-indigo-500 text-white shadow-md shadow-indigo-500/10 ring-1 ring-indigo-500/50"
                        : "bg-neutral-950/50 border-neutral-800 text-neutral-300 hover:border-neutral-700"
                    }`}
                  >
                    <span className="font-bold text-xs block">{item.label}</span>
                    <span className="text-[10px] text-neutral-400 block mt-0.5 leading-tight">{item.desc}</span>
                  </button>
                ))}
              </div>
            </div>

            {/* CUSTOM NUMBER INPUT & SLIDER */}
            <div className="bg-neutral-950 border border-neutral-800/80 p-4 rounded-2xl space-y-3">
              <div className="flex items-center justify-between">
                <span className="text-xs font-semibold text-neutral-300">Custom Slide Target</span>
                <div className="flex items-center space-x-2">
                  <input
                    id="input-target-slide-count"
                    type="number"
                    min={3}
                    max={30}
                    value={targetSlideCount}
                    onChange={(e) => {
                      const val = Math.max(3, Math.min(30, parseInt(e.target.value) || 7));
                      setTargetSlideCount(val);
                      setSelectedSlideOption("custom");
                    }}
                    className="w-14 bg-neutral-900 border border-neutral-750 text-white font-mono font-bold text-center text-xs py-1 rounded-lg focus:outline-none focus:border-indigo-500"
                  />
                  <span className="text-xs text-neutral-400 font-mono">pages</span>
                </div>
              </div>

              <input
                id="range-target-slide-count"
                type="range"
                min={3}
                max={25}
                value={targetSlideCount}
                onChange={(e) => {
                  setTargetSlideCount(parseInt(e.target.value));
                  setSelectedSlideOption("custom");
                }}
                className="w-full h-1.5 bg-neutral-800 rounded-lg appearance-none cursor-pointer accent-indigo-500"
              />
            </div>

            {/* QUALITY ASSURANCE NOTICE */}
            <div className="bg-indigo-950/20 border border-indigo-500/20 p-3 rounded-2xl flex items-start space-x-2.5">
              <Sparkles className="w-4 h-4 text-indigo-400 shrink-0 mt-0.5" />
              <div className="text-[11px] text-indigo-200/90 leading-relaxed">
                Output fixed at <strong>1920 × 1080</strong> Full HD with <strong>WCAG 2.1 contrast verification</strong> and a closing <strong>"Thank you for listening"</strong> slide.
              </div>
            </div>

            {/* MODAL ACTION BUTTONS */}
            <div className="flex flex-col-reverse sm:flex-row items-center justify-between gap-2.5 pt-2 border-t border-neutral-800/80">
              <button
                id="btn-skip-slide-count"
                type="button"
                onClick={() => handleGeneratePresentation(null)}
                className="w-full sm:w-auto px-4 py-2.5 border border-neutral-800 hover:bg-neutral-800 text-neutral-400 hover:text-white text-xs font-semibold rounded-xl transition cursor-pointer"
              >
                Skip (Auto-Detect Slide Count)
              </button>

              <button
                id="btn-confirm-generate-deck"
                type="button"
                onClick={() => handleGeneratePresentation(targetSlideCount)}
                className="w-full sm:w-auto px-6 py-2.5 bg-gradient-to-r from-indigo-600 to-pink-600 hover:opacity-95 text-white text-xs font-bold rounded-xl shadow-lg shadow-indigo-500/20 flex items-center justify-center space-x-1.5 transition cursor-pointer"
              >
                <Sparkles className="w-3.5 h-3.5 text-yellow-300" />
                <span>Generate {targetSlideCount} Slides</span>
              </button>
            </div>

          </div>
        </div>
      )}

      {/* FULL PRESENTATION PAGE COPIES FOR PDF CONTAINER PRINT-ONLY */}
      {presentation && (
        <div id="pdf-print-container" className="hidden print:block absolute inset-0 bg-white">
          {presentation.slides.map((s, index) => {
            const isTitle = s.layout === "title-slide";
            const presetTheme = THEME_PRESETS[presentation.themeId];
            
            const slideBg = s.bgColor || presentation.palette.background;
            const slideText = s.textColor || presentation.palette.text;
            const slideAccent = s.accentColor || presentation.palette.accent;
            const slidePrimary = s.primaryColor || presentation.palette.primary;
            const slideCardBg = s.cardBgColor || presentation.palette.cardBg;
            const slideBorder = s.borderColor || presentation.palette.border;

            const printScale = s.fontSize === "small" ? 0.82 : s.fontSize === "large" ? 1.18 : 1.0;
            const getPrintFontSizeStyle = (baseSizePx: number) => {
              return `${Math.round(baseSizePx * printScale)}px`;
            };

            return (
              <div
                key={s.id}
                className="print-page relative overflow-hidden"
                style={{
                  width: "1920px",
                  height: "1080px",
                  minWidth: "1920px",
                  minHeight: "1080px",
                  maxWidth: "1920px",
                  maxHeight: "1080px",
                  backgroundColor: slideBg,
                  color: slideText,
                  "--bg-color": slideBg,
                  "--text-color": slideText
                } as React.CSSProperties}
              >
                {/* DECORATIVE LIGHT SHAPES FOR AVANTGARDE */}
                {presentation.themeId === "creative-avantgarde" && (
                  <div className="absolute -left-12 -top-12 w-48 h-48 bg-pink-300/10 rounded-full blur-2xl pointer-events-none" />
                )}

                {/* DECORATIVE RADIAL GLOWS FOR COSMIC */}
                {presentation.themeId === "cosmic-slate" && (
                  <div className="absolute right-0 bottom-0 w-80 h-80 bg-indigo-500/5 rounded-full blur-3xl pointer-events-none" />
                )}

                <div className={`absolute inset-0 z-10 ${
                  presentation.themeId === "brutalist-mono" ? "border-[24px] border-black" : ""
                }`}>
                  {/* HEADER - FIXED POSITION */}
                  <div className="absolute top-[80px] left-[80px] right-[80px] h-[50px] flex justify-between items-center text-lg opacity-80 border-b border-neutral-700/10 pb-3 z-20">
                    <span className="uppercase font-mono font-bold tracking-widest text-[#4f46e5]" style={{ color: slideAccent }}>{presentation.title}</span>
                    <span className="font-mono text-base">Slide {index + 1} of {presentation.slides.length}</span>
                  </div>

                  {/* BODY CONTENT AREA - FIXED POSITION WITH SCALED HIGH-FIDELITY TYPOGRAPHY & LAYOUTS */}
                  <div className="absolute top-[170px] bottom-[170px] left-[80px] right-[80px] flex flex-col justify-center overflow-hidden z-20">
                    {isTitle ? (
                      /* LAYOUT 1: TITLE SLIDE (1920x1080) */
                      <div className="text-center space-y-8 py-6">
                        {s.badge && (
                          <div className="mb-2">
                            <span
                              className="text-sm font-mono font-bold uppercase tracking-widest px-4 py-2 rounded-full border shadow-md inline-block"
                              style={{
                                backgroundColor: `${slideAccent}1d`,
                                color: slideAccent,
                                borderColor: `${slideAccent}3a`
                              }}
                            >
                              {s.badge}
                            </span>
                          </div>
                        )}
                        <h1 className={`tracking-tight leading-tight font-black ${presetTheme.fontHeading}`} style={{ color: slidePrimary, fontSize: getPrintFontSizeStyle(72) }}>
                          {presentation.title}
                        </h1>
                        <div className="block mt-4">
                          <h2 className="font-mono font-semibold opacity-90 text-center" style={{ color: slideText, fontSize: getPrintFontSizeStyle(30) }}>
                            {s.title}
                          </h2>
                        </div>
                        {s.content.length > 0 && (
                          <div className="flex flex-wrap items-center justify-center gap-4 pt-6">
                            {s.content.map((point, idx) => (
                              <span
                                key={idx}
                                className="px-4 py-1.5 bg-neutral-500/10 border border-neutral-700/30 rounded-full font-medium"
                                style={{ color: slideAccent, fontSize: getPrintFontSizeStyle(18) }}
                              >
                                {point}
                              </span>
                            ))}
                          </div>
                        )}
                      </div>
                    ) : s.layout === "two-column" ? (
                      /* LAYOUT 2: TWO COLUMN (1920x1080) */
                      <div className="space-y-6 h-full flex flex-col justify-center animate-fade-in">
                        <div>
                          {s.badge && (
                            <div className="mb-2">
                              <span
                                className="text-xs font-mono font-bold uppercase tracking-wider px-3 py-1.5 rounded-full border shadow-sm inline-block"
                                style={{
                                  backgroundColor: `${slideAccent}1d`,
                                  color: slideAccent,
                                  borderColor: `${slideAccent}3a`
                                }}
                              >
                                {s.badge}
                              </span>
                            </div>
                          )}
                          <h2 className={`font-black ${presetTheme.fontHeading} mb-2`} style={{ color: slidePrimary, fontSize: getPrintFontSizeStyle(48) }}>{s.title}</h2>
                        </div>
                        <div className="grid grid-cols-2 gap-12 pt-4">
                          <div className="p-8 rounded-2xl border-2 space-y-4 shadow-md min-h-[360px]" style={{ borderColor: slideBorder, backgroundColor: slideCardBg }}>
                            <span className="h-2 w-8 rounded-full block" style={{ backgroundColor: slideAccent }} />
                            <ul className="space-y-4">
                              {s.content.slice(0, Math.ceil(s.content.length / 2)).map((bullet, idx) => {
                                const isListItem = bullet.trim().startsWith("•") || bullet.trim().startsWith("-") || bullet.trim().startsWith("*") || /^\d+\./.test(bullet.trim());
                                let displayText = bullet;
                                if (isListItem) {
                                  displayText = bullet.replace(/^[•\-\*\s]+/, '').replace(/^\d+\.\s*/, '');
                                }
                                return (
                                  <li key={idx} className={`flex items-start ${isListItem ? "space-x-3" : ""} leading-relaxed ${presetTheme.fontBody}`} style={{ color: slideText, fontSize: getPrintFontSizeStyle(20) }}>
                                    {isListItem && (
                                      <span className="mt-2.5 w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: slidePrimary }} />
                                    )}
                                    <span>{displayText}</span>
                                  </li>
                                );
                              })}
                            </ul>
                          </div>
                          <div className="p-8 rounded-2xl border-2 space-y-4 shadow-md min-h-[360px]" style={{ borderColor: slideBorder, backgroundColor: slideCardBg }}>
                            <span className="h-2 w-8 rounded-full block" style={{ backgroundColor: slideAccent }} />
                            <ul className="space-y-4">
                              {s.content.slice(Math.ceil(s.content.length / 2)).map((bullet, idx) => {
                                const isListItem = bullet.trim().startsWith("•") || bullet.trim().startsWith("-") || bullet.trim().startsWith("*") || /^\d+\./.test(bullet.trim());
                                let displayText = bullet;
                                if (isListItem) {
                                  displayText = bullet.replace(/^[•\-\*\s]+/, '').replace(/^\d+\.\s*/, '');
                                }
                                return (
                                  <li key={idx} className={`flex items-start ${isListItem ? "space-x-3" : ""} leading-relaxed ${presetTheme.fontBody}`} style={{ color: slideText, fontSize: getPrintFontSizeStyle(20) }}>
                                    {isListItem && (
                                      <span className="mt-2.5 w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: slideAccent }} />
                                    )}
                                    <span>{displayText}</span>
                                  </li>
                                );
                              })}
                            </ul>
                          </div>
                        </div>
                      </div>
                    ) : s.layout === "quote-slide" ? (
                      /* LAYOUT 3: QUOTE SLIDE (1920x1080) */
                      <div className="text-center max-w-5xl mx-auto space-y-8">
                        {s.badge && (
                          <div className="mb-2">
                            <span
                              className="text-xs font-mono font-bold uppercase tracking-wider px-3 py-1.5 rounded-full border shadow-sm inline-block"
                              style={{
                                backgroundColor: `${slideAccent}1d`,
                                color: slideAccent,
                                borderColor: `${slideAccent}3a`
                              }}
                            >
                              {s.badge}
                            </span>
                          </div>
                        )}
                        <span className="text-9xl leading-none font-black opacity-30 block" style={{ color: slideAccent }}>“</span>
                        <h2 className="font-black italic tracking-wide leading-relaxed text-center" style={{ color: slideText, fontSize: getPrintFontSizeStyle(40) }}>{s.title}</h2>
                        {s.content.length > 0 && (
                          <div className="block pt-4">
                            <p className="uppercase tracking-widest font-mono font-bold text-center opacity-85" style={{ color: slidePrimary, fontSize: getPrintFontSizeStyle(24) }}>- {s.content.join(" & ")}</p>
                          </div>
                        )}
                      </div>
                    ) : s.layout === "image-left" && s.imageUrl ? (
                      /* LAYOUT 4: IMAGE LEFT (1920x1080) */
                      <div className="grid grid-cols-2 gap-16 items-center">
                        <div className="relative h-[480px] rounded-3xl overflow-hidden shadow-2xl border-2 border-neutral-700/10">
                          <img src={s.imageUrl} className="absolute inset-0 w-full h-full object-cover rounded-3xl" referrerPolicy="no-referrer" onError={handleImageError} />
                          <div className="absolute inset-0 bg-neutral-950/20" />
                          {s.imageCaption && (
                            <div className="absolute bottom-4 left-4 bg-neutral-950/80 text-xs px-3 py-1.5 rounded-lg text-neutral-300">
                              {s.imageCaption}
                            </div>
                          )}
                        </div>
                        <div className="space-y-6">
                          <div>
                            {s.badge && (
                              <div className="mb-2">
                                <span
                                  className="text-xs font-mono font-bold uppercase tracking-wider px-3 py-1.5 rounded-full border shadow-sm inline-block"
                                  style={{
                                    backgroundColor: `${slideAccent}1d`,
                                    color: slideAccent,
                                    borderColor: `${slideAccent}3a`
                                  }}
                                >
                                  {s.badge}
                                </span>
                              </div>
                            )}
                            <h2 className={`font-black leading-tight ${presetTheme.fontHeading} mb-4`} style={{ color: slidePrimary, fontSize: getPrintFontSizeStyle(48) }}>{s.title}</h2>
                          </div>
                          <ul className="space-y-5">
                            {s.content.map((bullet, idx) => {
                              const isListItem = bullet.trim().startsWith("•") || bullet.trim().startsWith("-") || bullet.trim().startsWith("*") || /^\d+\./.test(bullet.trim());
                              let displayText = bullet;
                              if (isListItem) {
                                displayText = bullet.replace(/^[•\-\*\s]+/, '').replace(/^\d+\.\s*/, '');
                              }
                              return (
                                <li key={idx} className={`flex items-start ${isListItem ? "space-x-3" : ""} leading-relaxed ${presetTheme.fontBody}`} style={{ color: slideText, fontSize: getPrintFontSizeStyle(20) }}>
                                  {isListItem && (
                                    <span className="mt-2.5 w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: slideAccent }} />
                                  )}
                                  <span>{displayText}</span>
                                </li>
                              );
                            })}
                          </ul>
                        </div>
                      </div>
                    ) : s.layout === "image-right" && s.imageUrl ? (
                      /* LAYOUT 5: IMAGE RIGHT (1920x1080) */
                      <div className="grid grid-cols-2 gap-16 items-center">
                        <div className="space-y-6">
                          <div>
                            {s.badge && (
                              <div className="mb-2">
                                <span
                                  className="text-xs font-mono font-bold uppercase tracking-wider px-3 py-1.5 rounded-full border shadow-sm inline-block"
                                  style={{
                                    backgroundColor: `${slideAccent}1d`,
                                    color: slideAccent,
                                    borderColor: `${slideAccent}3a`
                                  }}
                                >
                                  {s.badge}
                                </span>
                              </div>
                            )}
                            <h2 className={`font-black leading-tight ${presetTheme.fontHeading} mb-4`} style={{ color: slidePrimary, fontSize: getPrintFontSizeStyle(48) }}>{s.title}</h2>
                          </div>
                          <ul className="space-y-5">
                            {s.content.map((bullet, idx) => {
                              const isListItem = bullet.trim().startsWith("•") || bullet.trim().startsWith("-") || bullet.trim().startsWith("*") || /^\d+\./.test(bullet.trim());
                              let displayText = bullet;
                              if (isListItem) {
                                displayText = bullet.replace(/^[•\-\*\s]+/, '').replace(/^\d+\.\s*/, '');
                              }
                              return (
                                <li key={idx} className={`flex items-start ${isListItem ? "space-x-3" : ""} leading-relaxed ${presetTheme.fontBody}`} style={{ color: slideText, fontSize: getPrintFontSizeStyle(20) }}>
                                  {isListItem && (
                                    <span className="mt-2.5 w-2.5 h-2.5 rounded-full shrink-0" style={{ backgroundColor: slideAccent }} />
                                  )}
                                  <span>{displayText}</span>
                                </li>
                              );
                            })}
                          </ul>
                        </div>
                        <div className="relative h-[480px] rounded-3xl overflow-hidden shadow-2xl border-2 border-neutral-700/10">
                          <img src={s.imageUrl} className="absolute inset-0 w-full h-full object-cover rounded-3xl" referrerPolicy="no-referrer" onError={handleImageError} />
                          <div className="absolute inset-0 bg-neutral-950/20" />
                          {s.imageCaption && (
                            <div className="absolute bottom-4 left-4 bg-neutral-950/80 text-xs px-3 py-1.5 rounded-lg text-neutral-300">
                              {s.imageCaption}
                            </div>
                          )}
                        </div>
                      </div>
                    ) : s.layout === "stats-bento" ? (
                      /* LAYOUT 6: STATS BENTO (1920x1080) */
                      <div className="space-y-6 h-full flex flex-col justify-center">
                        <div>
                          {s.badge && (
                            <div className="mb-2">
                              <span
                                className="text-xs font-mono font-bold uppercase tracking-wider px-3 py-1.5 rounded-full border shadow-sm inline-block"
                                style={{
                                  backgroundColor: `${slideAccent}1d`,
                                  color: slideAccent,
                                  borderColor: `${slideAccent}3a`
                                }}
                              >
                                {s.badge}
                              </span>
                            </div>
                          )}
                          <h2 className={`font-black leading-tight ${presetTheme.fontHeading} mb-4`} style={{ color: slidePrimary, fontSize: getPrintFontSizeStyle(48) }}>{s.title}</h2>
                        </div>
                        <div className="grid grid-cols-2 gap-8 pt-4">
                          {[0, 1, 2, 3].map((val) => {
                            const textVal = s.content[val] || "N/A metric value";
                            const spaceIdx = textVal.indexOf(" ");
                            let metricNum = textVal;
                            let metricLabel = "";
                            if (spaceIdx > 0) {
                              metricNum = textVal.substring(0, spaceIdx);
                              metricLabel = textVal.substring(spaceIdx + 1);
                            }

                            const numericStr = metricNum.replace(/[^0-9.]/g, '');
                            const numericVal = numericStr ? parseFloat(numericStr) : null;
                            const isPercent = metricNum.includes("%");

                            return (
                              <div
                                key={val}
                                className="p-8 rounded-2xl border-2 shadow-md flex flex-col justify-between h-[210px] transition-all"
                                style={{
                                  backgroundColor: slideCardBg,
                                  border: `2px solid ${slideBorder}`
                                }}
                              >
                                <div className="space-y-1.5">
                                  <span className="font-heading font-black block" style={{ color: slideAccent, fontSize: getPrintFontSizeStyle(42) }}>
                                    {metricNum}
                                  </span>
                                  <span className="font-mono font-bold block leading-normal text-neutral-400 uppercase tracking-widest" style={{ fontSize: getPrintFontSizeStyle(13) }}>
                                    {metricLabel || "Strategic target metrics"}
                                  </span>
                                </div>

                                {/* PRINTABLE BAR CHART OR DIAGRAM WIDGET */}
                                <div className="w-full pt-4 mt-auto">
                                  {isPercent && numericVal !== null ? (
                                    <div className="space-y-1.5">
                                      <div className="flex justify-between items-center text-[10px] font-mono opacity-60">
                                        <span>Progress Level</span>
                                        <span>{Math.min(100, Math.round(numericVal))}%</span>
                                      </div>
                                      <div className="w-full h-3 bg-neutral-900/60 rounded-full overflow-hidden p-[2px] border border-neutral-850">
                                        <div 
                                          className="h-full rounded-full transition-all duration-1000"
                                          style={{
                                            width: `${Math.min(100, Math.max(8, numericVal))}%`,
                                            background: `linear-gradient(to right, ${slidePrimary}, ${slideAccent})`,
                                            boxShadow: `0 0 6px ${slideAccent}40`
                                          }}
                                        />
                                      </div>
                                    </div>
                                  ) : numericVal !== null ? (
                                    <div className="space-y-1.5">
                                      <div className="flex justify-between items-center text-[10px] font-mono opacity-60">
                                        <span>Relative Performance Graph</span>
                                        <span>Scale {metricNum}</span>
                                      </div>
                                      <div className="h-10 flex items-end justify-between gap-[3px] w-full pt-1">
                                        {[25, 40, 55, 30, 65, 80, 50, 70, 95, 100].map((hValue, barIdx) => {
                                          const scaleFactor = Math.min(1.2, Math.max(0.3, numericVal / 100));
                                          const calcHeight = Math.min(100, Math.max(15, hValue * scaleFactor));
                                          return (
                                            <div 
                                              key={barIdx}
                                              className="w-full rounded-t-sm transition-all"
                                              style={{
                                                height: `${calcHeight}%`,
                                                backgroundColor: barIdx >= 8 ? slideAccent : `${slidePrimary}35`,
                                              }}
                                            />
                                          );
                                        })}
                                      </div>
                                    </div>
                                  ) : (
                                    <div className="space-y-1.5">
                                      <div className="flex justify-between items-center text-[10px] font-mono opacity-60">
                                        <span>Target Trajectory</span>
                                        <span>Strategic Target</span>
                                      </div>
                                      <div className="h-8 w-full flex items-center justify-center">
                                        <svg className="w-full h-full" viewBox="0 0 100 20" preserveAspectRatio="none">
                                          <path 
                                            d="M 0 15 Q 15 5, 30 14 T 60 4 T 90 12 L 100 8" 
                                            fill="none" 
                                            stroke={slideAccent} 
                                            strokeWidth="3" 
                                            strokeLinecap="round"
                                            className="opacity-95"
                                          />
                                          <path 
                                            d="M 0 15 Q 15 5, 30 14 T 60 4 T 90 12 L 100 8 L 100 20 L 0 20 Z" 
                                            fill={`url(#ambient-grad-print-${val})`}
                                            className="opacity-25"
                                          />
                                          <defs>
                                            <linearGradient id={`ambient-grad-print-${val}`} x1="0" y1="0" x2="0" y2="1">
                                              <stop offset="0%" stopColor={slideAccent} stopOpacity="0.4" />
                                              <stop offset="100%" stopColor="transparent" stopOpacity="0" />
                                            </linearGradient>
                                          </defs>
                                        </svg>
                                      </div>
                                    </div>
                                  )}
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    ) : s.layout === "comparison-table" ? (
                      /* LAYOUT 8: COMPARISON TABLE (1920x1080) */
                      <div className="space-y-8 h-full flex flex-col justify-center">
                        <div>
                          {s.badge && (
                            <div className="mb-2">
                              <span
                                className="text-xs font-mono font-bold uppercase tracking-wider px-3 py-1.5 rounded-full border shadow-sm inline-block"
                                style={{
                                  backgroundColor: `${slideAccent}1d`,
                                  color: slideAccent,
                                  borderColor: `${slideAccent}3a`
                                }}
                              >
                                {s.badge}
                              </span>
                            </div>
                          )}
                          <h2 className={`font-black leading-tight ${presetTheme.fontHeading} mb-4`} style={{ color: slidePrimary, fontSize: getPrintFontSizeStyle(44) }}>{s.title}</h2>
                        </div>

                        <div className="overflow-hidden rounded-3xl border-2 shadow-2xl" style={{ borderColor: slideBorder, backgroundColor: slideCardBg }}>
                          <table className="w-full text-left border-collapse table-fixed">
                            <thead>
                              <tr style={{ borderBottom: `3px solid ${slideBorder}`, backgroundColor: `${slidePrimary}0d` }}>
                                <th className="px-8 py-5 w-1/4 text-xs uppercase tracking-widest font-mono opacity-70" style={{ color: slideText, fontSize: getPrintFontSizeStyle(14) }}>Evaluation Dimension</th>
                                <th className="px-8 py-5 w-3/8 font-bold" style={{ color: slidePrimary, fontSize: getPrintFontSizeStyle(22) }}>Standard Practices</th>
                                <th className="px-8 py-5 w-3/8 font-bold flex items-center gap-3" style={{ color: slideAccent, fontSize: getPrintFontSizeStyle(22) }}>
                                  Slidesss Advantage
                                  <span className="text-xs font-mono uppercase bg-emerald-500/10 text-emerald-400 border border-emerald-400/20 px-3 py-1 rounded">Optimized Solution</span>
                                </th>
                              </tr>
                            </thead>
                            <tbody>
                              {(() => {
                                const rows: { aspect: string; left: string; right: string }[] = [];
                                const contentLen = s.content.length;
                                
                                let hasSeparator = false;
                                s.content.forEach(bullet => {
                                  if (bullet.includes(" vs ") || bullet.includes(" | ") || bullet.includes(" - ")) {
                                    hasSeparator = true;
                                  }
                                });

                                if (hasSeparator) {
                                  s.content.forEach((bullet, idx) => {
                                    let aspect = `Dimension ${idx + 1}`;
                                    let left = bullet;
                                    let right = "";
                                    const separators = [" vs ", " | ", " - "];
                                    for (const sep of separators) {
                                      if (bullet.includes(sep)) {
                                        const parts = bullet.split(sep);
                                        left = parts[0].trim();
                                        right = parts.slice(1).join(sep).trim();
                                        if (left.includes(": ")) {
                                          const aspParts = left.split(": ");
                                          aspect = aspParts[0].trim();
                                          left = aspParts.slice(1).join(": ").trim();
                                        }
                                        break;
                                      }
                                    }
                                    rows.push({ aspect, left, right });
                                  });
                                } else {
                                  const half = Math.ceil(contentLen / 2);
                                  for (let i = 0; i < half; i++) {
                                    const leftBullet = s.content[i] || "";
                                    const rightBullet = s.content[i + half] || "";
                                    let aspect = `Dimension ${i + 1}`;
                                    let leftText = leftBullet;
                                    let rightText = rightBullet;
                                    if (leftBullet.includes(": ")) {
                                      const parts = leftBullet.split(": ");
                                      aspect = parts[0].trim();
                                      leftText = parts.slice(1).join(": ").trim();
                                    }
                                    if (rightBullet.includes(": ")) {
                                      const parts = rightBullet.split(": ");
                                      rightText = parts.slice(1).join(": ").trim();
                                    }
                                    rows.push({ aspect, left: leftText, right: rightText });
                                  }
                                }

                                return rows.map((row, rIdx) => (
                                  <tr key={rIdx} className="hover:bg-neutral-900/10 transition-colors" style={{ borderBottom: rIdx < rows.length - 1 ? `2px solid ${slideBorder}` : 'none' }}>
                                    <td className="px-8 py-5 font-mono opacity-80 break-words" style={{ color: slideText, fontSize: getPrintFontSizeStyle(15) }}>{row.aspect}</td>
                                    <td className="px-8 py-5 opacity-85 break-words" style={{ color: slideText, fontSize: getPrintFontSizeStyle(18) }}>
                                      <div className="flex items-start gap-2">
                                        <span className="text-red-500 shrink-0 font-bold mt-1">✗</span>
                                        <span>{row.left || "—"}</span>
                                      </div>
                                    </td>
                                    <td className="px-8 py-5 font-semibold break-words" style={{ color: slidePrimary, fontSize: getPrintFontSizeStyle(18) }}>
                                      <div className="flex items-start gap-2">
                                        <span className="text-emerald-500 shrink-0 font-bold mt-1">✓</span>
                                        <span>{row.right || "—"}</span>
                                      </div>
                                    </td>
                                  </tr>
                                ));
                              })()}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    ) : (
                      /* LAYOUT 7: DEFAULT / HEADLINE BULLET (1920x1080) */
                      <div className="space-y-6">
                        <div>
                          {s.badge && (
                            <div className="mb-2">
                              <span
                                className="text-xs font-mono font-bold uppercase tracking-wider px-3 py-1.5 rounded-full border shadow-sm inline-block"
                                style={{
                                  backgroundColor: `${slideAccent}1d`,
                                  color: slideAccent,
                                  borderColor: `${slideAccent}3a`
                                }}
                              >
                                  {s.badge}
                              </span>
                            </div>
                          )}
                          <h2 className={`font-black leading-tight ${presetTheme.fontHeading} mb-4`} style={{ color: slidePrimary, fontSize: getPrintFontSizeStyle(48) }}>{s.title}</h2>
                        </div>
                        <ul className="space-y-5 pt-4">
                          {s.content.map((bullet, idx) => {
                            const isListItem = bullet.trim().startsWith("•") || bullet.trim().startsWith("-") || bullet.trim().startsWith("*") || /^\d+\./.test(bullet.trim());
                            let displayText = bullet;
                            if (isListItem) {
                              displayText = bullet.replace(/^[•\-\*\s]+/, '').replace(/^\d+\.\s*/, '');
                            }
                            return (
                              <li key={idx} className={`flex items-start ${isListItem ? "space-x-3" : ""} leading-relaxed ${presetTheme.fontBody}`} style={{ color: slideText, fontSize: getPrintFontSizeStyle(24) }}>
                                {isListItem && (
                                  <span className="mt-3 w-3.5 h-3.5 rounded-full shrink-0 animate-pulse" style={{ backgroundColor: slideAccent }} />
                                )}
                                <span>{displayText}</span>
                              </li>
                            );
                          })}
                        </ul>
                      </div>
                    )}
                  </div>

                  {/* FOOTER - FIXED POSITION */}
                  <div className="absolute bottom-[80px] left-[80px] right-[80px] h-[50px] flex justify-between items-center text-sm opacity-60 border-t border-neutral-700/15 pt-3 z-20">
                    <span style={{ color: slideText }}>Excelsior Systems © 2026</span>
                    <span className="font-mono uppercase tracking-widest text-xs">PREPARATION LICENSE KEY: SINGLE-USER CONVERTER</span>
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      )}

    </div>
  );
}
