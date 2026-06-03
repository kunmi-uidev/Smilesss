/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef } from "react";
import mammoth from "mammoth";
import {
  Upload, Sparkles, Wand2, Download, RefreshCw, Plus, Trash2, ArrowUp, ArrowDown,
  Image as ImageIcon, Layout, Palette, Search, FileText, Check, HelpCircle, ChevronRight,
  Monitor, Play, Printer, X, Sliders, ExternalLink
} from "lucide-react";
import { ThemeId, ColorPalette, SlideLayout, Slide, Presentation } from "./types";
import { THEME_PRESETS, CURATED_UNSPLASH_IMAGES, getKeywordImage, normalizeHex, hexToRgbA } from "./lib/themePresets";
import { exportToPPTX } from "./lib/pptxExport";

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
    "Selecting semantic professional slide layouts...",
    "Aligning color spaces to your brand guidelines...",
    "Formulating concise high-impact summaries...",
    "Synthesizing Unsplash metadata queries...",
    "Constructing designer presentation system..."
  ];

  // Auto-generate dynamic colors palette dynamically on brandColor change
  useEffect(() => {
    if (presentation) {
      const generatedPalette = THEME_PRESETS[presentation.themeId].defaultPalette(brandColor);
      setPresentation(prev => prev ? {
        ...prev,
        brandColor,
        palette: generatedPalette
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
    
    // Construct premium presentation slides
    const sampleSlides: Slide[] = [
      {
        id: `slide-demo-1-${Date.now()}`,
        title: "EXCELSIOR GLOBAL SYSTEMS Q3 MARKETING STRATEGY",
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
        title: "CORPORATE FOCUS DIRECTIVE",
        content: [
          "\"Authentic visual intelligence is no longer a corporate luxury—it is the single highest leverage vector for corporate clarity and authority.\""
        ],
        layout: "quote-slide",
        notes: "Read this quote with high conviction. Visual intelligence is our brand's authority signature."
      },
      {
        id: `slide-demo-5-${Date.now()}`,
        title: "EXECUTION TIMELINE",
        content: [
          "PHASE I (JUNE): Finalize full-stack Gemini JSON validation schemas with exponential backoff handlers.",
          "PHASE II (JULY): Ship public pptxgenjs design pipeline & visual theme editors.",
          "PHASE III (AUGUST): Release premium curated visual searching tools.",
          "PHASE IV (SEPTEMBER): Announce general public market launch."
        ],
        layout: "headline-bullet",
        notes: "Walk the board through our rapid execution deliverables. All items are on track for Q3 delivery."
      }
    ];

    // Inject Unsplash images based on standard queries
    const slidesWithImages = sampleSlides.map(slide => ({
      ...slide,
      imageUrl: getKeywordImage(slide.title.toLowerCase() || "business strategy")
    }));

    const demoPresentation: Presentation = {
      title: "Excelsior Q3 Strategy Deck",
      themeId: selectedThemeId,
      brandColor: brandColor,
      palette: THEME_PRESETS[selectedThemeId].defaultPalette(brandColor),
      slides: slidesWithImages
    };

    setPresentation(demoPresentation);
    setActiveSlideIndex(0);
    setErrorMessage(null);
    setActiveStep("editor");
  };

  // Submit parsed text to backend to build full presentation slides using Gemini
  const handleGeneratePresentation = async () => {
    if (!rawText.trim()) {
      setErrorMessage("Please select a Word document to upload or load our professional sample text outline.");
      return;
    }

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
        const msgIndex = Math.floor(next / 15) % loadingMessages.length;
        setStatusMessage(loadingMessages[msgIndex]);
        return next;
      });
    }, 1200);

    try {
      const res = await fetch("/api/gemini/generate-presentation", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          docText: rawText,
          brandColor: brandColor,
          themeId: selectedThemeId,
          docName: docName
        })
      });

      if (!res.ok) {
        const errPayload = await res.json();
        throw new Error(errPayload.error || "Generation endpoint refused to formulate slide layouts.");
      }

      const deck = await res.json();
      
      // Inject unique slide IDs and build standard initial presentation state
      const initialSlides = deck.slides.map((slide: any, i: number) => ({
        ...slide,
        id: `slide-${Date.now()}-${i}-${Math.random().toString(36).substr(2, 5)}`,
        imageUrl: getKeywordImage(slide.imageSearchQuery || "business teamwork outline")
      }));

      const activeThemeId = (deck.themeId as ThemeId) || selectedThemeId;

      const loadedPresentation: Presentation = {
        title: deck.title || "Untitled Presentation",
        themeId: activeThemeId,
        brandColor: brandColor,
        palette: THEME_PRESETS[activeThemeId].defaultPalette(brandColor),
        slides: initialSlides
      };

      setPresentation(loadedPresentation);
      setSelectedThemeId(activeThemeId);
      setActiveSlideIndex(0);
      setGenerationProgress(100);
      setStatusMessage("Pristine slide deck successfully crafted!");

      // Small pause for visual feedback
      setTimeout(() => {
        setActiveStep("editor");
        setIsGenerating(false);
      }, 800);

    } catch (err: any) {
      console.error(err);
      setErrorMessage(err.message || "PowerPoint deck creation failed. Please verify your internet connection or verify your Gemini credentials.");
      setIsGenerating(false);
    } finally {
      clearInterval(progressTimer);
    }
  };

  // Apply another global layout theme
  const handleSwitchTheme = (themeId: ThemeId) => {
    if (presentation) {
      const updatedPalette = THEME_PRESETS[themeId].defaultPalette(brandColor);
      setPresentation({
        ...presentation,
        themeId,
        palette: updatedPalette
      });
    }
    setSelectedThemeId(themeId);
  };

  // Slide navigation actions
  const activeSlide = presentation?.slides[activeSlideIndex] || null;

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
  const triggerUnsplashSearch = () => {
    setIsSearchingUnsplash(true);
    // Mimic real-time API filtering from curated assets fallback
    setTimeout(() => {
      const term = unsplashSearch.trim().toLowerCase();
      let results: any[] = [];
      if (term) {
        // Build dynamic high relevance Unsplash result cards based on search input
        results = [
          {
            id: "res-1",
            url: `https://images.unsplash.com/photo-1454165804606-c3d57bc86b40?auto=format&fit=crop&q=80&w=1200&sig=1`,
            thumbnail: `https://images.unsplash.com/photo-1454165804606-c3d57bc86b40?auto=format&fit=crop&q=80&w=300&sig=1`,
            title: `Marketing for '${term}'`,
            author: "Corporate Photographer"
          },
          {
            id: "res-2",
            url: `https://images.unsplash.com/photo-1551836022-d5d88e9218df?auto=format&fit=crop&q=80&w=1200&sig=2`,
            thumbnail: `https://images.unsplash.com/photo-1551836022-d5d88e9218df?auto=format&fit=crop&q=80&w=300&sig=2`,
            title: `'${term}' Workplace Collaboration`,
            author: "Studio Focus"
          },
          {
            id: "res-3",
            url: `https://images.unsplash.com/photo-1507537297725-24a1c029d3ca?auto=format&fit=crop&q=80&w=1200&sig=3`,
            thumbnail: `https://images.unsplash.com/photo-1507537297725-24a1c029d3ca?auto=format&fit=crop&q=80&w=300&sig=3`,
            title: `${term} Digital Strategy`,
            author: "Abstract Designer"
          },
          {
            id: "res-4",
            url: `https://images.unsplash.com/photo-1531403009284-440f080d1e12?auto=format&fit=crop&q=80&w=1200&sig=4`,
            thumbnail: `https://images.unsplash.com/photo-1531403009284-440f080d1e12?auto=format&fit=crop&q=80&w=300&sig=4`,
            title: `${term} Flowchart & Metrics`,
            author: "Analytics Core"
          }
        ];
      } else {
        results = CURATED_UNSPLASH_IMAGES[unsplashCategory] || [];
      }
      setUnsplashResults(results);
      setIsSearchingUnsplash(false);
    }, 500);
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

      if (!res.ok) {
        throw new Error("Design Suggestion server endpoint refused proposal.");
      }

      const suggestion = await res.json();
      setAiSuggestion(suggestion);
      setShowSuggestionModal(true);
    } catch (err: any) {
      console.error(err);
      setErrorMessage("Could not generate design suggest. Ensure server.ts API routes are responsive.");
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
      imageSearchQuery: aiSuggestion.imageSearchQuery
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
                width: 297mm;
                height: 210mm;
                max-width: 100%;
                max-height: 100%;
                page-break-after: always;
                break-after: page;
                box-sizing: border-box;
                position: relative;
                display: flex;
                flex-direction: column;
                justify-content: space-between;
                overflow: hidden;
              }
              @page {
                size: A4 landscape;
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
                  width: 297mm;
                  height: 210mm;
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

      {/* ERROR MESSAGE TOAST */}
      {errorMessage && (
        <div id="error-banner" className="no-print bg-red-950/60 border-b border-red-900/50 px-6 py-3 text-red-200 text-xs flex items-center justify-between">
          <span className="font-mono">{errorMessage}</span>
          <button onClick={() => setErrorMessage(null)} className="hover:text-white font-bold p-1">✕</button>
        </div>
      )}

      {/* MAIN SCREEN WRAP */}
      <main className="no-print flex-1 flex flex-col">
        {activeStep === "upload" ? (
          
          /* VIEW 1: LANDING & INITIAL CONVERSION STEP (CENTERED UX) */
          <div id="view-upload-sandbox" className="flex-1 max-w-3xl w-full mx-auto px-4 py-8 md:py-12 flex flex-col items-center justify-start space-y-8">
            
            {/* GRADIENT CONTAINER ENVELOPING THE BODY SECTION */}
            <div className="w-[90%] md:w-full bg-gradient-to-r from-indigo-400 via-pink-400 to-amber-400 p-[2px] rounded-[40px] shadow-2xl">
              <div className="bg-neutral-950 rounded-[38px] p-6 sm:p-10 space-y-8">
                
                {/* HERO INTRODUCTION ZONE */}
                <div className="text-center space-y-4 max-w-2xl px-2 sm:px-0 mx-auto">
                  <span className="px-3.5 py-1 bg-indigo-500/10 text-indigo-400 border border-indigo-500/20 text-xs font-mono rounded-full font-bold inline-block mx-auto">
                    No sign-up required.
                  </span>
                  <h2 className="font-heading font-black text-[32px] sm:text-[48px] leading-tight tracking-tight text-white max-w-xl mx-auto line-clamp-3 sm:line-clamp-none">
                    Instantly rewrite wordy files into <span className="text-transparent bg-gradient-to-r from-indigo-400 via-pink-400 to-amber-400 bg-clip-text">designer slides.</span>
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
                onClick={handleGeneratePresentation}
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
                    <span>Generate PDF/PPTX</span>
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
        </div>

            {/* DEPLOYMENT LINE FOOTER */}
            <div className="text-neutral-500 text-[10px] font-mono pt-2 flex items-center space-x-2 justify-center pb-24 sm:pb-0">
              <span className="inline-block w-2 h-2 bg-emerald-500 rounded-full animate-ping" />
              <span>Express Server Engine Ready  |  Gemini Core v3.5-flash Online</span>
            </div>

            {/* MOBILE STICKY FLOATING CTA COMPONENT */}
            <div className="block sm:hidden fixed bottom-[40px] left-4 right-4 z-50">
              <button
                id="btn-generate-deck-mobile"
                disabled={!rawText.trim() || isGenerating}
                onClick={handleGeneratePresentation}
                className={`w-full py-3.5 px-6 rounded-2xl font-bold text-sm tracking-wide flex items-center justify-center space-x-2 shadow-2xl border border-neutral-700/40 backdrop-blur-md transition-all ${
                  isGenerating || !rawText.trim()
                    ? "bg-neutral-900 border-neutral-800 text-neutral-500 cursor-not-allowed"
                    : "bg-gradient-to-r from-indigo-600 via-pink-600 to-indigo-600 text-white hover:opacity-95 shadow-indigo-600/40 cursor-pointer"
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
                    <span>Generate PDF/PPTX</span>
                  </>
                )}
              </button>
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
                        backgroundColor: presentation?.palette.background,
                        color: presentation?.palette.text,
                        "--bg-color": presentation?.palette.background,
                        "--text-color": presentation?.palette.text
                      } as React.CSSProperties}
                    >
                      
                      {/* ASYMMETRICAL DECORATIVE GRID LINES FOR BRUTALIST & SWISS DESIGN */}
                      {presentation?.themeId === "brutalist-mono" && (
                        <div className="absolute inset-0 pointer-events-none opacity-10 border-b border-r" style={{ borderColor: presentation.palette.text, backgroundSize: "40px 40px", backgroundImage: "linear-gradient(to right, gray 1px, transparent 1px), linear-gradient(to bottom, gray 1px, transparent 1px)" }} />
                      )}

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
                            style={{ color: presentation?.palette.accent }}
                          >
                            Excelsior Q3 Strategy Deck
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
                              <h3
                                className={`text-4xl lg:text-5xl tracking-tight leading-tight ${activePresetTheme.fontHeading}`}
                                style={{ color: presentation?.palette.primary }}
                              >
                                {presentation?.title || "BUSINESS PRESENTATION"}
                              </h3>
                              <p
                                className="text-lg font-mono font-medium max-w-2xl mx-auto"
                                style={{ color: presentation?.palette.text }}
                              >
                                {activeSlide.title}
                              </p>
                              {activeSlide.content.length > 0 && (
                                <div className="flex flex-wrap items-center justify-center gap-3 pt-4">
                                  {activeSlide.content.map((point, idx) => (
                                    <span
                                      key={idx}
                                      className="text-xs px-3 py-1 bg-neutral-500/5 border border-neutral-750/30 rounded-full font-medium"
                                      style={{ color: presentation?.palette.accent }}
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
                              <h3 className={`text-2xl font-bold ${activePresetTheme.fontHeading}`} style={{ color: presentation?.palette.primary }}>
                                {activeSlide.title}
                              </h3>
                              <div className="grid grid-cols-1 md:grid-cols-2 gap-6 pt-2">
                                <div className="p-4 rounded-xl space-y-3" style={{ backgroundColor: presentation?.palette.cardBg, border: `1px solid ${presentation?.palette.border}` }}>
                                  <span className="h-1.5 w-6 rounded-full block" style={{ backgroundColor: presentation?.palette.accent }} />
                                  <ul className="space-y-2.5">
                                    {activeSlide.content.slice(0, Math.ceil(activeSlide.content.length / 2)).map((bullet, idx) => (
                                      <li key={idx} className={`flex items-start space-x-2 ${activePresetTheme.fontBody}`} style={{ color: presentation?.palette.text }}>
                                        <span className="mt-1.5 w-1.5 h-1.5 rounded-full shrink-0" style={{ backgroundColor: presentation?.palette.primary }} />
                                        <span>{bullet}</span>
                                      </li>
                                    ))}
                                  </ul>
                                </div>
                                <div className="p-4 rounded-xl space-y-3" style={{ backgroundColor: presentation?.palette.cardBg, border: `1px solid ${presentation?.palette.border}` }}>
                                  <span className="h-1.5 w-6 rounded-full block" style={{ backgroundColor: presentation?.palette.primary }} />
                                  <ul className="space-y-2.5">
                                    {activeSlide.content.slice(Math.ceil(activeSlide.content.length / 2)).map((bullet, idx) => (
                                      <li key={idx} className={`flex items-start space-x-2 ${activePresetTheme.fontBody}`} style={{ color: presentation?.palette.text }}>
                                        <span className="mt-1.5 w-1.5 h-1.5 rounded-full shrink-0" style={{ backgroundColor: presentation?.palette.accent }} />
                                        <span>{bullet}</span>
                                      </li>
                                    ))}
                                  </ul>
                                </div>
                              </div>
                            </div>

                          ) : activeSlide.layout === "quote-slide" ? (
                            
                            /* LAYOUT 3: STYLIZED QUOTE */
                            <div className="text-center max-w-4xl mx-auto space-y-4 py-4">
                              <span className="text-5xl leading-none font-black opacity-30 block" style={{ color: presentation?.palette.accent }}>“</span>
                              <h4
                                className="text-xl lg:text-2xl font-black italic tracking-wide"
                                style={{ color: presentation?.palette.text }}
                              >
                                {activeSlide.title}
                              </h4>
                              {activeSlide.content.length > 0 && (
                                <p className="text-xs uppercase tracking-widest font-mono font-bold" style={{ color: presentation?.palette.primary }}>
                                  - {activeSlide.content.join(" & ")}
                                </p>
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
                                <h3 className={`text-2xl font-bold ${activePresetTheme.fontHeading}`} style={{ color: presentation?.palette.primary }}>
                                  {activeSlide.title}
                                </h3>
                                <ul className="space-y-3">
                                  {activeSlide.content.map((bullet, idx) => (
                                    <li key={idx} className={`flex items-start space-x-2.5 ${activePresetTheme.fontBody}`} style={{ color: presentation?.palette.text }}>
                                      <span className="mt-1.5 w-1.5 h-1.5 rounded-full shrink-0" style={{ backgroundColor: presentation?.palette.accent }} />
                                      <span>{bullet}</span>
                                    </li>
                                  ))}
                                </ul>
                              </div>
                            </div>

                          ) : activeSlide.layout === "image-right" && activeSlide.imageUrl ? (
                            
                            /* LAYOUT 5: VISUAL RIGHT SPLIT */
                            <div className="grid grid-cols-1 md:grid-cols-2 gap-8 items-center h-full">
                              <div className="space-y-4">
                                <h3 className={`text-2xl font-bold ${activePresetTheme.fontHeading}`} style={{ color: presentation?.palette.primary }}>
                                  {activeSlide.title}
                                </h3>
                                <ul className="space-y-3">
                                  {activeSlide.content.map((bullet, idx) => (
                                    <li key={idx} className={`flex items-start space-x-2.5 ${activePresetTheme.fontBody}`} style={{ color: presentation?.palette.text }}>
                                      <span className="mt-1.5 w-1.5 h-1.5 rounded-full shrink-0" style={{ backgroundColor: presentation?.palette.accent }} />
                                      <span>{bullet}</span>
                                    </li>
                                  ))}
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
                              <h3 className={`text-xl font-bold ${activePresetTheme.fontHeading}`} style={{ color: presentation?.palette.primary }}>
                                {activeSlide.title}
                              </h3>
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

                                  return (
                                    <div
                                      key={val}
                                      className="p-4 rounded-xl space-y-1 transition-all"
                                      style={{
                                        backgroundColor: presentation?.palette.cardBg,
                                        border: `1px solid ${presentation?.palette.border}`
                                      }}
                                    >
                                      <span className="text-2xl font-heading font-black block" style={{ color: presentation?.palette.accent }}>
                                        {metricNum}
                                      </span>
                                      <span className="text-[10px] font-mono font-medium block leading-tight text-neutral-400 uppercase">
                                        {metricLabel || "Strategic target metrics"}
                                      </span>
                                    </div>
                                  );
                                })}
                              </div>
                            </div>

                          ) : (
                            
                            /* LAYOUT 7: DEFAULT / HEADLINE BULLET */
                            <div className="space-y-4">
                              <h3 className={`text-2xl font-bold ${activePresetTheme.fontHeading}`} style={{ color: presentation?.palette.primary }}>
                                {activeSlide.title}
                              </h3>
                              <ul className="space-y-3.5">
                                {activeSlide.content.map((bullet, idx) => (
                                  <li key={idx} className={`flex items-start space-x-2.5 ${activePresetTheme.fontBody}`} style={{ color: presentation?.palette.text }}>
                                    <span className="mt-1.5 w-1.5 h-1.5 rounded-full shrink-0 animate-pulse" style={{ backgroundColor: presentation?.palette.accent }} />
                                    <span>{bullet}</span>
                                  </li>
                                ))}
                              </ul>
                            </div>
                          )}
                        </div>

                        {/* SLIDE FOOTER BRAND LINE */}
                        <div className="flex items-center justify-between mt-4 text-[9px] border-t border-neutral-700/10 pt-2 opacity-60">
                          <span style={{ color: presentation?.palette.text }}>Excelsior Systems © 2026</span>
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
                            { id: "stats-bento", label: "Stats Bento" }
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
                <h3 className="font-heading font-extrabold text-lg text-white">Brand Color Layout Alignment</h3>
              </div>
              <button
                id="btn-close-export-settings"
                onClick={() => setIsExportSettingsOpen(false)}
                className="text-neutral-500 hover:text-white text-sm"
              >
                ✕
              </button>
            </div>

            <p className="text-xs text-neutral-400 leading-relaxed">
              Verify your company's primary color guidelines below. The generated slide deck elements of <span className="text-white font-bold">{THEME_PRESETS[presentation.themeId].name}</span> theme will instantly skin to form uniform brand presentation consistency.
            </p>

            <div className="bg-neutral-950 border border-neutral-800 p-4 rounded-2xl space-y-4">
              <div className="flex items-center space-x-4">
                <div className="relative">
                  <input
                    id="input-hex-picker-modal"
                    type="color"
                    value={brandColor}
                    onChange={(e) => setBrandColor(e.target.value)}
                    className="w-14 h-14 rounded-xl cursor-pointer border border-neutral-750 bg-transparent p-1"
                  />
                </div>
                <div className="flex-1 space-y-1">
                  <label className="text-xs font-mono text-neutral-400 block uppercase">Brand HEX Color Code</label>
                  <input
                    id="input-hex-text-modal"
                    type="text"
                    value={brandColor}
                    onChange={(e) => setBrandColor(normalizeHex(e.target.value, brandColor))}
                    placeholder="#4F46E5"
                    className="bg-neutral-900 border border-neutral-800 px-3 py-1.5 text-sm rounded-lg text-white font-mono w-full focus:outline-none focus:border-indigo-500"
                  />
                </div>
              </div>

              {/* REAL TIME PALETTE PROJECTION PREVIEW */}
              <div className="space-y-1.5 pt-1">
                <span className="text-[10px] font-mono font-bold text-neutral-400 block uppercase">Real-Time Palette Projection</span>
                <div className="grid grid-cols-5 gap-1.5 p-1 bg-neutral-950 rounded-xl border border-neutral-850">
                  <div className="text-center p-1.5 rounded-lg border border-neutral-900" style={{ backgroundColor: presentation.palette.primary }}>
                    <span className="text-[8px] font-mono font-bold px-1" style={{ color: presentation.palette.text }}>Brand</span>
                  </div>
                  <div className="text-center p-1.5 rounded-lg border border-neutral-900" style={{ backgroundColor: presentation.palette.secondary }}>
                    <span className="text-[8px] font-mono font-bold px-1" style={{ color: presentation.palette.background }}>Sec</span>
                  </div>
                  <div className="text-center p-1.5 rounded-lg border border-neutral-900" style={{ backgroundColor: presentation.palette.accent }}>
                    <span className="text-[8px] font-mono font-bold px-1" style={{ color: presentation.palette.text }}>Acc</span>
                  </div>
                  <div className="text-center p-1.5 rounded-lg border border-neutral-900" style={{ backgroundColor: presentation.palette.cardBg }}>
                    <span className="text-[8px] font-mono font-bold px-1" style={{ color: presentation.palette.text }}>Card</span>
                  </div>
                  <div className="text-center p-1.5 rounded-lg border border-neutral-900" style={{ backgroundColor: presentation.palette.background }}>
                    <span className="text-[8px] font-mono font-bold px-1" style={{ color: presentation.palette.text }}>Canvas</span>
                  </div>
                </div>
              </div>
            </div>

            {/* ACTION TRIGGERS IN FOOTER */}
            <div className="flex items-center justify-end space-x-3 pt-3 border-t border-neutral-800/60">
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
                  Confirm & Export {pendingExportType === "pdf" ? "PDF" : "PPTX"}
                </span>
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
            return (
              <div
                key={s.id}
                className="print-page w-[297mm] h-[210mm] relative p-16 border-b border-neutral-100 flex flex-col justify-between overflow-hidden"
                style={{
                  backgroundColor: presentation.palette.background,
                  color: presentation.palette.text,
                  "--bg-color": presentation.palette.background,
                  "--text-color": presentation.palette.text
                } as React.CSSProperties}
              >
                {/* ASYMMETRICAL DECORATIVE GRID LINES FOR BRUTALIST & SWISS DESIGN */}
                {presentation.themeId === "brutalist-mono" && (
                  <div className="absolute inset-0 pointer-events-none opacity-10 border-b border-r" style={{ borderColor: presentation.palette.text, backgroundSize: "40px 40px", backgroundImage: "linear-gradient(to right, gray 1px, transparent 1px), linear-gradient(to bottom, gray 1px, transparent 1px)" }} />
                )}

                {/* DECORATIVE LIGHT SHAPES FOR AVANTGARDE */}
                {presentation.themeId === "creative-avantgarde" && (
                  <div className="absolute -left-12 -top-12 w-48 h-48 bg-pink-300/10 rounded-full blur-2xl pointer-events-none" />
                )}

                {/* DECORATIVE RADIAL GLOWS FOR COSMIC */}
                {presentation.themeId === "cosmic-slate" && (
                  <div className="absolute right-0 bottom-0 w-80 h-80 bg-indigo-500/5 rounded-full blur-3xl pointer-events-none" />
                )}

                <div className={`w-full h-full flex flex-col justify-between relative z-10 ${
                  presentation.themeId === "brutalist-mono" ? "border-4 border-black p-8" : ""
                }`}>
                  {/* Header */}
                  <div className="flex justify-between items-center text-xs opacity-60 border-b border-neutral-700/10 pb-2">
                    <span className="uppercase font-mono font-bold tracking-widest" style={{ color: presentation.palette.accent }}>{presentation.title}</span>
                    <span className="font-mono">Slide {index + 1} of {presentation.slides.length}</span>
                  </div>

                  {/* Body Content */}
                  <div className="my-auto flex flex-col justify-center py-4">
                    {isTitle ? (
                      <div className="text-center space-y-6">
                        <h1 className={`text-4xl lg:text-5xl tracking-tight leading-tight ${presetTheme.fontHeading}`} style={{ color: presentation.palette.primary }}>
                          {presentation.title}
                        </h1>
                        <h2 className="text-xl font-mono font-semibold opacity-90">
                          {s.title}
                        </h2>
                        {s.content.length > 0 && (
                          <div className="flex flex-wrap items-center justify-center gap-3 pt-4">
                            {s.content.map((point, idx) => (
                              <span
                                key={idx}
                                className="text-xs px-3 py-1 bg-neutral-500/10 border border-neutral-700/30 rounded-full font-medium"
                                style={{ color: presentation.palette.accent }}
                              >
                                {point}
                              </span>
                            ))}
                          </div>
                        )}
                      </div>
                    ) : s.layout === "two-column" ? (
                      <div className="space-y-6">
                        <h2 className={`text-2xl font-bold ${presetTheme.fontHeading}`} style={{ color: presentation.palette.primary }}>{s.title}</h2>
                        <div className="grid grid-cols-2 gap-8 pt-2">
                          <div className="p-5 rounded-xl border space-y-3" style={{ borderColor: presentation.palette.border, backgroundColor: presentation.palette.cardBg }}>
                            <span className="h-1.5 w-6 rounded-full block" style={{ backgroundColor: presentation.palette.accent }} />
                            <ul className="space-y-2.5">
                              {s.content.slice(0, Math.ceil(s.content.length / 2)).map((bullet, idx) => (
                                <li key={idx} className={`flex items-start space-x-2 ${presetTheme.fontBody}`} style={{ color: presentation.palette.text }}>
                                  <span className="mt-1.5 w-1.5 h-1.5 rounded-full shrink-0" style={{ backgroundColor: presentation.palette.primary }} />
                                  <span>{bullet}</span>
                                </li>
                              ))}
                            </ul>
                          </div>
                          <div className="p-5 rounded-xl border space-y-3" style={{ borderColor: presentation.palette.border, backgroundColor: presentation.palette.cardBg }}>
                            <span className="h-1.5 w-6 rounded-full block" style={{ backgroundColor: presentation.palette.primary }} />
                            <ul className="space-y-2.5">
                              {s.content.slice(Math.ceil(s.content.length / 2)).map((bullet, idx) => (
                                <li key={idx} className={`flex items-start space-x-2 ${presetTheme.fontBody}`} style={{ color: presentation.palette.text }}>
                                  <span className="mt-1.5 w-1.5 h-1.5 rounded-full shrink-0" style={{ backgroundColor: presentation.palette.accent }} />
                                  <span>{bullet}</span>
                                </li>
                              ))}
                            </ul>
                          </div>
                        </div>
                      </div>
                    ) : s.layout === "quote-slide" ? (
                      <div className="text-center max-w-2xl mx-auto space-y-4">
                        <span className="text-6xl leading-none font-black opacity-30 block" style={{ color: presentation.palette.accent }}>“</span>
                        <h2 className="text-xl lg:text-2xl font-black italic tracking-wide" style={{ color: presentation.palette.text }}>{s.title}</h2>
                        {s.content.length > 0 && (
                          <p className="text-xs uppercase tracking-widest font-mono font-bold" style={{ color: presentation.palette.primary }}>- {s.content.join(" & ")}</p>
                        )}
                      </div>
                    ) : s.layout === "image-left" && s.imageUrl ? (
                      <div className="grid grid-cols-2 gap-8 items-center">
                        <div className="relative h-64 rounded-xl overflow-hidden shadow">
                          <img src={s.imageUrl} className="absolute inset-0 w-full h-full object-cover rounded-xl" referrerPolicy="no-referrer" onError={handleImageError} />
                          <div className="absolute inset-0 bg-neutral-950/20" />
                          {s.imageCaption && (
                            <div className="absolute bottom-2 left-2 bg-neutral-950/80 text-[9px] px-2 py-0.5 rounded text-neutral-300">
                              {s.imageCaption}
                            </div>
                          )}
                        </div>
                        <div className="space-y-4">
                          <h2 className={`text-2xl font-bold ${presetTheme.fontHeading}`} style={{ color: presentation.palette.primary }}>{s.title}</h2>
                          <ul className="space-y-2.5">
                            {s.content.map((bullet, idx) => (
                              <li key={idx} className={`flex items-start space-x-2.5 ${presetTheme.fontBody}`} style={{ color: presentation.palette.text }}>
                                <span className="mt-1.5 w-1.5 h-1.5 rounded-full shrink-0" style={{ backgroundColor: presentation.palette.accent }} />
                                <span>{bullet}</span>
                              </li>
                            ))}
                          </ul>
                        </div>
                      </div>
                    ) : s.layout === "image-right" && s.imageUrl ? (
                      <div className="grid grid-cols-2 gap-8 items-center">
                        <div className="space-y-4">
                          <h2 className={`text-2xl font-bold ${presetTheme.fontHeading}`} style={{ color: presentation.palette.primary }}>{s.title}</h2>
                          <ul className="space-y-2.5">
                            {s.content.map((bullet, idx) => (
                              <li key={idx} className={`flex items-start space-x-2.5 ${presetTheme.fontBody}`} style={{ color: presentation.palette.text }}>
                                <span className="mt-1.5 w-1.5 h-1.5 rounded-full shrink-0" style={{ backgroundColor: presentation.palette.accent }} />
                                <span>{bullet}</span>
                              </li>
                            ))}
                          </ul>
                        </div>
                        <div className="relative h-64 rounded-xl overflow-hidden shadow">
                          <img src={s.imageUrl} className="absolute inset-0 w-full h-full object-cover rounded-xl" referrerPolicy="no-referrer" onError={handleImageError} />
                          <div className="absolute inset-0 bg-neutral-950/20" />
                          {s.imageCaption && (
                            <div className="absolute bottom-2 left-2 bg-neutral-950/80 text-[9px] px-2 py-0.5 rounded text-neutral-300">
                              {s.imageCaption}
                            </div>
                          )}
                        </div>
                      </div>
                    ) : s.layout === "stats-bento" ? (
                      <div className="space-y-4">
                        <h2 className={`text-xl font-bold ${presetTheme.fontHeading}`} style={{ color: presentation.palette.primary }}>{s.title}</h2>
                        <div className="grid grid-cols-2 gap-4">
                          {[0, 1, 2, 3].map((val) => {
                            const textVal = s.content[val] || "N/A metric value";
                            const spaceIdx = textVal.indexOf(" ");
                            let metricNum = textVal;
                            let metricLabel = "";
                            if (spaceIdx > 0) {
                              metricNum = textVal.substring(0, spaceIdx);
                              metricLabel = textVal.substring(spaceIdx + 1);
                            }
                            return (
                              <div
                                key={val}
                                className="p-4 rounded-xl space-y-1"
                                style={{
                                  backgroundColor: presentation.palette.cardBg,
                                  border: `1px solid ${presentation.palette.border}`
                                }}
                              >
                                <span className="text-2xl font-heading font-black block" style={{ color: presentation.palette.accent }}>
                                  {metricNum}
                                </span>
                                <span className="text-[10px] font-mono font-medium block leading-tight text-neutral-400 uppercase">
                                  {metricLabel || "Strategic target metrics"}
                                </span>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    ) : (
                      <div className="space-y-4">
                        <h2 className={`text-2xl font-bold ${presetTheme.fontHeading}`} style={{ color: presentation.palette.primary }}>{s.title}</h2>
                        <ul className="space-y-3.5">
                          {s.content.map((bullet, idx) => (
                            <li key={idx} className={`flex items-start space-x-2.5 ${presetTheme.fontBody}`} style={{ color: presentation.palette.text }}>
                              <span className="mt-1.5 w-1.5 h-1.5 rounded-full shrink-0" style={{ backgroundColor: presentation.palette.accent }} />
                              <span>{bullet}</span>
                            </li>
                          ))}
                        </ul>
                      </div>
                    )}
                  </div>

                  {/* Footer */}
                  <div className="flex justify-between items-center text-[10px] opacity-50 border-t pt-4">
                    <span>Excelsior Systems © 2026</span>
                    <span>PREPARATION LICENSE KEY: SINGLE-USER CONVERTER</span>
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
