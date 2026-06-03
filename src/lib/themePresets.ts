/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { ThemeId, ThemeConfig, ColorPalette } from "../types";

// Helper to validate and normalize HEX colors
export function normalizeHex(color: string, defaultColor = "#0066FF"): string {
  if (!color) return defaultColor;
  let clean = color.trim().replace(/^#/, "");
  if (clean.length === 3) {
    clean = clean.split("").map(c => c + c).join("");
  }
  if (clean.length !== 6) {
    return defaultColor;
  }
  return `#${clean}`;
}

// Convert a hex color with alpha
export function hexToRgbA(hex: string, alpha: number): string {
  try {
    const clean = normalizeHex(hex).replace("#", "");
    const r = parseInt(clean.substring(0, 2), 16);
    const g = parseInt(clean.substring(2, 4), 16);
    const b = parseInt(clean.substring(4, 6), 16);
    return `rgba(${r}, ${g}, ${b}, ${alpha})`;
  } catch (e) {
    return hex;
  }
}

export const THEME_PRESETS: Record<ThemeId, ThemeConfig> = {
  [ThemeId.COSMIC_SLATE]: {
    id: ThemeId.COSMIC_SLATE,
    name: "Cosmic Slate",
    description: "Immersive deep indigo canvas, modern space typography, and high-visibility neon details.",
    fontHeading: "font-sans tracking-tight font-extrabold",
    fontBody: "font-sans leading-relaxed text-sm",
    defaultPalette: (brandColor: string) => {
      const primary = normalizeHex(brandColor, "#6366F1");
      return {
        primary,
        secondary: "#818CF8",
        accent: "#EC4899", // Neon pink
        background: "#0B0F19", // Dark slate space
        text: "#F1F5F9", // Bright white
        textLight: "#94A3B8", // Muted purple-grey
        cardBg: "#161E2E", // Glassmorphism container
        border: "#1F2937",
      };
    }
  },
  [ThemeId.MINIMAL_SWISS]: {
    id: ThemeId.MINIMAL_SWISS,
    name: "Swiss Modern",
    description: "Highly structured minimalism, clean neo-grotesque type scales, and high-impact branding blocks.",
    fontHeading: "font-sans tracking-tighter font-black uppercase text-4xl",
    fontBody: "font-sans leading-normal text-slate-700 text-sm",
    defaultPalette: (brandColor: string) => {
      const primary = normalizeHex(brandColor, "#000000"); // Standard black brand color
      return {
        primary,
        secondary: "#334155",
        accent: "#EF4444", // Bright Swiss red
        background: "#FFFFFF", // Crisp white
        text: "#0F172A", // Deep slate
        textLight: "#64748B",
        cardBg: "#F8FAFC",
        border: "#E2E8F0",
      };
    }
  },
  [ThemeId.WARM_EDITORIAL]: {
    id: ThemeId.WARM_EDITORIAL,
    name: "Warm Editorial",
    description: "Refined literary warmth, graceful serifs, off-white vanilla page backgrounds, and delicate accents.",
    fontHeading: "font-serif tracking-tight font-semibold italic text-3xl",
    fontBody: "font-sans leading-relaxed text-stone-850 text-sm",
    defaultPalette: (brandColor: string) => {
      const primary = normalizeHex(brandColor, "#78350F");
      return {
        primary,
        secondary: "#451A03",
        accent: "#C2410C", // Ochre orange
        background: "#FCF9F2", // Fine book cream
        text: "#1C1917", // Stone charcoal
        textLight: "#78716C",
        cardBg: "#F5EFEB",
        border: "#E7E0D4",
      };
    }
  },
  [ThemeId.BRUTALIST_MONO]: {
    id: ThemeId.BRUTALIST_MONO,
    name: "Brutalist Mono",
    description: "Industrial grid structures, high contrast code text, heavy solid lines, and neon hazard alerts.",
    fontHeading: "font-mono font-bold uppercase tracking-tight",
    fontBody: "font-mono leading-relaxed text-xs",
    defaultPalette: (brandColor: string) => {
      const primary = normalizeHex(brandColor, "#F59E0B"); // Bright yellow block
      return {
        primary,
        secondary: "#000000",
        accent: primary,
        background: "#F3F4F6", // Concrete gray
        text: "#000000",
        textLight: "#4B5563",
        cardBg: "#FFFFFF",
        border: "#000000", // Stark thick borders
      };
    }
  },
  [ThemeId.CREATIVE_AVANTGARDE]: {
    id: ThemeId.CREATIVE_AVANTGARDE,
    name: "Avant-Garde",
    description: "Asymmetrical artistic layouts, playful pastel backdrops, organic shapes, and deep plum text tones.",
    fontHeading: "font-sans tracking-wide font-light text-3xl",
    fontBody: "font-sans leading-relaxed text-slate-800 text-sm",
    defaultPalette: (brandColor: string) => {
      const primary = normalizeHex(brandColor, "#EC4899");
      return {
        primary,
        secondary: "#5B21B6",
        accent: "#3B82F6",
        background: "#FAF5FF", // Pale lavender
        text: "#1E1B4B", // Midnight plum
        textLight: "#6E6B7B",
        cardBg: "#FFFFFF",
        border: "#F3E8FF",
      };
    }
  }
};

// Curated collections under Unsplash categories
export interface CuratedImage {
  id: string;
  url: string;
  thumbnail: string;
  title: string;
  author: string;
}

export const CURATED_UNSPLASH_IMAGES: Record<string, CuratedImage[]> = {
  "Business & Leadership": [
    {
      id: "lU6X7n5s-F4",
      url: "https://images.unsplash.com/photo-1519085360753-af0119f7cbe7?auto=format&fit=crop&q=80&w=1200",
      thumbnail: "https://images.unsplash.com/photo-1519085360753-af0119f7cbe7?auto=format&fit=crop&q=80&w=300",
      title: "Business leader outlining strategy",
      author: "Amy Hirschi"
    },
    {
      id: "MYu38MpfT_4",
      url: "https://images.unsplash.com/photo-1556761175-b413da4baf72?auto=format&fit=crop&q=80&w=1200",
      thumbnail: "https://images.unsplash.com/photo-1556761175-b413da4baf72?auto=format&fit=crop&q=80&w=300",
      title: "Collaborative design studio",
      author: "Campaign Creators"
    },
    {
      id: "gM3Y8yi2usY",
      url: "https://images.unsplash.com/photo-1486406146926-c627a92ad1ab?auto=format&fit=crop&q=80&w=1200",
      thumbnail: "https://images.unsplash.com/photo-1486406146926-c627a92ad1ab?auto=format&fit=crop&q=80&w=300",
      title: "Modern corporate glass headquarters",
      author: "Sean Pollock"
    },
    {
      id: "38MpfT_something",
      url: "https://images.unsplash.com/photo-1454165804606-c3d57bc86b40?auto=format&fit=crop&q=80&w=1200",
      thumbnail: "https://images.unsplash.com/photo-1454165804606-c3d57bc86b40?auto=format&fit=crop&q=80&w=300",
      title: "Corporate strategy mapping",
      author: "Scott Graham"
    }
  ],
  "Technology & Innovation": [
    {
      id: "Y98YyisS",
      url: "https://images.unsplash.com/photo-1451187580459-43490279c0fa?auto=format&fit=crop&q=80&w=1200",
      thumbnail: "https://images.unsplash.com/photo-1451187580459-43490279c0fa?auto=format&fit=crop&q=80&w=300",
      title: "Deep network digital hub",
      author: "Skyline Visuals"
    },
    {
      id: "W8YyisS",
      url: "https://images.unsplash.com/photo-1526374965328-7f61d4dc18c5?auto=format&fit=crop&q=80&w=1200",
      thumbnail: "https://images.unsplash.com/photo-1526374965328-7f61d4dc18c5?auto=format&fit=crop&q=80&w=300",
      title: "Server terminal data code",
      author: "Markus Spiske"
    },
    {
      id: "Z783-s",
      url: "https://images.unsplash.com/photo-1518770660439-4636190af475?auto=format&fit=crop&q=80&w=1200",
      thumbnail: "https://images.unsplash.com/photo-1518770660439-4636190af475?auto=format&fit=crop&q=80&w=300",
      title: "Microprocessor circuitry close-up",
      author: "Alexandre Debiève"
    }
  ],
  "Finance & Growth": [
    {
      id: "F93-8A",
      url: "https://images.unsplash.com/photo-1460925895917-afdab827c52f?auto=format&fit=crop&q=80&w=1200",
      thumbnail: "https://images.unsplash.com/photo-1460925895917-afdab827c52f?auto=format&fit=crop&q=80&w=300",
      title: "Analytics growth and graphs",
      author: "Lukas Blazek"
    },
    {
      id: "F93-8B",
      url: "https://images.unsplash.com/photo-1559526324-4b87b5e36e44?auto=format&fit=crop&q=80&w=1200",
      thumbnail: "https://images.unsplash.com/photo-1559526324-4b87b5e36e44?auto=format&fit=crop&q=80&w=300",
      title: "Financial overview notebook charts",
      author: "Carlos Muza"
    },
    {
      id: "F93-8C",
      url: "https://images.unsplash.com/photo-1590283603385-17ffb3a7f29f?auto=format&fit=crop&q=80&w=1200",
      thumbnail: "https://images.unsplash.com/photo-1590283603385-17ffb3a7f29f?auto=format&fit=crop&q=80&w=300",
      title: "Digital finance stock candle charts",
      author: "Maxim Hopman"
    }
  ],
  "Office & Collaboration": [
    {
      id: "Co1",
      url: "https://images.unsplash.com/photo-1522071820081-009f0129c71c?auto=format&fit=crop&q=80&w=1200",
      thumbnail: "https://images.unsplash.com/photo-1522071820081-009f0129c71c?auto=format&fit=crop&q=80&w=300",
      title: "Active team creative debate",
      author: "Jason Goodman"
    },
    {
      id: "Co2",
      url: "https://images.unsplash.com/photo-1531538606174-0f90ff5dce83?auto=format&fit=crop&q=80&w=1200",
      thumbnail: "https://images.unsplash.com/photo-1531538606174-0f90ff5dce83?auto=format&fit=crop&q=80&w=300",
      title: "Brainstorm session with note tags",
      author: "Kaleidico"
    },
    {
      id: "Co3",
      url: "https://images.unsplash.com/photo-1497366216548-37526070297c?auto=format&fit=crop&q=80&w=1200",
      thumbnail: "https://images.unsplash.com/photo-1497366216548-37526070297c?auto=format&fit=crop&q=80&w=300",
      title: "Modern spacious work space outline",
      author: "Hugo Jehanne"
    }
  ],
  "Creative & Abstract": [
    {
      id: "Cr1",
      url: "https://images.unsplash.com/photo-1541701494587-cb58502866ab?auto=format&fit=crop&q=80&w=1200",
      thumbnail: "https://images.unsplash.com/photo-1541701494587-cb58502866ab?auto=format&fit=crop&q=80&w=300",
      title: "Dynamic paint explosion design",
      author: "Joel Filipe"
    },
    {
      id: "Cr2",
      url: "https://images.unsplash.com/photo-1507525428034-b723cf961d3e?auto=format&fit=crop&q=80&w=1200",
      thumbnail: "https://images.unsplash.com/photo-1507525428034-b723cf961d3e?auto=format&fit=crop&q=80&w=300",
      title: "Serene wave shapes sunset",
      author: "Sean Oulashin"
    },
    {
      id: "Cr3",
      url: "https://images.unsplash.com/photo-1550684848-fac1c5b4e853?auto=format&fit=crop&q=80&w=1200",
      thumbnail: "https://images.unsplash.com/photo-1550684848-fac1c5b4e853?auto=format&fit=crop&q=80&w=300",
      title: "Symmetric dark waves graphic flow",
      author: "Sven Brandsma"
    }
  ]
};

// Return fallback image url when user enters any keyword (uses deterministic high-fidelity Picsum seeds)
export function getKeywordImage(query: string): string {
  const cleaned = (query || "").toLowerCase().trim();
  
  // Choose standard seed integers to retrieve gorgeous deterministic photos from Picsum
  let hash = 0;
  for (let i = 0; i < cleaned.length; i++) {
    hash = (hash << 5) - hash + cleaned.charCodeAt(i);
    hash |= 0;
  }
  const seedNum = (Math.abs(hash) % 250) + 1;
  
  // Picsum is guaranteed to render, fast, and works seamlessly in previews without CORS or hotlink limits
  return `https://picsum.photos/seed/slidecraft-${seedNum}/1200/800`;
}
