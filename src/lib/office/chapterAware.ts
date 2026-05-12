import type { ApplyTarget, PresetStyleKey, SkripsiPresetV1 } from "@/types/preset";
import { applyStyleToParagraph } from "@/lib/office/formatter";

const FIGURE_CAPTION_PATTERN = /^(figure|gambar)\s+([0-9]+|[ivxlcdm]+)([.:-]|\s)/i;
const TABLE_CAPTION_PATTERN = /^(table|tabel)\s+([0-9]+|[ivxlcdm]+)([.:-]|\s)/i;
const CHAPTER_HEADING_PATTERN = /^(bab\s+[ivxlcdm]+|chapter\s+\d+)\b/i;
const HEADING3_PATTERN = /^\d+\.\d+\.\d+\s+/;
const HEADING2_PATTERN = /^\d+\.\d+\s+/;
const HEADING2_ALT_PATTERN = /^\d+\.\s+/;

function normalizeText(text: string): string {
  return text.replace(/\s+/g, " ").trim();
}

function isLikelyShoutedHeading(text: string): boolean {
  if (!text || text.length > 80) {
    return false;
  }

  const letters = text.replace(/[^A-Za-z]/g, "");
  if (letters.length < 4) {
    return false;
  }

  return letters === letters.toUpperCase();
}

function mapBuiltInStyle(styleBuiltIn: string): PresetStyleKey | null {
  if (styleBuiltIn === "Heading1") {
    return "heading1";
  }
  if (styleBuiltIn === "Heading2") {
    return "heading2";
  }
  if (styleBuiltIn === "Heading3") {
    return "heading3";
  }
  if (styleBuiltIn === "Quote" || styleBuiltIn === "IntenseQuote") {
    return "quote";
  }
  return null;
}

export function classifyParagraphStyleKey(text: string, styleBuiltIn: string): PresetStyleKey {
  const normalized = normalizeText(text);
  if (!normalized) {
    return "body";
  }

  const byStyle = mapBuiltInStyle(styleBuiltIn);
  if (byStyle) {
    return byStyle;
  }

  if (FIGURE_CAPTION_PATTERN.test(normalized)) {
    return "captionFigure";
  }
  if (TABLE_CAPTION_PATTERN.test(normalized)) {
    return "captionTable";
  }

  if (CHAPTER_HEADING_PATTERN.test(normalized) || isLikelyShoutedHeading(normalized)) {
    return "heading1";
  }
  if (HEADING3_PATTERN.test(normalized)) {
    return "heading3";
  }
  if (HEADING2_PATTERN.test(normalized) || HEADING2_ALT_PATTERN.test(normalized)) {
    return "heading2";
  }

  if (/^>\s+/.test(normalized) || /^".+"$/.test(normalized)) {
    return "quote";
  }

  return "body";
}

export type ChapterAwareSummary = {
  total: number;
  body: number;
  heading1: number;
  heading2: number;
  heading3: number;
  quote: number;
  captionFigure: number;
  captionTable: number;
};

export async function applyChapterAwareFormatting(
  preset: SkripsiPresetV1,
  target: ApplyTarget
): Promise<ChapterAwareSummary> {
  return Word.run(async (context) => {
    const paragraphs =
      target === "selection" ? context.document.getSelection().paragraphs : context.document.body.paragraphs;

    paragraphs.load("items/text,items/styleBuiltIn");
    await context.sync();

    const summary: ChapterAwareSummary = {
      total: paragraphs.items.length,
      body: 0,
      heading1: 0,
      heading2: 0,
      heading3: 0,
      quote: 0,
      captionFigure: 0,
      captionTable: 0,
    };

    for (const paragraph of paragraphs.items) {
      const styleKey = classifyParagraphStyleKey(paragraph.text, String(paragraph.styleBuiltIn || ""));
      applyStyleToParagraph(paragraph, preset.styles[styleKey], styleKey);
      summary[styleKey] += 1;
    }

    await context.sync();
    return summary;
  });
}
