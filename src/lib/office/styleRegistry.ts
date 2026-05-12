import type { PresetStyleKey, SkripsiPresetV1, StylePreset } from "@/types/preset";
import { cmToPoints } from "@/lib/utils/units";

type StyleBinding = {
  builtIn: Word.Paragraph["styleBuiltIn"];
  fallbackName: string;
  styleNameCandidates: string[];
};

const STYLE_BINDINGS: Record<PresetStyleKey, StyleBinding> = {
  body: {
    builtIn: "Normal",
    fallbackName: "Normal",
    styleNameCandidates: ["Normal"],
  },
  heading1: {
    builtIn: "Heading1",
    fallbackName: "Heading 1",
    styleNameCandidates: ["Heading 1", "Heading1"],
  },
  heading2: {
    builtIn: "Heading2",
    fallbackName: "Heading 2",
    styleNameCandidates: ["Heading 2", "Heading2"],
  },
  heading3: {
    builtIn: "Heading3",
    fallbackName: "Heading 3",
    styleNameCandidates: ["Heading 3", "Heading3"],
  },
  quote: {
    builtIn: "Quote",
    fallbackName: "Quote",
    styleNameCandidates: ["Quote", "Intense Quote", "IntenseQuote"],
  },
  captionFigure: {
    builtIn: "Caption",
    fallbackName: "Caption",
    styleNameCandidates: ["Caption"],
  },
  captionTable: {
    builtIn: "Caption",
    fallbackName: "Caption",
    styleNameCandidates: ["Caption"],
  },
};

const STYLE_SYNC_ORDER: PresetStyleKey[] = [
  "body",
  "heading1",
  "heading2",
  "heading3",
  "quote",
  // Word has one built-in Caption style shared by figure and table captions.
  "captionFigure",
];

function finiteNumber(value: number, fallback: number): number {
  return Number.isFinite(value) ? value : fallback;
}

function sanitizeStyle(style: StylePreset): StylePreset {
  return {
    text: {
      ...style.text,
      fontSizePt: finiteNumber(style.text.fontSizePt, 12),
    },
    paragraph: {
      ...style.paragraph,
      lineSpacingPt: finiteNumber(style.paragraph.lineSpacingPt, 24),
      spaceBeforePt: finiteNumber(style.paragraph.spaceBeforePt, 0),
      spaceAfterPt: finiteNumber(style.paragraph.spaceAfterPt, 0),
      firstLineIndentCm: finiteNumber(style.paragraph.firstLineIndentCm, 0),
      leftIndentCm: finiteNumber(style.paragraph.leftIndentCm, 0),
      rightIndentCm: finiteNumber(style.paragraph.rightIndentCm, 0),
    },
  };
}

function supportsAllCaps(): boolean {
  try {
    if (typeof Office === "undefined") {
      return false;
    }
    return Boolean(
      Office.context?.requirements?.isSetSupported &&
        Office.context.requirements.isSetSupported("WordApiDesktop", "1.3")
    );
  } catch {
    return false;
  }
}

async function resolveStyleByCandidates(
  context: Word.RequestContext,
  styles: Word.StyleCollection,
  candidates: string[]
): Promise<Word.Style | null> {
  for (const candidate of candidates) {
    const style = styles.getByNameOrNullObject(candidate);
    style.load("isNullObject");
    await context.sync();
    if (!style.isNullObject) {
      return style;
    }
  }
  return null;
}

function applyPresetToStyleObject(style: Word.Style, preset: StylePreset): void {
  const safeStyle = sanitizeStyle(preset);

  style.font.name = safeStyle.text.fontName;
  style.font.size = safeStyle.text.fontSizePt;
  style.font.bold = safeStyle.text.bold;
  style.font.italic = safeStyle.text.italic;
  style.font.underline = safeStyle.text.underline;

  if (supportsAllCaps()) {
    try {
      style.font.allCaps = safeStyle.text.allCaps;
    } catch {
      // no-op
    }
  }

  style.paragraphFormat.alignment = safeStyle.paragraph.alignment;
  style.paragraphFormat.lineSpacing = safeStyle.paragraph.lineSpacingPt;
  style.paragraphFormat.spaceBefore = safeStyle.paragraph.spaceBeforePt;
  style.paragraphFormat.spaceAfter = safeStyle.paragraph.spaceAfterPt;
  style.paragraphFormat.firstLineIndent = cmToPoints(safeStyle.paragraph.firstLineIndentCm);
  style.paragraphFormat.leftIndent = cmToPoints(safeStyle.paragraph.leftIndentCm);
  style.paragraphFormat.rightIndent = cmToPoints(safeStyle.paragraph.rightIndentCm);
}

export function getBuiltInStyleLabel(styleKey: PresetStyleKey): string {
  return STYLE_BINDINGS[styleKey].fallbackName;
}

export function applyBuiltInStyleForPreset(
  paragraph: Word.Paragraph,
  styleKey: PresetStyleKey
): void {
  const binding = STYLE_BINDINGS[styleKey];
  try {
    paragraph.styleBuiltIn = binding.builtIn;
  } catch {
    paragraph.style = binding.fallbackName;
  }
}

export async function syncPresetToWordBuiltInStyles(
  preset: SkripsiPresetV1
): Promise<{
  synced: number;
  skippedStyleKeys: PresetStyleKey[];
}> {
  return Word.run(async (context) => {
    const styles = context.document.getStyles();
    const skippedStyleKeys: PresetStyleKey[] = [];
    let synced = 0;

    for (const styleKey of STYLE_SYNC_ORDER) {
      const binding = STYLE_BINDINGS[styleKey];
      const styleObject = await resolveStyleByCandidates(
        context,
        styles,
        binding.styleNameCandidates
      );

      if (!styleObject) {
        skippedStyleKeys.push(styleKey);
        continue;
      }

      applyPresetToStyleObject(styleObject, preset.styles[styleKey]);
      synced += 1;
    }

    if (synced > 0) {
      await context.sync();
    }

    return { synced, skippedStyleKeys };
  });
}
