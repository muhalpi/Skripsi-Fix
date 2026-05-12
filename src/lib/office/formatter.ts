import type { ApplyTarget, PresetStyleKey, StylePreset } from "@/types/preset";
import { cmToPoints } from "@/lib/utils/units";
import { applyBuiltInStyleForPreset } from "@/lib/office/styleRegistry";
import {
  buildTextPreview,
  clearLastOfficeDiagnostics,
  extractOfficeErrorDetails,
  getDiagnosticModeEnabled,
  setLastOfficeDiagnostics,
  type OfficeActionFailure,
} from "@/lib/office/diagnostics";

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

export function applyStyleToParagraph(
  paragraph: Word.Paragraph,
  style: StylePreset,
  styleKey?: PresetStyleKey
): void {
  const safeStyle = sanitizeStyle(style);

  if (styleKey) {
    applyBuiltInStyleForPreset(paragraph, styleKey);
  }

  paragraph.font.name = safeStyle.text.fontName;
  paragraph.font.size = safeStyle.text.fontSizePt;
  paragraph.font.bold = safeStyle.text.bold;
  paragraph.font.italic = safeStyle.text.italic;
  paragraph.font.underline = safeStyle.text.underline;

  // Word allCaps support is desktop-only (WordApiDesktop 1.3).
  if (supportsAllCaps()) {
    try {
      paragraph.font.allCaps = safeStyle.text.allCaps;
    } catch {
      // no-op
    }
  }

  paragraph.alignment = safeStyle.paragraph.alignment;
  paragraph.lineSpacing = safeStyle.paragraph.lineSpacingPt;
  paragraph.spaceBefore = safeStyle.paragraph.spaceBeforePt;
  paragraph.spaceAfter = safeStyle.paragraph.spaceAfterPt;
  paragraph.firstLineIndent = cmToPoints(safeStyle.paragraph.firstLineIndentCm);
  paragraph.leftIndent = cmToPoints(safeStyle.paragraph.leftIndentCm);
  paragraph.rightIndent = cmToPoints(safeStyle.paragraph.rightIndentCm);
}

async function getTargetParagraphs(
  context: Word.RequestContext,
  target: ApplyTarget
): Promise<Word.ParagraphCollection> {
  const paragraphs =
    target === "selection" ? context.document.getSelection().paragraphs : context.document.body.paragraphs;

  paragraphs.load("items/text");
  await context.sync();
  return paragraphs;
}

async function applyStyleWithFallback(
  context: Word.RequestContext,
  paragraphs: Word.ParagraphCollection,
  style: StylePreset,
  styleKey: PresetStyleKey | undefined,
  operation: string,
  target: string
): Promise<number> {
  const diagnosticMode = getDiagnosticModeEnabled();
  const failures: OfficeActionFailure[] = [];
  let fallbackUsed = false;
  let batchError: string | undefined;

  for (const paragraph of paragraphs.items) {
    applyStyleToParagraph(paragraph, style, styleKey);
  }

  try {
    await context.sync();

    if (diagnosticMode) {
      setLastOfficeDiagnostics({
        operation,
        target,
        attempted: paragraphs.items.length,
        updated: paragraphs.items.length,
        failed: 0,
        fallbackUsed: false,
        failures: [],
        timestamp: new Date().toISOString(),
      });
    } else {
      clearLastOfficeDiagnostics();
    }

    return paragraphs.items.length;
  } catch (error: unknown) {
    fallbackUsed = true;
    batchError = extractOfficeErrorDetails(error).message;
  }

  const chunkSize = 25;
  let updated = 0;
  for (let i = 0; i < paragraphs.items.length; i += chunkSize) {
    const chunk = paragraphs.items.slice(i, i + chunkSize);
    try {
      for (const paragraph of chunk) {
        applyStyleToParagraph(paragraph, style, styleKey);
      }
      await context.sync();
      updated += chunk.length;
    } catch (chunkError: unknown) {
      fallbackUsed = true;
      if (!batchError) {
        batchError = extractOfficeErrorDetails(chunkError).message;
      }

      // If chunk fails, isolate to per-paragraph only for this chunk.
      for (let j = 0; j < chunk.length; j += 1) {
        const paragraph = chunk[j];
        try {
          applyStyleToParagraph(paragraph, style, styleKey);
          await context.sync();
          updated += 1;
        } catch (paragraphError: unknown) {
          if (diagnosticMode) {
            const details = extractOfficeErrorDetails(paragraphError);
            failures.push({
              paragraphIndex: i + j + 1,
              textPreview: buildTextPreview(paragraph.text || ""),
              phase: "per-paragraph",
              error: details.message,
              errorLocation: details.errorLocation,
              statement: details.statement,
            });
          }
        }
      }
    }
  }

  if (diagnosticMode) {
    setLastOfficeDiagnostics({
      operation,
      target,
      attempted: paragraphs.items.length,
      updated,
      failed: failures.length,
      fallbackUsed,
      batchError,
      failures,
      timestamp: new Date().toISOString(),
    });
  } else {
    clearLastOfficeDiagnostics();
  }

  return updated;
}

export async function applyStylePresetToTarget(
  styleKey: PresetStyleKey,
  style: StylePreset,
  target: ApplyTarget
): Promise<number> {
  return Word.run(async (context) => {
    const paragraphs = await getTargetParagraphs(context, target);
    return applyStyleWithFallback(
      context,
      paragraphs,
      style,
      styleKey,
      "Apply style preset",
      target
    );
  });
}

export async function applyStylePresetToCurrentParagraph(style: StylePreset): Promise<void> {
  await Word.run(async (context) => {
    const paragraph = context.document.getSelection().paragraphs.getFirst();
    applyStyleToParagraph(paragraph, style);
    await context.sync();
  });
}

export async function normalizeDocumentWithStyle(style: StylePreset): Promise<number> {
  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/text");
    await context.sync();

    return applyStyleWithFallback(
      context,
      paragraphs,
      style,
      "body",
      "Normalize as body style",
      "document"
    );
  });
}
