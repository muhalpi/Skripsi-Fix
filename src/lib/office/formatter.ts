import type { ApplyTarget, StylePreset } from "@/types/preset";
import { cmToPoints } from "@/lib/utils/units";

export function applyStyleToParagraph(paragraph: Word.Paragraph, style: StylePreset): void {
  paragraph.font.name = style.text.fontName;
  paragraph.font.size = style.text.fontSizePt;
  paragraph.font.bold = style.text.bold;
  paragraph.font.italic = style.text.italic;
  paragraph.font.underline = style.text.underline;

  // All caps is desktop-only in some hosts; silently ignore if unsupported.
  try {
    paragraph.font.allCaps = style.text.allCaps;
  } catch {
    // no-op
  }

  paragraph.alignment = style.paragraph.alignment;
  paragraph.lineSpacing = style.paragraph.lineSpacingPt;
  paragraph.spaceBefore = style.paragraph.spaceBeforePt;
  paragraph.spaceAfter = style.paragraph.spaceAfterPt;
  paragraph.firstLineIndent = cmToPoints(style.paragraph.firstLineIndentCm);
  paragraph.leftIndent = cmToPoints(style.paragraph.leftIndentCm);
  paragraph.rightIndent = cmToPoints(style.paragraph.rightIndentCm);
}

async function getTargetParagraphs(
  context: Word.RequestContext,
  target: ApplyTarget
): Promise<Word.ParagraphCollection> {
  const paragraphs =
    target === "selection" ? context.document.getSelection().paragraphs : context.document.body.paragraphs;

  paragraphs.load("items");
  await context.sync();
  return paragraphs;
}

export async function applyStylePresetToTarget(style: StylePreset, target: ApplyTarget): Promise<number> {
  return Word.run(async (context) => {
    const paragraphs = await getTargetParagraphs(context, target);

    for (const paragraph of paragraphs.items) {
      applyStyleToParagraph(paragraph, style);
    }

    await context.sync();
    return paragraphs.items.length;
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
    paragraphs.load("items");
    await context.sync();

    for (const paragraph of paragraphs.items) {
      applyStyleToParagraph(paragraph, style);
    }

    await context.sync();
    return paragraphs.items.length;
  });
}
