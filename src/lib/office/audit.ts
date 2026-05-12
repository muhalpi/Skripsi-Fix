import type { AuditMismatch, AuditReport, StylePreset } from "@/types/preset";
import { almostEqual, cmToPoints } from "@/lib/utils/units";

function getMismatchReasons(paragraph: Word.Paragraph, expected: StylePreset): string[] {
  const reasons: string[] = [];

  if (paragraph.font.name !== expected.text.fontName) {
    reasons.push("font name");
  }
  if (!almostEqual(paragraph.font.size ?? 0, expected.text.fontSizePt)) {
    reasons.push("font size");
  }
  if (paragraph.alignment !== expected.paragraph.alignment) {
    reasons.push("alignment");
  }
  if (!almostEqual(paragraph.lineSpacing, expected.paragraph.lineSpacingPt)) {
    reasons.push("line spacing");
  }
  if (!almostEqual(paragraph.spaceBefore, expected.paragraph.spaceBeforePt)) {
    reasons.push("space before");
  }
  if (!almostEqual(paragraph.spaceAfter, expected.paragraph.spaceAfterPt)) {
    reasons.push("space after");
  }
  if (!almostEqual(paragraph.firstLineIndent, cmToPoints(expected.paragraph.firstLineIndentCm))) {
    reasons.push("first line indent");
  }
  if (!almostEqual(paragraph.leftIndent, cmToPoints(expected.paragraph.leftIndentCm))) {
    reasons.push("left indent");
  }
  if (!almostEqual(paragraph.rightIndent, cmToPoints(expected.paragraph.rightIndentCm))) {
    reasons.push("right indent");
  }

  return reasons;
}

function buildPreview(text: string): string {
  const compact = text.replace(/\s+/g, " ").trim();
  if (!compact) {
    return "(empty paragraph)";
  }
  return compact.length > 60 ? `${compact.slice(0, 60)}...` : compact;
}

export async function auditDocumentBody(expected: StylePreset): Promise<AuditReport> {
  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load(
      "items/text,items/alignment,items/lineSpacing,items/spaceBefore,items/spaceAfter,items/firstLineIndent,items/leftIndent,items/rightIndent,items/font/name,items/font/size"
    );

    await context.sync();

    const mismatches: AuditMismatch[] = [];

    paragraphs.items.forEach((paragraph, index) => {
      const reasons = getMismatchReasons(paragraph, expected);
      if (reasons.length > 0) {
        mismatches.push({
          index: index + 1,
          textPreview: buildPreview(paragraph.text),
          reasons,
        });
      }
    });

    return {
      totalParagraphs: paragraphs.items.length,
      mismatches,
    };
  });
}

export async function fixDocumentBodyMismatches(expected: StylePreset): Promise<number> {
  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load(
      "items/text,items/alignment,items/lineSpacing,items/spaceBefore,items/spaceAfter,items/firstLineIndent,items/leftIndent,items/rightIndent,items/font/name,items/font/size"
    );

    await context.sync();

    let fixed = 0;

    for (const paragraph of paragraphs.items) {
      const reasons = getMismatchReasons(paragraph, expected);
      if (reasons.length === 0) {
        continue;
      }

      paragraph.font.name = expected.text.fontName;
      paragraph.font.size = expected.text.fontSizePt;
      paragraph.font.bold = expected.text.bold;
      paragraph.font.italic = expected.text.italic;
      paragraph.font.underline = expected.text.underline;

      try {
        paragraph.font.allCaps = expected.text.allCaps;
      } catch {
        // no-op
      }

      paragraph.alignment = expected.paragraph.alignment;
      paragraph.lineSpacing = expected.paragraph.lineSpacingPt;
      paragraph.spaceBefore = expected.paragraph.spaceBeforePt;
      paragraph.spaceAfter = expected.paragraph.spaceAfterPt;
      paragraph.firstLineIndent = cmToPoints(expected.paragraph.firstLineIndentCm);
      paragraph.leftIndent = cmToPoints(expected.paragraph.leftIndentCm);
      paragraph.rightIndent = cmToPoints(expected.paragraph.rightIndentCm);
      fixed += 1;
    }

    if (fixed > 0) {
      await context.sync();
    }

    return fixed;
  });
}
