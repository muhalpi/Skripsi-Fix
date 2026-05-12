import type { AuditMismatch, AuditReport, StylePreset } from "@/types/preset";
import { almostEqual, cmToPoints } from "@/lib/utils/units";
import { classifyParagraphStyleKey } from "@/lib/office/chapterAware";
import { buildTextPreview } from "@/lib/office/diagnostics";

function getMismatchReasons(paragraph: Word.Paragraph, expected: StylePreset): string[] {
  const reasons: string[] = [];

  if (String(paragraph.styleBuiltIn || "") !== "Normal") {
    reasons.push("built-in style");
  }

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

export async function auditDocumentBody(expected: StylePreset): Promise<AuditReport> {
  return Word.run(async (context) => {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load(
      "items/text,items/styleBuiltIn,items/alignment,items/lineSpacing,items/spaceBefore,items/spaceAfter,items/firstLineIndent,items/leftIndent,items/rightIndent,items/font/name,items/font/size"
    );

    await context.sync();

    const mismatches: AuditMismatch[] = [];
    let auditedBodyParagraphs = 0;

    paragraphs.items.forEach((paragraph, index) => {
      const styleKey = classifyParagraphStyleKey(paragraph.text, String(paragraph.styleBuiltIn || ""));
      if (styleKey !== "body") {
        return;
      }

      auditedBodyParagraphs += 1;
      const reasons = getMismatchReasons(paragraph, expected);
      if (reasons.length > 0) {
        mismatches.push({
          index: index + 1,
          textPreview: buildTextPreview(paragraph.text),
          reasons,
        });
      }
    });

    return {
      totalParagraphs: auditedBodyParagraphs,
      mismatches,
    };
  });
}
