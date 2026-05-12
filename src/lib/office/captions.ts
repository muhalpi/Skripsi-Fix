import type { CaptionLabel, StylePreset } from "@/types/preset";
import { cmToPoints } from "@/lib/utils/units";
import { applyBuiltInStyleForPreset } from "@/lib/office/styleRegistry";

function toSentenceCase(value: string): string {
  if (!value.trim()) {
    return value;
  }

  const trimmed = value.trim();
  return `${trimmed.charAt(0).toUpperCase()}${trimmed.slice(1)}`;
}

function toTitleCase(value: string): string {
  return value
    .trim()
    .split(/\s+/)
    .map((word) => `${word.charAt(0).toUpperCase()}${word.slice(1).toLowerCase()}`)
    .join(" ");
}

export async function insertCaption(options: {
  label: CaptionLabel;
  separator: "." | ":" | "-";
  title: string;
  titleCase: "Sentence" | "Title";
  captionStyle: StylePreset;
}): Promise<void> {
  const normalizedTitle = options.titleCase === "Title" ? toTitleCase(options.title) : toSentenceCase(options.title);

  await Word.run(async (context) => {
    const selection = context.document.getSelection();

    const captionParagraph = selection.insertParagraph("", Word.InsertLocation.after);
    const captionRange = captionParagraph.getRange();

    captionRange.insertText(`${options.label} `, Word.InsertLocation.start);
    captionRange.insertField(Word.InsertLocation.end, "Seq", `${options.label} \\* ARABIC`, false);
    captionRange.insertText(`${options.separator} ${normalizedTitle}`, Word.InsertLocation.end);

    applyBuiltInStyleForPreset(
      captionParagraph,
      options.label === "Figure" ? "captionFigure" : "captionTable"
    );

    captionParagraph.font.name = options.captionStyle.text.fontName;
    captionParagraph.font.size = options.captionStyle.text.fontSizePt;
    captionParagraph.font.bold = options.captionStyle.text.bold;
    captionParagraph.font.italic = options.captionStyle.text.italic;
    captionParagraph.font.underline = options.captionStyle.text.underline;

    try {
      captionParagraph.font.allCaps = options.captionStyle.text.allCaps;
    } catch {
      // no-op
    }

    captionParagraph.alignment = options.captionStyle.paragraph.alignment;
    captionParagraph.lineSpacing = options.captionStyle.paragraph.lineSpacingPt;
    captionParagraph.spaceBefore = options.captionStyle.paragraph.spaceBeforePt;
    captionParagraph.spaceAfter = options.captionStyle.paragraph.spaceAfterPt;
    captionParagraph.firstLineIndent = cmToPoints(options.captionStyle.paragraph.firstLineIndentCm);
    captionParagraph.leftIndent = cmToPoints(options.captionStyle.paragraph.leftIndentCm);
    captionParagraph.rightIndent = cmToPoints(options.captionStyle.paragraph.rightIndentCm);

    await context.sync();
  });
}
