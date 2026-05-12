import type { ApplyTarget } from "@/types/preset";

const HEADING_STYLE_MAP: Record<1 | 2 | 3, string> = {
  1: "Heading1",
  2: "Heading2",
  3: "Heading3",
};

export function setHeadingStyle(paragraph: Word.Paragraph, level: 1 | 2 | 3): void {
  const builtInStyle = HEADING_STYLE_MAP[level];

  try {
    paragraph.styleBuiltIn = builtInStyle as Word.BuiltInStyleName;
  } catch {
    // Fallback to localized style name where available.
    paragraph.style = `Heading ${level}`;
  }
}

export async function applyHeadingStyle(level: 1 | 2 | 3, target: ApplyTarget): Promise<number> {
  return Word.run(async (context) => {
    const paragraphs =
      target === "selection" ? context.document.getSelection().paragraphs : context.document.body.paragraphs;

    paragraphs.load("items");
    await context.sync();

    for (const paragraph of paragraphs.items) {
      setHeadingStyle(paragraph, level);
    }

    await context.sync();
    return paragraphs.items.length;
  });
}
