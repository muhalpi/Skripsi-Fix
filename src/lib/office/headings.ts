import type { ApplyTarget, PresetStyleKey } from "@/types/preset";
import { applyBuiltInStyleForPreset } from "@/lib/office/styleRegistry";
import {
  buildTextPreview,
  clearLastOfficeDiagnostics,
  extractOfficeErrorDetails,
  getDiagnosticModeEnabled,
  setLastOfficeDiagnostics,
  type OfficeActionFailure,
} from "@/lib/office/diagnostics";

const HEADING_STYLE_MAP: Record<1 | 2 | 3, PresetStyleKey> = {
  1: "heading1",
  2: "heading2",
  3: "heading3",
};

export function setHeadingStyle(paragraph: Word.Paragraph, level: 1 | 2 | 3): void {
  applyBuiltInStyleForPreset(paragraph, HEADING_STYLE_MAP[level]);
}

export async function applyHeadingStyle(level: 1 | 2 | 3, target: ApplyTarget): Promise<number> {
  return Word.run(async (context) => {
    const paragraphs =
      target === "selection" ? context.document.getSelection().paragraphs : context.document.body.paragraphs;

    const diagnosticMode = getDiagnosticModeEnabled();
    const failures: OfficeActionFailure[] = [];
    let fallbackUsed = false;
    let batchError: string | undefined;

    paragraphs.load("items/text");
    await context.sync();

    for (const paragraph of paragraphs.items) {
      setHeadingStyle(paragraph, level);
    }

    try {
      await context.sync();

      if (diagnosticMode) {
        setLastOfficeDiagnostics({
          operation: "Apply heading style",
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
          setHeadingStyle(paragraph, level);
        }
        await context.sync();
        updated += chunk.length;
      } catch (chunkError: unknown) {
        fallbackUsed = true;
        if (!batchError) {
          batchError = extractOfficeErrorDetails(chunkError).message;
        }

        for (let j = 0; j < chunk.length; j += 1) {
          const paragraph = chunk[j];
          try {
            setHeadingStyle(paragraph, level);
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
        operation: "Apply heading style",
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
  });
}
