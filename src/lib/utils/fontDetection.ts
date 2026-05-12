const BASELINE_FONT_FAMILIES = ["monospace", "sans-serif", "serif"] as const;
const PROBE_TEXT = "AaBbCcXxYyZz0123456789mmmmmmmmmmlliWW";
const PROBE_SIZE_PX = 72;

export const COMMON_FONT_CANDIDATES: string[] = [
  "Times New Roman",
  "Arial",
  "Calibri",
  "Cambria",
  "Book Antiqua",
  "Bookman Old Style",
  "Garamond",
  "Georgia",
  "Palatino Linotype",
  "Segoe UI",
  "Tahoma",
  "Trebuchet MS",
  "Verdana",
  "Arial Narrow",
  "Courier New",
  "Consolas",
  "Arial Black",
  "Century Gothic",
  "Franklin Gothic Book",
  "Gill Sans MT",
  "Lucida Sans Unicode",
  "Corbel",
  "Candara",
  "Impact",
  "Roboto",
  "Open Sans",
  "Lato",
  "Poppins",
  "Inter",
  "Merriweather",
  "Ubuntu",
  "Fira Sans",
  "PT Sans",
  "Source Sans Pro",
  "Source Serif Pro",
  "Noto Sans",
  "Noto Serif",
  "Helvetica",
  "Helvetica Neue",
];

function quoteFontFamily(fontFamily: string): string {
  return `"${fontFamily.replace(/"/g, '\\"')}"`;
}

function uniqueSorted(values: string[]): string[] {
  return Array.from(new Set(values.filter((value) => value.trim().length > 0))).sort((a, b) =>
    a.localeCompare(b)
  );
}

function createMeasurementContext(): CanvasRenderingContext2D | null {
  if (typeof document === "undefined") {
    return null;
  }

  const canvas = document.createElement("canvas");
  canvas.width = 400;
  canvas.height = 120;
  return canvas.getContext("2d");
}

function isFontDetectedByMetrics(
  context: CanvasRenderingContext2D,
  fontFamily: string,
  baselineWidths: Record<(typeof BASELINE_FONT_FAMILIES)[number], number>
): boolean {
  const quotedFamily = quoteFontFamily(fontFamily);
  for (const baseFamily of BASELINE_FONT_FAMILIES) {
    context.font = `${PROBE_SIZE_PX}px ${quotedFamily}, ${baseFamily}`;
    const width = context.measureText(PROBE_TEXT).width;
    if (Math.abs(width - baselineWidths[baseFamily]) > 0.1) {
      return true;
    }
  }

  return false;
}

export async function detectInstalledFonts(candidates: string[]): Promise<string[]> {
  if (typeof document === "undefined") {
    return [];
  }

  if (document.fonts?.ready) {
    try {
      await document.fonts.ready;
    } catch {
      // no-op
    }
  }

  const context = createMeasurementContext();
  if (!context) {
    return [];
  }

  const baselineWidths = {
    monospace: 0,
    "sans-serif": 0,
    serif: 0,
  } as Record<(typeof BASELINE_FONT_FAMILIES)[number], number>;

  for (const baseFamily of BASELINE_FONT_FAMILIES) {
    context.font = `${PROBE_SIZE_PX}px ${baseFamily}`;
    baselineWidths[baseFamily] = context.measureText(PROBE_TEXT).width;
  }

  const dedupedCandidates = uniqueSorted(candidates);
  const detected: string[] = [];
  for (const fontFamily of dedupedCandidates) {
    if (isFontDetectedByMetrics(context, fontFamily, baselineWidths)) {
      detected.push(fontFamily);
    }
  }

  return detected;
}
