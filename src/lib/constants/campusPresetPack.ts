import type { SkripsiPresetV1 } from "@/types/preset";
import { DEFAULT_PRESET } from "@/lib/constants/defaultPreset";

function clonePreset(preset: SkripsiPresetV1): SkripsiPresetV1 {
  return JSON.parse(JSON.stringify(preset)) as SkripsiPresetV1;
}

function buildVariant(
  id: string,
  name: string,
  mutate: (preset: SkripsiPresetV1) => void
): SkripsiPresetV1 {
  const next = clonePreset(DEFAULT_PRESET);
  next.id = id;
  next.name = name;
  mutate(next);
  return next;
}

export const CAMPUS_PRESET_PACK: SkripsiPresetV1[] = [
  DEFAULT_PRESET,
  buildVariant(
    "kampus-template-teknik-v1",
    "Campus Pack: Teknik Double Spacing",
    (preset) => {
      preset.styles.heading1.paragraph.alignment = "Centered";
      preset.styles.heading1.text.allCaps = true;
      preset.styles.heading2.paragraph.alignment = "Left";
      preset.styles.heading3.paragraph.alignment = "Left";
      preset.styles.body.paragraph.lineSpacingPt = 24;
      preset.styles.body.paragraph.firstLineIndentCm = 1.25;
      preset.styles.captionFigure.paragraph.alignment = "Centered";
      preset.styles.captionTable.paragraph.alignment = "Centered";
    }
  ),
  buildVariant(
    "kampus-template-soshum-v1",
    "Campus Pack: Soshum 1.5 Spacing",
    (preset) => {
      preset.styles.body.paragraph.lineSpacingPt = 18;
      preset.styles.body.paragraph.firstLineIndentCm = 1.25;
      preset.styles.heading1.text.fontSizePt = 13;
      preset.styles.heading1.paragraph.alignment = "Left";
      preset.styles.heading1.text.allCaps = false;
      preset.styles.heading2.text.fontSizePt = 12;
      preset.styles.quote.paragraph.lineSpacingPt = 16;
      preset.styles.captionFigure.text.fontSizePt = 10.5;
      preset.styles.captionTable.text.fontSizePt = 10.5;
    }
  ),
  buildVariant(
    "kampus-template-riset-v1",
    "Campus Pack: Riset Formal",
    (preset) => {
      preset.styles.body.text.fontName = "Arial";
      preset.styles.body.text.fontSizePt = 11;
      preset.styles.body.paragraph.lineSpacingPt = 18;
      preset.styles.body.paragraph.firstLineIndentCm = 0.75;

      preset.styles.heading1.text.fontName = "Arial";
      preset.styles.heading1.text.fontSizePt = 13;
      preset.styles.heading1.paragraph.alignment = "Left";
      preset.styles.heading1.text.allCaps = false;

      preset.styles.heading2.text.fontName = "Arial";
      preset.styles.heading3.text.fontName = "Arial";
      preset.styles.captionFigure.text.fontName = "Arial";
      preset.styles.captionTable.text.fontName = "Arial";
      preset.styles.quote.text.fontName = "Arial";
      preset.styles.captionFigure.paragraph.alignment = "Left";
      preset.styles.captionTable.paragraph.alignment = "Left";
    }
  ),
  buildVariant(
    "kampus-template-kesehatan-v1",
    "Campus Pack: Kesehatan Presisi",
    (preset) => {
      preset.styles.body.paragraph.lineSpacingPt = 20;
      preset.styles.body.paragraph.firstLineIndentCm = 1.25;
      preset.styles.heading1.paragraph.alignment = "Centered";
      preset.styles.heading1.text.fontSizePt = 14;
      preset.styles.heading2.paragraph.spaceBeforePt = 10;
      preset.styles.heading2.paragraph.spaceAfterPt = 4;
      preset.styles.heading3.paragraph.spaceBeforePt = 4;
      preset.styles.heading3.paragraph.spaceAfterPt = 4;
      preset.styles.captionFigure.paragraph.spaceBeforePt = 4;
      preset.styles.captionFigure.paragraph.spaceAfterPt = 4;
      preset.styles.captionTable.paragraph.spaceBeforePt = 4;
      preset.styles.captionTable.paragraph.spaceAfterPt = 4;
    }
  ),
];

export const BUILT_IN_PRESET_IDS = CAMPUS_PRESET_PACK.map((preset) => preset.id);
export const DEFAULT_PRESET_LIBRARY: SkripsiPresetV1[] = CAMPUS_PRESET_PACK;
export const PRESET_PACK_NOTICE =
  "Built-in campus packs are starter templates. Always confirm with your official faculty formatting guide.";
