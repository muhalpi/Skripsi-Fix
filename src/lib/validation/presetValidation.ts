import { z } from "zod";
import type { SkripsiPresetV1 } from "@/types/preset";

const alignmentSchema = z.enum(["Left", "Centered", "Right", "Justified"]);
const underlineSchema = z.enum(["None", "Single"]);

const textPresetSchema = z.object({
  fontName: z.string().min(1),
  fontSizePt: z.number().positive().max(72),
  bold: z.boolean(),
  italic: z.boolean(),
  underline: underlineSchema,
  allCaps: z.boolean(),
});

const paragraphPresetSchema = z.object({
  alignment: alignmentSchema,
  lineSpacingPt: z.number().positive().max(200),
  spaceBeforePt: z.number().min(0).max(200),
  spaceAfterPt: z.number().min(0).max(200),
  firstLineIndentCm: z.number().min(-5).max(10),
  leftIndentCm: z.number().min(0).max(20),
  rightIndentCm: z.number().min(0).max(20),
});

const stylePresetSchema = z.object({
  text: textPresetSchema,
  paragraph: paragraphPresetSchema,
});

const captionPresetSchema = z.object({
  label: z.enum(["Figure", "Table"]),
  separator: z.enum([".", ":", "-"]),
  titleCase: z.enum(["Sentence", "Title"]),
});

export const skripsiPresetSchema = z.object({
  id: z.string().min(1),
  name: z.string().min(1),
  version: z.literal("1.0.0"),
  locale: z.literal("id-ID"),
  styles: z.object({
    body: stylePresetSchema,
    heading1: stylePresetSchema,
    heading2: stylePresetSchema,
    heading3: stylePresetSchema,
    quote: stylePresetSchema,
    captionFigure: stylePresetSchema,
    captionTable: stylePresetSchema,
  }),
  captions: z.object({
    figure: captionPresetSchema,
    table: captionPresetSchema,
  }),
});

export function validatePreset(value: unknown): SkripsiPresetV1 {
  return skripsiPresetSchema.parse(value);
}

export function isValidPreset(value: unknown): value is SkripsiPresetV1 {
  return skripsiPresetSchema.safeParse(value).success;
}
