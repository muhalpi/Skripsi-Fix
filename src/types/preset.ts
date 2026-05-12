export type Alignment = "Left" | "Centered" | "Right" | "Justified";

export type Underline = "None" | "Single";

export type TextPreset = {
  fontName: string;
  fontSizePt: number;
  bold: boolean;
  italic: boolean;
  underline: Underline;
  allCaps: boolean;
};

export type ParagraphPreset = {
  alignment: Alignment;
  lineSpacingPt: number;
  spaceBeforePt: number;
  spaceAfterPt: number;
  firstLineIndentCm: number;
  leftIndentCm: number;
  rightIndentCm: number;
};

export type StylePreset = {
  text: TextPreset;
  paragraph: ParagraphPreset;
};

export type CaptionLabel = "Figure" | "Table";

export type CaptionPreset = {
  label: CaptionLabel;
  separator: "." | ":" | "-";
  titleCase: "Sentence" | "Title";
};

export type PresetStyleKey =
  | "body"
  | "heading1"
  | "heading2"
  | "heading3"
  | "quote"
  | "captionFigure"
  | "captionTable";

export type SkripsiPresetV1 = {
  id: string;
  name: string;
  version: "1.0.0";
  locale: "id-ID";
  styles: Record<PresetStyleKey, StylePreset>;
  captions: {
    figure: CaptionPreset;
    table: CaptionPreset;
  };
};

export type ApplyTarget = "selection" | "document";

export type AuditMismatch = {
  index: number;
  textPreview: string;
  reasons: string[];
};

export type AuditReport = {
  totalParagraphs: number;
  mismatches: AuditMismatch[];
};
