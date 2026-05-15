"use client";

import { useEffect, useMemo, useState } from "react";
import { DEFAULT_PRESET } from "@/lib/constants/defaultPreset";
import { insertCaption } from "@/lib/office/captions";
import { applyStylePresetToTarget } from "@/lib/office/formatter";
import { applyHeadingStyle } from "@/lib/office/headings";
import {
  applyAllLevelSettingsToSelectionList,
  applyLevelSettingsToSelectionList,
  getLevelFormatPreview,
  type FollowNumberWith,
  type LinkedHeadingStyle,
  type ListApplyScope,
  type ListLevelAlignment,
  type ListNumberingStyle,
  type MultiLevelLevelSettings,
} from "@/lib/office/multilevel";
import { isRequirementSetSupported, waitForOfficeReady } from "@/lib/office/runtime";
import {
  insertListOfFiguresAtSelection,
  insertListOfTablesAtSelection,
  insertTocAtSelection,
  updateAllFields,
  updateListOfFiguresFields,
  updateListOfTablesFields,
  updateTocFields,
} from "@/lib/office/toc";
import type { ApplyTarget, CaptionLabel, PresetStyleKey } from "@/types/preset";

type TabKey = "multilevel" | "heading" | "caption" | "toc";

type Notice = {
  type: "info" | "ok" | "error";
  text: string;
};

type RuntimeState = {
  isOfficeReady: boolean;
  isWordHost: boolean;
  hostName: string;
  details: string;
};

const TAB_ITEMS: Array<{ key: TabKey; label: string }> = [
  { key: "multilevel", label: "Multilevel" },
  { key: "heading", label: "Judul & Gaya" },
  { key: "caption", label: "Keterangan" },
  { key: "toc", label: "Daftar Isi" },
];

const STYLE_OPTIONS: Array<{ value: PresetStyleKey; label: string }> = [
  { value: "body", label: "Teks Utama" },
  { value: "heading1", label: "Judul 1" },
  { value: "heading2", label: "Judul 2" },
  { value: "heading3", label: "Judul 3" },
  { value: "quote", label: "Kutipan" },
  { value: "captionFigure", label: "Gaya Keterangan Gambar" },
  { value: "captionTable", label: "Gaya Keterangan Tabel" },
];

const NUMBER_STYLE_OPTIONS: Array<{ value: ListNumberingStyle; label: string }> = [
  { value: "UpperRoman", label: "I, II, III, ..." },
  { value: "Arabic", label: "1, 2, 3, ..." },
  { value: "UpperLetter", label: "A, B, C, ..." },
  { value: "LowerLetter", label: "a, b, c, ..." },
  { value: "LowerRoman", label: "i, ii, iii, ..." },
  { value: "None", label: "Tanpa nomor" },
];

const ALIGNMENT_OPTIONS: Array<{ value: ListLevelAlignment; label: string }> = [
  { value: "Left", label: "Kiri" },
  { value: "Centered", label: "Tengah" },
  { value: "Right", label: "Kanan" },
];

const FOLLOW_NUMBER_OPTIONS: Array<{ value: FollowNumberWith; label: string }> = [
  { value: "TrailingTab", label: "Karakter tab" },
  { value: "TrailingSpace", label: "Spasi" },
  { value: "TrailingNone", label: "Tanpa pemisah" },
];

const APPLY_CHANGES_OPTIONS: Array<{ value: ListApplyScope; label: string }> = [
  { value: "WholeList", label: "Seluruh daftar" },
  { value: "ThisPointForward", label: "Dari titik ini ke akhir daftar" },
  { value: "Selection", label: "Hanya paragraf terpilih" },
];

const LINKED_STYLE_OPTIONS: Array<{ value: LinkedHeadingStyle; label: string }> = [
  { value: "None", label: "(Tanpa gaya)" },
  { value: "Heading1", label: "Judul 1" },
  { value: "Heading2", label: "Judul 2" },
  { value: "Heading3", label: "Judul 3" },
  { value: "Heading4", label: "Judul 4" },
  { value: "Heading5", label: "Judul 5" },
  { value: "Heading6", label: "Judul 6" },
  { value: "Heading7", label: "Judul 7" },
  { value: "Heading8", label: "Judul 8" },
  { value: "Heading9", label: "Judul 9" },
];

const LEVEL_INDICES = Array.from({ length: 9 }, (_, index) => index);

function extractErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }
  return "Terjadi kesalahan yang tidak diketahui.";
}

function buildNumberFormatPattern(
  levelIndex: number,
  includeFromLevelIndex: number | null,
  levelSeparator: string,
  prefixText: string,
  suffixText: string
): string {
  const normalizedLevel = Math.max(0, Math.min(8, levelIndex));
  const includeFrom =
    includeFromLevelIndex === null
      ? normalizedLevel
      : Math.max(0, Math.min(normalizedLevel, includeFromLevelIndex));
  const separator = levelSeparator.length > 0 ? levelSeparator : ".";

  const levelTokens: string[] = [];
  for (let current = includeFrom; current <= normalizedLevel; current += 1) {
    levelTokens.push(`<L${current + 1}>`);
  }

  const numberingCore = levelTokens.join(separator);
  return `${prefixText}${numberingCore}${suffixText}`;
}

function toRoman(value: number): string {
  if (value <= 0) {
    return String(value);
  }

  const table: Array<{ value: number; symbol: string }> = [
    { value: 1000, symbol: "M" },
    { value: 900, symbol: "CM" },
    { value: 500, symbol: "D" },
    { value: 400, symbol: "CD" },
    { value: 100, symbol: "C" },
    { value: 90, symbol: "XC" },
    { value: 50, symbol: "L" },
    { value: 40, symbol: "XL" },
    { value: 10, symbol: "X" },
    { value: 9, symbol: "IX" },
    { value: 5, symbol: "V" },
    { value: 4, symbol: "IV" },
    { value: 1, symbol: "I" },
  ];

  let remaining = value;
  let result = "";
  for (const item of table) {
    while (remaining >= item.value) {
      result += item.symbol;
      remaining -= item.value;
    }
  }
  return result;
}

function toAlphabetic(value: number): string {
  if (value <= 0) {
    return String(value);
  }

  let number = value;
  let result = "";
  while (number > 0) {
    const remainder = (number - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    number = Math.floor((number - 1) / 26);
  }
  return result;
}

function toPreviewNumber(style: ListNumberingStyle, value: number): string {
  const normalized = Math.max(1, Number.isFinite(value) ? Math.floor(value) : 1);
  if (style === "UpperRoman") {
    return toRoman(normalized);
  }
  if (style === "LowerRoman") {
    return toRoman(normalized).toLowerCase();
  }
  if (style === "UpperLetter") {
    return toAlphabetic(normalized);
  }
  if (style === "LowerLetter") {
    return toAlphabetic(normalized).toLowerCase();
  }
  if (style === "None") {
    return "";
  }
  return String(normalized);
}

function getLinkedStylePreviewLabel(linkedStyle: LinkedHeadingStyle): string {
  if (linkedStyle === "None") {
    return "Paragraf";
  }
  return `Judul ${linkedStyle.replace("Heading", "")}`;
}

function getPatternFromSettings(settings: MultiLevelLevelSettings): string {
  const manualPattern = settings.numberFormatPattern.trim();
  if (manualPattern.length > 0) {
    return settings.numberFormatPattern;
  }
  return buildNumberFormatPattern(
    settings.levelIndex,
    settings.includeFromLevelIndex,
    settings.levelSeparator,
    settings.prefixText,
    settings.suffixText
  );
}

function buildRealLevelPreview(
  settings: MultiLevelLevelSettings,
  allSettings: MultiLevelLevelSettings[]
): string {
  const maxLevelIndex = Math.max(0, Math.min(8, settings.levelIndex));
  const pattern = getPatternFromSettings(settings);
  const baseText = pattern.replace(/<L([1-9])>/gi, (_match, rawLevel) => {
    const tokenLevelIndex = Number(rawLevel) - 1;
    if (!Number.isFinite(tokenLevelIndex) || tokenLevelIndex < 0 || tokenLevelIndex > maxLevelIndex) {
      return "";
    }

    const tokenLevelSettings = allSettings[tokenLevelIndex] ?? settings;
    return toPreviewNumber(tokenLevelSettings.numberStyle, tokenLevelSettings.startAt);
  });

  if (settings.followNumberWith === "TrailingSpace") {
    return `${baseText} `;
  }
  if (settings.followNumberWith === "TrailingTab") {
    return `${baseText}    `;
  }
  return baseText;
}

const PREVIEW_PX_PER_CM = 96 / 2.54;
const PREVIEW_NUMBER_CHAR_PX = 6.6;
const PREVIEW_NUMBER_GAP_PX = 8;

function toPreviewPx(cm: number): number {
  if (!Number.isFinite(cm)) {
    return 0;
  }
  return Math.max(0, cm) * PREVIEW_PX_PER_CM;
}

function toCssNumberAlign(alignment: ListLevelAlignment): "left" | "center" | "right" {
  if (alignment === "Centered") {
    return "center";
  }
  if (alignment === "Right") {
    return "right";
  }
  return "left";
}

function formatLocalizedDecimalInput(value: number): string {
  if (!Number.isFinite(value)) {
    return "0";
  }
  return String(value).replace(".", ",");
}

function tryParseNonNegativeDecimalInput(rawValue: string): number | null {
  const normalized = rawValue.trim().replace(",", ".");
  if (normalized.length === 0) {
    return null;
  }
  if (!/^(?:\d+|\d+\.\d*|\.\d+)$/.test(normalized)) {
    return null;
  }
  const parsedValue = Number(normalized);
  if (!Number.isFinite(parsedValue)) {
    return null;
  }
  return Math.max(0, parsedValue);
}

function parsePositiveIntegerInput(rawValue: string, fallback: number, minValue = 1): number {
  const normalized = rawValue.trim().replace(",", ".");
  if (normalized.length === 0) {
    return fallback;
  }
  const parsedValue = Number(normalized);
  if (!Number.isFinite(parsedValue)) {
    return fallback;
  }
  return Math.max(minValue, Math.floor(parsedValue));
}

function createDefaultLevelSettings(levelIndex: number): MultiLevelLevelSettings {
  const normalizedLevel = Math.max(0, Math.min(8, levelIndex));
  const headingLevel = normalizedLevel + 1;

  const linkedStyle: LinkedHeadingStyle =
    headingLevel <= 9 ? (`Heading${headingLevel}` as LinkedHeadingStyle) : "None";
  const includeFromLevelIndex = normalizedLevel === 0 ? null : 0;
  const levelSeparator = ".";
  const prefixText = normalizedLevel === 0 ? "BAB " : "";
  const suffixText = "";

  return {
    levelIndex: normalizedLevel,
    numberStyle: normalizedLevel === 0 ? "UpperRoman" : "Arabic",
    includeFromLevelIndex,
    levelSeparator,
    prefixText,
    suffixText,
    numberFormatPattern: buildNumberFormatPattern(
      normalizedLevel,
      includeFromLevelIndex,
      levelSeparator,
      prefixText,
      suffixText
    ),
    startAt: 1,
    alignment: "Left",
    alignedAtCm: 0,
    textIndentCm: 0.63,
    followNumberWith: "TrailingTab",
    addTabStopAt: false,
    tabStopAtCm: 0.63,
    legalStyleNumbering: false,
    restartListAfterLevelIndex: null,
    linkedStyle,
  };
}

function usesAdvancedDesktopOptions(settings: MultiLevelLevelSettings): boolean {
  return (
    settings.followNumberWith !== "TrailingTab" ||
    (settings.followNumberWith === "TrailingTab" && settings.addTabStopAt) ||
    settings.legalStyleNumbering ||
    settings.restartListAfterLevelIndex !== null
  );
}

export default function TaskpanePage() {
  const [activeTab, setActiveTab] = useState<TabKey>("multilevel");
  const [applyTarget, setApplyTarget] = useState<ApplyTarget>("selection");
  const [styleKey, setStyleKey] = useState<PresetStyleKey>("body");
  const [headingLevel, setHeadingLevel] = useState<1 | 2 | 3>(1);
  const [captionLabel, setCaptionLabel] = useState<CaptionLabel>("Figure");
  const [captionTitle, setCaptionTitle] = useState<string>("");

  const [selectedLevelIndex, setSelectedLevelIndex] = useState<number>(0);
  const [applyChangesTo, setApplyChangesTo] = useState<ListApplyScope>("WholeList");
  const [levelSettings, setLevelSettings] = useState<MultiLevelLevelSettings[]>(() =>
    LEVEL_INDICES.map((levelIndex) => createDefaultLevelSettings(levelIndex))
  );

  const [runtime, setRuntime] = useState<RuntimeState>({
    isOfficeReady: false,
    isWordHost: false,
    hostName: "Tidak diketahui",
    details: "Memeriksa runtime Office...",
  });
  const [busyAction, setBusyAction] = useState<string>("");
  const [notice, setNotice] = useState<Notice>({
    type: "info",
    text: "Versi baru dimulai dari nol. Pilih tab untuk menjalankan aksi format.",
  });
  const [alignedAtInput, setAlignedAtInput] = useState<string>(formatLocalizedDecimalInput(0));
  const [textIndentInput, setTextIndentInput] = useState<string>(
    formatLocalizedDecimalInput(0.63)
  );
  const [tabStopAtInput, setTabStopAtInput] = useState<string>(formatLocalizedDecimalInput(0.63));

  useEffect(() => {
    waitForOfficeReady()
      .then((status) => {
        setRuntime(status);
        setNotice({
          type: status.isWordHost ? "ok" : "info",
          text: status.details,
        });
      })
      .catch((error: unknown) => {
        setNotice({ type: "error", text: extractErrorMessage(error) });
      });
  }, []);

  const isWordReady = runtime.isOfficeReady && runtime.isWordHost;
  const isWordApi15Supported = isWordReady && isRequirementSetSupported("WordApi", "1.5");
  const isWordApiDesktopListSupported =
    isWordReady && isRequirementSetSupported("WordApiDesktop", "1.3");
  const isBusy = busyAction.length > 0;

  const activeCaptionPreset = useMemo(() => {
    return captionLabel === "Figure" ? DEFAULT_PRESET.captions.figure : DEFAULT_PRESET.captions.table;
  }, [captionLabel]);

  const activeCaptionStyle = useMemo(() => {
    return captionLabel === "Figure"
      ? DEFAULT_PRESET.styles.captionFigure
      : DEFAULT_PRESET.styles.captionTable;
  }, [captionLabel]);

  const selectedLevelSettings = levelSettings[selectedLevelIndex] ?? createDefaultLevelSettings(0);
  useEffect(() => {
    setAlignedAtInput(formatLocalizedDecimalInput(selectedLevelSettings.alignedAtCm));
    setTextIndentInput(formatLocalizedDecimalInput(selectedLevelSettings.textIndentCm));
    setTabStopAtInput(formatLocalizedDecimalInput(selectedLevelSettings.tabStopAtCm));
  }, [
    selectedLevelIndex,
    selectedLevelSettings.alignedAtCm,
    selectedLevelSettings.textIndentCm,
    selectedLevelSettings.tabStopAtCm,
  ]);

  const multilevelPreviewRows = useMemo(
    () =>
      LEVEL_INDICES.map((levelIndex) => {
        const settingsForLevel = levelSettings[levelIndex] ?? createDefaultLevelSettings(levelIndex);
        const previewText = buildRealLevelPreview(settingsForLevel, levelSettings);
        const numberOffsetPx = toPreviewPx(settingsForLevel.alignedAtCm);
        const textOffsetPx =
          settingsForLevel.followNumberWith === "TrailingTab" && settingsForLevel.addTabStopAt
            ? toPreviewPx(settingsForLevel.tabStopAtCm)
            : toPreviewPx(settingsForLevel.textIndentCm);
        const previewTextForWidth = previewText.trimEnd();
        const estimatedNumberWidthPx = Math.max(
          20,
          previewTextForWidth.length * PREVIEW_NUMBER_CHAR_PX + 6
        );
        const styleOffsetPx = Math.max(
          textOffsetPx,
          numberOffsetPx + estimatedNumberWidthPx + PREVIEW_NUMBER_GAP_PX
        );
        const numberBoxWidthPx = Math.max(30, styleOffsetPx - numberOffsetPx - 4);
        const ruleStartPx = styleOffsetPx + 56;
        const rowWidthPx = Math.max(
          260,
          numberOffsetPx + 90,
          styleOffsetPx + 170,
          numberOffsetPx + numberBoxWidthPx + 120
        );
        return {
          levelIndex,
          previewText,
          linkedStyleLabel: getLinkedStylePreviewLabel(settingsForLevel.linkedStyle),
          numberOffsetPx,
          styleOffsetPx,
          ruleStartPx,
          numberBoxWidthPx,
          rowWidthPx,
          numberTextAlign: toCssNumberAlign(settingsForLevel.alignment),
        };
      }),
    [levelSettings]
  );
  const restartAfterEnabled = selectedLevelSettings.restartListAfterLevelIndex !== null;
  const restartAfterLevelOptions = LEVEL_INDICES.filter(
    (levelIndex) => levelIndex < selectedLevelIndex
  );

  useEffect(() => {
    if (!isWordApiDesktopListSupported && applyChangesTo !== "WholeList") {
      setApplyChangesTo("WholeList");
    }
  }, [isWordApiDesktopListSupported, applyChangesTo]);

  function updateSelectedLevelSettings(
    updater: (previous: MultiLevelLevelSettings) => MultiLevelLevelSettings
  ): void {
    setLevelSettings((previous) =>
      previous.map((settings, index) =>
        index === selectedLevelIndex
          ? {
              ...updater(settings),
              levelIndex: selectedLevelIndex,
            }
          : settings
      )
    );
  }

  async function runAction(actionName: string, action: () => Promise<void>): Promise<void> {
    setBusyAction(actionName);
    try {
      await action();
      setNotice({ type: "ok", text: `${actionName} berhasil.` });
    } catch (error: unknown) {
      setNotice({ type: "error", text: `${actionName} gagal: ${extractErrorMessage(error)}` });
    } finally {
      setBusyAction("");
    }
  }

  return (
    <main className="app-root">
      <div className="tab-shell">
        <header className="tab-header">
          <p className="app-kicker">Skripsi-Fix</p>
          <h1>Formatter Skripsi Mode Tab</h1>
          <p>
            Fokus utama: editor multilevel mirip menu Word untuk atur level, format nomor, gaya
            judul, dan posisi indent.
          </p>
        </header>

        <section className="runtime-strip" aria-live="polite">
          Aplikasi: <strong>{runtime.hostName}</strong> | Siap Word: <strong>{isWordReady ? "Ya" : "Tidak"}</strong> |
          Dukungan WordApi 1.5: <strong>{isWordApi15Supported ? "Ya" : "Tidak"}</strong>
        </section>

        <nav className="tabs" role="tablist" aria-label="Menu format skripsi">
          {TAB_ITEMS.map((tab) => {
            const isActive = tab.key === activeTab;
            return (
              <button
                key={tab.key}
                className={isActive ? "tab-btn active" : "tab-btn"}
                onClick={() => setActiveTab(tab.key)}
                role="tab"
                aria-selected={isActive}
                aria-controls={`panel-${tab.key}`}
                id={`tab-${tab.key}`}
                type="button"
              >
                {tab.label}
              </button>
            );
          })}
        </nav>

        <section className="tab-panel" role="tabpanel" id={`panel-${activeTab}`} aria-labelledby={`tab-${activeTab}`}>
          {activeTab === "multilevel" ? (
            <>
              <h2>Atur Daftar Multilevel Baru</h2>
              <p>
                Klik level yang ingin diubah, lalu terapkan ke daftar pada seleksi. Menu ini mengikuti
                alur editor bawaan Word.
              </p>
              <p>
                Opsi lanjutan (`Ikuti nomor dengan`, `Tambahkan tab stop di`, `Penomoran gaya legal`,
                `Mulai ulang daftar setelah`) memerlukan dukungan WordApiDesktop 1.3:{" "}
                <strong>{isWordApiDesktopListSupported ? "Tersedia" : "Tidak tersedia"}</strong>.
              </p>

              <div className="level-editor">
                <div className="level-side">
                  <p className="mini-title">Pilih level untuk diubah</p>
                  <div className="level-grid">
                    {LEVEL_INDICES.map((levelIndex) => {
                      const isSelected = selectedLevelIndex === levelIndex;
                      return (
                        <button
                          key={`level-${levelIndex}`}
                          type="button"
                          className={isSelected ? "level-btn active" : "level-btn"}
                          onClick={() => setSelectedLevelIndex(levelIndex)}
                        >
                          {levelIndex + 1}
                        </button>
                      );
                    })}
                  </div>
                </div>

                <div className="level-main">
                  <div className="preview-showcase">
                    <p className="option-group-title">Pratinjau Hasil Akhir</p>
                    <div className="preview-canvas">
                      {multilevelPreviewRows.map((row) => (
                        <div
                          key={`preview-level-${row.levelIndex}`}
                          className={selectedLevelIndex === row.levelIndex ? "preview-item active" : "preview-item"}
                          onClick={() => setSelectedLevelIndex(row.levelIndex)}
                        >
                          <div className="preview-item-main">
                            <div
                              className="preview-item-track"
                              style={{ minWidth: `${row.rowWidthPx}px` }}
                            >
                              <span
                                className="preview-item-number"
                                style={{
                                  left: `${row.numberOffsetPx}px`,
                                  minWidth: `${row.numberBoxWidthPx}px`,
                                  textAlign: row.numberTextAlign,
                                }}
                              >
                                {row.previewText.length > 0 ? row.previewText : "(tanpa nomor)"}
                              </span>
                              <span
                                className="preview-item-style"
                                style={{ left: `${row.styleOffsetPx}px` }}
                              >
                                {row.linkedStyleLabel}
                              </span>
                              <span
                                className="preview-item-rule"
                                style={{ left: `${row.ruleStartPx}px` }}
                              />
                            </div>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>

                  <div className="option-group">
                    <p className="option-group-title">Pengaturan Umum</p>
                    <label htmlFor="apply-changes">Terapkan perubahan ke</label>
                    <select
                      id="apply-changes"
                      value={applyChangesTo}
                      onChange={(event) => setApplyChangesTo(event.target.value as ListApplyScope)}
                    >
                      {APPLY_CHANGES_OPTIONS.map((option) => (
                        <option
                          key={option.value}
                          value={option.value}
                          disabled={option.value !== "WholeList" && !isWordApiDesktopListSupported}
                        >
                          {option.label}
                        </option>
                      ))}
                    </select>
                    {applyChangesTo !== "WholeList" ? (
                      <p className="form-hint">
                        Mode cakupan ini akan memakai pemisahan daftar via WordApiDesktop 1.3.
                      </p>
                    ) : null}

                    <label htmlFor="linked-style">Hubungkan level ke gaya</label>
                    <select
                      id="linked-style"
                      value={selectedLevelSettings.linkedStyle}
                      onChange={(event) =>
                        updateSelectedLevelSettings((previous) => ({
                          ...previous,
                          linkedStyle: event.target.value as LinkedHeadingStyle,
                        }))
                      }
                    >
                      {LINKED_STYLE_OPTIONS.map((option) => (
                        <option key={option.value} value={option.value}>
                          {option.label}
                        </option>
                      ))}
                    </select>

                  </div>

                  <div className="option-group">
                    <p className="option-group-title">Format Nomor</p>
                    <label htmlFor="number-formatting">Masukkan format untuk nomor</label>
                    <input
                      id="number-formatting"
                      value={selectedLevelSettings.numberFormatPattern}
                      onChange={(event) =>
                        updateSelectedLevelSettings((previous) => ({
                          ...previous,
                          numberFormatPattern: event.target.value,
                        }))
                      }
                      placeholder="Contoh: BAB <L1> atau <L1>.<L2>"
                    />
                    <p className="form-hint">
                      Gunakan token level <code>&lt;L1&gt;</code> sampai <code>&lt;L9&gt;</code>.
                      Contoh level 3: <code>&lt;L1&gt;.&lt;L2&gt;.&lt;L3&gt;</code>
                    </p>

                    <div className="inline-pair">
                      <div>
                        <label htmlFor="number-style">Gaya nomor untuk level ini</label>
                        <select
                          id="number-style"
                          value={selectedLevelSettings.numberStyle}
                          onChange={(event) =>
                            updateSelectedLevelSettings((previous) => ({
                              ...previous,
                              numberStyle: event.target.value as ListNumberingStyle,
                            }))
                          }
                        >
                          {NUMBER_STYLE_OPTIONS.map((option) => (
                            <option key={option.value} value={option.value}>
                              {option.label}
                            </option>
                          ))}
                        </select>
                      </div>
                      <div>
                        <label htmlFor="include-from">Sertakan nomor level dari</label>
                        <select
                          id="include-from"
                          value={
                            selectedLevelSettings.includeFromLevelIndex === null
                              ? "none"
                              : String(selectedLevelSettings.includeFromLevelIndex)
                          }
                          onChange={(event) =>
                            updateSelectedLevelSettings((previous) => {
                              const includeFromLevelIndex =
                                event.target.value === "none" ? null : Number(event.target.value);
                              return {
                                ...previous,
                                includeFromLevelIndex,
                                numberFormatPattern: buildNumberFormatPattern(
                                  selectedLevelIndex,
                                  includeFromLevelIndex,
                                  previous.levelSeparator,
                                  previous.prefixText,
                                  previous.suffixText
                                ),
                              };
                            })
                          }
                        >
                          <option value="none">(Tidak ada)</option>
                          {LEVEL_INDICES.filter((levelIndex) => levelIndex <= selectedLevelIndex).map(
                            (levelIndex) => (
                              <option key={`include-${levelIndex}`} value={levelIndex}>
                                Level {levelIndex + 1}
                              </option>
                            )
                          )}
                        </select>
                      </div>
                    </div>

                    <div className="inline-pair">
                      <div>
                        <label htmlFor="separator-text">Pemisah level</label>
                        <input
                          id="separator-text"
                          value={selectedLevelSettings.levelSeparator}
                          onChange={(event) =>
                            updateSelectedLevelSettings((previous) => {
                              const levelSeparator = event.target.value;
                              return {
                                ...previous,
                                levelSeparator,
                                numberFormatPattern: buildNumberFormatPattern(
                                  selectedLevelIndex,
                                  previous.includeFromLevelIndex,
                                  levelSeparator,
                                  previous.prefixText,
                                  previous.suffixText
                                ),
                              };
                            })
                          }
                          placeholder="."
                        />
                      </div>
                      <div>
                        <label htmlFor="start-at">Mulai dari</label>
                        <input
                          id="start-at"
                          type="number"
                          min={1}
                          value={selectedLevelSettings.startAt}
                          onChange={(event) =>
                            updateSelectedLevelSettings((previous) => ({
                              ...previous,
                              startAt: parsePositiveIntegerInput(event.target.value, previous.startAt),
                            }))
                          }
                        />
                      </div>
                    </div>

                    <div className="inline-pair">
                      <div>
                        <label htmlFor="follow-number-with">Ikuti nomor dengan</label>
                        <select
                          id="follow-number-with"
                          value={selectedLevelSettings.followNumberWith}
                          disabled={!isWordApiDesktopListSupported}
                          onChange={(event) =>
                            updateSelectedLevelSettings((previous) => ({
                              ...previous,
                              followNumberWith: event.target.value as FollowNumberWith,
                              addTabStopAt:
                                event.target.value === "TrailingTab" ? previous.addTabStopAt : false,
                            }))
                          }
                        >
                          {FOLLOW_NUMBER_OPTIONS.map((option) => (
                            <option key={option.value} value={option.value}>
                              {option.label}
                            </option>
                          ))}
                        </select>
                      </div>
                      <div>
                        <label htmlFor="restart-after-select">Mulai ulang daftar setelah</label>
                        <select
                          id="restart-after-select"
                          value={
                            restartAfterEnabled
                              ? String(selectedLevelSettings.restartListAfterLevelIndex)
                              : "none"
                          }
                          disabled={
                            !isWordApiDesktopListSupported ||
                            !restartAfterEnabled ||
                            restartAfterLevelOptions.length === 0
                          }
                          onChange={(event) =>
                            updateSelectedLevelSettings((previous) => ({
                              ...previous,
                              restartListAfterLevelIndex:
                                event.target.value === "none"
                                  ? null
                                  : Number(event.target.value),
                            }))
                          }
                        >
                          <option value="none">(Tidak ada)</option>
                          {restartAfterLevelOptions.map((levelIndex) => (
                            <option key={`restart-${levelIndex}`} value={levelIndex}>
                              Level {levelIndex + 1}
                            </option>
                          ))}
                        </select>
                      </div>
                    </div>

                    <label className="checkbox-row">
                      <input
                        type="checkbox"
                        checked={restartAfterEnabled}
                        disabled={
                          !isWordApiDesktopListSupported || restartAfterLevelOptions.length === 0
                        }
                        onChange={(event) =>
                          updateSelectedLevelSettings((previous) => ({
                            ...previous,
                            restartListAfterLevelIndex: event.target.checked
                              ? Math.max(0, selectedLevelIndex - 1)
                              : null,
                          }))
                        }
                      />
                      Aktifkan restart numbering setelah level lebih tinggi
                    </label>

                    <label className="checkbox-row">
                      <input
                        type="checkbox"
                        checked={selectedLevelSettings.legalStyleNumbering}
                        disabled={!isWordApiDesktopListSupported}
                        onChange={(event) =>
                          updateSelectedLevelSettings((previous) => ({
                            ...previous,
                            legalStyleNumbering: event.target.checked,
                          }))
                        }
                      />
                      Penomoran gaya legal (ubah format turunan jadi angka legal)
                    </label>

                    <div className="format-preview">
                      Format aktif: <code>{getLevelFormatPreview(selectedLevelSettings)}</code>
                    </div>
                  </div>

                  <div className="option-group">
                    <p className="option-group-title">Posisi</p>
                    <div className="inline-pair">
                      <div>
                        <label htmlFor="number-alignment">Perataan nomor</label>
                        <select
                          id="number-alignment"
                          value={selectedLevelSettings.alignment}
                          onChange={(event) =>
                            updateSelectedLevelSettings((previous) => ({
                              ...previous,
                              alignment: event.target.value as ListLevelAlignment,
                            }))
                          }
                        >
                          {ALIGNMENT_OPTIONS.map((option) => (
                            <option key={option.value} value={option.value}>
                              {option.label}
                            </option>
                          ))}
                        </select>
                      </div>
                      <div>
                        <label htmlFor="aligned-at">Rata pada (cm)</label>
                        <input
                          id="aligned-at"
                          type="text"
                          inputMode="decimal"
                          value={alignedAtInput}
                          onChange={(event) => {
                            const rawValue = event.target.value;
                            setAlignedAtInput(rawValue);
                            const parsedValue = tryParseNonNegativeDecimalInput(rawValue);
                            if (parsedValue === null) {
                              return;
                            }
                            updateSelectedLevelSettings((previous) => ({
                              ...previous,
                              alignedAtCm: parsedValue,
                            }));
                          }}
                          onBlur={() =>
                            setAlignedAtInput(
                              formatLocalizedDecimalInput(selectedLevelSettings.alignedAtCm)
                            )
                          }
                        />
                      </div>
                    </div>

                    <label htmlFor="text-indent">Indentasi teks pada (cm)</label>
                    <input
                      id="text-indent"
                      type="text"
                      inputMode="decimal"
                      value={textIndentInput}
                      onChange={(event) => {
                        const rawValue = event.target.value;
                        setTextIndentInput(rawValue);
                        const parsedValue = tryParseNonNegativeDecimalInput(rawValue);
                        if (parsedValue === null) {
                          return;
                        }
                        updateSelectedLevelSettings((previous) => ({
                          ...previous,
                          textIndentCm: parsedValue,
                        }));
                      }}
                      onBlur={() =>
                        setTextIndentInput(
                          formatLocalizedDecimalInput(selectedLevelSettings.textIndentCm)
                        )
                      }
                    />

                    <div className="inline-pair">
                      <div>
                        <label className="checkbox-row">
                          <input
                            type="checkbox"
                            checked={selectedLevelSettings.addTabStopAt}
                            disabled={
                              !isWordApiDesktopListSupported ||
                              selectedLevelSettings.followNumberWith !== "TrailingTab"
                            }
                            onChange={(event) =>
                              updateSelectedLevelSettings((previous) => ({
                                ...previous,
                                addTabStopAt: event.target.checked,
                              }))
                            }
                          />
                          Tambahkan tab stop di
                        </label>
                      </div>
                      <div>
                        <label htmlFor="tab-stop-at">Posisi tab stop (cm)</label>
                        <input
                          id="tab-stop-at"
                          type="text"
                          inputMode="decimal"
                          value={tabStopAtInput}
                          disabled={
                            !isWordApiDesktopListSupported ||
                            selectedLevelSettings.followNumberWith !== "TrailingTab" ||
                            !selectedLevelSettings.addTabStopAt
                          }
                          onChange={(event) => {
                            const rawValue = event.target.value;
                            setTabStopAtInput(rawValue);
                            const parsedValue = tryParseNonNegativeDecimalInput(rawValue);
                            if (parsedValue === null) {
                              return;
                            }
                            updateSelectedLevelSettings((previous) => ({
                              ...previous,
                              tabStopAtCm: parsedValue,
                            }));
                          }}
                          onBlur={() =>
                            setTabStopAtInput(
                              formatLocalizedDecimalInput(selectedLevelSettings.tabStopAtCm)
                            )
                          }
                        />
                      </div>
                    </div>
                  </div>
                </div>
              </div>

              <div className="button-grid">
                <button
                  disabled={!isWordReady || isBusy}
                  onClick={() =>
                    runAction(`Terapkan Pengaturan Level ${selectedLevelIndex + 1}`, async () => {
                      const result = await applyLevelSettingsToSelectionList(selectedLevelSettings, {
                        applyTo: applyChangesTo,
                        applyScopeLevelIndex: selectedLevelIndex,
                      });
                      const linkedText =
                        result.linkedParagraphs > 0
                          ? ` ${result.linkedParagraphs} paragraf dihubungkan ke gaya ${selectedLevelSettings.linkedStyle}.`
                          : "";
                      const applyScopeText =
                        applyChangesTo !== "WholeList" && !result.applyScopeApplied
                          ? " Cakupan 'Terapkan perubahan ke' tidak diterapkan karena WordApiDesktop 1.3 tidak tersedia."
                          : "";
                      const advancedRequested = usesAdvancedDesktopOptions(selectedLevelSettings);
                      const advancedText =
                        advancedRequested && !result.desktopLevelOptionsApplied
                          ? " Opsi lanjutan tidak diterapkan karena WordApiDesktop 1.3 tidak tersedia."
                          : "";

                      setNotice({
                        type: "ok",
                        text:
                          `Pengaturan level ${result.configuredLevel + 1} diterapkan ke daftar ${result.listId}.` +
                          linkedText +
                          applyScopeText +
                          advancedText,
                      });
                    })
                  }
                >
                  Terapkan Pengaturan Level {selectedLevelIndex + 1}
                </button>

                <button
                  disabled={!isWordReady || isBusy}
                  onClick={() =>
                    runAction("Terapkan Semua Pengaturan Level", async () => {
                      const payload = levelSettings.map((settings, levelIndex) => ({
                        ...settings,
                        levelIndex,
                      }));
                      const advancedRequestedCount = payload.filter((settings) =>
                        usesAdvancedDesktopOptions(settings)
                      ).length;

                      const result = await applyAllLevelSettingsToSelectionList(payload, {
                        applyTo: applyChangesTo,
                        applyScopeLevelIndex: selectedLevelIndex,
                      });
                      const applyScopeText =
                        applyChangesTo !== "WholeList" && !result.applyScopeApplied
                          ? " Cakupan 'Terapkan perubahan ke' tidak diterapkan karena WordApiDesktop 1.3 tidak tersedia."
                          : "";
                      const advancedText =
                        advancedRequestedCount > 0 &&
                        result.desktopLevelOptionsAppliedCount < advancedRequestedCount
                          ? ` Opsi lanjutan hanya diterapkan pada ${result.desktopLevelOptionsAppliedCount}/${advancedRequestedCount} level karena dukungan WordApiDesktop terbatas.`
                          : "";
                      setNotice({
                        type: "ok",
                        text:
                          `${result.configuredLevels} level berhasil diterapkan ke daftar ${result.listId}. ` +
                          `Paragraf yang dihubungkan ke gaya: ${result.linkedParagraphs}.` +
                          applyScopeText +
                          advancedText,
                      });
                    })
                  }
                >
                  Terapkan Semua Level (1-9)
                </button>
              </div>
            </>
          ) : null}

          {activeTab === "heading" ? (
            <>
              <h2>Judul dan Gaya</h2>
              <p>Terapkan gaya dari preset default untuk menjaga konsistensi format.</p>

              <label htmlFor="heading-target">Target</label>
              <select
                id="heading-target"
                value={applyTarget}
                onChange={(event) => setApplyTarget(event.target.value as ApplyTarget)}
              >
                <option value="selection">Seleksi</option>
                <option value="document">Seluruh Dokumen</option>
              </select>

              <label htmlFor="style-key">Gaya</label>
              <select
                id="style-key"
                value={styleKey}
                onChange={(event) => setStyleKey(event.target.value as PresetStyleKey)}
              >
                {STYLE_OPTIONS.map((option) => (
                  <option key={option.value} value={option.value}>
                    {option.label}
                  </option>
                ))}
              </select>

              <label htmlFor="heading-level">Level Judul</label>
              <select
                id="heading-level"
                value={headingLevel}
                onChange={(event) => setHeadingLevel(Number(event.target.value) as 1 | 2 | 3)}
              >
                <option value={1}>Judul 1</option>
                <option value={2}>Judul 2</option>
                <option value={3}>Judul 3</option>
              </select>

              <div className="button-grid">
                <button
                  className="primary"
                  disabled={!isWordReady || isBusy}
                  onClick={() =>
                    runAction("Terapkan Gaya", async () => {
                      const count = await applyStylePresetToTarget(
                        styleKey,
                        DEFAULT_PRESET.styles[styleKey],
                        applyTarget
                      );
                      setNotice({ type: "ok", text: `${count} paragraf berhasil diperbarui.` });
                    })
                  }
                >
                  Terapkan Gaya Terpilih
                </button>
                <button
                  disabled={!isWordReady || isBusy}
                  onClick={() =>
                    runAction("Terapkan Judul", async () => {
                      const count = await applyHeadingStyle(headingLevel, applyTarget);
                      setNotice({ type: "ok", text: `${count} paragraf berhasil dijadikan judul.` });
                    })
                  }
                >
                  Terapkan Judul {headingLevel}
                </button>
              </div>
            </>
          ) : null}

          {activeTab === "caption" ? (
            <>
              <h2>Keterangan Gambar/Tabel</h2>
              <p>Sisipkan keterangan konsisten yang siap dipakai untuk daftar gambar dan daftar tabel.</p>

              <label htmlFor="caption-label">Label</label>
              <select
                id="caption-label"
                value={captionLabel}
                onChange={(event) => setCaptionLabel(event.target.value as CaptionLabel)}
              >
                <option value="Figure">Gambar</option>
                <option value="Table">Tabel</option>
              </select>

              <label htmlFor="caption-title">Judul Keterangan</label>
              <input
                id="caption-title"
                value={captionTitle}
                onChange={(event) => setCaptionTitle(event.target.value)}
                placeholder="Contoh: Diagram Arsitektur Sistem"
              />

              <button
                className="primary"
                disabled={!isWordReady || !isWordApi15Supported || isBusy}
                onClick={() =>
                  runAction("Sisipkan Keterangan", async () => {
                    if (!captionTitle.trim()) {
                      throw new Error("Judul keterangan tidak boleh kosong.");
                    }

                    await insertCaption({
                      label: captionLabel,
                      separator: activeCaptionPreset.separator,
                      title: captionTitle,
                      titleCase: activeCaptionPreset.titleCase,
                      captionStyle: activeCaptionStyle,
                    });

                    setCaptionTitle("");
                  })
                }
              >
                Sisipkan Keterangan
              </button>
            </>
          ) : null}

          {activeTab === "toc" ? (
            <>
              <h2>Daftar Isi dan Daftar</h2>
              <p>Buat lalu perbarui daftar isi, daftar gambar, dan daftar tabel.</p>

              <div className="button-grid">
                <button
                  className="primary"
                  disabled={!isWordReady || !isWordApi15Supported || isBusy}
                  onClick={() => runAction("Sisipkan Daftar Isi", insertTocAtSelection)}
                >
                  Sisipkan Daftar Isi
                </button>
                <button
                  disabled={!isWordReady || !isWordApi15Supported || isBusy}
                  onClick={() => runAction("Sisipkan Daftar Gambar", insertListOfFiguresAtSelection)}
                >
                  Sisipkan Daftar Gambar
                </button>
                <button
                  disabled={!isWordReady || !isWordApi15Supported || isBusy}
                  onClick={() => runAction("Sisipkan Daftar Tabel", insertListOfTablesAtSelection)}
                >
                  Sisipkan Daftar Tabel
                </button>
                <button
                  disabled={!isWordReady || !isWordApi15Supported || isBusy}
                  onClick={() =>
                    runAction("Perbarui Daftar Isi", async () => {
                      const count = await updateTocFields();
                      setNotice({ type: "ok", text: `${count} field daftar isi diperbarui.` });
                    })
                  }
                >
                  Perbarui Daftar Isi
                </button>
                <button
                  disabled={!isWordReady || !isWordApi15Supported || isBusy}
                  onClick={() =>
                    runAction("Perbarui Daftar Gambar", async () => {
                      const count = await updateListOfFiguresFields();
                      setNotice({ type: "ok", text: `${count} field daftar gambar diperbarui.` });
                    })
                  }
                >
                  Perbarui Daftar Gambar
                </button>
                <button
                  disabled={!isWordReady || !isWordApi15Supported || isBusy}
                  onClick={() =>
                    runAction("Perbarui Daftar Tabel", async () => {
                      const count = await updateListOfTablesFields();
                      setNotice({ type: "ok", text: `${count} field daftar tabel diperbarui.` });
                    })
                  }
                >
                  Perbarui Daftar Tabel
                </button>
                <button
                  disabled={!isWordReady || !isWordApi15Supported || isBusy}
                  onClick={() =>
                    runAction("Perbarui Semua Field", async () => {
                      const count = await updateAllFields();
                      setNotice({ type: "ok", text: `${count} field berhasil diperbarui.` });
                    })
                  }
                >
                  Perbarui Semua Field
                </button>
              </div>
            </>
          ) : null}
        </section>

        <section className={`notice ${notice.type}`}>
          {isBusy ? `Menjalankan: ${busyAction}...` : notice.text}
        </section>
      </div>
    </main>
  );
}
