"use client";

import { ChangeEvent, useCallback, useEffect, useMemo, useState } from "react";
import { DEFAULT_PRESET } from "@/lib/constants/defaultPreset";
import { PRESET_PACK_NOTICE } from "@/lib/constants/campusPresetPack";
import {
  deleteLocalPreset,
  exportPresetsJson,
  isBuiltInPresetId,
  importPresetsJson,
  loadLocalPresets,
  resetLocalPresetsToBuiltIns,
  upsertLocalPreset,
} from "@/lib/storage/localPresets";
import {
  clearDocumentPreset,
  loadDocumentPreset,
  saveDocumentPreset,
} from "@/lib/storage/documentSettings";
import { waitForOfficeReady, isRequirementSetSupported } from "@/lib/office/runtime";
import { applyStylePresetToTarget } from "@/lib/office/formatter";
import { applyHeadingStyle } from "@/lib/office/headings";
import { applyChapterAwareFormatting } from "@/lib/office/chapterAware";
import {
  getBuiltInStyleLabel,
  syncPresetToWordBuiltInStyles,
} from "@/lib/office/styleRegistry";
import { insertCaption } from "@/lib/office/captions";
import {
  insertListOfFiguresAtSelection,
  insertListOfTablesAtSelection,
  insertTocAtSelection,
  updateAllFields,
  updateListOfFiguresFields,
  updateListOfTablesFields,
  updateTocFields,
} from "@/lib/office/toc";
import { auditDocumentBody } from "@/lib/office/audit";
import {
  getDiagnosticModeEnabled,
  getLastOfficeDiagnostics,
  setDiagnosticModeEnabled as persistDiagnosticMode,
  summarizeDiagnostics,
  type OfficeActionDiagnostics,
} from "@/lib/office/diagnostics";
import {
  COMMON_FONT_CANDIDATES,
  detectInstalledFonts,
} from "@/lib/utils/fontDetection";
import type {
  Alignment,
  ApplyTarget,
  AuditReport,
  CaptionLabel,
  PresetStyleKey,
  SkripsiPresetV1,
} from "@/types/preset";

const STYLE_OPTIONS: Array<{ value: PresetStyleKey; label: string }> = [
  { value: "body", label: "Teks Utama" },
  { value: "heading1", label: "Judul 1" },
  { value: "heading2", label: "Judul 2" },
  { value: "heading3", label: "Judul 3" },
  { value: "quote", label: "Kutipan" },
  { value: "captionFigure", label: "Keterangan Gambar" },
  { value: "captionTable", label: "Keterangan Tabel" },
];

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

function clonePreset(preset: SkripsiPresetV1): SkripsiPresetV1 {
  return JSON.parse(JSON.stringify(preset)) as SkripsiPresetV1;
}

function slugify(value: string): string {
  return value
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function createPresetId(name: string): string {
  const slug = slugify(name) || "preset";
  return `${slug}-${Date.now()}`;
}

function getPresetFontNames(preset: SkripsiPresetV1): string[] {
  const names = STYLE_OPTIONS.map((item) => preset.styles[item.value].text.fontName.trim()).filter(
    (value) => value.length > 0
  );
  return Array.from(new Set(names)).sort((a, b) => a.localeCompare(b));
}

function extractErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }
  return "Terjadi kesalahan yang tidak diketahui.";
}

function downloadFile(filename: string, content: string): void {
  const blob = new Blob([content], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
}

export default function TaskpanePage() {
  const [runtime, setRuntime] = useState<RuntimeState>({
    isOfficeReady: false,
    isWordHost: false,
    hostName: "Tidak diketahui",
    details: "Memeriksa runtime Office...",
  });
  const [busyAction, setBusyAction] = useState<string>("");
  const [notice, setNotice] = useState<Notice>({
    type: "info",
    text: "Buka halaman ini dari task pane add-in Word untuk mulai memformat.",
  });

  const [presets, setPresets] = useState<SkripsiPresetV1[]>([]);
  const [selectedPresetId, setSelectedPresetId] = useState<string>(DEFAULT_PRESET.id);
  const [draftPreset, setDraftPreset] = useState<SkripsiPresetV1 | null>(null);

  const [applyTarget, setApplyTarget] = useState<ApplyTarget>("selection");
  const [styleKey, setStyleKey] = useState<PresetStyleKey>("body");
  const [styleEditorKey, setStyleEditorKey] = useState<PresetStyleKey>("body");
  const [headingLevel, setHeadingLevel] = useState<1 | 2 | 3>(1);
  const [captionLabel, setCaptionLabel] = useState<CaptionLabel>("Figure");
  const [captionTitle, setCaptionTitle] = useState<string>("");

  const [importText, setImportText] = useState<string>("");
  const [auditReport, setAuditReport] = useState<AuditReport | null>(null);
  const [diagnosticMode, setDiagnosticMode] = useState<boolean>(false);
  const [lastDiagnostics, setLastDiagnostics] = useState<OfficeActionDiagnostics | null>(null);
  const [availableFonts, setAvailableFonts] = useState<string[]>([]);
  const [fontScanBusy, setFontScanBusy] = useState<boolean>(false);

  useEffect(() => {
    const local = loadLocalPresets();
    setPresets(local);
    setSelectedPresetId(local[0]?.id ?? DEFAULT_PRESET.id);
    setDiagnosticMode(getDiagnosticModeEnabled());
    setLastDiagnostics(getLastOfficeDiagnostics());

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

  const selectedPreset = useMemo(() => {
    return presets.find((preset) => preset.id === selectedPresetId) ?? presets[0] ?? DEFAULT_PRESET;
  }, [presets, selectedPresetId]);

  useEffect(() => {
    setDraftPreset(clonePreset(selectedPreset));
  }, [selectedPreset]);

  const workingPreset = draftPreset ?? selectedPreset;
  const selectedPresetIsBuiltIn = isBuiltInPresetId(selectedPreset.id);
  const isWordReady = runtime.isOfficeReady && runtime.isWordHost;
  const isWordApi15Supported = isWordReady && isRequirementSetSupported("WordApi", "1.5");

  const captionPreset =
    captionLabel === "Figure" ? workingPreset.captions.figure : workingPreset.captions.table;
  const editingStyle = workingPreset.styles[styleEditorKey];

  const refreshDetectedFonts = useCallback(async (sourcePreset: SkripsiPresetV1): Promise<void> => {
    setFontScanBusy(true);
    try {
      const detected = await detectInstalledFonts(COMMON_FONT_CANDIDATES);
      const merged = Array.from(
        new Set([...detected, ...getPresetFontNames(sourcePreset)])
      ).sort((a, b) => a.localeCompare(b));
      setAvailableFonts(merged);
    } catch {
      setAvailableFonts(getPresetFontNames(sourcePreset));
    } finally {
      setFontScanBusy(false);
    }
  }, []);

  useEffect(() => {
    void refreshDetectedFonts(selectedPreset);
  }, [selectedPreset, refreshDetectedFonts]);

  async function runAction(actionName: string, action: () => Promise<void>): Promise<void> {
    setBusyAction(actionName);
    const previousDiagnosticTimestamp = getLastOfficeDiagnostics()?.timestamp;
    try {
      await action();
      const diagnostics = getLastOfficeDiagnostics();
      setLastDiagnostics(diagnostics);

      const hasNewDiagnostics =
        Boolean(diagnostics) && diagnostics?.timestamp !== previousDiagnosticTimestamp;

      if (diagnosticMode && hasNewDiagnostics && diagnostics?.failed) {
        setNotice({
          type: "error",
          text:
            `${actionName} selesai dengan ${diagnostics.failed} paragraf gagal diproses. ` +
            summarizeDiagnostics(diagnostics),
        });
        return;
      }

      if (diagnosticMode && hasNewDiagnostics && diagnostics?.fallbackUsed) {
        setNotice({
          type: "info",
          text: `${actionName} selesai menggunakan jalur fallback. ${summarizeDiagnostics(diagnostics)}`,
        });
        return;
      }

      setNotice({ type: "ok", text: `${actionName} berhasil.` });
    } catch (error: unknown) {
      setNotice({ type: "error", text: `${actionName} gagal: ${extractErrorMessage(error)}` });
    } finally {
      setBusyAction("");
    }
  }

  function persistPreset(nextPreset: SkripsiPresetV1, successMessage: string): void {
    const next = upsertLocalPreset(nextPreset);
    setPresets(next);
    setSelectedPresetId(nextPreset.id);
    setDraftPreset(clonePreset(nextPreset));
    setNotice({ type: "ok", text: successMessage });
  }

  function updateStyleText<K extends keyof SkripsiPresetV1["styles"]["body"]["text"]>(
    styleName: PresetStyleKey,
    key: K,
    value: SkripsiPresetV1["styles"]["body"]["text"][K]
  ): void {
    setDraftPreset((previous) => {
      const source = previous ?? clonePreset(selectedPreset);
      source.styles[styleName].text[key] = value;
      return { ...source };
    });
  }

  function updateStyleParagraph<K extends keyof SkripsiPresetV1["styles"]["body"]["paragraph"]>(
    styleName: PresetStyleKey,
    key: K,
    value: SkripsiPresetV1["styles"]["body"]["paragraph"][K]
  ): void {
    setDraftPreset((previous) => {
      const source = previous ?? clonePreset(selectedPreset);
      source.styles[styleName].paragraph[key] = value;
      return { ...source };
    });
  }

  function updateCaptionSeparator(label: CaptionLabel, value: "." | ":" | "-"): void {
    setDraftPreset((previous) => {
      const source = previous ?? clonePreset(selectedPreset);
      if (label === "Figure") {
        source.captions.figure.separator = value;
      } else {
        source.captions.table.separator = value;
      }
      return { ...source };
    });
  }

  function handleCreatePresetCopy(): void {
    const base = clonePreset(workingPreset);
    const nextName = `${base.name} Salinan`;
    const nextPreset: SkripsiPresetV1 = {
      ...base,
      id: createPresetId(nextName),
      name: nextName,
    };

    persistPreset(nextPreset, "Salinan preset berhasil dibuat.");
  }

  function handleDeletePreset(): void {
    if (selectedPresetIsBuiltIn) {
      setNotice({
        type: "info",
        text: "Preset bawaan kampus dilindungi. Buat salinan dulu sebelum menghapus.",
      });
      return;
    }

    const next = deleteLocalPreset(selectedPreset.id);
    setPresets(next);
    setSelectedPresetId(next[0]?.id ?? DEFAULT_PRESET.id);
    setNotice({ type: "ok", text: "Preset berhasil dihapus." });
  }

  function handleSavePresetToLibrary(): void {
    const cleaned = clonePreset(workingPreset);
    if (!cleaned.id || isBuiltInPresetId(cleaned.id)) {
      cleaned.id = createPresetId(cleaned.name);
      const lowerName = cleaned.name.toLowerCase();
      if (!lowerName.includes("custom") && !lowerName.includes("kustom")) {
        cleaned.name = `${cleaned.name} Kustom`;
      }
    }
    persistPreset(cleaned, "Preset berhasil disimpan ke pustaka lokal.");
  }

  function handleImportFile(event: ChangeEvent<HTMLInputElement>): void {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    const reader = new FileReader();
    reader.onload = () => {
      const text = String(reader.result ?? "");
      setImportText(text);
    };
    reader.readAsText(file);
  }

  function toggleDiagnosticMode(enabled: boolean): void {
    setDiagnosticMode(enabled);
    persistDiagnosticMode(enabled);
    setLastDiagnostics(getLastOfficeDiagnostics());
    setNotice({
      type: "info",
      text: enabled
        ? "Mode diagnostik aktif. Detail paragraf yang gagal akan direkam."
        : "Mode diagnostik nonaktif.",
    });
  }

  return (
    <main>
      <div className="shell">
        <section className="status info compact-meta">
          Host: <strong>{runtime.hostName}</strong> | Siap: <strong>{isWordReady ? "Ya" : "Tidak"}</strong> |
          WordApi 1.5: <strong>{isWordApi15Supported ? "Ya" : "Tidak"}</strong>
        </section>

        <section className="card">
          <h1 className="panel-title">Skripsi Helper</h1>
          <p className="panel-subtitle">Kontrol ringkas untuk task pane Word.</p>

          <div className="row">
            <label htmlFor="preset-select">Preset</label>
            <select
              id="preset-select"
              value={selectedPresetId}
              onChange={(event) => setSelectedPresetId(event.target.value)}
            >
              {presets.map((preset) => (
                <option key={preset.id} value={preset.id}>
                  {preset.name}
                </option>
              ))}
            </select>
            <div className="footer-note">
              Tipe: <strong>{selectedPresetIsBuiltIn ? "Bawaan" : "Kustom"}</strong>
            </div>
          </div>

          <div className="row inline">
            <div>
              <label htmlFor="apply-target">Target</label>
              <select
                id="apply-target"
                value={applyTarget}
                onChange={(event) => setApplyTarget(event.target.value as ApplyTarget)}
              >
                <option value="selection">Seleksi</option>
                <option value="document">Dokumen</option>
              </select>
            </div>
            <div>
              <label htmlFor="style-key">Gaya</label>
              <select
                id="style-key"
                value={styleKey}
                onChange={(event) => setStyleKey(event.target.value as PresetStyleKey)}
              >
                {STYLE_OPTIONS.map((item) => (
                  <option key={item.value} value={item.value}>
                    {item.label}
                  </option>
                ))}
              </select>
            </div>
          </div>

          <div className="actions">
            <button
              onClick={() =>
                runAction("Autofix berbasis bab", async () => {
                  const summary = await applyChapterAwareFormatting(workingPreset, applyTarget);
                  setNotice({
                    type: "ok",
                    text:
                      `Autofix berbasis bab selesai pada ${summary.total} paragraf. ` +
                      `J1:${summary.heading1}, J2:${summary.heading2}, J3:${summary.heading3}, ` +
                      `Utama:${summary.body}, KetGambar:${summary.captionFigure}, ` +
                      `KetTabel:${summary.captionTable}, Kutipan:${summary.quote}.`,
                  });
                })
              }
              disabled={!isWordReady || busyAction.length > 0}
            >
              Autofix Bab
            </button>
          </div>

          <div className="row inline">
            <div>
              <label htmlFor="heading-level">Level judul</label>
              <select
                id="heading-level"
                value={headingLevel}
                onChange={(event) => setHeadingLevel(Number(event.target.value) as 1 | 2 | 3)}
              >
                <option value={1}>Judul 1</option>
                <option value={2}>Judul 2</option>
                <option value={3}>Judul 3</option>
              </select>
            </div>
            <button
              onClick={() =>
                runAction("Terapkan gaya judul", async () => {
                  const count = await applyHeadingStyle(headingLevel, applyTarget);
                  setNotice({
                    type: "ok",
                    text: `Penerapan gaya judul selesai. ${count} paragraf diperbarui.`,
                  });
                })
              }
              disabled={!isWordReady || busyAction.length > 0}
            >
              Terapkan Judul
            </button>
          </div>
        </section>

        <details className="card details-card" open>
          <summary>Pustaka Preset</summary>
          <p>{PRESET_PACK_NOTICE}</p>

          <div className="row inline">
            <button onClick={handleCreatePresetCopy}>Buat Salinan</button>
            <button onClick={handleDeletePreset} disabled={selectedPresetIsBuiltIn}>
              Hapus
            </button>
          </div>
          <div className="row">
            <button
              onClick={() => {
                const reset = resetLocalPresetsToBuiltIns();
                setPresets(reset);
                setSelectedPresetId(reset[0]?.id ?? DEFAULT_PRESET.id);
                setNotice({ type: "ok", text: "Pustaka lokal direset ke paket bawaan kampus." });
              }}
            >
              Reset ke Paket Bawaan
            </button>
          </div>

          <div className="row">
            <label htmlFor="preset-name">Nama preset</label>
            <input
              id="preset-name"
              value={workingPreset.name}
              onChange={(event) =>
                setDraftPreset((previous) => {
                  const source = previous ?? clonePreset(selectedPreset);
                  source.name = event.target.value;
                  return { ...source };
                })
              }
            />
          </div>

          <div className="row inline">
            <button onClick={handleSavePresetToLibrary}>Simpan Pustaka</button>
            <button
              onClick={() =>
                runAction("Simpan preset ke dokumen", async () => {
                  await saveDocumentPreset(workingPreset);
                })
              }
              disabled={!isWordReady || busyAction.length > 0}
            >
              Simpan Dok
            </button>
          </div>

          <div className="row inline">
            <button
              onClick={() =>
                runAction("Muat preset dari dokumen", async () => {
                  const fromDoc = loadDocumentPreset();
                  if (!fromDoc) {
                    throw new Error("Preset tidak ditemukan di dokumen ini.");
                  }
                  persistPreset(fromDoc, "Preset berhasil dimuat dari dokumen.");
                })
              }
              disabled={!isWordReady || busyAction.length > 0}
            >
              Muat Dok
            </button>
            <button
              onClick={() => runAction("Hapus preset dokumen", clearDocumentPreset)}
              disabled={!isWordReady || busyAction.length > 0}
            >
              Bersihkan Dok
            </button>
          </div>

          <div className="row inline">
            <button
              onClick={() => downloadFile("skripsi-presets.json", exportPresetsJson())}
              disabled={busyAction.length > 0}
            >
              Ekspor JSON
            </button>
            <label style={{ marginBottom: 0 }}>
              <span style={{ display: "block", marginBottom: 4 }}>Impor file</span>
              <input type="file" accept="application/json,.json" onChange={handleImportFile} />
            </label>
          </div>

          <div className="row">
            <label htmlFor="import-json">Impor teks JSON</label>
            <textarea
              id="import-json"
              value={importText}
              onChange={(event) => setImportText(event.target.value)}
              placeholder="Tempel array JSON preset di sini"
            />
            <button
              onClick={() => {
                try {
                  const next = importPresetsJson(importText);
                  setPresets(next);
                  setSelectedPresetId(next[0]?.id ?? DEFAULT_PRESET.id);
                  setNotice({ type: "ok", text: "Preset berhasil diimpor dari JSON." });
                } catch (error: unknown) {
                  setNotice({ type: "error", text: `Impor gagal: ${extractErrorMessage(error)}` });
                }
              }}
            >
              Impor Preset
            </button>
          </div>
        </details>

        <details className="card details-card">
          <summary>Keterangan + Daftar Isi</summary>
          <div className="row inline">
            <div>
              <label htmlFor="caption-label">Label keterangan</label>
              <select
                id="caption-label"
                value={captionLabel}
                onChange={(event) => setCaptionLabel(event.target.value as CaptionLabel)}
              >
                <option value="Figure">Gambar</option>
                <option value="Table">Tabel</option>
              </select>
            </div>
            <div>
              <label htmlFor="caption-title">Judul keterangan</label>
              <input
                id="caption-title"
                value={captionTitle}
                onChange={(event) => setCaptionTitle(event.target.value)}
                placeholder="Judul keterangan"
              />
            </div>
          </div>

          <div className="actions">
            <button
              className="primary"
              onClick={() =>
                runAction("Sisipkan keterangan", async () => {
                  if (!captionTitle.trim()) {
                    throw new Error("Judul keterangan tidak boleh kosong.");
                  }

                  await insertCaption({
                    label: captionLabel,
                    separator: captionPreset.separator,
                    title: captionTitle,
                    titleCase: captionPreset.titleCase,
                    captionStyle:
                      captionLabel === "Figure"
                        ? workingPreset.styles.captionFigure
                        : workingPreset.styles.captionTable,
                  });
                  setCaptionTitle("");
                })
              }
              disabled={!isWordReady || busyAction.length > 0 || !isWordApi15Supported}
            >
              Sisipkan Keterangan
            </button>
            <button
              onClick={() => runAction("Sisipkan daftar isi di seleksi", insertTocAtSelection)}
              disabled={!isWordReady || busyAction.length > 0 || !isWordApi15Supported}
            >
              Sisipkan Daftar Isi
            </button>
            <button
              onClick={() => runAction("Sisipkan daftar gambar", insertListOfFiguresAtSelection)}
              disabled={!isWordReady || busyAction.length > 0 || !isWordApi15Supported}
            >
              Sisipkan Daftar Gambar
            </button>
            <button
              onClick={() => runAction("Sisipkan daftar tabel", insertListOfTablesAtSelection)}
              disabled={!isWordReady || busyAction.length > 0 || !isWordApi15Supported}
            >
              Sisipkan Daftar Tabel
            </button>
            <button
              onClick={() =>
                runAction("Perbarui field daftar isi", async () => {
                  const count = await updateTocFields();
                  setNotice({ type: "ok", text: `${count} field daftar isi diperbarui.` });
                })
              }
              disabled={!isWordReady || busyAction.length > 0 || !isWordApi15Supported}
            >
              Perbarui Daftar Isi
            </button>
            <button
              onClick={() =>
                runAction("Perbarui field daftar gambar", async () => {
                  const count = await updateListOfFiguresFields();
                  setNotice({ type: "ok", text: `${count} field daftar gambar diperbarui.` });
                })
              }
              disabled={!isWordReady || busyAction.length > 0 || !isWordApi15Supported}
            >
              Perbarui Daftar Gambar
            </button>
            <button
              onClick={() =>
                runAction("Perbarui field daftar tabel", async () => {
                  const count = await updateListOfTablesFields();
                  setNotice({ type: "ok", text: `${count} field daftar tabel diperbarui.` });
                })
              }
              disabled={!isWordReady || busyAction.length > 0 || !isWordApi15Supported}
            >
              Perbarui Daftar Tabel
            </button>
            <button
              onClick={() =>
                runAction("Perbarui semua field", async () => {
                  const count = await updateAllFields();
                  setNotice({ type: "ok", text: `${count} field di isi dokumen diperbarui.` });
                })
              }
              disabled={!isWordReady || busyAction.length > 0 || !isWordApi15Supported}
            >
              Perbarui Semua Field
            </button>
          </div>
        </details>

        <details className="card details-card">
          <summary>Editor Gaya</summary>
          <div className="row">
            <label htmlFor="style-editor-key">Gaya yang diedit</label>
            <select
              id="style-editor-key"
              value={styleEditorKey}
              onChange={(event) => setStyleEditorKey(event.target.value as PresetStyleKey)}
            >
              {STYLE_OPTIONS.map((item) => (
                <option key={item.value} value={item.value}>
                  {item.label}
                </option>
              ))}
            </select>
            <div className="footer-note">
              Pemetaan bawaan: <strong>{getBuiltInStyleLabel(styleEditorKey)}</strong> | Font:
              <strong> {availableFonts.length}</strong>
            </div>
            <button
              onClick={() => {
                void refreshDetectedFonts(workingPreset);
              }}
              disabled={fontScanBusy || busyAction.length > 0}
            >
              {fontScanBusy ? "Memindai Font..." : "Pindai Ulang Font"}
            </button>
          </div>

          <div className="row inline">
            <div>
              <label htmlFor="font-name">Font</label>
              <select
                id="font-name"
                value={editingStyle.text.fontName}
                onChange={(event) =>
                  updateStyleText(styleEditorKey, "fontName", event.target.value)
                }
              >
                {availableFonts.map((fontName) => (
                  <option key={fontName} value={fontName}>
                    {fontName}
                  </option>
                ))}
              </select>
            </div>
            <div>
              <label htmlFor="font-size">Ukuran font (pt)</label>
              <input
                id="font-size"
                type="number"
                value={editingStyle.text.fontSizePt}
                onChange={(event) =>
                  updateStyleText(styleEditorKey, "fontSizePt", Number(event.target.value))
                }
              />
            </div>
          </div>

          <div className="row inline">
            <label style={{ marginBottom: 0 }}>
              <input
                type="checkbox"
                checked={editingStyle.text.bold}
                onChange={(event) => updateStyleText(styleEditorKey, "bold", event.target.checked)}
                style={{ width: "auto", marginRight: 8 }}
              />
              Tebal
            </label>
            <label style={{ marginBottom: 0 }}>
              <input
                type="checkbox"
                checked={editingStyle.text.italic}
                onChange={(event) =>
                  updateStyleText(styleEditorKey, "italic", event.target.checked)
                }
                style={{ width: "auto", marginRight: 8 }}
              />
              Miring
            </label>
          </div>

          <div className="row inline">
            <label style={{ marginBottom: 0 }}>
              <input
                type="checkbox"
                checked={editingStyle.text.underline === "Single"}
                onChange={(event) =>
                  updateStyleText(
                    styleEditorKey,
                    "underline",
                    event.target.checked ? "Single" : "None"
                  )
                }
                style={{ width: "auto", marginRight: 8 }}
              />
              Garis bawah
            </label>
            <label style={{ marginBottom: 0 }}>
              <input
                type="checkbox"
                checked={editingStyle.text.allCaps}
                onChange={(event) =>
                  updateStyleText(styleEditorKey, "allCaps", event.target.checked)
                }
                style={{ width: "auto", marginRight: 8 }}
              />
              Huruf kapital semua
            </label>
          </div>

          <div className="row inline">
            <div>
              <label htmlFor="alignment">Perataan</label>
              <select
                id="alignment"
                value={editingStyle.paragraph.alignment}
                onChange={(event) =>
                  updateStyleParagraph(styleEditorKey, "alignment", event.target.value as Alignment)
                }
              >
                <option value="Left">Kiri</option>
                <option value="Centered">Tengah</option>
                <option value="Right">Kanan</option>
                <option value="Justified">Rata kiri-kanan</option>
              </select>
            </div>
            <div>
              <label htmlFor="line-spacing">Spasi baris (pt)</label>
              <input
                id="line-spacing"
                type="number"
                value={editingStyle.paragraph.lineSpacingPt}
                onChange={(event) =>
                  updateStyleParagraph(styleEditorKey, "lineSpacingPt", Number(event.target.value))
                }
              />
            </div>
          </div>

          <div className="row inline">
            <div>
              <label htmlFor="space-before">Jarak sebelum (pt)</label>
              <input
                id="space-before"
                type="number"
                value={editingStyle.paragraph.spaceBeforePt}
                onChange={(event) =>
                  updateStyleParagraph(styleEditorKey, "spaceBeforePt", Number(event.target.value))
                }
              />
            </div>
            <div>
              <label htmlFor="space-after">Jarak sesudah (pt)</label>
              <input
                id="space-after"
                type="number"
                value={editingStyle.paragraph.spaceAfterPt}
                onChange={(event) =>
                  updateStyleParagraph(styleEditorKey, "spaceAfterPt", Number(event.target.value))
                }
              />
            </div>
          </div>

          <div className="row inline">
            <div>
              <label htmlFor="first-line-indent">Indent baris pertama (cm)</label>
              <input
                id="first-line-indent"
                type="number"
                step="0.01"
                value={editingStyle.paragraph.firstLineIndentCm}
                onChange={(event) =>
                  updateStyleParagraph(
                    styleEditorKey,
                    "firstLineIndentCm",
                    Number(event.target.value)
                  )
                }
              />
            </div>
            <div>
              <label htmlFor="left-indent">Indent kiri (cm)</label>
              <input
                id="left-indent"
                type="number"
                step="0.01"
                value={editingStyle.paragraph.leftIndentCm}
                onChange={(event) =>
                  updateStyleParagraph(styleEditorKey, "leftIndentCm", Number(event.target.value))
                }
              />
            </div>
          </div>

          <div className="row inline">
            <div>
              <label htmlFor="right-indent">Indent kanan (cm)</label>
              <input
                id="right-indent"
                type="number"
                step="0.01"
                value={editingStyle.paragraph.rightIndentCm}
                onChange={(event) =>
                  updateStyleParagraph(styleEditorKey, "rightIndentCm", Number(event.target.value))
                }
              />
            </div>
            <button
              onClick={() =>
                runAction("Sinkronkan preset ke gaya bawaan Word", async () => {
                  await syncPresetToWordBuiltInStyles(workingPreset);
                })
              }
              disabled={!isWordReady || busyAction.length > 0}
            >
              Sinkron ke Gaya Word
            </button>
          </div>

          <div className="row inline">
            <div>
              <label htmlFor="caption-separator-figure">Pemisah gambar</label>
              <select
                id="caption-separator-figure"
                value={workingPreset.captions.figure.separator}
                onChange={(event) =>
                  updateCaptionSeparator("Figure", event.target.value as "." | ":" | "-")
                }
              >
                <option value=".">.</option>
                <option value=":">:</option>
                <option value="-">-</option>
              </select>
            </div>
            <div>
              <label htmlFor="caption-separator-table">Pemisah tabel</label>
              <select
                id="caption-separator-table"
                value={workingPreset.captions.table.separator}
                onChange={(event) =>
                  updateCaptionSeparator("Table", event.target.value as "." | ":" | "-")
                }
              >
                <option value=".">.</option>
                <option value=":">:</option>
                <option value="-">-</option>
              </select>
            </div>
          </div>
        </details>

        <details className="card details-card">
          <summary>Audit</summary>
          <div className="actions">
            <button
              className="primary"
              onClick={() =>
                runAction("Audit dokumen", async () => {
                  const report = await auditDocumentBody(workingPreset.styles.body);
                  setAuditReport(report);
                  setNotice({
                    type: "ok",
                    text: `Audit selesai. ${report.mismatches.length} ketidaksesuaian dari ${report.totalParagraphs} paragraf.`,
                  });
                })
              }
              disabled={!isWordReady || busyAction.length > 0}
            >
              Jalankan Audit
            </button>
          </div>

          {auditReport ? (
            <div className="audit-list">
              <div className="audit-item">
                Paragraf utama diaudit: <strong>{auditReport.totalParagraphs}</strong> | Ketidaksesuaian:
                <strong> {auditReport.mismatches.length}</strong>
              </div>
              {auditReport.mismatches.slice(0, 20).map((item) => (
                <div key={`${item.index}-${item.textPreview}`} className="audit-item">
                  #{item.index} - {item.textPreview}
                  <br />
                  Masalah: {item.reasons.join(", ")}
                </div>
              ))}
              {auditReport.mismatches.length > 20 ? (
                <div className="audit-item">Menampilkan 20 ketidaksesuaian pertama.</div>
              ) : null}
            </div>
          ) : null}
        </details>

        <details className="card details-card">
          <summary>Diagnostik</summary>
          <p>Rekam kegagalan tingkat paragraf untuk aksi apply dan heading.</p>
          <div className="row">
            <label>
              <input
                type="checkbox"
                checked={diagnosticMode}
                onChange={(event) => toggleDiagnosticMode(event.target.checked)}
                style={{ width: "auto", marginRight: 8 }}
              />
              Aktifkan mode diagnostik
            </label>
          </div>
          <div className="row">
            <button
              onClick={() => {
                const diagnostics = getLastOfficeDiagnostics();
                if (!diagnostics) {
                  setNotice({ type: "info", text: "Belum ada data diagnostik yang direkam." });
                  return;
                }

                setLastDiagnostics(diagnostics);
                downloadFile(
                  "skripsi-helper-diagnostic-report.json",
                  JSON.stringify(diagnostics, null, 2)
                );
                setNotice({ type: "ok", text: "Laporan diagnostik berhasil diekspor." });
              }}
              disabled={busyAction.length > 0}
            >
              Ekspor JSON Diagnostik
            </button>
          </div>
          {diagnosticMode && lastDiagnostics ? (
            <div className="audit-list">
              <div className="audit-item">
                Aksi terakhir: <strong>{lastDiagnostics.operation}</strong> | Target:
                <strong> {lastDiagnostics.target}</strong> | Diperbarui:
                <strong> {lastDiagnostics.updated}</strong> | Gagal:
                <strong> {lastDiagnostics.failed}</strong>
              </div>
              {lastDiagnostics.batchError ? (
                <div className="audit-item">Error batch: {lastDiagnostics.batchError}</div>
              ) : null}
              {lastDiagnostics.failures.slice(0, 10).map((item) => (
                <div
                  key={`${item.paragraphIndex}-${item.phase}-${item.statement ?? item.error}`}
                  className="audit-item"
                >
                  #{item.paragraphIndex} ({item.phase}) - {item.textPreview}
                  <br />
                  Error: {item.error}
                  {item.errorLocation ? <span> | Lokasi: {item.errorLocation}</span> : null}
                </div>
              ))}
              {lastDiagnostics.failures.length > 10 ? (
                <div className="audit-item">Menampilkan 10 paragraf gagal pertama.</div>
              ) : null}
            </div>
          ) : null}
        </details>

        <section className={`status ${notice.type}`}>
          {busyAction ? `Menjalankan: ${busyAction}...` : notice.text}
        </section>

        <div className="sticky-bar">
          <button
            className="primary"
            onClick={() =>
              runAction("Terapkan preset gaya", async () => {
                const count = await applyStylePresetToTarget(
                  styleKey,
                  workingPreset.styles[styleKey],
                  applyTarget
                );
                setNotice({
                  type: "ok",
                  text: `Penerapan preset gaya selesai. ${count} paragraf diperbarui.`,
                });
              })
            }
            disabled={!isWordReady || busyAction.length > 0}
          >
            Terapkan Gaya
          </button>
        </div>

        <p className="footer-note">
          Field keterangan dan daftar isi memerlukan dukungan WordApi 1.5 pada host Word Anda.
        </p>
      </div>
    </main>
  );
}
