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
  { value: "body", label: "Body" },
  { value: "heading1", label: "Heading 1" },
  { value: "heading2", label: "Heading 2" },
  { value: "heading3", label: "Heading 3" },
  { value: "quote", label: "Quote" },
  { value: "captionFigure", label: "Caption Figure" },
  { value: "captionTable", label: "Caption Table" },
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
  return "Unknown error.";
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
    hostName: "Unknown",
    details: "Checking Office runtime...",
  });
  const [busyAction, setBusyAction] = useState<string>("");
  const [notice, setNotice] = useState<Notice>({
    type: "info",
    text: "Load this page from Word add-in task pane to start formatting.",
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
            `${actionName} completed with ${diagnostics.failed} failed paragraph(s). ` +
            summarizeDiagnostics(diagnostics),
        });
        return;
      }

      if (diagnosticMode && hasNewDiagnostics && diagnostics?.fallbackUsed) {
        setNotice({
          type: "info",
          text: `${actionName} completed using fallback path. ${summarizeDiagnostics(diagnostics)}`,
        });
        return;
      }

      setNotice({ type: "ok", text: `${actionName} completed.` });
    } catch (error: unknown) {
      setNotice({ type: "error", text: `${actionName} failed: ${extractErrorMessage(error)}` });
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
    const nextName = `${base.name} Copy`;
    const nextPreset: SkripsiPresetV1 = {
      ...base,
      id: createPresetId(nextName),
      name: nextName,
    };

    persistPreset(nextPreset, "Preset copy created.");
  }

  function handleDeletePreset(): void {
    if (selectedPresetIsBuiltIn) {
      setNotice({
        type: "info",
        text: "Built-in campus presets are protected. Create a copy before deleting.",
      });
      return;
    }

    const next = deleteLocalPreset(selectedPreset.id);
    setPresets(next);
    setSelectedPresetId(next[0]?.id ?? DEFAULT_PRESET.id);
    setNotice({ type: "ok", text: "Preset deleted." });
  }

  function handleSavePresetToLibrary(): void {
    const cleaned = clonePreset(workingPreset);
    if (!cleaned.id || isBuiltInPresetId(cleaned.id)) {
      cleaned.id = createPresetId(cleaned.name);
      if (!cleaned.name.toLowerCase().includes("custom")) {
        cleaned.name = `${cleaned.name} Custom`;
      }
    }
    persistPreset(cleaned, "Preset saved to local library.");
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
        ? "Diagnostic mode enabled. Failed paragraph details will be captured."
        : "Diagnostic mode disabled.",
    });
  }

  return (
    <main>
      <div className="shell">
        <section className="status info compact-meta">
          Host: <strong>{runtime.hostName}</strong> | Ready: <strong>{String(isWordReady)}</strong> |
          WordApi 1.5: <strong>{String(isWordApi15Supported)}</strong>
        </section>

        <section className="card">
          <h1 className="panel-title">Skripsi Helper</h1>
          <p className="panel-subtitle">Compact controls for Word task pane.</p>

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
              Type: <strong>{selectedPresetIsBuiltIn ? "Built-in" : "Custom"}</strong>
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
                <option value="selection">Selection</option>
                <option value="document">Document</option>
              </select>
            </div>
            <div>
              <label htmlFor="style-key">Style</label>
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
                runAction("Chapter-aware autofix", async () => {
                  const summary = await applyChapterAwareFormatting(workingPreset, applyTarget);
                  setNotice({
                    type: "ok",
                    text:
                      `Chapter-aware autofix completed on ${summary.total} paragraph(s). ` +
                      `H1:${summary.heading1}, H2:${summary.heading2}, H3:${summary.heading3}, ` +
                      `Body:${summary.body}, FigCaption:${summary.captionFigure}, ` +
                      `TableCaption:${summary.captionTable}, Quote:${summary.quote}.`,
                  });
                })
              }
              disabled={!isWordReady || busyAction.length > 0}
            >
              Autofix Chapter
            </button>
          </div>

          <div className="row inline">
            <div>
              <label htmlFor="heading-level">Heading level</label>
              <select
                id="heading-level"
                value={headingLevel}
                onChange={(event) => setHeadingLevel(Number(event.target.value) as 1 | 2 | 3)}
              >
                <option value={1}>Heading 1</option>
                <option value={2}>Heading 2</option>
                <option value={3}>Heading 3</option>
              </select>
            </div>
            <button
              onClick={() =>
                runAction("Apply heading style", async () => {
                  const count = await applyHeadingStyle(headingLevel, applyTarget);
                  setNotice({
                    type: "ok",
                    text: `Apply heading style completed. Updated ${count} paragraph(s).`,
                  });
                })
              }
              disabled={!isWordReady || busyAction.length > 0}
            >
              Apply Heading
            </button>
          </div>
        </section>

        <details className="card details-card" open>
          <summary>Preset Library</summary>
          <p>{PRESET_PACK_NOTICE}</p>

          <div className="row inline">
            <button onClick={handleCreatePresetCopy}>Create Copy</button>
            <button onClick={handleDeletePreset} disabled={selectedPresetIsBuiltIn}>
              Delete
            </button>
          </div>
          <div className="row">
            <button
              onClick={() => {
                const reset = resetLocalPresetsToBuiltIns();
                setPresets(reset);
                setSelectedPresetId(reset[0]?.id ?? DEFAULT_PRESET.id);
                setNotice({ type: "ok", text: "Reset local library to built-in campus pack." });
              }}
            >
              Reset to Built-In Pack
            </button>
          </div>

          <div className="row">
            <label htmlFor="preset-name">Preset name</label>
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
            <button onClick={handleSavePresetToLibrary}>Save Library</button>
            <button
              onClick={() =>
                runAction("Save preset to document", async () => {
                  await saveDocumentPreset(workingPreset);
                })
              }
              disabled={!isWordReady || busyAction.length > 0}
            >
              Save Doc
            </button>
          </div>

          <div className="row inline">
            <button
              onClick={() =>
                runAction("Load preset from document", async () => {
                  const fromDoc = loadDocumentPreset();
                  if (!fromDoc) {
                    throw new Error("No preset found in this document.");
                  }
                  persistPreset(fromDoc, "Loaded preset from document.");
                })
              }
              disabled={!isWordReady || busyAction.length > 0}
            >
              Load Doc
            </button>
            <button
              onClick={() => runAction("Clear document preset", clearDocumentPreset)}
              disabled={!isWordReady || busyAction.length > 0}
            >
              Clear Doc
            </button>
          </div>

          <div className="row inline">
            <button
              onClick={() => downloadFile("skripsi-presets.json", exportPresetsJson())}
              disabled={busyAction.length > 0}
            >
              Export JSON
            </button>
            <label style={{ marginBottom: 0 }}>
              <span style={{ display: "block", marginBottom: 4 }}>Import file</span>
              <input type="file" accept="application/json,.json" onChange={handleImportFile} />
            </label>
          </div>

          <div className="row">
            <label htmlFor="import-json">Import JSON text</label>
            <textarea
              id="import-json"
              value={importText}
              onChange={(event) => setImportText(event.target.value)}
              placeholder="Paste preset JSON array here"
            />
            <button
              onClick={() => {
                try {
                  const next = importPresetsJson(importText);
                  setPresets(next);
                  setSelectedPresetId(next[0]?.id ?? DEFAULT_PRESET.id);
                  setNotice({ type: "ok", text: "Presets imported from JSON." });
                } catch (error: unknown) {
                  setNotice({ type: "error", text: `Import failed: ${extractErrorMessage(error)}` });
                }
              }}
            >
              Import Presets
            </button>
          </div>
        </details>

        <details className="card details-card">
          <summary>Captions + TOC</summary>
          <div className="row inline">
            <div>
              <label htmlFor="caption-label">Caption label</label>
              <select
                id="caption-label"
                value={captionLabel}
                onChange={(event) => setCaptionLabel(event.target.value as CaptionLabel)}
              >
                <option value="Figure">Figure</option>
                <option value="Table">Table</option>
              </select>
            </div>
            <div>
              <label htmlFor="caption-title">Caption title</label>
              <input
                id="caption-title"
                value={captionTitle}
                onChange={(event) => setCaptionTitle(event.target.value)}
                placeholder="Caption title"
              />
            </div>
          </div>

          <div className="actions">
            <button
              className="primary"
              onClick={() =>
                runAction("Insert caption", async () => {
                  if (!captionTitle.trim()) {
                    throw new Error("Caption title cannot be empty.");
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
              Insert Caption
            </button>
            <button
              onClick={() => runAction("Insert TOC at selection", insertTocAtSelection)}
              disabled={!isWordReady || busyAction.length > 0 || !isWordApi15Supported}
            >
              Insert TOC
            </button>
            <button
              onClick={() => runAction("Insert list of figures", insertListOfFiguresAtSelection)}
              disabled={!isWordReady || busyAction.length > 0 || !isWordApi15Supported}
            >
              Insert List Figures
            </button>
            <button
              onClick={() => runAction("Insert list of tables", insertListOfTablesAtSelection)}
              disabled={!isWordReady || busyAction.length > 0 || !isWordApi15Supported}
            >
              Insert List Tables
            </button>
            <button
              onClick={() =>
                runAction("Update TOC fields", async () => {
                  const count = await updateTocFields();
                  setNotice({ type: "ok", text: `Updated ${count} TOC field(s).` });
                })
              }
              disabled={!isWordReady || busyAction.length > 0 || !isWordApi15Supported}
            >
              Update TOC
            </button>
            <button
              onClick={() =>
                runAction("Update list of figures fields", async () => {
                  const count = await updateListOfFiguresFields();
                  setNotice({ type: "ok", text: `Updated ${count} list-of-figures field(s).` });
                })
              }
              disabled={!isWordReady || busyAction.length > 0 || !isWordApi15Supported}
            >
              Update List Figures
            </button>
            <button
              onClick={() =>
                runAction("Update list of tables fields", async () => {
                  const count = await updateListOfTablesFields();
                  setNotice({ type: "ok", text: `Updated ${count} list-of-tables field(s).` });
                })
              }
              disabled={!isWordReady || busyAction.length > 0 || !isWordApi15Supported}
            >
              Update List Tables
            </button>
            <button
              onClick={() =>
                runAction("Update all fields", async () => {
                  const count = await updateAllFields();
                  setNotice({ type: "ok", text: `Updated ${count} field(s) in document body.` });
                })
              }
              disabled={!isWordReady || busyAction.length > 0 || !isWordApi15Supported}
            >
              Update All Fields
            </button>
          </div>
        </details>

        <details className="card details-card">
          <summary>Style Editor</summary>
          <div className="row">
            <label htmlFor="style-editor-key">Editing style</label>
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
              Built-in mapping: <strong>{getBuiltInStyleLabel(styleEditorKey)}</strong> | Fonts:
              <strong> {availableFonts.length}</strong>
            </div>
            <button
              onClick={() => {
                void refreshDetectedFonts(workingPreset);
              }}
              disabled={fontScanBusy || busyAction.length > 0}
            >
              {fontScanBusy ? "Scanning Fonts..." : "Rescan Fonts"}
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
              <label htmlFor="font-size">Font size (pt)</label>
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
              Bold
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
              Italic
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
              Underline
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
              All Caps
            </label>
          </div>

          <div className="row inline">
            <div>
              <label htmlFor="alignment">Alignment</label>
              <select
                id="alignment"
                value={editingStyle.paragraph.alignment}
                onChange={(event) =>
                  updateStyleParagraph(styleEditorKey, "alignment", event.target.value as Alignment)
                }
              >
                <option value="Left">Left</option>
                <option value="Centered">Centered</option>
                <option value="Right">Right</option>
                <option value="Justified">Justified</option>
              </select>
            </div>
            <div>
              <label htmlFor="line-spacing">Line spacing (pt)</label>
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
              <label htmlFor="space-before">Space before (pt)</label>
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
              <label htmlFor="space-after">Space after (pt)</label>
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
              <label htmlFor="first-line-indent">First line indent (cm)</label>
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
              <label htmlFor="left-indent">Left indent (cm)</label>
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
              <label htmlFor="right-indent">Right indent (cm)</label>
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
                runAction("Sync preset to Word built-in styles", async () => {
                  await syncPresetToWordBuiltInStyles(workingPreset);
                })
              }
              disabled={!isWordReady || busyAction.length > 0}
            >
              Sync to Word Styles
            </button>
          </div>

          <div className="row inline">
            <div>
              <label htmlFor="caption-separator-figure">Figure separator</label>
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
              <label htmlFor="caption-separator-table">Table separator</label>
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
                runAction("Audit document", async () => {
                  const report = await auditDocumentBody(workingPreset.styles.body);
                  setAuditReport(report);
                  setNotice({
                    type: "ok",
                    text: `Audit complete. ${report.mismatches.length} mismatch(es) out of ${report.totalParagraphs} paragraph(s).`,
                  });
                })
              }
              disabled={!isWordReady || busyAction.length > 0}
            >
              Run Audit
            </button>
          </div>

          {auditReport ? (
            <div className="audit-list">
              <div className="audit-item">
                Audited body paragraphs: <strong>{auditReport.totalParagraphs}</strong> | Mismatches:
                <strong> {auditReport.mismatches.length}</strong>
              </div>
              {auditReport.mismatches.slice(0, 20).map((item) => (
                <div key={`${item.index}-${item.textPreview}`} className="audit-item">
                  #{item.index} - {item.textPreview}
                  <br />
                  Issues: {item.reasons.join(", ")}
                </div>
              ))}
              {auditReport.mismatches.length > 20 ? (
                <div className="audit-item">Showing first 20 mismatches.</div>
              ) : null}
            </div>
          ) : null}
        </details>

        <details className="card details-card">
          <summary>Diagnostics</summary>
          <p>Capture paragraph-level failures for apply and heading actions.</p>
          <div className="row">
            <label>
              <input
                type="checkbox"
                checked={diagnosticMode}
                onChange={(event) => toggleDiagnosticMode(event.target.checked)}
                style={{ width: "auto", marginRight: 8 }}
              />
              Enable diagnostic mode
            </label>
          </div>
          <div className="row">
            <button
              onClick={() => {
                const diagnostics = getLastOfficeDiagnostics();
                if (!diagnostics) {
                  setNotice({ type: "info", text: "No diagnostic data captured yet." });
                  return;
                }

                setLastDiagnostics(diagnostics);
                downloadFile(
                  "skripsi-helper-diagnostic-report.json",
                  JSON.stringify(diagnostics, null, 2)
                );
                setNotice({ type: "ok", text: "Diagnostic report exported." });
              }}
              disabled={busyAction.length > 0}
            >
              Export Diagnostic JSON
            </button>
          </div>
          {diagnosticMode && lastDiagnostics ? (
            <div className="audit-list">
              <div className="audit-item">
                Last action: <strong>{lastDiagnostics.operation}</strong> | Target:
                <strong> {lastDiagnostics.target}</strong> | Updated:
                <strong> {lastDiagnostics.updated}</strong> | Failed:
                <strong> {lastDiagnostics.failed}</strong>
              </div>
              {lastDiagnostics.batchError ? (
                <div className="audit-item">Batch error: {lastDiagnostics.batchError}</div>
              ) : null}
              {lastDiagnostics.failures.slice(0, 10).map((item) => (
                <div
                  key={`${item.paragraphIndex}-${item.phase}-${item.statement ?? item.error}`}
                  className="audit-item"
                >
                  #{item.paragraphIndex} ({item.phase}) - {item.textPreview}
                  <br />
                  Error: {item.error}
                  {item.errorLocation ? <span> | Location: {item.errorLocation}</span> : null}
                </div>
              ))}
              {lastDiagnostics.failures.length > 10 ? (
                <div className="audit-item">Showing first 10 failed paragraphs.</div>
              ) : null}
            </div>
          ) : null}
        </details>

        <section className={`status ${notice.type}`}>
          {busyAction ? `Running: ${busyAction}...` : notice.text}
        </section>

        <div className="sticky-bar">
          <button
            className="primary"
            onClick={() =>
              runAction("Apply style preset", async () => {
                const count = await applyStylePresetToTarget(
                  styleKey,
                  workingPreset.styles[styleKey],
                  applyTarget
                );
                setNotice({
                  type: "ok",
                  text: `Apply style preset completed. Updated ${count} paragraph(s).`,
                });
              })
            }
            disabled={!isWordReady || busyAction.length > 0}
          >
            Apply Style
          </button>
        </div>

        <p className="footer-note">
          Caption and TOC fields require WordApi 1.5 support in your Word host.
        </p>
      </div>
    </main>
  );
}
