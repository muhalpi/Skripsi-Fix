import type { SkripsiPresetV1 } from "@/types/preset";
import {
  BUILT_IN_PRESET_IDS,
  DEFAULT_PRESET_LIBRARY,
} from "@/lib/constants/campusPresetPack";
import { validatePreset } from "@/lib/validation/presetValidation";

const STORAGE_KEY = "skripsi-helper.local-presets.v1";

function getStorageKey(): string {
  const partitionKey =
    typeof Office !== "undefined" && "partitionKey" in Office.context
      ? (Office.context as Office.Context & { partitionKey?: string }).partitionKey
      : undefined;
  return partitionKey ? `${partitionKey}:${STORAGE_KEY}` : STORAGE_KEY;
}

function getStorage(): Storage | null {
  if (typeof window === "undefined") {
    return null;
  }
  return window.localStorage;
}

export function getBuiltInPresets(): SkripsiPresetV1[] {
  return [...DEFAULT_PRESET_LIBRARY];
}

export function isBuiltInPresetId(presetId: string): boolean {
  return BUILT_IN_PRESET_IDS.includes(presetId);
}

export function loadLocalPresets(): SkripsiPresetV1[] {
  const storage = getStorage();
  if (!storage) {
    return [...DEFAULT_PRESET_LIBRARY];
  }

  const raw = storage.getItem(getStorageKey());
  if (!raw) {
    return [...DEFAULT_PRESET_LIBRARY];
  }

  try {
    const parsed = JSON.parse(raw) as unknown[];
    const safe = parsed.map((item) => validatePreset(item));

    const existingIds = new Set(safe.map((preset) => preset.id));
    for (const preset of DEFAULT_PRESET_LIBRARY) {
      if (!existingIds.has(preset.id)) {
        safe.push(preset);
      }
    }

    return safe;
  } catch {
    return [...DEFAULT_PRESET_LIBRARY];
  }
}

export function saveLocalPresets(presets: SkripsiPresetV1[]): void {
  const storage = getStorage();
  if (!storage) {
    return;
  }
  storage.setItem(getStorageKey(), JSON.stringify(presets));
}

export function upsertLocalPreset(preset: SkripsiPresetV1): SkripsiPresetV1[] {
  const current = loadLocalPresets();
  const next = current.filter((item) => item.id !== preset.id);
  next.push(preset);
  saveLocalPresets(next);
  return next;
}

export function deleteLocalPreset(presetId: string): SkripsiPresetV1[] {
  if (isBuiltInPresetId(presetId)) {
    return loadLocalPresets();
  }

  const next = loadLocalPresets().filter((preset) => preset.id !== presetId);
  saveLocalPresets(next);
  return next;
}

export function exportPresetsJson(): string {
  return JSON.stringify(loadLocalPresets(), null, 2);
}

export function importPresetsJson(json: string): SkripsiPresetV1[] {
  const parsed = JSON.parse(json) as unknown[];
  const next = parsed.map((value) => validatePreset(value));
  saveLocalPresets(next);
  return next;
}

export function resetLocalPresetsToBuiltIns(): SkripsiPresetV1[] {
  const next = [...DEFAULT_PRESET_LIBRARY];
  saveLocalPresets(next);
  return next;
}
