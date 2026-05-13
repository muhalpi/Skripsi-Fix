import type { SkripsiPresetV1 } from "@/types/preset";
import { validatePreset } from "@/lib/validation/presetValidation";

const DOC_PRESET_KEY = "skripsi-fix.document.active-preset.v1";

function ensureOfficeDocumentSettings(): Office.Settings {
  if (typeof Office === "undefined" || !Office.context?.document?.settings) {
    throw new Error("Office document settings are not available in this environment.");
  }
  return Office.context.document.settings;
}

export async function saveDocumentPreset(preset: SkripsiPresetV1): Promise<void> {
  const settings = ensureOfficeDocumentSettings();
  settings.set(DOC_PRESET_KEY, JSON.stringify(preset));

  await new Promise<void>((resolve, reject) => {
    settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
        return;
      }

      reject(new Error(result.error?.message || "Failed to save document preset."));
    });
  });
}

export function loadDocumentPreset(): SkripsiPresetV1 | null {
  const settings = ensureOfficeDocumentSettings();
  const raw = settings.get(DOC_PRESET_KEY);

  if (!raw) {
    return null;
  }

  if (typeof raw !== "string") {
    return null;
  }

  try {
    return validatePreset(JSON.parse(raw));
  } catch {
    return null;
  }
}

export async function clearDocumentPreset(): Promise<void> {
  const settings = ensureOfficeDocumentSettings();
  settings.remove(DOC_PRESET_KEY);

  await new Promise<void>((resolve, reject) => {
    settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
        return;
      }

      reject(new Error(result.error?.message || "Failed to clear document preset."));
    });
  });
}
