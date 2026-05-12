export type OfficeActionFailure = {
  paragraphIndex: number;
  textPreview: string;
  phase: string;
  error: string;
  errorLocation?: string;
  statement?: string;
};

export type OfficeActionDiagnostics = {
  operation: string;
  target: string;
  attempted: number;
  updated: number;
  failed: number;
  fallbackUsed: boolean;
  batchError?: string;
  failures: OfficeActionFailure[];
  timestamp: string;
};

const DIAGNOSTIC_MODE_KEY = "skripsi-helper.diagnostic-mode.v1";
let lastDiagnostics: OfficeActionDiagnostics | null = null;

function getStorage(): Storage | null {
  if (typeof window === "undefined") {
    return null;
  }

  try {
    return window.localStorage;
  } catch {
    return null;
  }
}

export function getDiagnosticModeEnabled(): boolean {
  const storage = getStorage();
  if (!storage) {
    return false;
  }

  try {
    return storage.getItem(DIAGNOSTIC_MODE_KEY) === "1";
  } catch {
    return false;
  }
}

export function setDiagnosticModeEnabled(enabled: boolean): void {
  const storage = getStorage();
  if (!storage) {
    return;
  }

  try {
    storage.setItem(DIAGNOSTIC_MODE_KEY, enabled ? "1" : "0");
  } catch {
    // no-op
  }
}

export function setLastOfficeDiagnostics(value: OfficeActionDiagnostics | null): void {
  lastDiagnostics = value;
}

export function getLastOfficeDiagnostics(): OfficeActionDiagnostics | null {
  return lastDiagnostics;
}

export function clearLastOfficeDiagnostics(): void {
  lastDiagnostics = null;
}

export function buildTextPreview(text: string): string {
  const compact = text.replace(/\s+/g, " ").trim();
  if (!compact) {
    return "(empty paragraph)";
  }

  return compact.length > 80 ? `${compact.slice(0, 80)}...` : compact;
}

export function extractOfficeErrorDetails(error: unknown): {
  message: string;
  errorLocation?: string;
  statement?: string;
} {
  if (error instanceof Error) {
    const maybeRichError = error as Error & {
      debugInfo?: {
        message?: string;
        errorLocation?: string;
        statement?: string;
      };
    };

    return {
      message: maybeRichError.debugInfo?.message || error.message,
      errorLocation: maybeRichError.debugInfo?.errorLocation,
      statement: maybeRichError.debugInfo?.statement,
    };
  }

  return {
    message: "Unknown error.",
  };
}

export function summarizeDiagnostics(diag: OfficeActionDiagnostics): string {
  if (diag.failed === 0) {
    return `${diag.operation}: no paragraph-level failures.`;
  }

  const sample = diag.failures
    .slice(0, 3)
    .map((item) => `#${item.paragraphIndex} (${item.phase})`)
    .join(", ");

  return `${diag.operation}: ${diag.failed} paragraph(s) failed. Sample: ${sample}.`;
}
