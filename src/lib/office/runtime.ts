export type OfficeRuntimeStatus = {
  isOfficeReady: boolean;
  isWordHost: boolean;
  hostName: string;
  details: string;
};

export async function waitForOfficeReady(): Promise<OfficeRuntimeStatus> {
  try {
    if (typeof Office === "undefined" || !Office.onReady) {
      return {
        isOfficeReady: false,
        isWordHost: false,
        hostName: "Tidak diketahui",
        details: "Office.js tidak tersedia. Buka halaman ini di task pane add-in Word.",
      };
    }

    const info = await Office.onReady();
    const host = info?.host;
    const wordHostType = Office.HostType?.Word;
    const isWordHost = Boolean(host && wordHostType && host === wordHostType);

    return {
      isOfficeReady: true,
      isWordHost,
      hostName: host ? String(host) : "Browser",
      details: isWordHost
        ? "Terhubung ke Microsoft Word."
        : host
          ? `Terhubung ke host Office: ${String(host)}. Add-in ini dirancang untuk Word.`
          : "Berjalan di luar host Office (mode pratinjau browser).",
    };
  } catch {
    return {
      isOfficeReady: false,
      isWordHost: false,
      hostName: "Tidak diketahui",
      details: "Runtime Office masih inisialisasi. Coba lagi dari task pane Word.",
    };
  }
}

export function isRequirementSetSupported(requirementSet: string, version: string): boolean {
  try {
    if (typeof Office === "undefined") {
      return false;
    }

    const requirementSupport = Office.context?.requirements;
    if (!requirementSupport || typeof requirementSupport.isSetSupported !== "function") {
      return false;
    }

    return requirementSupport.isSetSupported(requirementSet, version);
  } catch {
    return false;
  }
}
