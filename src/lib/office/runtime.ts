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
        hostName: "Unknown",
        details: "Office.js is not available. Open this page inside Word Add-in task pane.",
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
        ? "Connected to Microsoft Word."
        : host
          ? `Connected to Office host: ${String(host)}. This add-in is designed for Word.`
          : "Running outside Office host (browser preview mode).",
    };
  } catch {
    return {
      isOfficeReady: false,
      isWordHost: false,
      hostName: "Unknown",
      details: "Office runtime is still initializing. Please retry in Word task pane.",
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
