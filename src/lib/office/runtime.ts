export type OfficeRuntimeStatus = {
  isOfficeReady: boolean;
  isWordHost: boolean;
  hostName: string;
  details: string;
};

export async function waitForOfficeReady(): Promise<OfficeRuntimeStatus> {
  if (typeof Office === "undefined" || !Office.onReady) {
    return {
      isOfficeReady: false,
      isWordHost: false,
      hostName: "Unknown",
      details: "Office.js is not available. Open this page inside Word Add-in task pane.",
    };
  }

  const info = await Office.onReady();
  const isWordHost = info.host === Office.HostType.Word;

  return {
    isOfficeReady: true,
    isWordHost,
    hostName: String(info.host),
    details: isWordHost
      ? "Connected to Microsoft Word."
      : `Connected to Office host: ${String(info.host)}. This add-in is designed for Word.`,
  };
}

export function isRequirementSetSupported(requirementSet: string, version: string): boolean {
  if (typeof Office === "undefined") {
    return false;
  }

  return Office.context.requirements.isSetSupported(requirementSet, version);
}
