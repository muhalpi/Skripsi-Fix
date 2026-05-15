import { cmToPoints } from "@/lib/utils/units";

export type ListNumberingStyle =
  | "None"
  | "Arabic"
  | "UpperRoman"
  | "LowerRoman"
  | "UpperLetter"
  | "LowerLetter";

export type ListLevelAlignment = "Left" | "Centered" | "Right";

export type FollowNumberWith = "TrailingTab" | "TrailingSpace" | "TrailingNone";
export type ListApplyScope = "WholeList" | "ThisPointForward" | "Selection";

export type LinkedHeadingStyle =
  | "None"
  | "Heading1"
  | "Heading2"
  | "Heading3"
  | "Heading4"
  | "Heading5"
  | "Heading6"
  | "Heading7"
  | "Heading8"
  | "Heading9";

export type MultiLevelLevelSettings = {
  levelIndex: number;
  numberStyle: ListNumberingStyle;
  includeFromLevelIndex: number | null;
  levelSeparator: string;
  prefixText: string;
  suffixText: string;
  numberFormatPattern: string;
  startAt: number;
  alignment: ListLevelAlignment;
  alignedAtCm: number;
  textIndentCm: number;
  followNumberWith: FollowNumberWith;
  addTabStopAt: boolean;
  tabStopAtCm: number;
  legalStyleNumbering: boolean;
  restartListAfterLevelIndex: number | null;
  linkedStyle: LinkedHeadingStyle;
};

export type EnsureSelectionListResult = {
  listId: number;
  paragraphCount: number;
  createdFromSelection: boolean;
};

type ApplySettingsOptions = {
  applyTo?: ListApplyScope;
  applyScopeLevelIndex?: number;
};

function clampLevel(levelIndex: number): number {
  return Math.max(0, Math.min(8, Math.floor(levelIndex)));
}

function clampStartAt(startAt: number): number {
  const normalized = Number.isFinite(startAt) ? Math.floor(startAt) : 1;
  return Math.max(1, normalized);
}

function finiteNumber(value: number, fallback: number): number {
  return Number.isFinite(value) ? value : fallback;
}

function supportsDesktopListTemplateApi(): boolean {
  try {
    if (typeof Office === "undefined") {
      return false;
    }

    const requirements = Office.context?.requirements;
    if (!requirements || typeof requirements.isSetSupported !== "function") {
      return false;
    }

    return requirements.isSetSupported("WordApiDesktop", "1.3");
  } catch {
    return false;
  }
}

function mapFollowNumberWith(
  value: FollowNumberWith
): Word.TrailingCharacter | "TrailingTab" | "TrailingSpace" | "TrailingNone" {
  if (value === "TrailingNone") {
    return "TrailingNone";
  }
  if (value === "TrailingSpace") {
    return "TrailingSpace";
  }
  return "TrailingTab";
}

function mapNumberStyleToBuiltInStyle(
  numberStyle: ListNumberingStyle,
  legalStyleNumbering: boolean
):
  | Word.ListBuiltInNumberStyle
  | "None"
  | "Arabic"
  | "UpperRoman"
  | "LowerRoman"
  | "UpperLetter"
  | "LowerLetter"
  | "Legal" {
  if (legalStyleNumbering) {
    return "Legal";
  }

  if (numberStyle === "Arabic") {
    return "Arabic";
  }
  if (numberStyle === "UpperRoman") {
    return "UpperRoman";
  }
  if (numberStyle === "LowerRoman") {
    return "LowerRoman";
  }
  if (numberStyle === "UpperLetter") {
    return "UpperLetter";
  }
  if (numberStyle === "LowerLetter") {
    return "LowerLetter";
  }
  return "None";
}

function mapRestartOnHigher(
  levelIndex: number,
  restartListAfterLevelIndex: number | null
): number {
  if (restartListAfterLevelIndex === null || levelIndex <= 0) {
    return 0;
  }

  const maxRestartIndex = levelIndex - 1;
  const normalized = Math.max(0, Math.min(maxRestartIndex, Math.floor(restartListAfterLevelIndex)));
  return normalized + 1;
}

function normalizeApplyScope(applyTo: ListApplyScope | undefined): ListApplyScope {
  if (applyTo === "Selection") {
    return "Selection";
  }
  if (applyTo === "ThisPointForward") {
    return "ThisPointForward";
  }
  return "WholeList";
}

function toLegacyFormatTokens(settings: MultiLevelLevelSettings): Array<string | number> {
  const levelIndex = clampLevel(settings.levelIndex);
  const includeFrom =
    settings.includeFromLevelIndex === null
      ? levelIndex
      : Math.max(0, Math.min(levelIndex, Math.floor(settings.includeFromLevelIndex)));

  const tokens: Array<string | number> = [];

  if (settings.prefixText.length > 0) {
    tokens.push(settings.prefixText);
  }

  const separator = settings.levelSeparator.length > 0 ? settings.levelSeparator : ".";
  for (let currentLevel = includeFrom; currentLevel <= levelIndex; currentLevel += 1) {
    tokens.push(currentLevel);
    if (currentLevel < levelIndex) {
      tokens.push(separator);
    }
  }

  if (settings.suffixText.length > 0) {
    tokens.push(settings.suffixText);
  }

  return tokens;
}

function parseNumberFormatPattern(
  pattern: string,
  maxLevelIndex: number
): Array<string | number> {
  if (pattern.trim().length === 0) {
    return [];
  }

  const tokens: Array<string | number> = [];
  const levelTokenRegex = /<L([1-9])>/gi;
  let lastMatchIndex = 0;
  let hasLevelToken = false;
  let match = levelTokenRegex.exec(pattern);

  while (match) {
    if (match.index > lastMatchIndex) {
      tokens.push(pattern.slice(lastMatchIndex, match.index));
    }

    const matchedLevel = Math.floor(Number(match[1])) - 1;
    if (!Number.isFinite(matchedLevel)) {
      throw new Error("Token level tidak valid pada format nomor.");
    }
    if (matchedLevel > maxLevelIndex) {
      throw new Error(
        `Token <L${matchedLevel + 1}> tidak boleh melebihi level ${maxLevelIndex + 1}.`
      );
    }

    tokens.push(matchedLevel);
    hasLevelToken = true;
    lastMatchIndex = match.index + match[0].length;
    match = levelTokenRegex.exec(pattern);
  }

  if (lastMatchIndex < pattern.length) {
    tokens.push(pattern.slice(lastMatchIndex));
  }

  if (!hasLevelToken) {
    throw new Error("Format nomor harus memuat minimal satu token level, misalnya <L1>.");
  }

  return tokens.filter((token) => token !== "");
}

function toFormatTokens(settings: MultiLevelLevelSettings): Array<string | number> {
  const levelIndex = clampLevel(settings.levelIndex);
  const pattern = settings.numberFormatPattern.trim();
  if (pattern.length > 0) {
    return parseNumberFormatPattern(settings.numberFormatPattern, levelIndex);
  }

  return toLegacyFormatTokens(settings);
}

export function getLevelFormatPreview(settings: MultiLevelLevelSettings): string {
  const prefix = (() => {
    try {
      return toFormatTokens(settings)
        .map((token) => (typeof token === "number" ? `<L${token + 1}>` : token))
        .join("");
    } catch {
      return settings.numberFormatPattern.trim().length > 0
        ? settings.numberFormatPattern
        : toLegacyFormatTokens(settings)
            .map((token) => (typeof token === "number" ? `<L${token + 1}>` : token))
            .join("");
    }
  })();

  if (settings.followNumberWith === "TrailingNone") {
    return prefix;
  }
  if (settings.followNumberWith === "TrailingSpace") {
    return `${prefix}<SPACE>`;
  }
  if (settings.addTabStopAt) {
    const tabStopAtCm = Math.max(0, finiteNumber(settings.tabStopAtCm, 0.63));
    return `${prefix}<TAB@${tabStopAtCm.toFixed(2)}cm>`;
  }
  return `${prefix}<TAB>`;
}

async function ensureSelectionInSingleListInternal(
  context: Word.RequestContext,
  baseLevelIndex: number,
  forceSelectedLevel: boolean
): Promise<{
  list: Word.List;
  paragraphs: Word.ParagraphCollection;
  createdFromSelection: boolean;
}> {
  const paragraphs = context.document.getSelection().paragraphs;
  paragraphs.load("items/isListItem");
  await context.sync();

  if (paragraphs.items.length === 0) {
    throw new Error("Tidak ada paragraf pada seleksi.");
  }

  const firstParagraph = paragraphs.items[0];
  const normalizedLevel = clampLevel(baseLevelIndex);
  let createdFromSelection = false;
  let activeList: Word.List;

  if (firstParagraph.isListItem) {
    activeList = firstParagraph.list;
    activeList.load("id");
    await context.sync();
  } else {
    activeList = firstParagraph.startNewList();
    activeList.load("id");
    if (forceSelectedLevel) {
      firstParagraph.listItem.level = normalizedLevel;
    }
    createdFromSelection = true;
    await context.sync();
  }

  for (const paragraph of paragraphs.items) {
    if (paragraph.isListItem) {
      paragraph.list.load("id");
    }
  }
  await context.sync();

  for (const paragraph of paragraphs.items) {
    if (!paragraph.isListItem) {
      paragraph.attachToList(activeList.id, normalizedLevel);
      paragraph.listItem.level = normalizedLevel;
    } else if (paragraph.list.id !== activeList.id) {
      paragraph.detachFromList();
      paragraph.attachToList(activeList.id, normalizedLevel);
      paragraph.listItem.level = normalizedLevel;
    } else if (forceSelectedLevel) {
      paragraph.listItem.level = normalizedLevel;
    }
  }
  await context.sync();

  return {
    list: activeList,
    paragraphs,
    createdFromSelection,
  };
}

async function applyLinkedStyleToLevelParagraphs(
  context: Word.RequestContext,
  list: Word.List,
  levelIndex: number,
  linkedStyle: LinkedHeadingStyle
): Promise<number> {
  if (linkedStyle === "None") {
    return 0;
  }

  const levelParagraphs = list.getLevelParagraphs(clampLevel(levelIndex));
  levelParagraphs.load("items");
  await context.sync();

  let updated = 0;
  for (const paragraph of levelParagraphs.items) {
    paragraph.styleBuiltIn = linkedStyle;
    updated += 1;
  }

  return updated;
}

async function applyDesktopLevelOptions(
  context: Word.RequestContext,
  settings: MultiLevelLevelSettings
): Promise<boolean> {
  if (!supportsDesktopListTemplateApi()) {
    return false;
  }

  const levelIndex = clampLevel(settings.levelIndex);
  const selection = context.document.getSelection();
  const listTemplate = selection.listFormat.listTemplate;
  const listLevels = listTemplate.listLevels;

  listLevels.load("items");
  await context.sync();

  if (listLevels.items.length <= levelIndex) {
    return false;
  }

  const listLevel = listLevels.items[levelIndex];
  listLevel.numberStyle = mapNumberStyleToBuiltInStyle(
    settings.numberStyle,
    settings.legalStyleNumbering
  );
  listLevel.trailingCharacter = mapFollowNumberWith(settings.followNumberWith);
  if (settings.followNumberWith === "TrailingTab" && settings.addTabStopAt) {
    listLevel.tabPosition = cmToPoints(Math.max(0, finiteNumber(settings.tabStopAtCm, 0.63)));
  } else {
    listLevel.tabPosition = 0;
  }
  listLevel.resetOnHigher = mapRestartOnHigher(levelIndex, settings.restartListAfterLevelIndex);

  await context.sync();
  return true;
}

async function applyListTemplateScope(
  context: Word.RequestContext,
  applyScopeLevelIndex: number,
  applyTo: ListApplyScope
): Promise<boolean> {
  const normalizedApplyTo = normalizeApplyScope(applyTo);
  if (normalizedApplyTo === "WholeList") {
    return true;
  }

  if (!supportsDesktopListTemplateApi()) {
    return false;
  }

  const selection = context.document.getSelection();
  const listTemplate = selection.listFormat.listTemplate;
  selection.listFormat.applyListTemplateWithLevel(listTemplate, {
    applyLevel: clampLevel(applyScopeLevelIndex) + 1,
    applyTo: normalizedApplyTo,
    continuePreviousList: true,
  });
  await context.sync();
  return true;
}

export async function ensureSelectionInSingleList(
  baseLevelIndex: number
): Promise<EnsureSelectionListResult> {
  return Word.run(async (context) => {
    const { list, paragraphs, createdFromSelection } = await ensureSelectionInSingleListInternal(
      context,
      baseLevelIndex,
      true
    );

    list.load("id");
    await context.sync();

    return {
      listId: list.id,
      paragraphCount: paragraphs.items.length,
      createdFromSelection,
    };
  });
}

export async function setSelectionParagraphLevel(
  levelIndex: number
): Promise<{
  listId: number;
  updatedParagraphs: number;
  createdFromSelection: boolean;
}> {
  return Word.run(async (context) => {
    const { list, paragraphs, createdFromSelection } = await ensureSelectionInSingleListInternal(
      context,
      levelIndex,
      true
    );

    return {
      listId: list.id,
      updatedParagraphs: paragraphs.items.length,
      createdFromSelection,
    };
  });
}

export async function applyLevelSettingsToSelectionList(
  settings: MultiLevelLevelSettings,
  options?: ApplySettingsOptions
): Promise<{
  listId: number;
  configuredLevel: number;
  linkedParagraphs: number;
  createdFromSelection: boolean;
  desktopLevelOptionsApplied: boolean;
  applyScopeApplied: boolean;
}> {
  return Word.run(async (context) => {
    const normalizedLevel = clampLevel(settings.levelIndex);
    const { createdFromSelection } = await ensureSelectionInSingleListInternal(
      context,
      normalizedLevel,
      false
    );
    const applyScopeApplied = await applyListTemplateScope(
      context,
      options?.applyScopeLevelIndex ?? normalizedLevel,
      options?.applyTo ?? "WholeList"
    );
    const { list } = await ensureSelectionInSingleListInternal(context, normalizedLevel, false);

    const textIndentPoints = cmToPoints(finiteNumber(settings.textIndentCm, 0.63));
    const alignedAtPoints = cmToPoints(finiteNumber(settings.alignedAtCm, 0));
    const firstLineIndentPoints = alignedAtPoints - textIndentPoints;

    list.setLevelNumbering(normalizedLevel, settings.numberStyle, toFormatTokens(settings));
    list.setLevelStartingNumber(normalizedLevel, clampStartAt(settings.startAt));
    list.setLevelAlignment(normalizedLevel, settings.alignment);
    list.setLevelIndents(normalizedLevel, textIndentPoints, firstLineIndentPoints);

    await context.sync();

    const linkedParagraphs = await applyLinkedStyleToLevelParagraphs(
      context,
      list,
      normalizedLevel,
      settings.linkedStyle
    );
    const desktopLevelOptionsApplied = await applyDesktopLevelOptions(context, settings);
    await context.sync();

    return {
      listId: list.id,
      configuredLevel: normalizedLevel,
      linkedParagraphs,
      createdFromSelection,
      desktopLevelOptionsApplied,
      applyScopeApplied,
    };
  });
}

export async function applyAllLevelSettingsToSelectionList(
  levelSettings: MultiLevelLevelSettings[],
  options?: ApplySettingsOptions
): Promise<{
  listId: number;
  configuredLevels: number;
  linkedParagraphs: number;
  createdFromSelection: boolean;
  desktopLevelOptionsAppliedCount: number;
  applyScopeApplied: boolean;
}> {
  if (levelSettings.length === 0) {
    throw new Error("Pengaturan level kosong.");
  }

  return Word.run(async (context) => {
    const { createdFromSelection } = await ensureSelectionInSingleListInternal(
      context,
      0,
      false
    );
    const applyScopeApplied = await applyListTemplateScope(
      context,
      options?.applyScopeLevelIndex ?? 0,
      options?.applyTo ?? "WholeList"
    );
    const { list } = await ensureSelectionInSingleListInternal(context, 0, false);

    for (const settings of levelSettings) {
      const normalizedLevel = clampLevel(settings.levelIndex);
      const textIndentPoints = cmToPoints(finiteNumber(settings.textIndentCm, 0.63));
      const alignedAtPoints = cmToPoints(finiteNumber(settings.alignedAtCm, 0));
      const firstLineIndentPoints = alignedAtPoints - textIndentPoints;

      list.setLevelNumbering(normalizedLevel, settings.numberStyle, toFormatTokens(settings));
      list.setLevelStartingNumber(normalizedLevel, clampStartAt(settings.startAt));
      list.setLevelAlignment(normalizedLevel, settings.alignment);
      list.setLevelIndents(normalizedLevel, textIndentPoints, firstLineIndentPoints);
    }

    await context.sync();

    let linkedParagraphs = 0;
    let desktopLevelOptionsAppliedCount = 0;
    for (const settings of levelSettings) {
      linkedParagraphs += await applyLinkedStyleToLevelParagraphs(
        context,
        list,
        settings.levelIndex,
        settings.linkedStyle
      );
      if (await applyDesktopLevelOptions(context, settings)) {
        desktopLevelOptionsAppliedCount += 1;
      }
    }

    await context.sync();

    return {
      listId: list.id,
      configuredLevels: levelSettings.length,
      linkedParagraphs,
      createdFromSelection,
      desktopLevelOptionsAppliedCount,
      applyScopeApplied,
    };
  });
}
