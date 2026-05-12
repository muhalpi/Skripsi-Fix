function isTocCode(code: string): boolean {
  return /\bTOC\b/i.test(code);
}

function isFigureListCode(code: string): boolean {
  return /\\c\s+"Figure"/i.test(code);
}

function isTableListCode(code: string): boolean {
  return /\\c\s+"Table"/i.test(code);
}

async function updateFieldsByPredicate(predicate: (code: string) => boolean): Promise<number> {
  return Word.run(async (context) => {
    const fields = context.document.body.fields;
    fields.load("items/code");
    await context.sync();

    let updated = 0;

    for (const field of fields.items) {
      if (predicate(field.code)) {
        field.updateResult();
        updated += 1;
      }
    }

    if (updated > 0) {
      await context.sync();
    }

    return updated;
  });
}

export async function insertTocAtSelection(): Promise<void> {
  await Word.run(async (context) => {
    const range = context.document.getSelection().getRange();
    range.insertField(Word.InsertLocation.replace, "TOC", '\\o "1-3" \\h \\z \\u', false);
    await context.sync();
  });
}

export async function updateTocFields(): Promise<number> {
  return updateFieldsByPredicate((code) => isTocCode(code) && !isFigureListCode(code) && !isTableListCode(code));
}

export async function insertListOfFiguresAtSelection(): Promise<void> {
  await Word.run(async (context) => {
    const range = context.document.getSelection().getRange();
    range.insertField(Word.InsertLocation.replace, "TOC", '\\h \\z \\c "Figure"', false);
    await context.sync();
  });
}

export async function updateListOfFiguresFields(): Promise<number> {
  return updateFieldsByPredicate((code) => isTocCode(code) && isFigureListCode(code));
}

export async function insertListOfTablesAtSelection(): Promise<void> {
  await Word.run(async (context) => {
    const range = context.document.getSelection().getRange();
    range.insertField(Word.InsertLocation.replace, "TOC", '\\h \\z \\c "Table"', false);
    await context.sync();
  });
}

export async function updateListOfTablesFields(): Promise<number> {
  return updateFieldsByPredicate((code) => isTocCode(code) && isTableListCode(code));
}

export async function updateAllFields(): Promise<number> {
  return Word.run(async (context) => {
    const fields = context.document.body.fields;
    fields.load("items");
    await context.sync();

    for (const field of fields.items) {
      field.updateResult();
    }

    if (fields.items.length > 0) {
      await context.sync();
    }

    return fields.items.length;
  });
}
