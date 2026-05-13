# Skripsi-Fix QA Checklist

Test document: `test-assets/Skripsi-Fix-Dummy-Test-Doc.docx`

## 1. Startup and Runtime

- [ ] Open Word desktop, sideload add-in, open the dummy doc.
- [ ] Confirm runtime status shows Word connected.
- [ ] Confirm no add-in startup error dialog appears.

Expected:
- Add-in pane loads fully.
- Buttons requiring Word host are enabled.

## 2. Preset Manager

- [ ] Change preset from dropdown and verify editor values update.
- [ ] Click `Create Copy`, rename it, click `Save to Local Library`.
- [ ] Delete the custom preset.
- [ ] Click `Reset to Built-In Campus Pack`.

Expected:
- Built-in presets cannot be deleted.
- Custom preset appears/disappears correctly.

## 3. Document Preset Save/Load/Clear

- [ ] Click `Save to This Document`.
- [ ] Change active preset to another one.
- [ ] Click `Load from Document`.
- [ ] Click `Clear Document Preset`.
- [ ] Click `Load from Document` again.

Expected:
- Load restores saved preset.
- After clear, load shows "No preset found" style error.

## 4. Style Editor

- [ ] Change values for `Body` style.
- [ ] Switch `Editing style` to `Heading 1`, change at least one value.
- [ ] Click `Sync Preset to Word Styles`.
- [ ] Click `Save to Local Library`.

Expected:
- New values persist when re-selecting that preset.
- Sync action completes without error.

## 5. Apply Style Preset

- [ ] Select 2-3 paragraphs in document.
- [ ] Target=`Selection`, style=`body`, click `Apply Style Preset`.
- [ ] Target=`Whole document`, style=`body`, click `Apply Style Preset`.

Expected:
- Selection mode updates only selected paragraphs.
- Document mode updates full document.

## 6. Chapter-Aware

- [ ] Click `Chapter-Aware Autofix`.

Expected:
- `BAB I`, `BAB II`, etc. become heading-like formatting.
- Numbered sections (`1.1`, `2.2.1`) map to heading levels.
- Caption-like lines map to caption styles.
- Heading/body/quote/caption built-in styles are applied automatically.

## 7. Heading Enforcer

- [ ] Select one normal paragraph.
- [ ] Choose `Heading 1`, click `Enforce Heading`.
- [ ] Repeat for Heading 2 and Heading 3.

Expected:
- Word built-in heading style is applied each time.

## 8. Captions

- [ ] Cursor on `[Figure Placeholder]` line, set label=`Figure`, enter title, click `Insert Caption`.
- [ ] Cursor on `[Table Placeholder]` line, set label=`Table`, enter title, click `Insert Caption`.
- [ ] Try empty title and click `Insert Caption`.

Expected:
- Figure/Table caption inserted with sequence field.
- Empty title is blocked with error notice.

## 9. TOC and Lists

- [ ] Cursor on TOC placeholder, click `Insert TOC Field`.
- [ ] Cursor on figure list placeholder, click `Insert List of Figures`.
- [ ] Cursor on table list placeholder, click `Insert List of Tables`.
- [ ] Click `Update TOC`, `Update List of Figures`, `Update List of Tables`, then `Update All Fields`.

Expected:
- Fields inserted at cursor position.
- Update actions report non-zero counts after insertion.

## 10. Audit

- [ ] Click `Run Audit` before normalization (or after manual edits).
- [ ] Confirm mismatch list appears.

Expected:
- First audit finds mismatches.
- Mismatch list clearly identifies remaining non-compliant body-like paragraphs.

## 11. JSON Export/Import

- [ ] Click `Export JSON`.
- [ ] Import that same file with file input.
- [ ] Paste JSON text in text box and click `Import Presets`.
- [ ] Try invalid JSON.

Expected:
- Valid JSON import succeeds.
- Invalid JSON shows an error without crashing.

## 12. Negative/Edge Cases

- [ ] Open `https://localhost:3000/taskpane` in browser.
- [ ] Confirm page stays stable (no client-side exception page).
- [ ] In Word, rapidly click 2-3 actions while one is running.

Expected:
- Browser preview remains stable.
- Busy action state prevents conflicting actions.
