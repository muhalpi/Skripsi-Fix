# Skripsi-Fix Smoke Test (5-10 Minutes)

Base doc: `test-assets/Skripsi-Fix-Dummy-Test-Doc.docx`

## A. Launch Health

- [ ] Run `npm run dev`.
- [ ] Open Word desktop and load add-in.
- [ ] Confirm taskpane loads without error dialog.
- [ ] Confirm status section shows Word connected.

Pass criteria:
- Add-in opens and remains stable for at least 30 seconds.

## B. Core Action Sanity

- [ ] In `Preset Manager`, click `Create Copy` then `Save to Local Library`.
- [ ] Select 2 paragraphs in doc, set target `Selection`, click `Apply Style Preset`.
- [ ] Click `Chapter-Aware Autofix`.
- [ ] Enter caption title and click `Insert Caption` once for Figure.
- [ ] Click `Run Audit`.

Pass criteria:
- Every action returns success notice.
- No crash, no frozen UI.

## C. Field Tools

- [ ] Place cursor at TOC placeholder and click `Insert TOC Field`.
- [ ] Click `Update TOC`.
- [ ] Click `Update All Fields`.

Pass criteria:
- Field insert/update actions return non-error result.

## D. Browser Stability

- [ ] Open `https://localhost:3000/taskpane` in browser.
- [ ] Wait ~20 seconds and interact with one control.

Pass criteria:
- No "Application error: a client-side exception has occurred".

## E. Exit Check

- [ ] Close and reopen Word.
- [ ] Open same document and add-in again.

Pass criteria:
- Add-in can reopen successfully after restart.
