# Diagnostic Mode Guide

Use this mode to identify which paragraph fails during formatting actions.

## How to Enable

1. Open add-in taskpane in Word.
2. In `Diagnostic Mode` card, check `Enable diagnostic mode`.

## What It Captures

- Action name (`Apply style preset`, `Chapter-aware autofix`, `Apply heading style`)
- Target (`selection` or `document`)
- Attempted paragraph count
- Updated paragraph count
- Failed paragraph count
- Per-failure details:
  - Paragraph index
  - Text preview
  - Error message
  - Error location/statement (if available from Office.js)

## Exporting Report

1. Run the action that fails.
2. Click `Export Diagnostic JSON`.
3. Share the exported JSON for debugging.

## Interpreting Results

- `fallbackUsed = false`: batch operation succeeded normally.
- `fallbackUsed = true`: batch failed; add-in switched to safe fallback path.
- `failed > 0`: some paragraphs could not be updated even in fallback mode.

## Typical Root Causes

- Protected or restricted paragraph
- Field-generated content (TOC/list fields)
- Unsupported formatting property on current Word host/build
- Paragraph context limitation (table/header/footer/structured area)
