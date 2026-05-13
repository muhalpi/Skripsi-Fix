# Skripsi-Fix Button Function Guide

This guide explains what each button does, when to use it, and what result to expect.

## Runtime Status Area

- `Host / Word Ready / WordApi 1.5`
  - Purpose: Shows whether add-in is attached to Word host and whether advanced field APIs are available.
  - Use case: Quick diagnosis when buttons appear disabled.

## Preset Manager

- `Create Copy`
  - Purpose: Clone selected preset into a new editable custom preset.
  - Use case: Start customization without modifying built-in campus template.

- `Delete Selected`
  - Purpose: Delete current custom preset from local library.
  - Use case: Clean unused custom presets.
  - Note: Built-in presets are protected and cannot be deleted.

- `Reset to Built-In Campus Pack`
  - Purpose: Restore local preset library to default built-in set.
  - Use case: Recover from broken/undesired custom presets.

- `Save to Local Library`
  - Purpose: Save current draft values into local browser/WebView storage.
  - Use case: Keep personalized formatting profile for future docs.

- `Save to This Document`
  - Purpose: Store active preset into Word document settings.
  - Use case: Share consistent formatting context with the document file.

- `Load from Document`
  - Purpose: Read preset from current Word document settings.
  - Use case: Re-open a doc and restore its formatting profile quickly.

- `Clear Document Preset`
  - Purpose: Remove saved document-level preset metadata.
  - Use case: Reset document-specific preset binding.

- `Export JSON`
  - Purpose: Download all local presets as JSON.
  - Use case: Backup/migrate preset library.

- `Import Presets`
  - Purpose: Replace local preset library from imported JSON text/file.
  - Use case: Restore backup or distribute standardized preset pack.

## Style Editor

- `Editing style` dropdown
  - Purpose: Pick which style profile you are editing (`Body`, `Heading 1/2/3`, `Quote`, `Caption Figure`, `Caption Table`).
  - Use case: Customize each style independently, like Word styles.

- Editable controls: `Font`, `Font size`, `Bold`, `Italic`, `Underline`, `All Caps`, `Alignment`, `Line spacing`, `Space before/after`, `First line indent`, `Left/Right indent`.
  - Purpose: Fine-tune the selected style profile before applying.
  - Use case: Match faculty-specific formatting for every paragraph class.

- `Sync Preset to Word Styles`
  - Purpose: Push current preset values into Word built-in styles (`Normal`, `Heading 1/2/3`, `Quote`, `Caption`).
  - Use case: Keep Word-native style gallery and TOC behavior aligned with your preset.

## Formatting Actions

- `Apply Style Preset`
  - Purpose: Apply selected style key (`Body`, `Heading 1/2/3`, `Quote`, `Caption`) to target.
  - Target options:
    - `Selection`: selected paragraphs only.
    - `Whole document`: all body paragraphs.
  - Use case: Controlled, manual formatting.

- `Chapter-Aware Autofix`
  - Purpose: Auto-classify paragraphs (heading/body/quote/caption) and apply corresponding style preset.
  - Built-in behavior: Heading/body/quote/caption built-in styles are also applied automatically.
  - Use case: Fast first-pass cleanup of draft thesis chapters.

- `Enforce Heading`
  - Purpose: Apply chosen Word Heading level to selected target.
  - Use case: Fix heading structure manually.

Note:
- `Caption Figure` and `Caption Table` share Word built-in style `Caption`. Keep both caption presets similar when you need strict Word built-in style sync.

## Captions + TOC

- `Insert Caption`
  - Purpose: Insert Figure/Table caption with sequence field at current selection.
  - Requirement: Caption title must not be empty.
  - Use case: Proper auto-numbered captions.

- `Insert TOC Field`
  - Purpose: Insert table of contents field at cursor.
  - Use case: Build/update TOC from Heading styles.

- `Insert List of Figures`
  - Purpose: Insert list-of-figures field at cursor.
  - Use case: Generate figure index automatically.

- `Insert List of Tables`
  - Purpose: Insert list-of-tables field at cursor.
  - Use case: Generate table index automatically.

- `Update TOC`
  - Purpose: Refresh TOC fields in document.

- `Update List of Figures`
  - Purpose: Refresh list-of-figures fields.

- `Update List of Tables`
  - Purpose: Refresh list-of-tables fields.

- `Update All Fields`
  - Purpose: Refresh all fields in document body in one action.

## Audit

- `Run Audit`
  - Purpose: Compare body-like paragraphs against current body style settings.
  - Scope: Skips headings/captions/quotes so chapter structure is preserved.
  - Output: mismatch count + reason list (font, spacing, indent, alignment, etc.).
  - Use case: Verify compliance before submission.


## Notice Bar

- Status messages (`Running...`, `completed`, `failed`)
  - Purpose: Action feedback and error visibility.
  - Use case: Confirm operation success and troubleshoot quickly.
