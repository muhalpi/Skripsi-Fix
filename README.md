# Skripsi Helper Word Add-in (No Auth, No DB)

Word task-pane add-in for skripsi formatting presets, caption workflows, TOC/list updates, and document audit/fix.

## Features

- Built-in campus preset pack (protected starter templates)
- Local preset library (`localStorage` partition-aware)
- Reset local library back to built-in campus pack
- Per-document preset save/load (`Office.context.document.settings`)
- Apply style preset to selection or whole document
- Chapter-aware autofix (auto-classifies heading/body/caption/quote)
- Heading 1/2/3 enforcement
- Figure/Table caption insertion
- TOC + list of figures + list of tables field helpers
- Document audit and bulk fix for body paragraph mismatches
- JSON export/import for presets

## Stack

- Next.js + React + TypeScript
- Office.js Word API
- Deployable to Vercel

## Quick start

1. Install dependencies.

```bash
npm install
```

2. Start dev server.

```bash
npm run dev
```

3. Keep manifest on localhost for dev (`https://localhost:3000`).

4. Sideload `public/manifest.xml` in Word.

## Deploy to Vercel

1. Deploy app to Vercel.
2. Update manifest URLs from localhost to Vercel domain.

```powershell
./scripts/update-manifest-url.ps1 -BaseUrl "https://YOUR-APP.vercel.app"
```

3. Re-sideload updated `public/manifest.xml`.

## Notes

- Caption and TOC/list field actions rely on Word API field support (`WordApi 1.5`).
- Works best on Word desktop; web host behavior can vary by build.
- Built-in campus packs are starter profiles. Validate against your official campus/faculty guide.
- `public/manifest.xml` is a template. Confirm `Id`, domain, and support URL before distribution.
