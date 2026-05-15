# Skripsi-Fix Word Add-in (No Auth, No DB)

Word task-pane add-in untuk mempercepat format multilevel, style heading, caption, dan Table of Contents pada dokumen skripsi.

## Features

- UI baru mode tab (MultiLevel, Heading & Style, Caption, Table of Contents)
- Chapter-aware autofix untuk merapikan struktur multilevel
- Terapkan Heading 1/2/3 dan style preset default ke seleksi atau dokumen
- Sisipkan caption Figure/Table dengan format konsisten
- Sisipkan dan perbarui TOC, daftar gambar, daftar tabel, serta semua field

## Stack

- Next.js + React + TypeScript
- Office.js Word API
- Deployable to Vercel

## Quick start

1. Install dependencies.

```bash
npm install
```

2. Start HTTPS dev server (required by Office add-in runtime).

```bash
npm run dev
```

If Word still rejects startup because of certificate trust, open `https://localhost:3000/taskpane` in Edge/Chrome and accept/trust the local certificate first.

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
