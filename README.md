# GuitarTabber

Web app for generating first-pass fingerstyle tabs from MusicXML, PDF, and sheet-image uploads.

`MusicXML -> parse/analyze -> simple fingerstyle tab -> HTML`

## Run locally

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python3 app.py
```

Open `http://127.0.0.1:5000` and upload `.musicxml`, `.xml`, `.mxl`, `.pdf`, or sheet-image files.

## What it does now

- Upload `.musicxml`, `.xml`, `.mxl`, `.pdf`, `.png`, `.jpg`, `.jpeg`, `.webp`
- PDF/image OMR conversion via Audiveris
- Melody + bass extraction with SATB-aware voice handling
- Chord labels from explicit MusicXML chord symbols + inferred harmony fallback
- Measure-grouped tab rows (not one endless line)
- Section/repeat markers (when detected in MusicXML)
- Basic playability sanity filter (limits extreme simultaneous fret spans)
- Key estimation + capo suggestion
- Arrangement permalink/history + `.txt` download
- Transpose preview and **Save As New Arrangement**
- Tab display controls:
  - text-size slider
  - fit-to-screen button (auto-shrinks rows to avoid horizontal scroll)

## PDF/Image Uploads (OMR)

- PDF/image support uses Audiveris to convert sheet music into MusicXML first.
- Install Audiveris and ensure `audiveris` is available in `PATH`.
- Or set `AUDIVERIS_BIN` to your Audiveris executable path.
- If Audiveris is unavailable, MusicXML uploads still work.

## Railway Deploy (With Audiveris)

- This repo now includes a `Dockerfile` that installs Audiveris automatically.
- In Railway, deploy from this repo with Docker enabled (it will build from `Dockerfile`).
- Keep your `DATABASE_URL` variable set.
- Optional: set `AUDIVERIS_BIN` only if your binary path differs from `/usr/bin/audiveris`.
- After deploy, test with a PDF/image upload and check logs for any Audiveris conversion errors.

## Database

- Railway: set `DATABASE_URL` (or `Database_URL`) from your Postgres service.
- Local fallback (no env var): SQLite file `guitartabber.db`.
- Tables (`songs`, `arrangements`) are auto-created at app startup.
- Visit `/history` to revisit saved tab outputs.

## Next steps

- Improve fingering/position optimization across phrase context (not just per-slot)
- Better chord voicing and rhythm grouping for dense piano reductions
- Add stronger section labeling (Verse/Chorus/Refrain) from direction text
- Add confidence/debug panel (voice selection, detected key/chords, dropped-note reasons)
