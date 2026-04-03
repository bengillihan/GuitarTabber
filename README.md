# GuitarTabber

Web app for generating guitar tabs from MusicXML, PDF, and sheet-image uploads.

`Sheet music -> parse/analyze -> guitar tab -> HTML`

## Run locally

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python3 app.py
```

Open `http://127.0.0.1:5000` and upload a sheet music file.

## What it does

- Upload `.musicxml`, `.xml`, `.mxl`, `.pdf`, `.png`, `.jpg`, `.jpeg`, `.webp`, `.heic`, `.heif`
- PDF/image OMR conversion via Audiveris (with ImageMagick preprocessing for camera photos)
- HEIC/HEIF support — iPhone camera photos upload directly (`brew install imagemagick` required)
- Melody + bass extraction with SATB-aware voice handling
- Chord labels from explicit MusicXML chord symbols + inferred harmony fallback
- Measure-grouped tab rows with empty-measure filtering
- Section/repeat markers (when detected in MusicXML)
- Playability filter — limits simultaneous fret spans to prevent unplayable positions
- Key estimation + capo suggestion

### Arrangement styles

Choose the style that fits how you want to play the song:

| Style | What's included |
|---|---|
| **Melody Only** | Single-note melody line on upper strings; chord labels above |
| **Chords Only** | Chord voicings + bass root notes; no single-note melody |
| **Fingerstyle** | Melody (upper strings) + alternating bass (lower strings) — classic fingerpicking |
| **Chords & Melody** | Full chord shapes with melody on top — chord-melody style |

### Complexity levels

Applied on top of the style:

- **Make Easier** — first-position bias, simplified chord labels, lower fret span, auto-capo to an open key
- **Standard** — balanced voicings, up to fret 17 on melody
- **Make More Complete** — adds inner voices, fuller harmonic placement, wider fret range

### Arrangement controls

- **Change key** — transpose and preview in any key, or save as a new arrangement
- **Change style or complexity** — re-render any saved song without re-uploading
- **Permalink** — every arrangement gets a stable URL
- **Download .txt** — plain-text tab for copying into your notes app

### Tab display controls

- Text-size slider (8–20 px)
- Fit-to-screen button — auto-shrinks to avoid horizontal scroll

## PDF / Image Uploads (OMR)

- PDF and image uploads require [Audiveris](https://github.com/Audiveris/audiveris) for sheet music recognition.
- Install and ensure `audiveris` is in your `PATH`, or set the `AUDIVERIS_BIN` environment variable.
- Camera photos (HEIC, JPG) are automatically preprocessed with ImageMagick before OMR — grayscale conversion, contrast normalization, and sharpening improve note detection.
- If Audiveris is unavailable, MusicXML uploads still work.

## Railway Deploy (with Audiveris)

- The included `Dockerfile` installs Audiveris automatically.
- In Railway, deploy from this repo with Docker enabled.
- Set `DATABASE_URL` from your Postgres service.
- Set `AUDIVERIS_BIN` only if your binary path differs from `/usr/bin/audiveris`.

## Database

- Railway: set `DATABASE_URL` (or `Database_URL`) from your Postgres service.
- Local fallback: SQLite file `guitartabber.db`.
- Tables (`songs`, `arrangements`) are auto-created at startup.
- Schema migrations for new columns run automatically — existing data is preserved.
- Visit `/history` to browse saved arrangements.

## Next steps

- Improve fingering/position optimization across phrase context (not just per-slot)
- Better chord voicing and rhythm grouping for dense piano reductions
- Add confidence/debug panel (voice selection, detected key/chords, dropped-note reasons)
- Lyrics display alongside tab to help with placement and chord alignment
