# GuitarTabber

Web app for generating guitar tabs from MusicXML, PDF, and sheet-image uploads.

`Sheet music → OMR / parse → guitar tab → HTML`

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
- **MuseScore** (primary OMR) + **Audiveris** (fallback) for PDF/image → MusicXML conversion
- Camera photo detection with a separate, heavier preprocessing pipeline (see below)
- SATB-aware voice extraction: lyric-attached part identified as melody source
- Iterative multi-pass melody extraction: each pass fills in slots missing from earlier passes
- Chord labels from explicit MusicXML symbols; root-deduplication prevents walking-bass inversions (C/E, C/G…) flooding the display when the chord hasn't changed
- Measure-grouped tab rows with empty-measure filtering
- Section/repeat markers (when detected in MusicXML)
- Playability filter — fret span limit prevents unplayable simultaneous positions
- Key estimation + capo suggestion
- Built-in quality scoring per arrangement:
  - **Accuracy score** (melody/chord coverage)
  - **Playability score** (fret-span/fret-range friendliness)

### Arrangement styles

| Style | What's included |
|---|---|
| **Solo (Melody Only)** | Single-note melody line (melody objective only) |
| **Chords (Accompaniment Only)** | Chord-shape accompaniment (no melody objective) |
| **Melody + Bass (Light Fingerstyle)** | Melody-first with supportive bass movement on strong beats |
| **Chords + Melody Fills** | Chord changes with melody fills between chord events |

### Complexity levels

Applied on top of the style:

- **Make Easier** — first-position bias, simplified chord labels, lower fret span, auto-capo to an open key
- **Standard** — balanced voicings, melody up to fret 17
- **Make More Complete** — adds inner voices, fuller harmonic placement

### Arrangement controls

- **Change key** — transpose and preview in any key, or save as a new arrangement
- **Change style or complexity** — re-render any saved song without re-uploading
- **Reprocess source** — rerun OMR/parsing from the original uploaded file with new style/complexity settings
- **Permalink** — every arrangement gets a stable URL
- **Download .txt** — plain-text tab
- **Download original file** — recover the original upload from a saved arrangement
- **Transcribed Notes Editor** — preview/save a version with selected notes removed (interactive note pruning)

### Tab display

- Text-size slider (8–20 px)
- Fit-to-screen button — auto-shrinks to avoid horizontal scroll
- Clickable fret annotations in HTML output for note-edit workflows

## Upload limits

- Default max upload size: **20 MB**
- Configurable via `MAX_UPLOAD_MB` environment variable (clamped to `5..100`)
- Friendly `413` response when files exceed the configured limit

## PDF / Image / Camera Photo Uploads (OMR)

### OMR engines

The app tries OMR engines in this order, stopping at the first success:

1. **MuseScore CLI** — best note and chord-symbol recognition for professionally typeset PDFs.
   Install: `brew install musescore` (macOS) or download from [musescore.org](https://musescore.org).
   The app also checks common macOS app bundle paths automatically.
2. **Audiveris** — fallback for when MuseScore is not installed.
   Install and add `audiveris` to `PATH`, or set `AUDIVERIS_BIN`.

If neither is installed, MusicXML uploads still work.

### Clean scans vs. camera photos

The app automatically detects whether an uploaded image is a **camera photograph** or a **clean scan / typeset PDF** and applies a different preprocessing pipeline:

| Input type | Detection | Preprocessing |
|---|---|---|
| HEIC/HEIF | Always a camera photo | Full camera pipeline |
| JPG/PNG with EXIF camera metadata (Make/Model) | Camera photo | Full camera pipeline |
| JPG/PNG without camera EXIF | Clean scan | Scan pipeline |

**Scan pipeline** (ImageMagick): grayscale → normalize → sharpen

**Camera photo pipeline**:

1. **Perspective correction** (optional, requires `opencv-python-headless`): detects the page boundary as a quadrilateral and applies a perspective warp to flatten book curl and correct camera angle
2. **ImageMagick local adaptive threshold** (`-lat 80x80-5%`): handles uneven lighting and binding shadows without washing out note heads — much more effective than global normalize for photos
3. Median denoise → sharpen → deskew (rotation correction) → trim borders

To enable perspective correction:

```bash
pip install opencv-python-headless numpy
```

Without it, the rest of the camera pipeline still runs and gives meaningful improvement over the scan defaults.

**Tips for better camera photo results:**

- Shoot straight down (directly over the page), not at an angle
- Use good, even lighting — avoid shadows from your hand or a lamp to one side
- A flatbed scanner always produces better OMR results than a camera photo

## Railway Deploy (with Audiveris)

- The included `Dockerfile` installs Audiveris automatically.
- In Railway, deploy from this repo with Docker enabled.
- Set `DATABASE_URL` from your Postgres service.
- Set `AUDIVERIS_BIN` only if your binary path differs from `/usr/bin/audiveris`.
- MuseScore is not included in the Docker image by default; add it to the Dockerfile if needed.

## Database

- Railway: set `DATABASE_URL` (or `Database_URL`) from your Postgres service.
- Local fallback: SQLite file `guitartabber.db`.
- Tables (`songs`, `arrangements`) are auto-created at startup.
- Schema migrations for new columns run automatically — existing data is preserved.
- Visit `/history` to browse saved arrangements.
- Saved songs include both:
  - normalized MusicXML bytes used for rendering
  - original uploaded file bytes for reprocess/download flows

## Next steps

- Add MuseScore to the Railway Dockerfile for production OMR quality
- Add confidence/debug panel (detected parts, note counts per pass, key/chord sources)
