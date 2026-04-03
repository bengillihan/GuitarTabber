# GuitarTabber MVP

Minimal web app for the first pipeline step:

`MusicXML -> parse/analyze -> simple fingerstyle tab -> HTML`

## Run locally

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python3 app.py
```

Open `http://127.0.0.1:5000` and upload `.musicxml`, `.xml`, `.mxl`, `.pdf`, or sheet-image files.

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

## Current scope

- Upload MusicXML, PDF, or sheet images
- Parse score with `music21`
- Build a basic arrangement:
  - melody biased to high strings
  - bass biased to low strings
  - lightweight chord labels from vertical sonorities
- Render ASCII tab in an Ultimate Guitar-style `<pre>` block

## Next steps

- Improve fingering/position optimization
- Better chord voicing and rhythm grouping
- Add section detection (Verse/Chorus)
- Add transpose + capo suggestions
