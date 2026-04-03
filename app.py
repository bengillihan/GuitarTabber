import os
import shutil
import subprocess
from pathlib import Path
from typing import Optional
from uuid import uuid4

from flask import Flask, abort, render_template_string, request, url_for
from flask_sqlalchemy import SQLAlchemy
from music21 import chord, converter, harmony, note, pitch, stream
from werkzeug.utils import secure_filename


UPLOAD_DIR = Path("uploads")
MUSICXML_EXTENSIONS = {"musicxml", "xml", "mxl"}
PDF_EXTENSIONS = {"pdf"}
IMAGE_EXTENSIONS = {"png", "jpg", "jpeg", "webp", "bmp", "tif", "tiff"}
ALLOWED_EXTENSIONS = MUSICXML_EXTENSIONS | PDF_EXTENSIONS | IMAGE_EXTENSIONS

# MIDI note numbers for standard tuning, low E to high E.
STANDARD_TUNING = [40, 45, 50, 55, 59, 64]
STRING_NAMES = ["E", "A", "D", "G", "B", "E"]
SLOT_WIDTH = 3  # 16th-note slot width in monospace characters
MAX_SLOTS = 320  # Keep output readable for large files
OMR_TIMEOUT_SECONDS = 120


BASE_PAGE = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{{ page_title }}</title>
  <style>
    body {
      margin: 0;
      padding: 2rem;
      font-family: Georgia, serif;
      background: #f7f4ea;
      color: #242424;
    }
    main {
      max-width: 920px;
      margin: 0 auto;
      background: #fff;
      border: 1px solid #e0dac4;
      border-radius: 10px;
      padding: 1.5rem;
      box-shadow: 0 6px 24px rgba(0, 0, 0, 0.08);
    }
    h1 {
      margin-top: 0;
      font-size: 1.9rem;
    }
    a {
      color: #1f5d35;
      text-decoration: none;
      border-bottom: 1px solid transparent;
    }
    a:hover {
      border-bottom-color: #1f5d35;
    }
    p {
      line-height: 1.5;
    }
    .topnav {
      display: flex;
      justify-content: space-between;
      gap: 1rem;
      margin-bottom: 1rem;
      align-items: center;
      flex-wrap: wrap;
    }
    .hint {
      color: #5f5a47;
      margin-top: 0.5rem;
      font-size: 0.95rem;
    }
    form {
      display: flex;
      gap: 0.75rem;
      align-items: center;
      margin: 1.25rem 0;
      flex-wrap: wrap;
    }
    input[type="file"] {
      border: 1px solid #d4ccb1;
      background: #fff;
      border-radius: 6px;
      padding: 0.4rem;
    }
    button {
      border: 0;
      background: #2e6f40;
      color: #fff;
      border-radius: 6px;
      padding: 0.6rem 1rem;
      font-size: 0.95rem;
      cursor: pointer;
    }
    button:hover {
      background: #245734;
    }
    .error {
      border: 1px solid #d89f9f;
      background: #fff1f1;
      color: #7d1a1a;
      border-radius: 6px;
      padding: 0.65rem;
      margin: 1rem 0;
    }
    .warning {
      border: 1px solid #d7bf8a;
      background: #fff8e7;
      color: #745b25;
      border-radius: 6px;
      padding: 0.65rem;
      margin: 1rem 0;
    }
    .result {
      border-top: 1px solid #e6dfc9;
      margin-top: 1.25rem;
      padding-top: 1.25rem;
    }
    .meta {
      color: #4f4936;
      font-size: 0.95rem;
    }
    .history-list {
      list-style: none;
      margin: 1rem 0;
      padding: 0;
    }
    .history-list li {
      padding: 0.7rem 0;
      border-bottom: 1px solid #ece5ce;
    }
    pre {
      background: #f8f6ef;
      border: 1px solid #e2dcc6;
      border-radius: 8px;
      padding: 0.9rem;
      overflow-x: auto;
      font-family: "Courier New", Courier, monospace;
      font-size: 0.93rem;
      line-height: 1.35;
    }
    #processing {
      display: none;
      margin-top: 0.75rem;
      color: #2f5f3b;
      font-weight: 600;
    }
  </style>
</head>
<body>
  <main>
    {{ body|safe }}
  </main>
</body>
</html>
"""


HOME_BODY = """
<div class="topnav">
  <h1>GuitarTabber MVP</h1>
  <a href="{{ url_for('history') }}">View History</a>
</div>
<p>Upload a MusicXML file and get a first-pass fingerstyle tab.</p>
<p class="hint">Supported formats: .musicxml, .xml, .mxl, .pdf, .png, .jpg, .jpeg, .webp</p>

{% if omr_warning %}
  <div class="warning">{{ omr_warning }}</div>
{% endif %}

<form method="post" enctype="multipart/form-data" id="upload-form">
  <input type="file" name="music_file" accept=".musicxml,.xml,.mxl,.pdf,.png,.jpg,.jpeg,.webp" required>
  <button type="submit">Generate Tab</button>
</form>
<div id="processing">Processing upload... OMR on PDFs/images can take up to a minute.</div>

{% if error %}
  <div class="error">{{ error }}</div>
{% endif %}

{% if result %}
  <section class="result">
    <h2>{{ result.title }}</h2>
    <p class="meta"><strong>Uploaded file:</strong> {{ result.filename }}</p>
    <p class="meta"><strong>Estimated key:</strong> {{ result.key_name }} | <strong>Capo suggestion:</strong> {{ result.capo_suggestion }}</p>
    {% if result.truncation_warning %}
      <div class="warning">{{ result.truncation_warning }}</div>
    {% endif %}
    <p class="meta"><strong>Saved arrangement:</strong> <a href="{{ result.arrangement_url }}">Open permalink</a></p>
    <pre>{{ result.tab }}</pre>
  </section>
{% endif %}

<script>
  const form = document.getElementById("upload-form");
  const processing = document.getElementById("processing");
  if (form && processing) {
    form.addEventListener("submit", () => {
      processing.style.display = "block";
    });
  }
</script>
"""


HISTORY_BODY = """
<div class="topnav">
  <h1>Saved Arrangements</h1>
  <a href="{{ url_for('index') }}">Upload New Song</a>
</div>

{% if rows %}
  <ul class="history-list">
    {% for row in rows %}
      <li>
        <a href="{{ url_for('view_arrangement', arrangement_id=row.id) }}">{{ row.song_title }}</a>
        <div class="meta">File: {{ row.original_filename }} | {{ row.created_at }}</div>
      </li>
    {% endfor %}
  </ul>
{% else %}
  <p>No saved arrangements yet. Upload one from the home page.</p>
{% endif %}
"""


ARRANGEMENT_BODY = """
<div class="topnav">
  <h1>{{ row.song_title }}</h1>
  <a href="{{ url_for('history') }}">Back To History</a>
</div>
<p class="meta"><strong>Original file:</strong> {{ row.original_filename }}</p>
<p class="meta"><strong>Estimated key:</strong> {{ row.key_name }}</p>
<p class="meta"><strong>Saved:</strong> {{ row.created_at }}</p>
<pre>{{ row.tab_text }}</pre>
"""


def normalize_database_url(raw_url: Optional[str]) -> str:
    if not raw_url:
        return "sqlite:///guitartabber.db"
    if raw_url.startswith("postgres://"):
        return raw_url.replace("postgres://", "postgresql://", 1)
    return raw_url


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB upload limit
raw_db_url = os.getenv("DATABASE_URL") or os.getenv("Database_URL")
app.config["SQLALCHEMY_DATABASE_URI"] = normalize_database_url(raw_db_url)
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)


class Song(db.Model):
    __tablename__ = "songs"

    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(255), nullable=False)
    original_filename = db.Column(db.String(255), nullable=False)
    mime_type = db.Column(db.String(120), nullable=False)
    file_data = db.Column(db.LargeBinary, nullable=False)
    created_at = db.Column(db.DateTime(timezone=True), server_default=db.func.now(), nullable=False)


class Arrangement(db.Model):
    __tablename__ = "arrangements"

    id = db.Column(db.Integer, primary_key=True)
    song_id = db.Column(db.Integer, db.ForeignKey("songs.id", ondelete="CASCADE"), nullable=False)
    key_name = db.Column(db.String(80), nullable=False)
    tab_text = db.Column(db.Text, nullable=False)
    created_at = db.Column(db.DateTime(timezone=True), server_default=db.func.now(), nullable=False)


with app.app_context():
    db.create_all()


class OMRConversionError(Exception):
    pass


class ScoreParseError(Exception):
    pass


def render_page(body_template: str, **context: object) -> str:
    body = render_template_string(body_template, **context)
    return render_template_string(BASE_PAGE, page_title="GuitarTabber", body=body)


def is_allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def file_extension(filename: str) -> str:
    if "." not in filename:
        return ""
    return filename.rsplit(".", 1)[1].lower()


def omr_input_needs_conversion(filename: str) -> bool:
    ext = file_extension(filename)
    return ext in PDF_EXTENSIONS or ext in IMAGE_EXTENSIONS


def is_omr_available() -> bool:
    audiveris_bin = os.getenv("AUDIVERIS_BIN", "audiveris")
    resolved = shutil.which(audiveris_bin)
    return resolved is not None


def convert_sheet_to_musicxml(source_path: Path, work_dir: Path) -> list[Path]:
    """Convert PDF/image sheet music to MusicXML using Audiveris."""
    audiveris_bin = os.getenv("AUDIVERIS_BIN", "audiveris")
    output_dir = work_dir / "omr_exports"
    output_dir.mkdir(parents=True, exist_ok=True)

    cmd = [audiveris_bin, "-batch", "-export", "-output", str(output_dir), str(source_path)]
    try:
        completed = subprocess.run(cmd, capture_output=True, text=True, check=False, timeout=OMR_TIMEOUT_SECONDS)
    except subprocess.TimeoutExpired as exc:
        raise OMRConversionError(
            f"OMR timed out after {OMR_TIMEOUT_SECONDS}s. Try a smaller/cropped PDF page."
        ) from exc
    except FileNotFoundError as exc:
        raise OMRConversionError(
            "Audiveris is not installed or not found. Set AUDIVERIS_BIN or upload MusicXML directly."
        ) from exc

    if completed.returncode != 0:
        stderr = (completed.stderr or "").strip()
        stdout = (completed.stdout or "").strip()
        detail = stderr or stdout or f"exit code {completed.returncode}"
        raise OMRConversionError(f"Audiveris conversion failed: {detail}")

    generated = []
    for pattern in ("*.musicxml", "*.mxl", "*.xml"):
        generated.extend(output_dir.rglob(pattern))

    if not generated:
        raise OMRConversionError("Audiveris ran but no MusicXML output was produced.")

    return sorted(generated)


def quarter_to_slot(quarter_length: float) -> int:
    return max(0, int(round(quarter_length * 4)))


def find_position(midi_value: int, preferred_strings: list[int], max_fret: int = 14) -> Optional[tuple[int, int]]:
    candidates: list[tuple[int, int]] = []
    for string_index in preferred_strings:
        fret = midi_value - STANDARD_TUNING[string_index]
        if 0 <= fret <= max_fret:
            candidates.append((string_index, fret))
    if not candidates:
        return None
    # Prefer open strings, then lower frets, then the caller's string priority order.
    priority_index = {s: i for i, s in enumerate(preferred_strings)}
    candidates.sort(key=lambda c: (0 if c[1] == 0 else 1, c[1], priority_index.get(c[0], 99)))
    return candidates[0]


def slot_is_free(line: list[str], slot: int) -> bool:
    start = slot * SLOT_WIDTH
    return all(ch == "-" for ch in line[start : start + SLOT_WIDTH])


def place_token(line: list[str], slot: int, token: str) -> None:
    start = slot * SLOT_WIDTH
    for idx, ch in enumerate(token[:SLOT_WIDTH]):
        line[start + idx] = ch


def place_measure_dividers(lines: dict[int, list[str]], measure_slots: list[int]) -> None:
    for slot in measure_slots:
        if slot <= 0:
            continue
        for line in lines.values():
            start = slot * SLOT_WIDTH
            if start < len(line):
                line[start] = "|"


def build_chord_line(total_slots: int, chord_events: list[tuple[int, str]]) -> str:
    line = [" "] * (total_slots * SLOT_WIDTH)
    used_slots: set[int] = set()

    for slot, label in chord_events:
        if slot in used_slots or not label:
            continue
        used_slots.add(slot)
        start = slot * SLOT_WIDTH
        for idx, ch in enumerate(label):
            pos = start + idx
            if pos < len(line):
                line[pos] = ch

    return "".join(line).rstrip()


def gather_events(score: stream.Score) -> tuple[list[tuple[int, int]], list[tuple[int, int]], list[int], list[tuple[int, str]], int, bool]:
    flat = score.flatten()

    melody_events: list[tuple[int, int]] = []
    bass_events: list[tuple[int, int]] = []
    total_slots = 0

    for element in flat.notesAndRests:
        slot = quarter_to_slot(float(element.offset))
        total_slots = max(total_slots, slot + 1)

        if isinstance(element, note.Note):
            midi_value = int(element.pitch.midi)
            melody_events.append((slot, midi_value))
            if midi_value <= 57:
                bass_events.append((slot, midi_value))
        elif isinstance(element, chord.Chord):
            midi_values = sorted(int(p.midi) for p in element.pitches)
            if midi_values:
                melody_events.append((slot, midi_values[-1]))
                bass_events.append((slot, midi_values[0]))

    measure_slots: list[int] = []
    for measure in flat.getElementsByClass(stream.Measure):
        slot = quarter_to_slot(float(measure.offset))
        if slot not in measure_slots:
            measure_slots.append(slot)

    chord_events: list[tuple[int, str]] = []
    max_offset = float(flat.highestTime)
    beat = 0.0
    while beat <= max_offset:
        beat_slot = quarter_to_slot(beat)
        vertical = flat.notes.getElementsByOffset(beat, mustBeginInSpan=True, includeEndBoundary=False)
        midi_values: list[int] = []
        for item in vertical:
            if isinstance(item, note.Note):
                midi_values.append(int(item.pitch.midi))
            elif isinstance(item, chord.Chord):
                midi_values.extend(int(p.midi) for p in item.pitches)

        if len(midi_values) >= 3:
            guessed = chord.Chord([pitch.Pitch(midi=v) for v in sorted(set(midi_values))])
            symbol = harmony.chordSymbolFromChord(guessed)
            label = symbol.figure if symbol and symbol.figure else ""
            if label:
                chord_events.append((beat_slot, label))

        beat += 1.0

    unclamped_slots = max(total_slots + 1, 16)
    was_truncated = unclamped_slots > MAX_SLOTS
    total_slots = min(unclamped_slots, MAX_SLOTS)
    return melody_events, bass_events, measure_slots, chord_events, total_slots, was_truncated


def arrange_tab(score: stream.Score) -> tuple[str, bool]:
    melody_events, bass_events, measure_slots, chord_events, total_slots, was_truncated = gather_events(score)

    lines = {idx: ["-"] * (total_slots * SLOT_WIDTH) for idx in range(6)}

    for slot, midi_value in melody_events:
        if slot >= total_slots:
            continue
        pos = find_position(midi_value, preferred_strings=[5, 4, 3, 2, 1, 0])
        if pos is None:
            continue
        string_index, fret = pos
        if slot_is_free(lines[string_index], slot):
            place_token(lines[string_index], slot, str(fret))

    for slot, midi_value in bass_events:
        if slot >= total_slots:
            continue
        pos = find_position(midi_value, preferred_strings=[0, 1, 2, 3, 4, 5], max_fret=10)
        if pos is None:
            continue
        string_index, fret = pos
        if slot_is_free(lines[string_index], slot):
            place_token(lines[string_index], slot, str(fret))

    place_measure_dividers(lines, measure_slots)

    rendered = []
    chord_line = build_chord_line(total_slots, chord_events)
    if chord_line:
        rendered.append(chord_line)

    for string_index in [5, 4, 3, 2, 1, 0]:
        rendered.append(f"{STRING_NAMES[string_index]}|{''.join(lines[string_index])}|")

    return "\n".join(rendered), was_truncated


def suggest_capo(key_name: str) -> str:
    try:
        tonic_name, mode = key_name.split(" ", 1)
    except ValueError:
        return "No suggestion"

    open_majors = {"C", "G", "D", "A", "E"}
    open_minors = {"A", "E", "D"}
    tonic = pitch.Pitch(tonic_name)

    for capo in range(0, 8):
        shifted = pitch.Pitch()
        shifted.midi = tonic.midi - capo
        candidate = shifted.name
        if mode == "major" and candidate in open_majors:
            return "No capo needed" if capo == 0 else f"Capo {capo} (play in {candidate})"
        if mode == "minor" and candidate in open_minors:
            return "No capo needed" if capo == 0 else f"Capo {capo} (play in {candidate}m)"
    return "No strong capo suggestion"


def parse_sheet_to_tab(saved_path: Path, safe_name: str, work_dir: Path) -> dict[str, str]:
    parse_paths: list[Path] = [saved_path]
    source_label = safe_name
    if omr_input_needs_conversion(safe_name):
        parse_paths = convert_sheet_to_musicxml(saved_path, work_dir)
        if len(parse_paths) == 1:
            source_label = f"{safe_name} (via OMR: {parse_paths[0].name})"
        else:
            source_label = f"{safe_name} (via OMR: {len(parse_paths)} exported files, using first)"

    try:
        score = converter.parse(str(parse_paths[0]))
    except Exception as exc:
        raise ScoreParseError(f"MusicXML parsing failed: {exc}") from exc

    title = safe_name
    if score.metadata and score.metadata.title:
        title = str(score.metadata.title)

    key_name = "Unknown"
    try:
        analyzed_key = score.analyze("key")
        key_name = f"{analyzed_key.tonic.name} {analyzed_key.mode}"
    except Exception:
        pass

    tab, was_truncated = arrange_tab(score)
    capo_suggestion = suggest_capo(key_name)
    header = [
        f"# {title}",
        f"# Source file: {source_label}",
        f"# Estimated key: {key_name}",
        "# Layout: basic melody (high strings) + bass (low strings)",
        "",
    ]

    return {
        "song_title": title,
        "key_name": key_name,
        "capo_suggestion": capo_suggestion,
        "truncation_warning": (
            f"This score was truncated for display at {MAX_SLOTS} tab slots. Split into sections for full output."
            if was_truncated
            else ""
        ),
        "tab": "\n".join(header) + tab,
    }


@app.route("/", methods=["GET", "POST"])
def index():
    error = None
    result = None
    omr_warning = None
    if not is_omr_available():
        omr_warning = "OMR is currently unavailable on this server. PDF/image uploads may fail; MusicXML uploads still work."

    if request.method == "POST":
        upload = request.files.get("music_file")

        if upload is None or not upload.filename:
            error = "Please choose a MusicXML, PDF, or image file to upload."
        elif not is_allowed_file(upload.filename):
            error = "Unsupported file type. Upload MusicXML, PDF, or sheet-music image formats."
        else:
            request_id = uuid4().hex[:10]
            request_dir = UPLOAD_DIR / "requests" / request_id
            request_dir.mkdir(parents=True, exist_ok=True)
            safe_name = secure_filename(upload.filename)
            saved_path = request_dir / safe_name
            upload.save(saved_path)

            try:
                parsed = parse_sheet_to_tab(saved_path, safe_name, request_dir)
                file_bytes = saved_path.read_bytes()
                song = Song(
                    title=parsed["song_title"],
                    original_filename=safe_name,
                    mime_type=upload.mimetype or "application/octet-stream",
                    file_data=file_bytes,
                )
                db.session.add(song)
                db.session.flush()

                arrangement = Arrangement(
                    song_id=song.id,
                    key_name=parsed["key_name"],
                    tab_text=parsed["tab"],
                )
                db.session.add(arrangement)
                db.session.commit()

                result = {
                    "title": "Easy Fingerstyle Tab (Saved)",
                    "filename": safe_name,
                    "key_name": parsed["key_name"],
                    "capo_suggestion": parsed["capo_suggestion"],
                    "truncation_warning": parsed["truncation_warning"],
                    "tab": parsed["tab"],
                    "arrangement_url": url_for("view_arrangement", arrangement_id=arrangement.id),
                }
            except OMRConversionError as exc:
                db.session.rollback()
                error = f"OMR conversion failed: {exc}"
            except ScoreParseError as exc:
                db.session.rollback()
                error = str(exc)
            except Exception as exc:
                db.session.rollback()
                error = f"Could not generate tab from this file: {exc}"
            finally:
                shutil.rmtree(request_dir, ignore_errors=True)

    return render_page(HOME_BODY, error=error, result=result, omr_warning=omr_warning)


@app.route("/history", methods=["GET"])
def history():
    rows = (
        db.session.query(
            Arrangement.id.label("id"),
            Song.title.label("song_title"),
            Song.original_filename.label("original_filename"),
            Arrangement.created_at.label("created_at"),
        )
        .join(Song, Arrangement.song_id == Song.id)
        .order_by(Arrangement.created_at.desc(), Arrangement.id.desc())
        .limit(100)
        .all()
    )
    return render_page(HISTORY_BODY, rows=rows)


@app.route("/arrangement/<int:arrangement_id>", methods=["GET"])
def view_arrangement(arrangement_id: int):
    row = (
        db.session.query(
            Arrangement.id.label("id"),
            Arrangement.tab_text.label("tab_text"),
            Arrangement.key_name.label("key_name"),
            Arrangement.created_at.label("created_at"),
            Song.title.label("song_title"),
            Song.original_filename.label("original_filename"),
        )
        .join(Song, Arrangement.song_id == Song.id)
        .filter(Arrangement.id == arrangement_id)
        .first()
    )

    if row is None:
        abort(404)

    return render_page(ARRANGEMENT_BODY, row=row)


if __name__ == "__main__":
    port = int(os.getenv("PORT", "8080"))
    app.run(host="0.0.0.0", port=port, debug=False)
