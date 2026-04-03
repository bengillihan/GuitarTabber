import os
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

<form method="post" enctype="multipart/form-data">
  <input type="file" name="music_file" accept=".musicxml,.xml,.mxl,.pdf,.png,.jpg,.jpeg,.webp" required>
  <button type="submit">Generate Tab</button>
</form>

{% if error %}
  <div class="error">{{ error }}</div>
{% endif %}

{% if result %}
  <section class="result">
    <h2>{{ result.title }}</h2>
    <p class="meta"><strong>Uploaded file:</strong> {{ result.filename }}</p>
    <p class="meta"><strong>Saved arrangement:</strong> <a href="{{ result.arrangement_url }}">Open permalink</a></p>
    <pre>{{ result.tab }}</pre>
  </section>
{% endif %}
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


def convert_sheet_to_musicxml(source_path: Path) -> Path:
    """Convert PDF/image sheet music to MusicXML using Audiveris."""
    audiveris_bin = os.getenv("AUDIVERIS_BIN", "audiveris")
    output_dir = UPLOAD_DIR / "omr_exports" / f"{source_path.stem}-{uuid4().hex[:8]}"
    output_dir.mkdir(parents=True, exist_ok=True)

    cmd = [audiveris_bin, "-batch", "-export", "-output", str(output_dir), str(source_path)]
    try:
        completed = subprocess.run(cmd, capture_output=True, text=True, check=False)
    except FileNotFoundError as exc:
        raise RuntimeError(
            "Audiveris is not installed or not found. Set AUDIVERIS_BIN or upload MusicXML directly."
        ) from exc

    if completed.returncode != 0:
        stderr = (completed.stderr or "").strip()
        stdout = (completed.stdout or "").strip()
        detail = stderr or stdout or f"exit code {completed.returncode}"
        raise RuntimeError(f"Audiveris conversion failed: {detail}")

    generated = []
    for pattern in ("*.musicxml", "*.mxl", "*.xml"):
        generated.extend(output_dir.rglob(pattern))

    if not generated:
        raise RuntimeError("Audiveris ran but no MusicXML output was produced.")

    return max(generated, key=lambda p: p.stat().st_mtime)


def quarter_to_slot(quarter_length: float) -> int:
    return max(0, int(round(quarter_length * 4)))


def find_position(midi_value: int, preferred_strings: list[int], max_fret: int = 14) -> Optional[tuple[int, int]]:
    for string_index in preferred_strings:
        fret = midi_value - STANDARD_TUNING[string_index]
        if 0 <= fret <= max_fret:
            return string_index, fret
    return None


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


def gather_events(score: stream.Score) -> tuple[list[tuple[int, int]], list[tuple[int, int]], list[int], list[tuple[int, str]], int]:
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

    total_slots = min(max(total_slots + 1, 16), MAX_SLOTS)
    return melody_events, bass_events, measure_slots, chord_events, total_slots


def arrange_tab(score: stream.Score) -> str:
    melody_events, bass_events, measure_slots, chord_events, total_slots = gather_events(score)

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

    return "\n".join(rendered)


def parse_sheet_to_tab(saved_path: Path, safe_name: str) -> dict[str, str]:
    parse_path = saved_path
    source_label = safe_name
    if omr_input_needs_conversion(safe_name):
        parse_path = convert_sheet_to_musicxml(saved_path)
        source_label = f"{safe_name} (via OMR: {parse_path.name})"

    score = converter.parse(str(parse_path))

    title = safe_name
    if score.metadata and score.metadata.title:
        title = str(score.metadata.title)

    key_name = "Unknown"
    try:
        analyzed_key = score.analyze("key")
        key_name = f"{analyzed_key.tonic.name} {analyzed_key.mode}"
    except Exception:
        pass

    tab = arrange_tab(score)
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
        "tab": "\n".join(header) + tab,
    }


@app.route("/", methods=["GET", "POST"])
def index():
    error = None
    result = None

    if request.method == "POST":
        upload = request.files.get("music_file")

        if upload is None or not upload.filename:
            error = "Please choose a MusicXML, PDF, or image file to upload."
        elif not is_allowed_file(upload.filename):
            error = "Unsupported file type. Upload MusicXML, PDF, or sheet-music image formats."
        else:
            UPLOAD_DIR.mkdir(exist_ok=True)
            safe_name = secure_filename(upload.filename)
            saved_path = UPLOAD_DIR / safe_name
            upload.save(saved_path)

            try:
                parsed = parse_sheet_to_tab(saved_path, safe_name)
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
                    "tab": parsed["tab"],
                    "arrangement_url": url_for("view_arrangement", arrangement_id=arrangement.id),
                }
            except Exception as exc:
                db.session.rollback()
                error = f"Could not generate tab from this file: {exc}"

    return render_page(HOME_BODY, error=error, result=result)


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
