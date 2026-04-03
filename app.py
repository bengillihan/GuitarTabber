import os
import shutil
import subprocess
import html
from types import SimpleNamespace
from pathlib import Path
from typing import Any, Optional
from uuid import uuid4

from flask import Flask, Response, abort, render_template_string, request, url_for
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import text
from music21 import bar, chord, converter, expressions, harmony, meter, note, pitch, stream
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
MEASURES_PER_ROW = 4
# Approximate max monospace characters (tab columns) per rendered row before wrapping.
# Row wrapping always happens at measure boundaries.
TAB_TARGET_CHARS_PER_ROW = 192
MAX_FRETTED_SPAN = 5
KEY_TONICS = ["C", "C#", "D", "Eb", "E", "F", "F#", "G", "Ab", "A", "Bb", "B"]
KEY_MODES = ["major", "minor"]


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
    .transpose-form {
      margin: 0.9rem 0;
    }
    .transpose-form select {
      border: 1px solid #d0c8af;
      background: #fff;
      border-radius: 6px;
      padding: 0.45rem 0.5rem;
      font-size: 0.95rem;
      color: #2b2b2b;
      min-width: 10rem;
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
    .tab-container {
      display: flex;
      flex-direction: column;
      gap: 1rem;
      margin-top: 0.75rem;
    }
    .tab-row {
      border: 1px solid #e2dcc6;
      border-radius: 8px;
      background: #faf8f1;
      padding: 0.7rem 0.8rem;
      overflow-x: auto;
    }
    .tab-notes {
      color: #6f654d;
      font-style: italic;
      white-space: pre;
      font-family: Menlo, Monaco, Consolas, "Courier New", monospace;
      margin-bottom: 0.25rem;
    }
    .tab-chords {
      color: #2e7d32;
      font-weight: 700;
      white-space: pre;
      font-family: Menlo, Monaco, Consolas, "Courier New", monospace;
      margin-bottom: 0.35rem;
    }
    .tab-strings {
      display: flex;
      flex-direction: column;
      gap: 0.1rem;
      font-family: Menlo, Monaco, Consolas, "Courier New", monospace;
      font-size: 0.95rem;
      white-space: pre;
      min-width: max-content;
    }
    .sl {
      line-height: 1.2;
      color: #1f2937;
    }
    .sn {
      color: #5f5a47;
      display: inline-block;
      width: 1.4rem;
    }
    @media (max-width: 700px) {
      body {
        padding: 1rem;
      }
      main {
        padding: 1rem;
      }
      .tab-chords,
      .tab-strings {
        font-size: 0.82rem;
      }
      .tab-row {
        padding: 0.55rem 0.6rem;
      }
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
<p>Upload sheet music (MusicXML, PDF, or image) and get a first-pass fingerstyle tab.</p>
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
    {% if result.multi_page_warning %}
      <div class="warning">{{ result.multi_page_warning }}</div>
    {% endif %}
    {% if result.truncation_warning %}
      <div class="warning">{{ result.truncation_warning }}</div>
    {% endif %}
    <p class="meta"><strong>Saved arrangement:</strong> <a href="{{ result.arrangement_url }}">Open permalink</a> | <a href="{{ result.download_url }}">Download .txt</a></p>
    {{ result.tab_html|safe }}
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
<p class="meta"><strong>Capo suggestion:</strong> {{ row.capo_suggestion }}</p>
<p class="meta"><strong>Saved:</strong> {{ row.created_at }}</p>
<p class="meta"><a href="{{ url_for('download_arrangement', arrangement_id=row.id) }}">Download tab as .txt</a></p>
<form method="post" class="transpose-form">
  <label for="target_key"><strong>Change key:</strong></label>
  <select id="target_key" name="target_key">
    {% for option in key_options %}
      <option value="{{ option }}" {% if option == selected_key %}selected{% endif %}>{{ option }}</option>
    {% endfor %}
  </select>
  <button type="submit">Update Tab</button>
</form>
{% if transpose_error %}
  <div class="error">{{ transpose_error }}</div>
{% endif %}
{% if transpose_note %}
  <p class="meta">{{ transpose_note }}</p>
{% endif %}
{% if row.tab_html %}
  {{ row.tab_html|safe }}
{% else %}
  <pre>{{ row.tab_text }}</pre>
{% endif %}
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
    capo_suggestion = db.Column(db.String(120), nullable=False, default="No suggestion")
    tab_text = db.Column(db.Text, nullable=False)
    tab_html = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime(timezone=True), server_default=db.func.now(), nullable=False)


with app.app_context():
    db.create_all()
    # Lightweight migrations for existing deployments.
    def _safe_add_column(sql_if_not_exists: str, sql_plain: str) -> None:
        try:
            db.session.execute(text(sql_if_not_exists))
        except Exception:
            try:
                db.session.execute(text(sql_plain))
            except Exception:
                # Already exists or unsupported syntax; continue.
                pass

    _safe_add_column(
        "ALTER TABLE arrangements ADD COLUMN IF NOT EXISTS capo_suggestion VARCHAR(120) NOT NULL DEFAULT 'No suggestion'",
        "ALTER TABLE arrangements ADD COLUMN capo_suggestion VARCHAR(120) NOT NULL DEFAULT 'No suggestion'",
    )
    _safe_add_column(
        "ALTER TABLE arrangements ADD COLUMN IF NOT EXISTS tab_html TEXT",
        "ALTER TABLE arrangements ADD COLUMN tab_html TEXT",
    )
    db.session.commit()


class OMRConversionError(Exception):
    pass


class ScoreParseError(Exception):
    pass


def render_page(body_template: str, page_title: str = "GuitarTabber", **context: object) -> str:
    body = render_template_string(body_template, **context)
    return render_template_string(BASE_PAGE, page_title=page_title, body=body)


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
    positions = find_positions(midi_value, preferred_strings, max_fret=max_fret)
    return positions[0] if positions else None


def find_positions(midi_value: int, preferred_strings: list[int], max_fret: int = 14) -> list[tuple[int, int]]:
    candidates: list[tuple[int, int]] = []
    for string_index in preferred_strings:
        fret = midi_value - STANDARD_TUNING[string_index]
        if 0 <= fret <= max_fret:
            candidates.append((string_index, fret))
    if not candidates:
        return []
    # Prefer open strings, then lower frets, then the caller's string priority order.
    priority_index = {s: i for i, s in enumerate(preferred_strings)}
    candidates.sort(key=lambda c: (0 if c[1] == 0 else 1, c[1], priority_index.get(c[0], 99)))
    return candidates


def can_place_fret_at_slot(
    slot_frets: list[int],
    fret: int,
    max_fretted_span: int = MAX_FRETTED_SPAN,
) -> bool:
    fretted = [f for f in slot_frets if f > 0]
    if fret > 0:
        fretted.append(fret)
    if len(fretted) < 2:
        return True
    return (max(fretted) - min(fretted)) <= max_fretted_span


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


def gather_events(
    score: stream.Score,
) -> tuple[
    list[tuple[int, int]],
    list[tuple[int, int]],
    list[int],
    list[tuple[int, str]],
    list[tuple[int, str]],
    int,
    bool,
]:
    full_flat = score.flatten()
    parts = list(score.parts)

    # Voice-aware extraction for SATB/hymn scores:
    # melody from first part voice 1 (soprano), bass from last part voice 2.
    melody_source = parts[0] if parts else full_flat
    bass_source = parts[-1] if parts else full_flat

    def choose_voice(source: stream.Stream, preferred_voice: str) -> Optional[str]:
        voice_ids = {
            str(voice.id)
            for voice in source.recurse().getElementsByClass(stream.Voice)
            if voice.id is not None
        }
        if preferred_voice in voice_ids:
            return preferred_voice
        if len(voice_ids) == 1:
            return next(iter(voice_ids))
        return None

    melody_voice = choose_voice(melody_source, "1")
    bass_voice = choose_voice(bass_source, "2")

    def collect_part_events(source: stream.Stream, mode: str, voice_filter: Optional[str] = None) -> tuple[list[tuple[int, int]], int]:
        events: list[tuple[int, int]] = []
        max_slot = 0
        for element in source.recurse().notesAndRests:
            if voice_filter is not None:
                parent_voice = element.getContextByClass(stream.Voice)
                if parent_voice is None or str(parent_voice.id) != voice_filter:
                    continue

            slot = quarter_to_slot(float(element.getOffsetInHierarchy(source)))
            max_slot = max(max_slot, slot + 1)
            if isinstance(element, note.Note):
                events.append((slot, int(element.pitch.midi)))
            elif isinstance(element, chord.Chord):
                midi_values = sorted(int(p.midi) for p in element.pitches)
                if midi_values:
                    events.append((slot, midi_values[-1] if mode == "high" else midi_values[0]))
        return events, max_slot

    melody_events, melody_max_slot = collect_part_events(melody_source, mode="high", voice_filter=melody_voice)
    bass_events, bass_max_slot = collect_part_events(bass_source, mode="low", voice_filter=bass_voice)
    total_slots = max(melody_max_slot, bass_max_slot)

    measure_slots: list[int] = []
    for measure in score.recurse().getElementsByClass(stream.Measure):
        slot = quarter_to_slot(float(measure.offset))
        if slot not in measure_slots:
            measure_slots.append(slot)
    measure_slots.sort()

    chord_events: list[tuple[int, str]] = []
    max_offset = float(full_flat.highestTime)
    beat = 0.0
    while beat <= max_offset:
        beat_slot = quarter_to_slot(beat)
        vertical = full_flat.notes.getElementsByOffset(beat, mustBeginInSpan=True, includeEndBoundary=False)
        midi_values: list[int] = []
        durations: list[float] = []
        for item in vertical:
            if isinstance(item, note.Note):
                midi_values.append(int(item.pitch.midi))
                durations.append(float(item.quarterLength))
            elif isinstance(item, chord.Chord):
                midi_values.extend(int(p.midi) for p in item.pitches)
                durations.append(float(item.quarterLength))

        if len(midi_values) >= 3 and max(durations or [0.0]) >= 1.0:
            guessed = chord.Chord([pitch.Pitch(midi=v) for v in sorted(set(midi_values))])
            symbol = harmony.chordSymbolFromChord(guessed)
            label = symbol.figure if symbol and symbol.figure else ""
            if label:
                chord_events.append((beat_slot, label))

        beat += 1.0

    # Extract section/repeat markers and map them to measure starts.
    section_events: list[tuple[int, str]] = []
    seen_sections: set[tuple[int, str]] = set()
    for measure in score.recurse().getElementsByClass(stream.Measure):
        measure_slot = quarter_to_slot(float(measure.offset))

        labels: list[str] = []
        for mark in measure.recurse().getElementsByClass(expressions.RehearsalMark):
            value = str(getattr(mark, "content", None) or mark).strip()
            if value:
                labels.append(f"[{value}]")

        for text_item in measure.recurse().getElementsByClass(expressions.TextExpression):
            value = str(getattr(text_item, "content", None) or text_item).strip()
            lowered = value.lower()
            if value and len(value) <= 24 and any(ch.isalpha() for ch in value):
                if any(token in lowered for token in ("refrain", "chorus", "verse", "bridge", "intro", "outro")):
                    labels.append(f"[{value}]")

        left_bar = measure.leftBarline
        right_bar = measure.rightBarline
        if isinstance(left_bar, bar.Repeat) and getattr(left_bar, "direction", None) == "start":
            labels.append("[|: Repeat]")
        if isinstance(right_bar, bar.Repeat) and getattr(right_bar, "direction", None) == "end":
            labels.append("[:| Repeat]")

        for label in labels:
            key = (measure_slot, label)
            if key in seen_sections:
                continue
            seen_sections.add(key)
            section_events.append(key)

    total_slots = max(total_slots, quarter_to_slot(max_offset) + 1)
    unclamped_slots = max(total_slots + 1, 16)
    was_truncated = unclamped_slots > MAX_SLOTS
    total_slots = min(unclamped_slots, MAX_SLOTS)

    # Fallback: some OMR exports flatten measure metadata. Synthesize bar starts
    # from time signature so wrapping still happens at real measure boundaries.
    if len(measure_slots) <= 1:
        bar_quarters = 4.0
        first_ts = None
        for ts in score.recurse().getElementsByClass(meter.TimeSignature):
            first_ts = ts
            break
        if first_ts is not None and first_ts.barDuration is not None:
            bar_quarters = float(first_ts.barDuration.quarterLength)
        bar_slots = max(1, quarter_to_slot(bar_quarters))
        measure_slots = list(range(0, total_slots, bar_slots))

    section_events.sort(key=lambda item: item[0])
    return melody_events, bass_events, measure_slots, chord_events, section_events, total_slots, was_truncated


def _build_tab_row_html(chord_chunk: str, row_lines: list[tuple[str, str]], row_note: str = "") -> str:
    row_note_markup = f'<div class="tab-notes">{html.escape(row_note)}</div>' if row_note else ""
    chord_markup = html.escape(chord_chunk) if chord_chunk else ""
    line_markup: list[str] = []
    for string_name, line_text in row_lines:
        line_markup.append(
            f'<div class="sl"><span class="sn">{html.escape(string_name)}|</span>{html.escape(line_text)}</div>'
        )
    return (
        '<div class="tab-row">'
        f"{row_note_markup}"
        f'<div class="tab-chords">{chord_markup}</div>'
        f'<div class="tab-strings">{"".join(line_markup)}</div>'
        "</div>"
    )


def arrange_tab(score: stream.Score) -> tuple[str, str, bool]:
    melody_events, bass_events, measure_slots, chord_events, section_events, total_slots, was_truncated = gather_events(score)

    lines = {idx: ["-"] * (total_slots * SLOT_WIDTH) for idx in range(6)}
    slot_fretted_notes: dict[int, list[int]] = {}

    for slot, midi_value in melody_events:
        if slot >= total_slots:
            continue
        candidates = find_positions(midi_value, preferred_strings=[5, 4, 3, 2], max_fret=14)
        for string_index, fret in candidates:
            if not slot_is_free(lines[string_index], slot):
                continue
            existing_frets = slot_fretted_notes.get(slot, [])
            if not can_place_fret_at_slot(existing_frets, fret):
                continue
            place_token(lines[string_index], slot, str(fret))
            slot_fretted_notes.setdefault(slot, []).append(fret)
            break

    for slot, midi_value in bass_events:
        if slot >= total_slots:
            continue
        candidates = find_positions(midi_value, preferred_strings=[0, 1, 2, 3], max_fret=10)
        for string_index, fret in candidates:
            if not slot_is_free(lines[string_index], slot):
                continue
            existing_frets = slot_fretted_notes.get(slot, [])
            if not can_place_fret_at_slot(existing_frets, fret):
                continue
            place_token(lines[string_index], slot, str(fret))
            slot_fretted_notes.setdefault(slot, []).append(fret)
            break

    place_measure_dividers(lines, measure_slots)

    chord_line = build_chord_line(total_slots, chord_events)
    measure_starts = sorted({slot for slot in measure_slots if 0 <= slot < total_slots})
    if not measure_starts or measure_starts[0] != 0:
        measure_starts = [0] + measure_starts

    # Build row boundaries by keeping complete measures together while targeting
    # a readable on-screen width.
    row_ranges: list[tuple[int, int]] = []
    idx = 0
    max_slots_per_row = max(1, TAB_TARGET_CHARS_PER_ROW // SLOT_WIDTH)
    while idx < len(measure_starts):
        row_start_idx = idx
        row_start_slot = measure_starts[row_start_idx]
        row_end_idx = row_start_idx + 1

        # Keep adding whole measures while we stay within both configured limits.
        while row_end_idx < len(measure_starts):
            if (row_end_idx - row_start_idx) >= MEASURES_PER_ROW:
                break
            candidate_end_slot = measure_starts[row_end_idx]
            if (candidate_end_slot - row_start_slot) > max_slots_per_row:
                break
            row_end_idx += 1

        if row_end_idx < len(measure_starts):
            row_end_slot = measure_starts[row_end_idx]
        else:
            row_end_slot = total_slots

        if row_end_slot <= row_start_slot:
            row_end_slot = min(total_slots, row_start_slot + 1)

        row_ranges.append((row_start_slot, row_end_slot))
        idx = row_end_idx

    html_rows: list[str] = []
    plain_rows: list[str] = []

    for start_slot, end_slot in row_ranges:
        char_start = start_slot * SLOT_WIDTH
        char_end = end_slot * SLOT_WIDTH
        chord_chunk = chord_line[char_start:char_end].rstrip()
        row_note = "  ".join(label for slot, label in section_events if start_slot <= slot < end_slot)

        row_lines: list[tuple[str, str]] = []
        plain_block: list[str] = []
        if row_note:
            plain_block.append(row_note)
        if chord_chunk:
            plain_block.append(chord_chunk)

        for string_index in [5, 4, 3, 2, 1, 0]:
            chunk = "".join(lines[string_index][char_start:char_end])
            row_lines.append((STRING_NAMES[string_index], f"{chunk}|"))
            plain_block.append(f"{STRING_NAMES[string_index]}|{chunk}|")

        html_rows.append(_build_tab_row_html(chord_chunk, row_lines, row_note=row_note))
        plain_rows.append("\n".join(plain_block))

    tab_html = f'<div class="tab-container">{"".join(html_rows)}</div>'
    tab_plain = "\n\n".join(plain_rows)
    return tab_html, tab_plain, was_truncated


def suggest_capo(key_name: str) -> str:
    try:
        tonic_name, mode = key_name.split(" ", 1)
    except ValueError:
        return "No suggestion"

    try:
        tonic = pitch.Pitch(tonic_name)
    except Exception:
        return "No suggestion"

    open_majors_pc = {0: "C", 7: "G", 2: "D", 9: "A", 4: "E"}
    open_minors_pc = {9: "A", 4: "E", 2: "D"}

    for capo in range(0, 8):
        target_pc = (tonic.pitchClass - capo) % 12
        if mode == "major" and target_pc in open_majors_pc:
            candidate = open_majors_pc[target_pc]
            return "No capo needed" if capo == 0 else f"Capo {capo} (play in {candidate})"
        if mode == "minor" and target_pc in open_minors_pc:
            candidate = open_minors_pc[target_pc]
            return "No capo needed" if capo == 0 else f"Capo {capo} (play in {candidate}m)"
    return "No strong capo suggestion"


def build_key_options() -> list[str]:
    return [f"{tonic} {mode}" for mode in KEY_MODES for tonic in KEY_TONICS]


def parse_key_name(key_name: str) -> Optional[tuple[pitch.Pitch, str]]:
    try:
        tonic_name, mode = key_name.split(" ", 1)
    except ValueError:
        return None
    mode = mode.strip().lower()
    if mode not in KEY_MODES:
        return None
    try:
        tonic = pitch.Pitch(tonic_name.strip())
    except Exception:
        return None
    return tonic, mode


def render_score_to_tab_payload(score: stream.Score, title: str, source_label: str) -> dict[str, str]:
    key_name = "Unknown"
    try:
        analyzed_key = score.analyze("key")
        key_name = f"{analyzed_key.tonic.name} {analyzed_key.mode}"
    except Exception:
        pass

    tab_html, tab_plain, was_truncated = arrange_tab(score)
    header = [
        f"# {title}",
        f"# Source file: {source_label}",
        f"# Estimated key: {key_name}",
        "# Layout: basic melody (high strings) + bass (low strings)",
        "",
    ]
    return {
        "key_name": key_name,
        "capo_suggestion": suggest_capo(key_name),
        "tab_text": "\n".join(header) + tab_plain,
        "tab_html": tab_html,
        "truncation_warning": (
            f"This score was truncated for display at {MAX_SLOTS} tab slots. Split into sections for full output."
            if was_truncated
            else ""
        ),
    }


def parse_musicxml_bytes(file_bytes: bytes) -> stream.Score:
    try:
        return converter.parseData(file_bytes)
    except Exception:
        try:
            return converter.parseData(file_bytes.decode("utf-8", errors="ignore"))
        except Exception as exc:
            raise ScoreParseError(f"Saved source file could not be parsed: {exc}") from exc


def transpose_score_between_keys(score: stream.Score, source_key_name: str, target_key_name: str) -> stream.Score:
    source = parse_key_name(source_key_name)
    target = parse_key_name(target_key_name)
    if source is None or target is None:
        return score
    source_tonic, _ = source
    target_tonic, _ = target
    semitones = (target_tonic.pitchClass - source_tonic.pitchClass) % 12
    if semitones > 6:
        semitones -= 12
    return score.transpose(semitones, inPlace=False)


def parse_sheet_to_tab(saved_path: Path, safe_name: str, work_dir: Path) -> dict[str, Any]:
    parse_paths: list[Path] = [saved_path]
    source_label = safe_name
    multi_page_warning = ""
    if omr_input_needs_conversion(safe_name):
        parse_paths = convert_sheet_to_musicxml(saved_path, work_dir)
        if len(parse_paths) == 1:
            source_label = f"{safe_name} (via OMR: {parse_paths[0].name})"
        else:
            source_label = f"{safe_name} (via OMR: {len(parse_paths)} exported files, using first)"
            multi_page_warning = (
                f"Detected {len(parse_paths)} OMR-exported files. Currently using the first export only."
            )

    try:
        score = converter.parse(str(parse_paths[0]))
    except Exception as exc:
        raise ScoreParseError(f"MusicXML parsing failed: {exc}") from exc

    title = safe_name
    if score.metadata and score.metadata.title:
        title = str(score.metadata.title)

    rendered = render_score_to_tab_payload(score, title, source_label)

    return {
        "song_title": title,
        "key_name": rendered["key_name"],
        "capo_suggestion": rendered["capo_suggestion"],
        "musicxml_bytes": parse_paths[0].read_bytes(),
        "truncation_warning": rendered["truncation_warning"],
        "multi_page_warning": multi_page_warning,
        "tab_text": rendered["tab_text"],
        "tab_html": rendered["tab_html"],
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
                file_bytes = parsed["musicxml_bytes"]
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
                    capo_suggestion=parsed["capo_suggestion"],
                    tab_text=parsed["tab_text"],
                    tab_html=parsed["tab_html"],
                )
                db.session.add(arrangement)
                db.session.commit()

                result = {
                    "title": "Easy Fingerstyle Tab (Saved)",
                    "filename": safe_name,
                    "key_name": parsed["key_name"],
                    "capo_suggestion": parsed["capo_suggestion"],
                    "truncation_warning": parsed["truncation_warning"],
                    "multi_page_warning": parsed["multi_page_warning"],
                    "tab_html": parsed["tab_html"],
                    "arrangement_url": url_for("view_arrangement", arrangement_id=arrangement.id),
                    "download_url": url_for("download_arrangement", arrangement_id=arrangement.id),
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
    formatted_rows = []
    for row in rows:
        created_at = row.created_at
        created_label = created_at.strftime("%Y-%m-%d %H:%M") if created_at else ""
        formatted_rows.append(
            {
                "id": row.id,
                "song_title": row.song_title,
                "original_filename": row.original_filename,
                "created_at": created_label,
            }
        )
    return render_page(HISTORY_BODY, rows=formatted_rows)


@app.route("/arrangement/<int:arrangement_id>", methods=["GET", "POST"])
def view_arrangement(arrangement_id: int):
    row = (
        db.session.query(
            Arrangement.id.label("id"),
            Arrangement.tab_text.label("tab_text"),
            Arrangement.tab_html.label("tab_html"),
            Arrangement.key_name.label("key_name"),
            Arrangement.capo_suggestion.label("capo_suggestion"),
            Arrangement.created_at.label("created_at"),
            Song.title.label("song_title"),
            Song.original_filename.label("original_filename"),
            Song.file_data.label("file_data"),
        )
        .join(Song, Arrangement.song_id == Song.id)
        .filter(Arrangement.id == arrangement_id)
        .first()
    )

    if row is None:
        abort(404)

    key_options = build_key_options()
    selected_key = row.key_name if row.key_name in key_options else key_options[0]
    transpose_error = None
    transpose_note = None
    created_label = row.created_at.strftime("%Y-%m-%d %H:%M") if row.created_at else ""
    display = {
        "id": row.id,
        "tab_text": row.tab_text,
        "tab_html": row.tab_html,
        "key_name": row.key_name,
        "capo_suggestion": row.capo_suggestion,
        "created_at": created_label,
        "song_title": row.song_title,
        "original_filename": row.original_filename,
    }

    if request.method == "POST":
        selected_key = (request.form.get("target_key") or "").strip()
        if selected_key not in key_options:
            transpose_error = "Please choose a valid target key."
        else:
            try:
                score = parse_musicxml_bytes(row.file_data)
                transposed = transpose_score_between_keys(score, row.key_name, selected_key)
                source_label = f"{row.original_filename} (transposed to {selected_key})"
                rendered = render_score_to_tab_payload(transposed, row.song_title, source_label)
                display["tab_text"] = rendered["tab_text"]
                display["tab_html"] = rendered["tab_html"]
                display["key_name"] = rendered["key_name"]
                display["capo_suggestion"] = rendered["capo_suggestion"]
                transpose_note = f"Showing transposed preview in {selected_key}. Saved arrangement remains unchanged."
                if rendered["truncation_warning"]:
                    transpose_note = f"{transpose_note} {rendered['truncation_warning']}"
            except Exception as exc:
                transpose_error = f"Could not transpose this arrangement: {exc}"

    return render_page(
        ARRANGEMENT_BODY,
        page_title=f"{display['song_title']} | GuitarTabber",
        row=SimpleNamespace(**display),
        key_options=key_options,
        selected_key=selected_key,
        transpose_error=transpose_error,
        transpose_note=transpose_note,
    )


@app.route("/arrangement/<int:arrangement_id>/download", methods=["GET"])
def download_arrangement(arrangement_id: int):
    row = (
        db.session.query(
            Arrangement.id.label("id"),
            Arrangement.tab_text.label("tab_text"),
            Song.title.label("song_title"),
        )
        .join(Song, Arrangement.song_id == Song.id)
        .filter(Arrangement.id == arrangement_id)
        .first()
    )
    if row is None:
        abort(404)

    safe_title = secure_filename(row.song_title or f"arrangement-{arrangement_id}") or f"arrangement-{arrangement_id}"
    filename = f"{safe_title}.txt"
    return Response(
        row.tab_text,
        mimetype="text/plain; charset=utf-8",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


if __name__ == "__main__":
    port = int(os.getenv("PORT", "8080"))
    app.run(host="0.0.0.0", port=port, debug=False)
