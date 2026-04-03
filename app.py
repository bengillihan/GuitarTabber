from pathlib import Path
from typing import Optional

from flask import Flask, render_template_string, request
from music21 import chord, converter, harmony, note, pitch, stream
from werkzeug.utils import secure_filename


UPLOAD_DIR = Path("uploads")
ALLOWED_EXTENSIONS = {"musicxml", "xml", "mxl"}

# MIDI note numbers for standard tuning, low E to high E.
STANDARD_TUNING = [40, 45, 50, 55, 59, 64]
STRING_NAMES = ["E", "A", "D", "G", "B", "E"]
SLOT_WIDTH = 3  # 16th-note slot width in monospace characters
MAX_SLOTS = 320  # Keep output readable for large files

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB upload limit


PAGE_TEMPLATE = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>GuitarTabber MVP</title>
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
    p {
      line-height: 1.5;
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
    <h1>GuitarTabber MVP</h1>
    <p>Upload a MusicXML file and get a first-pass fingerstyle tab.</p>
    <p class="hint">Supported formats: .musicxml, .xml, .mxl</p>

    <form method="post" enctype="multipart/form-data">
      <input type="file" name="music_file" accept=".musicxml,.xml,.mxl" required>
      <button type="submit">Generate Tab</button>
    </form>

    {% if error %}
      <div class="error">{{ error }}</div>
    {% endif %}

    {% if result %}
      <section class="result">
        <h2>{{ result.title }}</h2>
        <p><strong>Uploaded file:</strong> {{ result.filename }}</p>
        <pre>{{ result.tab }}</pre>
      </section>
    {% endif %}
  </main>
</body>
</html>
"""


def is_allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


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

    # Keep output strings indexed low-to-high internally; format high-to-low for display.
    lines = {idx: ["-"] * (total_slots * SLOT_WIDTH) for idx in range(6)}

    # Melody preference: high E, B, G, D, A, E.
    for slot, midi_value in melody_events:
        if slot >= total_slots:
            continue
        pos = find_position(midi_value, preferred_strings=[5, 4, 3, 2, 1, 0])
        if pos is None:
            continue
        string_index, fret = pos
        if slot_is_free(lines[string_index], slot):
            place_token(lines[string_index], slot, str(fret))

    # Bass preference: low E, A, D, then fallback higher if needed.
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


def build_result(saved_path: Path, safe_name: str) -> dict[str, str]:
    score = converter.parse(str(saved_path))

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
        f"# Source file: {safe_name}",
        f"# Estimated key: {key_name}",
        "# Layout: basic melody (high strings) + bass (low strings)",
        "",
    ]

    return {
        "title": "Easy Fingerstyle Tab (MVP v1)",
        "filename": safe_name,
        "tab": "\n".join(header) + tab,
    }


@app.route("/", methods=["GET", "POST"])
def index():
    error = None
    result = None

    if request.method == "POST":
        upload = request.files.get("music_file")

        if upload is None or not upload.filename:
            error = "Please choose a MusicXML file to upload."
        elif not is_allowed_file(upload.filename):
            error = "Unsupported file type. Please upload .musicxml, .xml, or .mxl."
        else:
            UPLOAD_DIR.mkdir(exist_ok=True)
            safe_name = secure_filename(upload.filename)
            saved_path = UPLOAD_DIR / safe_name
            upload.save(saved_path)

            try:
                result = build_result(saved_path, safe_name)
            except Exception as exc:
                error = f"Could not parse MusicXML: {exc}"

    return render_template_string(PAGE_TEMPLATE, error=error, result=result)


if __name__ == "__main__":
    app.run(debug=True)
