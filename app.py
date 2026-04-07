import os
import shutil
import subprocess
import html
import tempfile
import re
import copy
from bisect import bisect_right
from dataclasses import dataclass
from enum import Enum
from types import SimpleNamespace
from pathlib import Path
from typing import Any, Optional
from uuid import uuid4

from flask import Flask, Response, abort, redirect, render_template_string, request, url_for
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import text
from music21 import bar, chord, converter, expressions, harmony, meter, note, pitch, stream
from werkzeug.utils import secure_filename


UPLOAD_DIR = Path("uploads")
MUSICXML_EXTENSIONS = {"musicxml", "xml", "mxl"}
PDF_EXTENSIONS = {"pdf"}
IMAGE_EXTENSIONS = {"png", "jpg", "jpeg", "webp", "bmp", "tif", "tiff", "heic", "heif"}
ALLOWED_EXTENSIONS = MUSICXML_EXTENSIONS | PDF_EXTENSIONS | IMAGE_EXTENSIONS

# MIDI note numbers for standard tuning, low E to high E.
STANDARD_TUNING = [40, 45, 50, 55, 59, 64]
STRING_NAMES = ["E", "A", "D", "G", "B", "E"]
SLOT_WIDTH = 3  # 16th-note slot width in monospace characters
MAX_SLOTS = 960  # Allow full multi-verse hymns without truncating the song
OMR_TIMEOUT_SECONDS = 120
MEASURES_PER_ROW = 4
# Approximate max monospace characters (tab columns) per rendered row before wrapping.
# Row wrapping always happens at measure boundaries.
TAB_TARGET_CHARS_PER_ROW = 192
MAX_FRETTED_SPAN = 5
KEY_TONICS = ["C", "C#", "D", "Eb", "E", "F", "F#", "G", "Ab", "A", "Bb", "B"]
KEY_MODES = ["major", "minor"]
GUITAR_MIN_MIDI = 40
GUITAR_MAX_MIDI = 79
DEFAULT_MAX_UPLOAD_MB = 20
CHORD_TEXT_RE = re.compile(
    r"^[A-G](?:b|#)?(?:(?:maj|min|m|dim|aug|sus|add)?\d*)?(?:/[A-G](?:b|#)?)?$"
)
CHORD_TEXT_SEARCH_RE = re.compile(
    r"([A-G](?:b|#)?(?:(?:maj|min|m|dim|aug|sus|add)?\d*)?(?:/[A-G](?:b|#)?)?)",
    re.IGNORECASE,
)


class TabDifficulty(str, Enum):
    EASY = "easy"
    STANDARD = "standard"
    COMPLETE = "complete"


class TabStyle(str, Enum):
    MELODY = "melody"           # Solo: single-note melody only, exact pitch/rhythm
    CHORDS = "chords"           # Accompaniment: full chord shapes, no melody
    FINGERSTYLE = "fingerstyle" # Melody + Bass: melody always top, bass on strong beats
    CHORDS_AND_MELODY = "chords_and_melody"  # Strum + Licks: chords at changes, melody fills between


# Standard open-position guitar chord shapes.
# Each value is a list of 6 fret numbers (low E → high e).
# None = string not played.  0 = open string.
# These are used for the CHORDS style to generate proper guitar voicings
# instead of trying to place raw bass-staff notes.
_CHORD_SHAPES: dict[str, list[Optional[int]]] = {
    # Major
    "C":   [None, 3, 2, 0, 1, 0],
    "D":   [None, None, 0, 2, 3, 2],
    "E":   [0, 2, 2, 1, 0, 0],
    "F":   [1, 3, 3, 2, 1, 1],
    "G":   [3, 2, 0, 0, 0, 3],
    "A":   [None, 0, 2, 2, 2, 0],
    "B":   [None, 2, 4, 4, 4, 2],
    "Bb":  [None, 1, 3, 3, 3, 1],
    "Ab":  [4, 6, 6, 5, 4, 4],
    "Eb":  [None, None, 1, 3, 4, 3],
    "Db":  [None, None, 3, 1, 2, 1],
    "Gb":  [2, 4, 4, 3, 2, 2],
    # Minor
    "Cm":  [None, 3, 5, 5, 4, 3],
    "Dm":  [None, None, 0, 2, 3, 1],
    "Em":  [0, 2, 2, 0, 0, 0],
    "Fm":  [1, 3, 3, 1, 1, 1],
    "Gm":  [3, 5, 5, 3, 3, 3],
    "Am":  [None, 0, 2, 2, 1, 0],
    "Bm":  [None, 2, 4, 4, 3, 2],
    "Bbm": [None, 1, 3, 3, 2, 1],
    # Dominant 7th
    "C7":  [None, 3, 2, 3, 1, 0],
    "D7":  [None, None, 0, 2, 1, 2],
    "E7":  [0, 2, 0, 1, 0, 0],
    "F7":  [1, 3, 1, 2, 1, 1],
    "G7":  [3, 2, 0, 0, 0, 1],
    "A7":  [None, 0, 2, 0, 2, 0],
    "B7":  [None, 2, 1, 2, 0, 2],
    # Major 7th
    "Cmaj7": [None, 3, 2, 0, 0, 0],
    "Dmaj7": [None, None, 0, 2, 2, 2],
    "Emaj7": [0, 2, 1, 1, 0, 0],
    "Fmaj7": [1, 3, 2, 2, 1, 0],
    "Gmaj7": [3, 2, 0, 0, 0, 2],
    "Amaj7": [None, 0, 2, 1, 2, 0],
    # Minor 7th
    "Am7":  [None, 0, 2, 0, 1, 0],
    "Dm7":  [None, None, 0, 2, 1, 1],
    "Em7":  [0, 2, 2, 0, 3, 0],
    "Bm7":  [None, 2, 4, 2, 3, 2],
    # Suspended
    "Csus2":  [None, 3, 0, 0, 1, 3],
    "Gsus4":  [3, 3, 0, 0, 1, 3],
    "Asus4":  [None, 0, 2, 2, 3, 0],
    # Common slash chords
    "G/D":  [None, None, 0, 0, 0, 3],
    "G/B":  [None, 2, 0, 0, 0, 3],
    "G7/B": [None, 2, 0, 0, 0, 1],
    "C/E":  [0, 3, 2, 0, 1, 0],
    "C/G":  [3, 3, 2, 0, 1, 0],
    "D/F#": [2, None, 0, 2, 3, 2],
    "F/C":  [None, 3, 3, 2, 1, 1],
    "A/C#": [None, 4, 2, 2, 2, 0],
    "E/G#": [4, None, 2, 1, 0, 0],
}


def chord_label_to_midi(label: str) -> list[int]:
    """
    Return MIDI note values for the standard guitar voicing of `label`.
    Falls back to music21 chord construction if no shape is in the library.
    Returns an empty list if the chord cannot be resolved.
    """
    # Direct lookup, then try stripping quality variations.
    shape = _CHORD_SHAPES.get(label)
    if shape is None:
        # Try the root only (e.g. "C#" → look for "Db" enharmonic).
        root_only = re.match(r"^([A-G][b#]?)(.*)$", label)
        if root_only:
            root, quality = root_only.groups()
            # Enharmonic mapping for roots.
            enharmonic = {"C#": "Db", "D#": "Eb", "F#": "Gb", "G#": "Ab", "A#": "Bb",
                          "Db": "C#", "Eb": "D#", "Gb": "F#", "Ab": "G#", "Bb": "A#"}
            alt_root = enharmonic.get(root)
            if alt_root:
                shape = _CHORD_SHAPES.get(f"{alt_root}{quality}")

    if shape is not None:
        midi_values: list[int] = []
        for string_index, fret in enumerate(shape):
            if fret is None:
                continue
            midi_values.append(STANDARD_TUNING[string_index] + fret)
        return midi_values

    # Library miss — fall back to music21 chord construction.
    try:
        sym = harmony.ChordSymbol(label)
        pitches = [int(p.midi) for p in sym.pitches]
        # Build a rough 6-string voicing from the chord tones.
        return pitches[:6] if pitches else []
    except Exception:
        return []


BASE_PAGE = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{{ page_title }}</title>
  <link rel="icon" type="image/svg+xml" href='data:image/svg+xml,%3Csvg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64"%3E%3Crect width="64" height="64" rx="12" fill="%232e6f40"/%3E%3Ctext x="32" y="41" text-anchor="middle" font-size="26" font-family="Georgia,serif" fill="white"%3EGT%3C/text%3E%3C/svg%3E'>
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
    .tab-controls {
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      gap: 0.65rem;
      margin: 0.75rem 0 0.4rem;
    }
    .note-editor {
      margin-top: 1rem;
      border: 1px solid #e2dcc6;
      border-radius: 8px;
      background: #faf8f1;
      padding: 0.75rem;
    }
    .note-editor h3 {
      margin: 0 0 0.6rem 0;
      font-size: 1.05rem;
    }
    .note-editor-list {
      max-height: 280px;
      overflow: auto;
      border: 1px solid #e2dcc6;
      border-radius: 6px;
      background: #fff;
      padding: 0.5rem;
      display: grid;
      gap: 0.25rem;
    }
    .note-editor-item {
      display: flex;
      align-items: center;
      gap: 0.5rem;
      font-family: Menlo, Monaco, Consolas, "Courier New", monospace;
      font-size: 0.86rem;
      color: #2f2f2f;
    }
    .tab-controls label {
      color: #4f4936;
      font-size: 0.95rem;
    }
    .tab-size-range {
      width: 220px;
      max-width: 100%;
    }
    .tab-size-value {
      font-weight: 700;
      color: #2f5f3b;
      min-width: 2.1rem;
      display: inline-block;
      text-align: right;
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
      --tab-font-size: 15px;
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
      font-size: var(--tab-font-size);
      margin-bottom: 0.25rem;
    }
    .tab-chords {
      color: #2e7d32;
      font-weight: 700;
      white-space: pre;
      font-family: Menlo, Monaco, Consolas, "Courier New", monospace;
      font-size: var(--tab-font-size);
      margin-bottom: 0.35rem;
    }
    .tab-strings {
      display: flex;
      flex-direction: column;
      gap: 0.1rem;
      font-family: Menlo, Monaco, Consolas, "Courier New", monospace;
      font-size: var(--tab-font-size);
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
    .tab-lyrics {
      font-family: Menlo, Monaco, Consolas, "Courier New", monospace;
      font-size: var(--tab-font-size);
      white-space: pre;
      min-width: max-content;
      color: #5a3e7a;
      font-style: italic;
      margin-top: 0.3rem;
      padding-left: 1.4rem;
    }
    .ft {
      cursor: pointer;
      border-radius: 2px;
    }
    .ft:hover {
      background: #fee2e2;
      color: #b91c1c;
      text-decoration: line-through;
    }
    .tab-edit-hint {
      font-size: 0.78rem;
      color: #9ca3af;
      margin-top: 0.4rem;
    }
    .score-badge {
      display: inline-block;
      padding: 0.1rem 0.45rem;
      border-radius: 4px;
      font-weight: 700;
      font-size: 0.88rem;
    }
    .score-good { background: #dcfce7; color: #166534; }
    .score-ok   { background: #fef9c3; color: #713f12; }
    .score-low  { background: #fee2e2; color: #991b1b; }
    .score-hint { font-size: 0.78rem; color: #9ca3af; }
    .reprocess-bar {
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      gap: 0.5rem;
      margin: 0.75rem 0 0.25rem;
      padding: 0.6rem 0.75rem;
      background: #f5f3ec;
      border: 1px solid #e2dcc6;
      border-radius: 6px;
    }
    .reprocess-label {
      font-size: 0.88rem;
      color: #5f5a47;
      white-space: nowrap;
    }
    .reprocess-form {
      display: flex;
      flex-wrap: wrap;
      gap: 0.4rem;
      align-items: center;
    }
    .reprocess-form select {
      font-size: 0.88rem;
      padding: 0.25rem 0.4rem;
      border-radius: 4px;
      border: 1px solid #c8c3b0;
      background: #fff;
    }
    .reprocess-form button {
      font-size: 0.88rem;
      padding: 0.25rem 0.75rem;
      border-radius: 4px;
      border: none;
      cursor: pointer;
      background: #4a7c59;
      color: #fff;
    }
    .reprocess-form button:hover { background: #3a6347; }
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
  <script>
    (function () {
      const tabContainer = document.querySelector(".tab-container");
      if (!tabContainer) return;

      const sizeInput = document.querySelector(".tab-size-range");
      const sizeValue = document.querySelector(".tab-size-value");
      const fitButton = document.querySelector(".tab-fit-btn");
      const STORAGE_KEY = "guitartabber_tab_font_px";

      function updateSizeLabel(px) {
        if (sizeValue) {
          sizeValue.textContent = Number(px).toFixed(1);
        }
      }

      function applyTabSize(px) {
        const normalized = Math.max(8, Math.min(20, Number(px) || 15));
        document.querySelectorAll(".tab-container").forEach((container) => {
          container.style.setProperty("--tab-font-size", normalized + "px");
        });
        if (sizeInput) sizeInput.value = String(normalized);
        updateSizeLabel(normalized);
        try {
          localStorage.setItem(STORAGE_KEY, String(normalized));
        } catch (e) {
          // Ignore storage failures.
        }
      }

      function contentFitsRow(row) {
        const available = row.clientWidth - 4;
        if (available <= 0) return true;
        const blocks = row.querySelectorAll(".tab-notes, .tab-chords, .tab-strings, .tab-lyrics");
        if (!blocks.length) return true;
        for (const block of blocks) {
          if (block.scrollWidth > available) {
            return false;
          }
        }
        return true;
      }

      function allRowsFit() {
        const rows = document.querySelectorAll(".tab-row");
        if (!rows.length) return true;
        for (const row of rows) {
          if (!contentFitsRow(row)) return false;
        }
        return true;
      }

      function fitTabsToScreen() {
        const minSize = Number(sizeInput ? sizeInput.min : 8) || 8;
        let size = Number(sizeInput ? sizeInput.value : 15) || 15;
        applyTabSize(size);
        while (size > minSize && !allRowsFit()) {
          size = Math.max(minSize, size - 0.5);
          applyTabSize(size);
        }
      }

      let initialSize = 15;
      try {
        const stored = localStorage.getItem(STORAGE_KEY);
        if (stored) initialSize = Number(stored) || initialSize;
      } catch (e) {
        // Ignore storage failures.
      }
      applyTabSize(initialSize);

      if (sizeInput) {
        sizeInput.addEventListener("input", () => applyTabSize(sizeInput.value));
      }

      if (fitButton) {
        fitButton.addEventListener("click", fitTabsToScreen);
      }

      requestAnimationFrame(() => {
        if (!allRowsFit()) fitTabsToScreen();
      });

      window.addEventListener("resize", () => {
        if (!allRowsFit()) fitTabsToScreen();
      });

      // --- Click-to-remove notes ---
      // Clicking a fret number replaces it (and any digits sharing its slot)
      // with dashes so the note disappears from the tab.  The action is
      // reversible by reloading the page.
      document.querySelectorAll(".tab-container").forEach((container) => {
        container.addEventListener("click", (e) => {
          const ft = e.target.closest(".ft");
          if (!ft) return;
          const len = parseInt(ft.dataset.len, 10) || ft.textContent.length;
          const dashes = "-".repeat(len);
          ft.textContent = dashes;
          ft.classList.remove("ft");
          ft.style.cursor = "";
          ft.removeAttribute("data-nid");
          ft.removeAttribute("data-len");
          ft.removeAttribute("title");
        });
      });
    })();
  </script>
</body>
</html>
"""


HOME_BODY = """
<div class="topnav">
  <h1>GuitarTabber</h1>
  <a href="{{ url_for('history') }}">View History</a>
</div>
<p>Upload sheet music (MusicXML, PDF, or image) and get a guitar tab.</p>
<p class="hint">Supported formats: .musicxml, .xml, .mxl, .pdf, .png, .jpg, .jpeg, .webp, .heic, .heif</p>
<p class="hint">Max upload size: {{ max_upload_mb }} MB</p>

{% if omr_warning %}
  <div class="warning">{{ omr_warning }}</div>
{% endif %}

<form method="post" enctype="multipart/form-data" id="upload-form">
  <input type="file" name="music_file" accept=".musicxml,.xml,.mxl,.pdf,.png,.jpg,.jpeg,.webp,.heic,.heif" required>
  <label for="style"><strong>Style:</strong></label>
  <select id="style" name="style">
    {% for option in style_options %}
      <option value="{{ option.value }}" title="{{ option.description }}" {% if option.value == selected_style %}selected{% endif %}>{{ option.label }}</option>
    {% endfor %}
  </select>
  <label for="difficulty"><strong>Complexity:</strong></label>
  <select id="difficulty" name="difficulty">
    {% for option in difficulty_options %}
      <option value="{{ option.value }}" {% if option.value == selected_difficulty %}selected{% endif %}>{{ option.label }}</option>
    {% endfor %}
  </select>
  <button type="submit">Generate Tab</button>
</form>
<p class="hint"><strong>Mode intent:</strong> Solo = melody only, Chords = accompaniment only, Melody + Bass = melody-first with supportive bass, Chords + Melody Fills = rhythm chords with melodic fills.</p>
<div id="processing">Processing upload... OMR on PDFs/images can take up to a minute.</div>

{% if error %}
  <div class="error">{{ error }}</div>
{% endif %}

{% if result %}
  <section class="result">
    <h2>{{ result.title }}</h2>
    <p class="meta"><strong>Uploaded file:</strong> {{ result.filename }}</p>
    <p class="meta"><strong>Style:</strong> {{ result.style_label }} &nbsp;|&nbsp; <strong>Complexity:</strong> {{ result.difficulty_label }}</p>
    <p class="meta"><strong>Style objective:</strong> {{ result.style_goal }}</p>
    <p class="meta"><strong>Estimated key:</strong> {{ result.key_name }} | <strong>Capo suggestion:</strong> {{ result.capo_suggestion }}</p>
    {% if result.multi_page_warning %}
      <div class="warning">{{ result.multi_page_warning }}</div>
    {% endif %}
    {% if result.truncation_warning %}
      <div class="warning">{{ result.truncation_warning }}</div>
    {% endif %}
<p class="meta"><strong>Saved arrangement:</strong> <a href="{{ result.arrangement_url }}">Open permalink</a> | <a href="{{ result.download_url }}">Download tab (.txt)</a> | <a href="{{ result.download_original_url }}">Download original file</a></p>
{% if result.accuracy_score is not none %}
<p class="meta">
  <strong>Coverage:</strong>
  <span class="score-badge {% if result.accuracy_score >= 80 %}score-good{% elif result.accuracy_score >= 50 %}score-ok{% else %}score-low{% endif %}">{{ result.accuracy_score }}%</span>
  &nbsp;
  <strong>Playability:</strong>
  <span class="score-badge {% if result.playability_score >= 80 %}score-good{% elif result.playability_score >= 50 %}score-ok{% else %}score-low{% endif %}">{{ result.playability_score }}%</span>
</p>
{% endif %}
<div class="reprocess-bar">
  <span class="reprocess-label">Try a different style or complexity with the same file:</span>
  <form method="post" action="{{ result.arrangement_url }}" class="reprocess-form">
    <select name="target_style">
      {% for option in style_options %}
        <option value="{{ option.value }}" title="{{ option.description }}" {% if option.value == result.selected_style %}selected{% endif %}>{{ option.label }}</option>
      {% endfor %}
    </select>
    <select name="target_difficulty">
      {% for option in difficulty_options %}
        <option value="{{ option.value }}" {% if option.value == result.selected_difficulty %}selected{% endif %}>{{ option.label }}</option>
      {% endfor %}
    </select>
    <input type="hidden" name="target_key" value="{{ result.key_name }}">
    <button type="submit" name="source_action" value="reprocess_preview">Preview</button>
    <button type="submit" name="source_action" value="reprocess_save">Save New</button>
  </form>
</div>
    <div class="tab-controls">
      <label>Tab text size: <span class="tab-size-value">15.0</span>px</label>
      <input class="tab-size-range" type="range" min="8" max="20" step="0.5" value="15">
      <button type="button" class="tab-fit-btn">Fit To Screen</button>
    </div>
    <p class="tab-edit-hint">Click any fret number to remove that note.</p>
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
<p class="meta"><strong>Style:</strong> {{ row.style_label }} &nbsp;|&nbsp; <strong>Complexity:</strong> {{ row.difficulty_label }}</p>
<p class="meta"><strong>Style objective:</strong> {{ row.style_goal }}</p>
<p class="meta"><strong>Capo suggestion:</strong> {{ row.capo_suggestion }}</p>
{% if row.accuracy_score is not none and row.playability_score is not none %}
<p class="meta">
  <strong>Coverage:</strong>
  <span class="score-badge {% if row.accuracy_score >= 80 %}score-good{% elif row.accuracy_score >= 50 %}score-ok{% else %}score-low{% endif %}">{{ row.accuracy_score }}%</span>
  &nbsp;
  <strong>Playability:</strong>
  <span class="score-badge {% if row.playability_score >= 80 %}score-good{% elif row.playability_score >= 50 %}score-ok{% else %}score-low{% endif %}">{{ row.playability_score }}%</span>
  <span class="score-hint">— scores update when you reprocess or save a new arrangement</span>
</p>
{% endif %}
<p class="meta"><strong>Saved:</strong> {{ row.created_at }}</p>
<p class="meta">
  <a href="{{ url_for('download_arrangement', arrangement_id=row.id) }}">Download tab (.txt)</a>
  {% if row.has_original %} | <a href="{{ url_for('download_original', arrangement_id=row.id) }}">Download original file ({{ row.original_filename }})</a>{% endif %}
</p>
<form method="post" class="transpose-form">
  <label for="target_key"><strong>Key:</strong></label>
  <select id="target_key" name="target_key">
    {% for option in key_options %}
      <option value="{{ option }}" {% if option == selected_key %}selected{% endif %}>{{ option }}</option>
    {% endfor %}
  </select>
  <label for="target_style"><strong>Style:</strong></label>
  <select id="target_style" name="target_style">
    {% for option in style_options %}
      <option value="{{ option.value }}" title="{{ option.description }}" {% if option.value == selected_style %}selected{% endif %}>{{ option.label }}</option>
    {% endfor %}
  </select>
  <label for="target_difficulty"><strong>Complexity:</strong></label>
  <select id="target_difficulty" name="target_difficulty">
    {% for option in difficulty_options %}
      <option value="{{ option.value }}" {% if option.value == selected_difficulty %}selected{% endif %}>{{ option.label }}</option>
    {% endfor %}
  </select>
  <button type="submit" name="transpose_action" value="preview">Update Tab</button>
  <button type="submit" name="transpose_action" value="save">Save As New Arrangement</button>
  <button type="submit" name="source_action" value="reprocess_preview">Reprocess Source (Preview)</button>
  <button type="submit" name="source_action" value="reprocess_save">Reprocess Source (Save New)</button>
</form>
{% if transpose_error %}
  <div class="error">{{ transpose_error }}</div>
{% endif %}
{% if transpose_note %}
  <p class="meta">{{ transpose_note }}</p>
{% endif %}
{% if note_edit_error %}
  <div class="error">{{ note_edit_error }}</div>
{% endif %}
{% if note_edit_note %}
  <p class="meta">{{ note_edit_note }}</p>
{% endif %}
{% if row.tab_html %}
  <div class="tab-controls">
    <label>Tab text size: <span class="tab-size-value">15.0</span>px</label>
    <input class="tab-size-range" type="range" min="8" max="20" step="0.5" value="15">
    <button type="button" class="tab-fit-btn">Fit To Screen</button>
  </div>
  <p class="tab-edit-hint">Click any fret number to remove that note.</p>
  {{ row.tab_html|safe }}
{% else %}
  <pre>{{ row.tab_text }}</pre>
{% endif %}
{% if note_catalog %}
  <section class="note-editor">
    <h3>Transcribed Notes Editor</h3>
    <p class="meta">Check notes to remove, then preview or save a new arrangement.</p>
    <form method="post" class="transpose-form">
      <input type="hidden" name="target_key" value="{{ selected_key }}">
      <input type="hidden" name="target_style" value="{{ selected_style }}">
      <input type="hidden" name="target_difficulty" value="{{ selected_difficulty }}">
      <div class="note-editor-list">
        {% for note in note_catalog %}
          <label class="note-editor-item">
            <input type="checkbox" name="remove_note_ids" value="{{ note.id }}" {% if note.id in selected_remove_ids %}checked{% endif %}>
            <span>{{ note.label }}</span>
          </label>
        {% endfor %}
      </div>
      <button type="submit" name="note_edit_action" value="preview">Preview Without Selected Notes</button>
      <button type="submit" name="note_edit_action" value="save">Save As New Arrangement</button>
    </form>
  </section>
{% endif %}
"""


def normalize_database_url(raw_url: Optional[str]) -> str:
    if not raw_url:
        return "sqlite:///guitartabber.db"
    if raw_url.startswith("postgres://"):
        return raw_url.replace("postgres://", "postgresql://", 1)
    return raw_url


app = Flask(__name__)

def resolve_max_upload_mb() -> int:
    raw = os.getenv("MAX_UPLOAD_MB", str(DEFAULT_MAX_UPLOAD_MB))
    try:
        mb = int(raw)
    except Exception:
        mb = DEFAULT_MAX_UPLOAD_MB
    return max(5, min(100, mb))


MAX_UPLOAD_MB = resolve_max_upload_mb()
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024
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
    original_file_data = db.Column(db.LargeBinary, nullable=True)
    original_file_mime_type = db.Column(db.String(120), nullable=True)
    created_at = db.Column(db.DateTime(timezone=True), server_default=db.func.now(), nullable=False)


class Arrangement(db.Model):
    __tablename__ = "arrangements"

    id = db.Column(db.Integer, primary_key=True)
    song_id = db.Column(db.Integer, db.ForeignKey("songs.id", ondelete="CASCADE"), nullable=False)
    key_name = db.Column(db.String(80), nullable=False)
    difficulty = db.Column(db.String(20), nullable=False, default=TabDifficulty.STANDARD.value)
    style = db.Column(db.String(30), nullable=False, default=TabStyle.FINGERSTYLE.value)
    capo_suggestion = db.Column(db.String(120), nullable=False, default="No suggestion")
    tab_text = db.Column(db.Text, nullable=False)
    tab_html = db.Column(db.Text, nullable=True)
    accuracy_score = db.Column(db.Integer, nullable=True)
    playability_score = db.Column(db.Integer, nullable=True)
    created_at = db.Column(db.DateTime(timezone=True), server_default=db.func.now(), nullable=False)


@dataclass
class ScoreEvents:
    melody_events: list[tuple[int, int]]
    bass_events: list[tuple[int, int]]
    inner_events: list[tuple[int, int]]
    played_chord_events: list[tuple[int, list[int]]]
    measure_slots: list[int]
    chord_events: list[tuple[int, str]]
    section_events: list[tuple[int, str]]
    lyric_events: list[tuple[int, str]]
    total_slots: int
    was_truncated: bool


def _safe_add_column(sql_if_not_exists: str, sql_plain: str) -> None:
    try:
        db.session.execute(text(sql_if_not_exists))
    except Exception:
        try:
            db.session.execute(text(sql_plain))
        except Exception:
            # Already exists or unsupported syntax; continue.
            pass


with app.app_context():
    db.create_all()
    # Lightweight migrations for existing deployments.
    _safe_add_column(
        "ALTER TABLE arrangements ADD COLUMN IF NOT EXISTS capo_suggestion VARCHAR(120) NOT NULL DEFAULT 'No suggestion'",
        "ALTER TABLE arrangements ADD COLUMN capo_suggestion VARCHAR(120) NOT NULL DEFAULT 'No suggestion'",
    )
    _safe_add_column(
        "ALTER TABLE arrangements ADD COLUMN IF NOT EXISTS tab_html TEXT",
        "ALTER TABLE arrangements ADD COLUMN tab_html TEXT",
    )
    _safe_add_column(
        f"ALTER TABLE arrangements ADD COLUMN IF NOT EXISTS difficulty VARCHAR(20) NOT NULL DEFAULT '{TabDifficulty.STANDARD.value}'",
        f"ALTER TABLE arrangements ADD COLUMN difficulty VARCHAR(20) NOT NULL DEFAULT '{TabDifficulty.STANDARD.value}'",
    )
    _safe_add_column(
        f"ALTER TABLE arrangements ADD COLUMN IF NOT EXISTS style VARCHAR(30) NOT NULL DEFAULT '{TabStyle.FINGERSTYLE.value}'",
        f"ALTER TABLE arrangements ADD COLUMN style VARCHAR(30) NOT NULL DEFAULT '{TabStyle.FINGERSTYLE.value}'",
    )
    _safe_add_column(
        "ALTER TABLE songs ADD COLUMN IF NOT EXISTS original_file_data BYTEA",
        "ALTER TABLE songs ADD COLUMN original_file_data BLOB",
    )
    _safe_add_column(
        "ALTER TABLE songs ADD COLUMN IF NOT EXISTS original_file_mime_type VARCHAR(120)",
        "ALTER TABLE songs ADD COLUMN original_file_mime_type VARCHAR(120)",
    )
    _safe_add_column(
        "ALTER TABLE arrangements ADD COLUMN IF NOT EXISTS accuracy_score INTEGER",
        "ALTER TABLE arrangements ADD COLUMN accuracy_score INTEGER",
    )
    _safe_add_column(
        "ALTER TABLE arrangements ADD COLUMN IF NOT EXISTS playability_score INTEGER",
        "ALTER TABLE arrangements ADD COLUMN playability_score INTEGER",
    )
    db.session.commit()


class OMRConversionError(Exception):
    pass


class ScoreParseError(Exception):
    pass


def parse_difficulty(value: Optional[str]) -> TabDifficulty:
    raw = (value or "").strip().lower()
    for option in TabDifficulty:
        if option.value == raw:
            return option
    return TabDifficulty.STANDARD


def difficulty_label(difficulty: TabDifficulty) -> str:
    if difficulty == TabDifficulty.EASY:
        return "Make Easier"
    if difficulty == TabDifficulty.COMPLETE:
        return "Make More Complete"
    return "Standard"


def difficulty_options() -> list[dict[str, str]]:
    return [{"value": d.value, "label": difficulty_label(d)} for d in TabDifficulty]


def parse_style(value: Optional[str]) -> TabStyle:
    raw = (value or "").strip().lower()
    for option in TabStyle:
        if option.value == raw:
            return option
    return TabStyle.FINGERSTYLE


def style_label(style: TabStyle) -> str:
    return {
        TabStyle.MELODY: "Solo (Melody Only)",
        TabStyle.CHORDS: "Chords (Accompaniment)",
        TabStyle.FINGERSTYLE: "Melody + Bass",
        TabStyle.CHORDS_AND_MELODY: "Chords + Melody Fills",
    }[style]


_STYLE_DESCRIPTIONS: dict[TabStyle, str] = {
    TabStyle.MELODY: "Single-note melody only — hum/sing along test: you can recognize the song",
    TabStyle.CHORDS: "Full chord shapes only — works under a singer, no melody line",
    TabStyle.FINGERSTYLE: "Melody always on top, bass on strong beats — fullest solo sound",
    TabStyle.CHORDS_AND_MELODY: "Chord shapes at changes, melody fills between phrases — live accompaniment feel",
}


def style_options() -> list[dict[str, str]]:
    return [
        {"value": s.value, "label": style_label(s), "description": _STYLE_DESCRIPTIONS[s]}
        for s in TabStyle
    ]


def style_goal(style: TabStyle) -> str:
    return _STYLE_DESCRIPTIONS.get(style, "Melody-first guitar arrangement")


def render_page(body_template: str, page_title: str = "GuitarTabber", **context: object) -> str:
    body = render_template_string(body_template, **context)
    return render_template_string(BASE_PAGE, page_title=page_title, body=body)


def is_allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def file_extension(filename: str) -> str:
    if "." not in filename:
        return ""
    return filename.rsplit(".", 1)[1].lower()


_EXT_TO_MIME: dict[str, str] = {
    "pdf": "application/pdf",
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "webp": "image/webp",
    "bmp": "image/bmp",
    "tif": "image/tiff",
    "tiff": "image/tiff",
    "heic": "image/heic",
    "heif": "image/heif",
    "musicxml": "application/vnd.recordare.musicxml+xml",
    "xml": "application/xml",
    "mxl": "application/vnd.recordare.musicxml",
}


def ext_to_mime(ext: str) -> str:
    return _EXT_TO_MIME.get(ext.lower(), "application/octet-stream")


def omr_input_needs_conversion(filename: str) -> bool:
    ext = file_extension(filename)
    return ext in PDF_EXTENSIONS or ext in IMAGE_EXTENSIONS


def is_omr_available() -> bool:
    if find_musescore_bin() is not None:
        return True
    audiveris_bin = os.getenv("AUDIVERIS_BIN", "audiveris")
    return shutil.which(audiveris_bin) is not None


def collect_generated_musicxml(output_dir: Path) -> list[Path]:
    generated: list[Path] = []
    for pattern in ("*.musicxml", "*.mxl", "*.xml"):
        generated.extend(output_dir.rglob(pattern))
    return sorted(generated)


def run_audiveris_export(audiveris_bin: str, output_dir: Path, source_path: Path) -> tuple[int, str, str]:
    cmd = [audiveris_bin, "-batch", "-export", "-output", str(output_dir), str(source_path)]
    completed = subprocess.run(cmd, capture_output=True, text=True, check=False, timeout=OMR_TIMEOUT_SECONDS)
    return completed.returncode, completed.stdout or "", completed.stderr or ""


def _imagemagick_bin() -> Optional[str]:
    """Return the ImageMagick binary name available on this system, or None."""
    for candidate in ("magick", "convert"):
        if shutil.which(candidate):
            return candidate
    return None


def detect_is_camera_photo(source_path: Path) -> bool:
    """
    Return True when the file is likely a camera/phone photograph rather than a
    clean flatbed scan or computer-generated PDF.

    Checked in order:
      1. HEIC/HEIF format  → always a camera photo (iPhone/Android default)
      2. EXIF camera metadata (Make / Model / LensMake) embedded in the file
         — detected via ImageMagick identify if available
    """
    ext = file_extension(source_path.name)
    if ext in {"heic", "heif"}:
        return True

    im_bin = _imagemagick_bin()
    if im_bin and ext in IMAGE_EXTENSIONS:
        try:
            if im_bin == "magick":
                cmd = ["magick", "identify", "-verbose", str(source_path)]
            else:
                cmd = ["identify", "-verbose", str(source_path)]
            result = subprocess.run(cmd, capture_output=True, text=True, check=False, timeout=15)
            output = result.stdout.lower()
            if any(m in output for m in ("exif:make:", "exif:model:", "exif:lensmake:")):
                return True
        except Exception:
            pass

    return False


def _opencv_perspective_correct(source_path: Path, work_dir: Path) -> Optional[Path]:
    """
    Detect the sheet-music page boundary and apply a perspective transform to
    flatten it.  This corrects the camera angle and book-page curl that are
    common in hand-held photos of hymn books.

    Returns the corrected-image path on success, or None if OpenCV is not
    installed or no clear page quadrilateral can be found.

    Install: pip install opencv-python-headless numpy
    """
    try:
        import cv2  # type: ignore
        import numpy as np  # type: ignore
    except ImportError:
        return None

    img = cv2.imread(str(source_path))
    if img is None:
        return None

    h, w = img.shape[:2]
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Edge detection
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    edges = cv2.Canny(blurred, 50, 150)
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
    dilated = cv2.dilate(edges, kernel, iterations=2)

    # Find the largest quadrilateral that covers ≥ 20 % of the image area
    contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    page_quad = None
    for cnt in sorted(contours, key=cv2.contourArea, reverse=True)[:10]:
        peri = cv2.arcLength(cnt, True)
        approx = cv2.approxPolyDP(cnt, 0.02 * peri, True)
        if len(approx) == 4 and cv2.contourArea(cnt) > 0.20 * h * w:
            page_quad = approx.reshape(4, 2).astype("float32")
            break

    if page_quad is None:
        return None  # can't find a clear page boundary

    # Order corners: top-left, top-right, bottom-right, bottom-left
    s = page_quad.sum(axis=1)
    d = np.diff(page_quad, axis=1)
    rect = np.array(
        [page_quad[np.argmin(s)], page_quad[np.argmin(d)],
         page_quad[np.argmax(s)], page_quad[np.argmax(d)]],
        dtype="float32",
    )
    tl, tr, br, bl = rect

    out_w = int(max(np.linalg.norm(br - bl), np.linalg.norm(tr - tl)))
    out_h = int(max(np.linalg.norm(tr - br), np.linalg.norm(tl - bl)))
    if out_w < 100 or out_h < 100:
        return None  # degenerate result

    dst = np.array(
        [[0, 0], [out_w - 1, 0], [out_w - 1, out_h - 1], [0, out_h - 1]],
        dtype="float32",
    )
    M = cv2.getPerspectiveTransform(rect, dst)
    warped = cv2.warpPerspective(img, M, (out_w, out_h))

    out_dir = work_dir / "omr_dewarp"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"{source_path.stem}_dewarped.png"
    cv2.imwrite(str(out_path), warped)
    return out_path


def _imagemagick_preprocess(source_path: Path, out_path: Path, *, is_photo: bool) -> bool:
    """
    Run ImageMagick on source_path → out_path with a pipeline chosen for the
    input type.

    Clean scan / typeset PDF
      grayscale → normalize → sharpen

    Camera photo
      grayscale → auto-level → local adaptive threshold (handles shadows) →
      median denoise → slight blur+sharpen (recover antialiasing) →
      deskew (rotation correction) → trim border → density hint
    """
    im_bin = _imagemagick_bin()
    if im_bin is None:
        return False

    if im_bin == "magick":
        cmd = ["magick", "convert", str(source_path)]
    else:
        cmd = [im_bin, str(source_path)]

    if is_photo:
        # -lat WxH+offset% : Local Adaptive Threshold — handles uneven lighting
        # and shadows from book binding without washing out faint note heads.
        # 80×80 pixel neighbourhood is ~2-3 % of a 3000-pixel-wide phone photo.
        cmd += [
            "-colorspace", "Gray",
            "-auto-level",
            "-lat", "80x80-5%",
            "-median", "1",
            "-blur", "0x0.5",
            "-sharpen", "0x1.5",
            "-deskew", "40%",    # correct small rotations / tilt
            "-trim", "+repage",  # remove uniform (white/black) border strips
            "-density", "300",
            "-type", "Grayscale",
            str(out_path),
        ]
    else:
        cmd += [
            "-colorspace", "Gray",
            "-normalize",
            "-sharpen", "0x1",
            "-density", "300",
            "-type", "Grayscale",
            str(out_path),
        ]

    try:
        result = subprocess.run(cmd, capture_output=True, text=True, check=False, timeout=60)
        return result.returncode == 0 and out_path.exists() and out_path.stat().st_size > 0
    except (subprocess.TimeoutExpired, OSError):
        return False


def preprocess_image_for_omr(source_path: Path, work_dir: Path) -> Path:
    """
    Return a preprocessed PNG ready for OMR.

    Detects whether the input is a camera photograph or a clean scan and
    applies the appropriate pipeline:

    • Camera photo  → OpenCV perspective correction (optional, requires
                       opencv-python-headless) then ImageMagick local adaptive
                       threshold + deskew + trim.
    • Clean scan    → ImageMagick grayscale + normalize + sharpen.

    HEIC/HEIF always requires ImageMagick to convert; raises OMRConversionError
    if it is unavailable.
    """
    ext = file_extension(source_path.name)
    needs_conversion = ext in {"heic", "heif"}
    is_photo = detect_is_camera_photo(source_path)

    im_bin = _imagemagick_bin()
    if im_bin is None and not needs_conversion and not is_photo:
        return source_path  # nothing to do

    if im_bin is None:
        raise OMRConversionError(
            "HEIC/HEIF images require ImageMagick to convert. "
            "Install it with: brew install imagemagick"
        )

    # For camera photos, attempt OpenCV perspective correction first.
    working_path = source_path
    if is_photo:
        dewarped = _opencv_perspective_correct(source_path, work_dir)
        if dewarped is not None:
            working_path = dewarped

    out_dir = work_dir / "omr_preprocessed"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"{source_path.stem}_omr.png"

    if _imagemagick_preprocess(working_path, out_path, is_photo=is_photo):
        return out_path

    if not needs_conversion:
        return source_path  # preprocessing failed but original may still work

    raise OMRConversionError(
        "HEIC/HEIF images require ImageMagick to convert. "
        "Install it with: brew install imagemagick"
    )


def find_musescore_bin() -> Optional[str]:
    """
    Return the MuseScore CLI executable path, or None if not installed.
    Checks PATH entries and common macOS/Linux/Windows install locations.
    """
    # Common PATH names
    for candidate in ("mscore", "musescore", "mscore4", "mscore3", "MuseScore4", "MuseScore3"):
        if shutil.which(candidate):
            return shutil.which(candidate)

    # macOS app bundle locations
    mac_paths = [
        "/Applications/MuseScore 4.app/Contents/MacOS/mscore",
        "/Applications/MuseScore 3.app/Contents/MacOS/mscore",
        "/Applications/MuseScore4.app/Contents/MacOS/mscore",
        "/Applications/MuseScore3.app/Contents/MacOS/mscore",
    ]
    for p in mac_paths:
        if Path(p).exists():
            return p

    # Linux flatpak / snap
    for candidate in ("/usr/bin/mscore", "/usr/bin/musescore", "/usr/local/bin/mscore"):
        if Path(candidate).exists():
            return candidate

    return None


def convert_with_musescore(source_path: Path, work_dir: Path) -> Optional[Path]:
    """
    Use MuseScore CLI to convert a PDF or image to MusicXML.
    Returns the output MusicXML path on success, None on failure.
    MuseScore produces far more accurate note recognition than Audiveris
    for professionally typeset sheet music PDFs.
    """
    ms_bin = find_musescore_bin()
    if ms_bin is None:
        return None

    output_dir = work_dir / "ms_exports"
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{source_path.stem}.musicxml"

    cmd = [ms_bin, "-o", str(output_path), str(source_path)]
    try:
        result = subprocess.run(
            cmd, capture_output=True, text=True, check=False, timeout=OMR_TIMEOUT_SECONDS
        )
        if output_path.exists() and output_path.stat().st_size > 0:
            return output_path
    except (subprocess.TimeoutExpired, OSError):
        pass

    # MuseScore sometimes writes the file even on non-zero exit; check again.
    if output_path.exists() and output_path.stat().st_size > 0:
        return output_path

    return None


def rasterize_pdf_first_page(source_path: Path, work_dir: Path) -> Optional[Path]:
    gs_bin = shutil.which("gs")
    if gs_bin is None:
        return None
    image_dir = work_dir / "omr_raster"
    image_dir.mkdir(parents=True, exist_ok=True)
    output_png = image_dir / f"{source_path.stem}_page1_300dpi.png"
    cmd = [
        gs_bin,
        "-dSAFER",
        "-dBATCH",
        "-dNOPAUSE",
        "-sDEVICE=pnggray",
        "-r300",
        "-dFirstPage=1",
        "-dLastPage=1",
        f"-sOutputFile={output_png}",
        str(source_path),
    ]
    completed = subprocess.run(cmd, capture_output=True, text=True, check=False, timeout=OMR_TIMEOUT_SECONDS)
    if completed.returncode != 0 or not output_png.exists():
        return None
    return output_png


def convert_sheet_to_musicxml(source_path: Path, work_dir: Path) -> list[Path]:
    """
    Convert PDF/image sheet music to MusicXML.

    Strategy (most accurate first):
      1. MuseScore CLI  — best note/chord-symbol recognition for typeset PDFs
      2. Audiveris      — fallback when MuseScore is not installed
    """
    # --- MuseScore (primary, most accurate) ---
    ms_result = convert_with_musescore(source_path, work_dir)
    if ms_result is not None:
        return [ms_result]

    # --- Preprocess camera/phone images before handing to Audiveris. ---
    # Converts HEIC → PNG and applies contrast enhancement.
    if file_extension(source_path.name) in IMAGE_EXTENSIONS:
        source_path = preprocess_image_for_omr(source_path, work_dir)

    audiveris_bin = os.getenv("AUDIVERIS_BIN", "audiveris")

    output_dir = work_dir / "omr_exports"
    output_dir.mkdir(parents=True, exist_ok=True)
    try:
        returncode, stdout, stderr = run_audiveris_export(audiveris_bin, output_dir, source_path)
    except subprocess.TimeoutExpired as exc:
        raise OMRConversionError(
            f"OMR timed out after {OMR_TIMEOUT_SECONDS}s. Try a smaller/cropped PDF page."
        ) from exc
    except FileNotFoundError as exc:
        raise OMRConversionError(
            "Audiveris is not installed or not found. Set AUDIVERIS_BIN or upload MusicXML directly."
        ) from exc

    generated = collect_generated_musicxml(output_dir)

    if returncode != 0:
        stderr = stderr.strip()
        stdout = stdout.strip()
        detail = stderr or stdout or f"exit code {returncode}"
        # Audiveris can report a non-zero exit while still exporting usable XML
        # (for example when one processing stage flags warnings/errors).
        if generated:
            return generated
        # Retry path for camera-photo PDFs: rasterize first page and retry Audiveris
        # on a high-resolution grayscale PNG, which is often more robust than direct PDF.
        if file_extension(source_path.name) == "pdf":
            raster_png = rasterize_pdf_first_page(source_path, work_dir)
            if raster_png is not None:
                output_dir_retry = work_dir / "omr_exports_retry"
                output_dir_retry.mkdir(parents=True, exist_ok=True)
                try:
                    retry_code, retry_stdout, retry_stderr = run_audiveris_export(audiveris_bin, output_dir_retry, raster_png)
                except subprocess.TimeoutExpired:
                    retry_code = 1
                    retry_stdout = ""
                    retry_stderr = "timeout during retry from rasterized PDF"
                retry_generated = collect_generated_musicxml(output_dir_retry)
                if retry_generated:
                    return retry_generated
                retry_detail = (retry_stderr or retry_stdout or f"exit code {retry_code}").strip()
                raise OMRConversionError(
                    f"Audiveris conversion failed: {detail}. Retry from rasterized PDF also failed: {retry_detail}"
                )
        raise OMRConversionError(f"Audiveris conversion failed: {detail}")

    if not generated:
        if file_extension(source_path.name) == "pdf":
            raster_png = rasterize_pdf_first_page(source_path, work_dir)
            if raster_png is not None:
                output_dir_retry = work_dir / "omr_exports_retry"
                output_dir_retry.mkdir(parents=True, exist_ok=True)
                try:
                    run_audiveris_export(audiveris_bin, output_dir_retry, raster_png)
                except subprocess.TimeoutExpired:
                    pass
                retry_generated = collect_generated_musicxml(output_dir_retry)
                if retry_generated:
                    return retry_generated
        raise OMRConversionError("Audiveris ran but no MusicXML output was produced.")

    return generated


def quarter_to_slot(quarter_length: float) -> int:
    return max(0, int(round(quarter_length * 4)))


def find_position(
    midi_value: int,
    preferred_strings: list[int],
    max_fret: int = 14,
    prefer_open: bool = True,
) -> Optional[tuple[int, int]]:
    positions = find_positions(midi_value, preferred_strings, max_fret=max_fret, prefer_open=prefer_open)
    return positions[0] if positions else None


def find_positions(
    midi_value: int,
    preferred_strings: list[int],
    max_fret: int = 14,
    prefer_open: bool = True,
) -> list[tuple[int, int]]:
    candidates: list[tuple[int, int]] = []
    for string_index in preferred_strings:
        fret = midi_value - STANDARD_TUNING[string_index]
        if 0 <= fret <= max_fret:
            candidates.append((string_index, fret))
    if not candidates:
        return []
    # Prefer open strings, then lower frets, then the caller's string priority order.
    priority_index = {s: i for i, s in enumerate(preferred_strings)}
    if prefer_open:
        candidates.sort(key=lambda c: (0 if c[1] == 0 else 1, c[1], priority_index.get(c[0], 99)))
    else:
        # Deprioritize open strings (fret 0) so fretted positions are tried first.
        candidates.sort(key=lambda c: (1 if c[1] == 0 else 0, c[1], priority_index.get(c[0], 99)))
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


def build_lyric_line(total_slots: int, lyric_events: list[tuple[int, str]]) -> str:
    """Build a character-grid lyric line aligned to slot positions."""
    line = [" "] * (total_slots * SLOT_WIDTH)
    for slot, text in lyric_events:
        start = slot * SLOT_WIDTH
        for idx, ch in enumerate(text):
            pos = start + idx
            if pos < len(line):
                line[pos] = ch
    return "".join(line).rstrip()


def score_placement(
    lines: dict[int, list[str]],
    display_melody_events: list[tuple[int, int]],
    display_total_slots: int,
) -> tuple[int, int]:
    """Return (accuracy_score, playability_score) each 0–100.

    accuracy   — what fraction of melody slots have any fret placed on them.
    playability — penalises fret-span violations and large position jumps.
    """
    # --- Accuracy ---
    melody_slots_in_range = [s for s, _ in display_melody_events if s < display_total_slots]
    placed_count = 0
    for slot in melody_slots_in_range:
        char_start = slot * SLOT_WIDTH
        char_end = char_start + SLOT_WIDTH
        for string_idx in range(6):
            seg = lines[string_idx][char_start:char_end]
            if any(c.isdigit() for c in seg):
                placed_count += 1
                break
    accuracy = int(100 * placed_count / max(1, len(melody_slots_in_range)))

    # --- Playability ---
    # Build slot → list[fret] from all strings.
    slot_frets: dict[int, list[int]] = {}
    melody_fret_sequence: list[int] = []
    for string_idx in range(6):
        line = lines[string_idx]
        i = 0
        while i < len(line):
            if line[i].isdigit():
                j = i
                while j < len(line) and line[j].isdigit():
                    j += 1
                fret = int("".join(line[i:j]))
                slot = i // SLOT_WIDTH
                slot_frets.setdefault(slot, []).append(fret)
                i = j
            else:
                i += 1
    # Collect melody fret sequence (highest string with a digit per melody slot).
    for slot in sorted(melody_slots_in_range):
        char_start = slot * SLOT_WIDTH
        char_end = char_start + SLOT_WIDTH
        for string_idx in [5, 4, 3, 2, 1, 0]:
            seg = "".join(lines[string_idx][char_start:char_end]).strip("- |")
            if seg.isdigit():
                melody_fret_sequence.append(int(seg))
                break

    penalty = 0.0
    total_weight = 0.0

    # Span penalty: each slot with fretted notes more than MAX_FRETTED_SPAN apart.
    for slot, frets in slot_frets.items():
        fretted = [f for f in frets if f > 0]
        if len(fretted) >= 2:
            span = max(fretted) - min(fretted)
            if span > MAX_FRETTED_SPAN:
                penalty += (span - MAX_FRETTED_SPAN) * 3
                total_weight += 10

    # Jump penalty: large fret-position leaps in melody.
    for i in range(1, len(melody_fret_sequence)):
        jump = abs(melody_fret_sequence[i] - melody_fret_sequence[i - 1])
        total_weight += 5
        if jump > 7:
            penalty += (jump - 7) * 1.5
        elif jump > 4:
            penalty += (jump - 4) * 0.5

    # High-fret penalty.
    for fret in melody_fret_sequence:
        total_weight += 2
        if fret > 12:
            penalty += (fret - 12) * 0.5

    if total_weight == 0:
        playability = 100
    else:
        raw = max(0.0, 1.0 - penalty / total_weight)
        playability = int(raw * 100)

    return accuracy, playability


def simplify_chord_label(label: str) -> str:
    cleaned = re.sub(r"\([^)]*\)", "", label or "").strip()
    lowered = cleaned.lower()
    if "dim" in lowered:
        m = re.match(r"^([A-G][b#]?)(.*)$", cleaned)
        return f"{m.group(1)}dim" if m else cleaned
    if "aug" in lowered or "+" in cleaned:
        m = re.match(r"^([A-G][b#]?)(.*)$", cleaned)
        return f"{m.group(1)}aug" if m else cleaned
    m = re.match(r"^([A-G][b#]?)(m?)(.*)$", cleaned)
    if not m:
        return cleaned
    tonic, minorish, _ = m.groups()
    return f"{tonic}{'m' if minorish == 'm' else ''}"


def normalize_chord_label(label: str) -> str:
    cleaned = (label or "").strip().replace(" ", "")
    cleaned = re.sub(r"(?i)power", "", cleaned)
    cleaned = cleaned.replace("5/", "/")
    if cleaned.endswith("5") and "/" not in cleaned:
        cleaned = cleaned[:-1]
    return cleaned


def is_valid_chord_label(label: str) -> bool:
    cleaned = (label or "").strip()
    if not cleaned:
        return False
    lowered = cleaned.lower()
    return "cannotbeidentified" not in lowered


def pc_to_name(pc: int) -> str:
    names = {
        0: "C",
        1: "C#",
        2: "D",
        3: "Eb",
        4: "E",
        5: "F",
        6: "F#",
        7: "G",
        8: "Ab",
        9: "A",
        10: "Bb",
        11: "B",
    }
    return names.get(pc % 12, "C")


def infer_simple_chord_label(midi_values: list[int]) -> str:
    if not midi_values:
        return ""
    pcs = sorted({v % 12 for v in midi_values})
    if not pcs:
        return ""
    bass_pc = min(midi_values) % 12

    best_root = pcs[0]
    best_quality = ""
    best_score = -1
    for root in pcs:
        major_score = int(((root + 4) % 12) in pcs) + int(((root + 7) % 12) in pcs)
        minor_score = int(((root + 3) % 12) in pcs) + int(((root + 7) % 12) in pcs)
        quality = ""
        score = max(major_score, minor_score)
        if major_score > minor_score and major_score >= 1:
            quality = ""
        elif minor_score > major_score and minor_score >= 1:
            quality = "m"
        if score > best_score or (score == best_score and root == bass_pc):
            best_score = score
            best_root = root
            best_quality = quality

    label = f"{pc_to_name(best_root)}{best_quality}"
    if bass_pc != best_root:
        label = f"{label}/{pc_to_name(bass_pc)}"
    return label


def extract_chord_token(raw_value: str) -> Optional[str]:
    raw = (raw_value or "").strip()
    if not raw:
        return None
    match = CHORD_TEXT_SEARCH_RE.search(raw)
    if not match:
        return None
    token = normalize_chord_label(match.group(1))
    if not token:
        return None

    def normalize_root(root: str) -> str:
        if not root:
            return root
        if len(root) == 1:
            return root.upper()
        return root[0].upper() + root[1:]

    m = re.match(r"^([A-Ga-g](?:b|#)?)(.*)$", token)
    if not m:
        return token
    root, rest = m.groups()
    root = normalize_root(root)
    if "/" in rest:
        left, right = rest.split("/", 1)
        right = normalize_root(right)
        return f"{root}{left}/{right}"
    return f"{root}{rest}"


def keep_easy_chord_slot(slot: int, measure_starts: list[int], first_ts: Optional[meter.TimeSignature]) -> bool:
    if not measure_starts:
        return True
    measure_index = bisect_right(measure_starts, slot) - 1
    measure_start = measure_starts[max(0, measure_index)]
    in_measure = max(0, slot - measure_start)
    num = getattr(first_ts, "numerator", None) if first_ts else None
    den = getattr(first_ts, "denominator", None) if first_ts else None
    if num == 4 and den == 4:
        return in_measure in {0, 8}  # beats 1 and 3
    return in_measure == 0


def _part_has_lyrics(part: stream.Stream) -> bool:
    """Return True if any note in the part has an attached lyric."""
    for n in part.flatten().notes:
        if isinstance(n, note.Note):
            if getattr(n, "lyric", None):
                return True
            if any(getattr(lyr, "text", None) for lyr in (getattr(n, "lyrics", []) or [])):
                return True
    return False


def _melody_note_count(part: stream.Stream) -> int:
    """Count distinct melody-range notes (MIDI 55–84) in a part."""
    slots: set[int] = set()
    for n in part.flatten().notes:
        if isinstance(n, note.Note) and 55 <= int(n.pitch.midi) <= 84:
            slots.add(quarter_to_slot(float(n.offset)))
        elif isinstance(n, chord.Chord):
            top = max(int(p.midi) for p in n.pitches)
            if 55 <= top <= 84:
                slots.add(quarter_to_slot(float(n.offset)))
    return len(slots)


def gather_events(score: stream.Score, difficulty: TabDifficulty = TabDifficulty.STANDARD) -> ScoreEvents:
    full_flat = score.flatten()
    parts = list(score.parts)

    # For vocal+piano arrangements, prefer the part that has lyrics attached
    # (the vocal line), falling back to parts[0].
    if parts:
        lyric_part = next((p for p in parts if _part_has_lyrics(p)), None)
        melody_source = lyric_part if lyric_part is not None else parts[0]
    else:
        lyric_part = None
        melody_source = full_flat
    has_lyric_part = lyric_part is not None
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
        flattened = source.flatten()
        for element in flattened.notesAndRests:
            if voice_filter is not None:
                voice_id: Optional[str] = None
                parent_voice = element.getContextByClass(stream.Voice)
                if parent_voice is not None and parent_voice.id is not None:
                    voice_id = str(parent_voice.id)
                else:
                    raw_voice = getattr(element, "voice", None)
                    if raw_voice is not None:
                        voice_id = str(raw_voice)
                # If voice metadata is present, enforce the filter.
                # If it is missing, keep the note to avoid dropping most melody events.
                if voice_id is not None and voice_id != voice_filter:
                    continue

            slot = quarter_to_slot(float(element.offset))
            max_slot = max(max_slot, slot + 1)
            if isinstance(element, note.Note):
                events.append((slot, int(element.pitch.midi)))
            elif isinstance(element, chord.Chord):
                midi_values = sorted(int(p.midi) for p in element.pitches)
                if midi_values:
                    events.append((slot, midi_values[-1] if mode == "high" else midi_values[0]))
        return events, max_slot

    def collect_highest_per_slot(source: stream.Stream) -> tuple[list[tuple[int, int]], int]:
        by_slot: dict[int, int] = {}
        max_slot = 0
        for element in source.flatten().notesAndRests:
            slot = quarter_to_slot(float(element.offset))
            max_slot = max(max_slot, slot + 1)
            midi_value: Optional[int] = None
            if isinstance(element, note.Note):
                midi_value = int(element.pitch.midi)
            elif isinstance(element, chord.Chord):
                midi_values = [int(p.midi) for p in element.pitches]
                if midi_values:
                    midi_value = max(midi_values)
            if midi_value is None:
                continue
            current = by_slot.get(slot)
            if current is None or midi_value > current:
                by_slot[slot] = midi_value
        return sorted(by_slot.items()), max_slot

    def collect_lowest_per_slot(source: stream.Stream) -> tuple[list[tuple[int, int]], int]:
        by_slot: dict[int, int] = {}
        max_slot = 0
        for element in source.flatten().notesAndRests:
            slot = quarter_to_slot(float(element.offset))
            max_slot = max(max_slot, slot + 1)
            midi_value: Optional[int] = None
            if isinstance(element, note.Note):
                midi_value = int(element.pitch.midi)
            elif isinstance(element, chord.Chord):
                midi_values = [int(p.midi) for p in element.pitches]
                if midi_values:
                    midi_value = min(midi_values)
            if midi_value is None:
                continue
            current = by_slot.get(slot)
            if current is None or midi_value < current:
                by_slot[slot] = midi_value
        return sorted(by_slot.items()), max_slot

    def collect_lyric_melody(source: stream.Stream) -> tuple[list[tuple[int, int]], int]:
        """Collect melody notes that explicitly carry lyric syllables."""
        by_slot: dict[int, int] = {}
        max_slot = 0
        for n in source.flatten().notes:
            texts: list[str] = []
            if isinstance(n, note.Note):
                if getattr(n, "lyric", None):
                    texts.append(str(n.lyric))
                for lyr in getattr(n, "lyrics", []) or []:
                    text_val = getattr(lyr, "text", None)
                    if text_val:
                        texts.append(str(text_val))
                midi_val = int(n.pitch.midi)
            elif isinstance(n, chord.Chord):
                for lyr in getattr(n, "lyrics", []) or []:
                    text_val = getattr(lyr, "text", None)
                    if text_val:
                        texts.append(str(text_val))
                midi_values = [int(p.midi) for p in n.pitches]
                if not midi_values:
                    continue
                midi_val = max(midi_values)
            else:
                continue
            if not any(any(ch.isalpha() for ch in t) for t in texts):
                continue
            slot = quarter_to_slot(float(n.offset))
            max_slot = max(max_slot, slot + 1)
            prev = by_slot.get(slot)
            if prev is None or midi_val > prev:
                by_slot[slot] = midi_val
        return sorted(by_slot.items()), max_slot

    def collect_verticals(source: stream.Stream) -> dict[int, set[int]]:
        by_slot: dict[int, set[int]] = {}
        for item in source.flatten().notes:
            slot = quarter_to_slot(float(item.offset))
            bucket = by_slot.setdefault(slot, set())
            if isinstance(item, note.Note):
                bucket.add(int(item.pitch.midi))
            elif isinstance(item, chord.Chord):
                bucket.update(int(p.midi) for p in item.pitches)
        return by_slot

    melody_events, melody_max_slot = collect_part_events(melody_source, mode="high", voice_filter=melody_voice)
    bass_events, bass_max_slot = collect_part_events(bass_source, mode="low", voice_filter=bass_voice)
    comp_source: Optional[stream.Stream] = parts[1] if len(parts) >= 3 else None

    # Extract lyrics from the melody source so they can be displayed below the tab.
    lyric_events: list[tuple[int, str]] = []
    for _n in melody_source.flatten().notes:
        if not isinstance(_n, note.Note):
            continue
        _slot = quarter_to_slot(float(_n.offset))
        if _slot >= MAX_SLOTS:
            continue
        _text: Optional[str] = None
        for _lyr in (getattr(_n, "lyrics", None) or []):
            _t = getattr(_lyr, "text", None)
            if _t:
                _text = _t.strip()
                break
        if not _text:
            _t2 = getattr(_n, "lyric", None)
            if _t2:
                _text = str(_t2).strip()
        if _text:
            lyric_events.append((_slot, _text))
    lyric_events.sort(key=lambda x: x[0])

    # In Standard/Complete, use "highest note per slot" from top staff as a
    # fallback when voice-ID filtering is sparse — but prefer the voice-filtered
    # soprano events when they cover ≥ 75% of the top-staff slots, since voice 1
    # is more accurate than a brute-force pitch max (which can mix in alto notes
    # when the soprano holds while the alto moves).
    if parts and difficulty != TabDifficulty.EASY:
        top_staff_reference = melody_source if has_lyric_part else parts[0]
        top_staff_events, top_staff_max_slot = collect_highest_per_slot(top_staff_reference)
        voice_coverage = len(melody_events) / max(1, len(top_staff_events))
        if voice_coverage < 0.75 and top_staff_events:
            melody_events = top_staff_events
            melody_max_slot = max(melody_max_slot, top_staff_max_slot)
        elif top_staff_events:
            melody_max_slot = max(melody_max_slot, top_staff_max_slot)
        low_staff_events, low_staff_max_slot = collect_lowest_per_slot(bass_source)
        if low_staff_events and len(low_staff_events) >= max(3, len(bass_events) // 2):
            bass_events = low_staff_events
            bass_max_slot = max(bass_max_slot, low_staff_max_slot)

    def merge_melody(base: list[tuple[int, int]], additions: list[tuple[int, int]]) -> list[tuple[int, int]]:
        """Return base with any slots from additions that base doesn't already cover."""
        covered = {s for s, _ in base}
        extras = [(s, m) for s, m in additions if s not in covered]
        return sorted(base + extras, key=lambda x: x[0])

    def collapse_events_by_slot(events: list[tuple[int, int]], prefer: str = "high") -> list[tuple[int, int]]:
        by_slot: dict[int, int] = {}
        for slot, midi_value in events:
            current = by_slot.get(slot)
            if current is None:
                by_slot[slot] = midi_value
                continue
            if prefer == "low":
                by_slot[slot] = min(current, midi_value)
            else:
                by_slot[slot] = max(current, midi_value)
        return sorted(by_slot.items(), key=lambda x: x[0])

    if parts:
        all_noteheads = sum(
            1 if isinstance(el, note.Note) else len(el.pitches)
            for el in full_flat.notes
        )
        sparse_threshold = max(4, int(all_noteheads * 0.05))

        # Pass 1: unfiltered melody source (no voice filter).
        if len(melody_events) < sparse_threshold:
            p1, p1_max = collect_part_events(melody_source, mode="high", voice_filter=None)
            melody_events = merge_melody(melody_events, p1)
            melody_max_slot = max(melody_max_slot, p1_max)

        # If lyrics are present, treat lyric-bearing noteheads as high-confidence
        # melody anchors so syllabic melody notes are not lost in polyphonic OMR.
        lyric_melody, lyric_melody_max = collect_lyric_melody(melody_source)
        if lyric_melody:
            melody_events = merge_melody(melody_events, lyric_melody)
            melody_max_slot = max(melody_max_slot, lyric_melody_max)

        # Final melody recovery: ensure lyric-carrying time slots have a melody
        # note, even if OMR voice IDs are inconsistent. Pull from full-score
        # verticals at that slot (treble range only) when needed.
        lyric_slots = {s for s, _ in lyric_events}
        present_slots = {s for s, _ in melody_events}
        recovered: list[tuple[int, int]] = []
        for slot in sorted(lyric_slots - present_slots):
            offset = slot / 4.0
            vertical = full_flat.notes.getElementsByOffset(
                offset,
                mustBeginInSpan=True,
                includeEndBoundary=False,
            )
            candidates: list[int] = []
            for item in vertical:
                if isinstance(item, note.Note):
                    candidates.append(int(item.pitch.midi))
                elif isinstance(item, chord.Chord):
                    candidates.extend(int(p.midi) for p in item.pitches)
            candidates = [m for m in candidates if 55 <= m <= 84]
            if candidates:
                recovered.append((slot, max(candidates)))
        if recovered:
            melody_events = merge_melody(melody_events, recovered)
            melody_max_slot = max(melody_max_slot, max(s for s, _ in recovered) + 1)

        if has_lyric_part:
            # A vocal part was identified by lyrics — trust it exclusively.
            # Merging from other parts pulls in piano accompaniment patterns
            # and replaces the real melody with chord/bass filler.
            pass
        else:
            # No vocal part found — run the full multi-pass to extract melody
            # from whatever the OMR produced.

            # Pass 2: highest-per-slot from every part; fill in any still-missing slots.
            for part in parts:
                candidate, candidate_max = collect_highest_per_slot(part)
                melody_events = merge_melody(melody_events, candidate)
                melody_max_slot = max(melody_max_slot, candidate_max)
                if len(melody_events) >= sparse_threshold * 4:
                    break

            # Pass 3: global highest across all treble parts.
            all_treble_slots: dict[int, int] = {}
            for part in parts[:-1]:
                for s, midi_val in collect_highest_per_slot(part)[0]:
                    if s not in all_treble_slots or midi_val > all_treble_slots[s]:
                        all_treble_slots[s] = midi_val
            melody_events = merge_melody(melody_events, list(all_treble_slots.items()))
            if all_treble_slots:
                melody_max_slot = max(melody_max_slot, max(all_treble_slots) + 1)

    # Normalize melody MIDI values into the playable guitar range (GUITAR_MIN_MIDI–GUITAR_MAX_MIDI)
    # by octave-shifting rather than discarding notes entirely.
    def _to_guitar_range(midi: int) -> int:
        v = midi
        while v > GUITAR_MAX_MIDI:
            v -= 12
        while v < GUITAR_MIN_MIDI:
            v += 12
        return v

    melody_events = [(s, _to_guitar_range(m)) for s, m in melody_events]
    bass_events = [(s, _to_guitar_range(m)) for s, m in bass_events]

    # Keep one clear melodic pitch per slot (highest) and one bass pitch (lowest).
    # This avoids unstable placement when multiple extraction passes disagree.
    melody_events = collapse_events_by_slot(melody_events, prefer="high")
    bass_events = collapse_events_by_slot(bass_events, prefer="low")

    total_slots = max(melody_max_slot, bass_max_slot)

    measure_slots: list[int] = []
    for measure in score.recurse().getElementsByClass(stream.Measure):
        slot = quarter_to_slot(float(measure.offset))
        if slot not in measure_slots:
            measure_slots.append(slot)
    measure_slots.sort()

    chord_events: list[tuple[int, str]] = []
    seen_chord_slots: set[int] = set()
    played_chord_events: list[tuple[int, list[int]]] = []
    seen_played_slots: set[int] = set()

    def get_global_offset(source_obj: Any) -> Optional[float]:
        try:
            return float(source_obj.getOffsetInHierarchy(score))
        except Exception:
            measure_ctx = source_obj.getContextByClass(stream.Measure)
            if measure_ctx is not None:
                try:
                    return float(measure_ctx.getOffsetInHierarchy(score)) + float(getattr(source_obj, "offset", 0.0) or 0.0)
                except Exception:
                    return None
        return None

    def add_chord_event(raw_label: str, source_obj: Any) -> None:
        label = extract_chord_token(raw_label)
        if not label or not CHORD_TEXT_RE.match(label) or not is_valid_chord_label(label):
            return
        offset = get_global_offset(source_obj)
        if offset is None:
            return
        slot = quarter_to_slot(offset)
        if slot in seen_chord_slots:
            return
        seen_chord_slots.add(slot)
        chord_events.append((slot, label))

    # Prefer explicit MusicXML chord symbols when available and preserve labels verbatim.
    for symbol in score.recurse().getElementsByClass(harmony.ChordSymbol):
        add_chord_event(str(getattr(symbol, "figure", "") or "").strip(), symbol)

    # Fallback: many OMR exports parse chord text as generic text expressions.
    for text_item in score.recurse().getElementsByClass(expressions.TextExpression):
        raw_value = str(getattr(text_item, "content", None) or text_item or "").strip()
        add_chord_event(raw_value, text_item)

    # Additional fallback: some OMR exports chord names as rehearsal marks.
    for mark in score.recurse().getElementsByClass(expressions.RehearsalMark):
        raw_value = str(getattr(mark, "content", None) or mark or "").strip()
        add_chord_event(raw_value, mark)

    # Additional fallback: some OMR exports chord names as note/chord lyrics.
    for n in score.recurse().notes:
        raw_lyrics: list[str] = []
        if isinstance(n, note.Note):
            if getattr(n, "lyric", None):
                raw_lyrics.append(str(n.lyric))
            for lyr in getattr(n, "lyrics", []) or []:
                text_val = getattr(lyr, "text", None)
                if text_val:
                    raw_lyrics.append(str(text_val))
        elif isinstance(n, chord.Chord):
            for lyr in getattr(n, "lyrics", []) or []:
                text_val = getattr(lyr, "text", None)
                if text_val:
                    raw_lyrics.append(str(text_val))
        for raw in raw_lyrics:
            add_chord_event(raw.strip(), n)

    has_explicit_chords = len(chord_events) > 0

    # Capture explicit written chord attacks (especially in lower staves) so
    # they can be rendered as actual tab notes, not only labels.
    for chord_obj in bass_source.flatten().recurse().getElementsByClass(chord.Chord):
        slot = quarter_to_slot(float(chord_obj.offset))
        midi_values = sorted({int(p.midi) for p in chord_obj.pitches})
        if len(midi_values) < 2:
            continue
        if slot not in seen_played_slots:
            seen_played_slots.add(slot)
            played_chord_events.append((slot, midi_values))
        if (not has_explicit_chords) and slot not in seen_chord_slots:
            try:
                symbol = harmony.chordSymbolFromChord(chord_obj)
                inferred_label = normalize_chord_label(symbol.figure if symbol and symbol.figure else "")
            except Exception:
                inferred_label = ""
            if inferred_label and is_valid_chord_label(inferred_label):
                seen_chord_slots.add(slot)
                chord_events.append((slot, inferred_label))

    # Also reconstruct vertical lower-staff chords from simultaneous Note objects
    # when OMR does not emit chord.Chord containers.
    lower_verticals = collect_verticals(bass_source)
    for slot, values in lower_verticals.items():
        midi_values = sorted(values)
        if len(midi_values) < 2 or slot in seen_played_slots:
            continue
        seen_played_slots.add(slot)
        played_chord_events.append((slot, midi_values))

    # Add piano-treble accompaniment (middle voice chords) in Standard/Complete.
    if comp_source is not None and difficulty != TabDifficulty.EASY:
        comp_verticals = collect_verticals(comp_source)
        for slot, values in comp_verticals.items():
            midi_values = sorted(values)
            if len(midi_values) < 2:
                continue
            if slot not in seen_played_slots:
                seen_played_slots.add(slot)
                played_chord_events.append((slot, midi_values))

    # Also capture top-staff (treble) verticals as chord-hit data so alto/inner voices
    # appear in the tab rather than only notes from the bass staff.
    if len(parts) >= 1 and melody_source is not None:
        treble_verticals = collect_verticals(melody_source)
        for slot, values in treble_verticals.items():
            if slot >= total_slots or slot in seen_played_slots:
                continue
            midi_values = sorted(values)
            if len(midi_values) >= 2:
                seen_played_slots.add(slot)
                played_chord_events.append((slot, midi_values))

    max_offset = float(full_flat.highestTime)
    beat_step = 1.0
    first_ts = None
    for ts in score.recurse().getElementsByClass(meter.TimeSignature):
        first_ts = ts
        break
    if first_ts is not None:
        num = getattr(first_ts, "numerator", None)
        den = getattr(first_ts, "denominator", None)
        if isinstance(num, int) and isinstance(den, int) and den == 8 and num % 3 == 0 and num > 3:
            beat_step = 1.5
        elif getattr(first_ts, "beatDuration", None) is not None:
            try:
                candidate = float(first_ts.beatDuration.quarterLength)
                if candidate > 0:
                    beat_step = candidate
            except Exception:
                pass

    def collect_vertical_pitches_at_offset(offset: float) -> list[int]:
        vertical = bass_source.flatten().notes.getElementsByOffset(
            offset, mustBeginInSpan=True, includeEndBoundary=False
        )
        values: list[int] = []
        for item in vertical:
            if isinstance(item, note.Note):
                values.append(int(item.pitch.midi))
            elif isinstance(item, chord.Chord):
                values.extend(int(p.midi) for p in item.pitches)
        return sorted(set(values))

    # Meter-aware chord-hit capture to increase accompaniment density.
    beat = 0.0
    while beat <= max_offset:
        slot = quarter_to_slot(beat)
        midi_values = collect_vertical_pitches_at_offset(beat)
        if len(midi_values) >= 2:
            if difficulty == TabDifficulty.EASY and not keep_easy_chord_slot(slot, measure_slots, first_ts):
                beat += beat_step
                continue
            if slot not in seen_played_slots:
                seen_played_slots.add(slot)
                played_chord_events.append((slot, midi_values))
        beat += beat_step

    beat = 0.0
    last_inferred_label = ""
    last_inferred_root = ""
    while beat <= max_offset:
        beat_slot = quarter_to_slot(beat)
        if beat_slot in seen_chord_slots:
            beat += beat_step
            continue
        if has_explicit_chords:
            beat += beat_step
            continue
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

        min_duration = 0.25 if difficulty == TabDifficulty.COMPLETE else 0.5
        if len(midi_values) >= 2 and max(durations or [0.0]) >= min_duration:
            guessed = chord.Chord([pitch.Pitch(midi=v) for v in sorted(set(midi_values))])
            symbol = harmony.chordSymbolFromChord(guessed)
            label = symbol.figure if symbol and symbol.figure else ""
            label = normalize_chord_label(label)
            label_root = label.split("/")[0] if "/" in label else label
            if label and is_valid_chord_label(label) and label != last_inferred_label and label_root != last_inferred_root:
                chord_events.append((beat_slot, label))
                last_inferred_label = label
                last_inferred_root = label_root

        beat += beat_step

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
    bar_quarters = 4.0
    if first_ts is not None and first_ts.barDuration is not None:
        bar_quarters = float(first_ts.barDuration.quarterLength)
    bar_slots = max(1, quarter_to_slot(bar_quarters))
    synthetic_measure_slots = list(range(0, total_slots, bar_slots))
    measure_slots = sorted(set(measure_slots) | set(synthetic_measure_slots))

    if difficulty == TabDifficulty.EASY:
        chord_events = [(slot, simplify_chord_label(label)) for slot, label in chord_events]
        simplified_played: list[tuple[int, list[int]]] = []
        for slot, midi_values in played_chord_events:
            # Keep easier dyads (root + upper chord tone) for first-position feel.
            if len(midi_values) >= 2:
                simplified_played.append((slot, [midi_values[0], midi_values[-1]]))
            else:
                simplified_played.append((slot, midi_values))
        played_chord_events = simplified_played

    inner_events: list[tuple[int, int]] = []
    if difficulty != TabDifficulty.EASY and parts:
        inner_sources: list[tuple[stream.Stream, Optional[str], str]] = []
        # Voice 2 of the top staff = alto in SATB arrangements.
        inner_sources.append((parts[0], choose_voice(parts[0], "2"), "high"))
        if difficulty == TabDifficulty.COMPLETE:
            # Also pull voice 1 of the bass staff (tenor) in COMPLETE mode.
            if len(parts) > 1:
                inner_sources.append((parts[-1], choose_voice(parts[-1], "1"), "low"))
            for mid in parts[1:-1]:
                inner_sources.append((mid, None, "high"))
        elif comp_source is not None:
            inner_sources.append((comp_source, None, "high"))

        by_slot: dict[int, list[int]] = {}
        for src, voice_id, mode in inner_sources:
            evs, _ = collect_part_events(src, mode=mode, voice_filter=voice_id)
            for slot, midi_value in evs:
                if slot >= total_slots:
                    continue
                by_slot.setdefault(slot, []).append(midi_value)
        for slot in sorted(by_slot):
            values = sorted(by_slot[slot])
            if not values:
                continue
            inner_events.append((slot, values[len(values) // 2]))

    section_events.sort(key=lambda item: item[0])
    played_chord_events.sort(key=lambda item: item[0])

    # Final harmonic fallback: derive simple chord labels from played chord
    # events when explicit symbols/text are sparse or missing.
    # Only emit a label when the chord ROOT changes — this prevents walking bass
    # passing tones (e.g. C/E, C/G) from flooding the display when the chord
    # hasn't actually changed.
    last_fallback_label = ""
    last_fallback_root = ""
    for slot, midi_values in sorted(played_chord_events, key=lambda x: x[0]):
        if slot in seen_chord_slots:
            continue
        guessed = normalize_chord_label(infer_simple_chord_label(midi_values))
        if not guessed or not is_valid_chord_label(guessed):
            continue
        guessed_root = guessed.split("/")[0]
        if guessed == last_fallback_label or guessed_root == last_fallback_root:
            continue
        seen_chord_slots.add(slot)
        chord_events.append((slot, guessed))
        last_fallback_label = guessed
        last_fallback_root = guessed_root

    # Lyric-anchored labels: carry active harmony onto lyric-note slots so
    # users see chord names where words occur in lead sheets.
    lyric_slots: set[int] = set()
    if parts:
        for n in parts[0].flatten().notes:
            texts: list[str] = []
            if isinstance(n, note.Note):
                if getattr(n, "lyric", None):
                    texts.append(str(n.lyric))
                for lyr in getattr(n, "lyrics", []) or []:
                    text_val = getattr(lyr, "text", None)
                    if text_val:
                        texts.append(str(text_val))
            if any(any(ch.isalpha() for ch in t) for t in texts):
                lyric_slots.add(quarter_to_slot(float(n.offset)))

    chord_events.sort(key=lambda item: item[0])
    by_slot_label = {slot: label for slot, label in chord_events}
    played_by_slot = {slot: midi_values for slot, midi_values in played_chord_events}
    for slot in sorted(lyric_slots):
        if slot in by_slot_label:
            continue
        prior = [pair for pair in chord_events if pair[0] <= slot]
        active_label = prior[-1][1] if prior else ""
        if not active_label and slot in played_by_slot:
            active_label = normalize_chord_label(infer_simple_chord_label(played_by_slot[slot]))
        if active_label and is_valid_chord_label(active_label):
            by_slot_label[slot] = active_label

    chord_events = sorted(by_slot_label.items(), key=lambda item: item[0])
    chord_events.sort(key=lambda item: item[0])
    return ScoreEvents(
        melody_events=melody_events,
        bass_events=bass_events,
        inner_events=inner_events,
        played_chord_events=played_chord_events,
        measure_slots=measure_slots,
        chord_events=chord_events,
        section_events=section_events,
        lyric_events=lyric_events,
        total_slots=total_slots,
        was_truncated=was_truncated,
    )


def _annotate_frets_html(line_text: str, string_idx: int, row_idx: int, char_offset: int = 0) -> str:
    """Wrap each run of digits in line_text with a clickable span.

    Spans get a stable data-nid attribute (row/string/char-position) so JS can
    identify and remove individual notes without affecting the rest of the layout.
    """
    parts: list[str] = []
    i = 0
    while i < len(line_text):
        ch = line_text[i]
        if ch.isdigit():
            j = i
            while j < len(line_text) and line_text[j].isdigit():
                j += 1
            fret_str = line_text[i:j]
            nid = f"r{row_idx}s{string_idx}p{char_offset + i}"
            parts.append(
                f'<span class="ft" data-nid="{nid}" data-len="{len(fret_str)}" title="click to remove">'
                f"{html.escape(fret_str)}</span>"
            )
            i = j
        else:
            parts.append(html.escape(ch))
            i += 1
    return "".join(parts)


def _build_tab_row_html(chord_chunk: str, row_lines: list[tuple[str, str]], row_note: str = "", lyric_chunk: str = "", row_idx: int = 0) -> str:
    row_note_markup = f'<div class="tab-notes">{html.escape(row_note)}</div>' if row_note else ""
    chord_markup = html.escape(chord_chunk) if chord_chunk else ""
    line_markup: list[str] = []
    for str_pos, (string_name, line_text) in enumerate(row_lines):
        # string_name is "E","A","D","G","B","e" — string_pos maps to MIDI string index (5=high e … 0=low E)
        annotated = _annotate_frets_html(line_text, str_pos, row_idx)
        line_markup.append(
            f'<div class="sl"><span class="sn">{html.escape(string_name)}|</span>{annotated}</div>'
        )
    lyric_markup = f'<div class="tab-lyrics">{html.escape(lyric_chunk)}</div>' if lyric_chunk.strip() else ""
    return (
        '<div class="tab-row">'
        f"{row_note_markup}"
        f'<div class="tab-chords">{chord_markup}</div>'
        f'<div class="tab-strings">{"".join(line_markup)}</div>'
        f"{lyric_markup}"
        "</div>"
    )


def arrange_tab(
    score: stream.Score,
    difficulty: TabDifficulty = TabDifficulty.STANDARD,
    style: TabStyle = TabStyle.FINGERSTYLE,
) -> tuple[str, str, bool, int, int]:
    events = gather_events(score, difficulty=difficulty)
    measure_starts = sorted({slot for slot in events.measure_slots if 0 <= slot < events.total_slots})
    if not measure_starts or measure_starts[0] != 0:
        measure_starts = [0] + measure_starts

    # Visual spacing pass: sparse measures are stretched horizontally so quarter-note
    # rhythms remain readable even when OMR durations are noisy.
    onset_slots: set[int] = set()
    onset_slots.update(slot for slot, _ in events.melody_events)
    onset_slots.update(slot for slot, _ in events.bass_events)
    onset_slots.update(slot for slot, _ in events.inner_events)
    onset_slots.update(slot for slot, _ in events.played_chord_events)
    onset_slots = {slot for slot in onset_slots if 0 <= slot < events.total_slots}

    remap_ranges: list[tuple[int, int, int, int]] = []
    out_cursor = 0
    for idx, start_slot in enumerate(measure_starts):
        end_slot = measure_starts[idx + 1] if idx + 1 < len(measure_starts) else events.total_slots
        if end_slot <= start_slot:
            continue
        count = len([slot for slot in onset_slots if start_slot <= slot < end_slot])
        stretch = 2 if count <= 4 else 1
        width = (end_slot - start_slot) * stretch
        remap_ranges.append((start_slot, end_slot, out_cursor, stretch))
        out_cursor += width

    def remap_slot(slot: int) -> int:
        for start, end, out_start, stretch in remap_ranges:
            if start <= slot < end:
                return out_start + (slot - start) * stretch
        return slot

    display_total_slots = max(out_cursor, 1)
    display_melody_events = [(remap_slot(slot), midi_value) for slot, midi_value in events.melody_events]
    display_bass_events = [(remap_slot(slot), midi_value) for slot, midi_value in events.bass_events]
    display_inner_events = [(remap_slot(slot), midi_value) for slot, midi_value in events.inner_events]
    display_played_chord_events = [(remap_slot(slot), midi_values) for slot, midi_values in events.played_chord_events]
    display_chord_events = [(remap_slot(slot), label) for slot, label in events.chord_events]
    display_section_events = [(remap_slot(slot), label) for slot, label in events.section_events]
    display_lyric_events = [(remap_slot(slot), text) for slot, text in events.lyric_events]
    display_measure_slots = sorted({remap_slot(slot) for slot in measure_starts})

    lines = {idx: ["-"] * (display_total_slots * SLOT_WIDTH) for idx in range(6)}
    slot_fretted_notes: dict[int, list[int]] = {}
    last_melody_pos: Optional[tuple[int, int]] = None
    melody_max_fret = 5 if difficulty == TabDifficulty.EASY else 22
    bass_max_fret = 5 if difficulty == TabDifficulty.EASY else (11 if difficulty == TabDifficulty.COMPLETE else 8)
    inner_max_fret = 9 if difficulty == TabDifficulty.EASY else (14 if difficulty == TabDifficulty.COMPLETE else 10)
    max_span = 3 if difficulty == TabDifficulty.EASY else 5
    chord_hit_span = 4 if difficulty == TabDifficulty.EASY else (6 if difficulty == TabDifficulty.COMPLETE else 5)
    chord_hit_max_fret = 5 if difficulty == TabDifficulty.EASY else (12 if difficulty == TabDifficulty.COMPLETE else 10)

    def normalize_midi_to_guitar_range(midi_value: int) -> int:
        moved = int(midi_value)
        while moved > GUITAR_MAX_MIDI:
            moved -= 12
        while moved < GUITAR_MIN_MIDI:
            moved += 12
        return moved

    def try_place_midi_at_slot(
        slot: int,
        midi_value: int,
        preferred_strings: list[int],
        max_fret: int,
        span_limit: Optional[int] = None,
        prefer_open: bool = True,
        near_position: Optional[tuple[int, int]] = None,
        melody_ceiling: Optional[int] = None,
    ) -> Optional[tuple[int, int]]:
        if melody_ceiling is not None and midi_value > melody_ceiling:
            return None
        candidates = find_positions(
            midi_value,
            preferred_strings=preferred_strings,
            max_fret=max_fret,
            prefer_open=prefer_open,
        )
        if near_position is not None and candidates:
            last_string, last_fret = near_position
            candidates = sorted(candidates, key=lambda c: (abs(c[1] - last_fret), abs(c[0] - last_string), c[1]))
        for string_index, fret in candidates:
            if not slot_is_free(lines[string_index], slot):
                continue
            existing_frets = slot_fretted_notes.get(slot, [])
            span = max_span if span_limit is None else span_limit
            if not can_place_fret_at_slot(existing_frets, fret, max_fretted_span=span):
                continue
            place_token(lines[string_index], slot, str(fret))
            slot_fretted_notes.setdefault(slot, []).append(fret)
            return (string_index, fret)
        return None

    # Build shared lookup used by all modes.
    melody_slots = {slot for slot, _ in display_melody_events if 0 <= slot < display_total_slots}
    melody_pitch_by_slot = {
        slot: midi_value for slot, midi_value in display_melody_events if 0 <= slot < display_total_slots
    }

    # Determine strong beats (beats 1 and 3 in 4/4; beat 1 only otherwise).
    # Used by FINGERSTYLE mode to place bass only where it counts rhythmically.
    strong_beat_slots: set[int] = set()
    for ms_start in display_measure_slots:
        strong_beat_slots.add(ms_start)  # beat 1 always
        strong_beat_slots.add(ms_start + 8)  # beat 3 in 4/4 (8 slots = 2 quarter notes)

    def _place_melody_note(slot: int, midi_value: int) -> None:
        """Place a single melody note. Always succeeds — retries without span limit."""
        nonlocal last_melody_pos
        placed = try_place_midi_at_slot(
            slot, midi_value,
            preferred_strings=[5, 4, 3, 2, 1, 0],
            max_fret=melody_max_fret,
            prefer_open=False,
            near_position=last_melody_pos,
        )
        if placed is None:
            placed = try_place_midi_at_slot(
                slot, midi_value,
                preferred_strings=[5, 4, 3, 2, 1, 0],
                max_fret=melody_max_fret,
                span_limit=99,
                prefer_open=False,
            )
        if placed is not None:
            last_melody_pos = placed

    # -----------------------------------------------------------------------
    # MODE: SOLO (Melody Only)
    # Objective: one note at a time, exact melody pitch/rhythm, nothing else.
    # -----------------------------------------------------------------------
    if style == TabStyle.MELODY:
        for slot, midi_value in display_melody_events:
            if slot >= display_total_slots:
                continue
            _place_melody_note(slot, midi_value)

    # -----------------------------------------------------------------------
    # MODE: CHORDS (Accompaniment Only)
    # Objective: full chord shapes from the library on every chord change.
    #            No melody line — this is a strumming/backing track arrangement.
    # -----------------------------------------------------------------------
    elif style == TabStyle.CHORDS:
        placed_chord_slots: set[int] = set()
        for slot, chord_label in display_chord_events:
            if slot >= display_total_slots or slot in placed_chord_slots:
                continue
            shape_midi = chord_label_to_midi(chord_label)
            if not shape_midi:
                continue
            placed_chord_slots.add(slot)
            for midi_value in shape_midi:
                try_place_midi_at_slot(
                    slot, midi_value,
                    preferred_strings=[0, 1, 2, 3, 4, 5],
                    max_fret=chord_hit_max_fret,
                    span_limit=chord_hit_span,
                )
        # COMPLETE: also fill detected score chords at beat positions.
        if difficulty == TabDifficulty.COMPLETE:
            for slot, chord_label in display_chord_events:
                if slot >= display_total_slots:
                    continue
                try:
                    symbol = harmony.ChordSymbol(chord_label)
                    chord_midis = [int(p.midi) for p in symbol.pitches]
                except Exception:
                    continue
                for midi_value in chord_midis[:4]:
                    try_place_midi_at_slot(
                        slot, midi_value,
                        preferred_strings=[2, 3, 1, 4, 0, 5],
                        max_fret=inner_max_fret,
                        span_limit=chord_hit_span,
                    )

    # -----------------------------------------------------------------------
    # MODE: MELODY + BASS (Fingerstyle)
    # Objective: melody is ALWAYS the highest note and placed first.
    #            Bass only on strong beats (1 and 3), lower strings only.
    #            No chord blocks — just melody + walking/root bass.
    # -----------------------------------------------------------------------
    elif style == TabStyle.FINGERSTYLE:
        # Pass 1: melody — unconditional, always wins string/slot conflicts.
        for slot, midi_value in display_melody_events:
            if slot >= display_total_slots:
                continue
            _place_melody_note(slot, midi_value)

        # Pass 2: bass on strong beats only (root or lowest available note).
        last_bass_pos: Optional[tuple[int, int]] = None
        bass_by_slot = {slot: midi for slot, midi in display_bass_events if slot < display_total_slots}
        for slot in sorted(strong_beat_slots):
            if slot >= display_total_slots:
                continue
            midi_value = bass_by_slot.get(slot)
            if midi_value is None:
                # No explicit bass note — look for the nearest one within 2 slots.
                for offset in [0, 2, 1, -1, -2]:
                    midi_value = bass_by_slot.get(slot + offset)
                    if midi_value is not None:
                        break
            if midi_value is None:
                continue
            melody_ceiling = melody_pitch_by_slot.get(slot)
            placed = try_place_midi_at_slot(
                slot, midi_value,
                preferred_strings=[0, 1, 2],
                max_fret=bass_max_fret,
                near_position=last_bass_pos,
                melody_ceiling=melody_ceiling,
            )
            if placed is not None:
                last_bass_pos = placed

        # Pass 3 (COMPLETE only): add inner-voice notes on off-beats, staying below melody.
        if difficulty == TabDifficulty.COMPLETE:
            for slot, midi_value in display_inner_events:
                if slot >= display_total_slots or slot in strong_beat_slots:
                    continue
                try_place_midi_at_slot(
                    slot, midi_value,
                    preferred_strings=[3, 2, 4, 1],
                    max_fret=inner_max_fret,
                    span_limit=max_span,
                    melody_ceiling=melody_pitch_by_slot.get(slot),
                )

    # -----------------------------------------------------------------------
    # MODE: CHORDS + MELODY FILLS (Strum + Licks)
    # Objective: chord shapes at every chord change; melody notes fill the
    #            beats where no chord is changing (the "licks" between phrases).
    # -----------------------------------------------------------------------
    elif style == TabStyle.CHORDS_AND_MELODY:
        chord_change_slots: set[int] = {slot for slot, _ in display_chord_events if slot < display_total_slots}

        # Pass 1: full chord shapes at every chord change.
        for slot, chord_label in display_chord_events:
            if slot >= display_total_slots:
                continue
            shape_midi = chord_label_to_midi(chord_label)
            if shape_midi:
                for midi_value in shape_midi:
                    try_place_midi_at_slot(
                        slot, midi_value,
                        preferred_strings=[0, 1, 2, 3, 4, 5],
                        max_fret=chord_hit_max_fret,
                        span_limit=chord_hit_span,
                    )
            else:
                # Fallback: use detected chord notes if library has no shape.
                pass

        # Pass 2: melody notes on beats where no chord shape was placed.
        for slot, midi_value in display_melody_events:
            if slot >= display_total_slots:
                continue
            if slot in chord_change_slots:
                continue  # chord already owns this beat — don't overwrite
            _place_melody_note(slot, midi_value)

        # Pass 3 (COMPLETE): fill detected lower-staff chord notes on chord-change slots
        # so chord shapes have harmonic depth beyond open-position library voicings.
        if difficulty == TabDifficulty.COMPLETE:
            for slot, midi_values in display_played_chord_events:
                if slot >= display_total_slots or slot not in chord_change_slots:
                    continue
                unique_values = sorted({normalize_midi_to_guitar_range(v) for v in midi_values})
                for midi_value in unique_values[:4]:
                    try_place_midi_at_slot(
                        slot, midi_value,
                        preferred_strings=[0, 1, 2, 3],
                        max_fret=chord_hit_max_fret,
                        span_limit=chord_hit_span,
                        melody_ceiling=melody_pitch_by_slot.get(slot),
                    )

    place_measure_dividers(lines, display_measure_slots)

    chord_line = build_chord_line(display_total_slots, display_chord_events)
    lyric_line = build_lyric_line(display_total_slots, display_lyric_events)
    row_measure_starts = display_measure_slots[:]
    if not row_measure_starts or row_measure_starts[0] != 0:
        row_measure_starts = [0] + row_measure_starts

    # Build row boundaries by keeping complete measures together while targeting
    # a readable on-screen width.
    row_ranges: list[tuple[int, int]] = []
    idx = 0
    max_slots_per_row = max(1, TAB_TARGET_CHARS_PER_ROW // SLOT_WIDTH)
    while idx < len(row_measure_starts):
        row_start_idx = idx
        row_start_slot = row_measure_starts[row_start_idx]
        row_end_idx = row_start_idx + 1

        # Keep adding whole measures while we stay within both configured limits.
        while row_end_idx < len(row_measure_starts):
            if (row_end_idx - row_start_idx) >= MEASURES_PER_ROW:
                break
            candidate_end_slot = row_measure_starts[row_end_idx]
            if (candidate_end_slot - row_start_slot) > max_slots_per_row:
                break
            row_end_idx += 1

        if row_end_idx < len(row_measure_starts):
            row_end_slot = row_measure_starts[row_end_idx]
        else:
            row_end_slot = display_total_slots

        if row_end_slot <= row_start_slot:
            row_end_slot = min(display_total_slots, row_start_slot + 1)

        row_ranges.append((row_start_slot, row_end_slot))
        idx = row_end_idx

    html_rows: list[str] = []
    plain_rows: list[str] = []

    for row_idx, (start_slot, end_slot) in enumerate(row_ranges):
        # Build measure slices for this row and drop measures that contain no
        # frets/chords/labels so blank measures do not render as gaps.
        row_measure_starts_in_range = [slot for slot in row_measure_starts if start_slot <= slot < end_slot]
        if not row_measure_starts_in_range or row_measure_starts_in_range[0] != start_slot:
            row_measure_starts_in_range = [start_slot] + row_measure_starts_in_range

        measure_ranges: list[tuple[int, int]] = []
        for idx, measure_start in enumerate(row_measure_starts_in_range):
            measure_end = (
                row_measure_starts_in_range[idx + 1]
                if idx + 1 < len(row_measure_starts_in_range)
                else end_slot
            )
            if measure_end > measure_start:
                measure_ranges.append((measure_start, measure_end))

        kept_measure_ranges: list[tuple[int, int]] = []
        for measure_start, measure_end in measure_ranges:
            seg_char_start = measure_start * SLOT_WIDTH
            seg_char_end = measure_end * SLOT_WIDTH
            seg_has_frets = False
            for string_index in [5, 4, 3, 2, 1, 0]:
                seg = "".join(lines[string_index][seg_char_start:seg_char_end])
                if any(ch.isdigit() for ch in seg):
                    seg_has_frets = True
                    break
            seg_chord = chord_line[seg_char_start:seg_char_end].strip()
            seg_note = any(measure_start <= slot < measure_end for slot, _ in display_section_events)
            if seg_has_frets or seg_chord or seg_note:
                kept_measure_ranges.append((measure_start, measure_end))

        if not kept_measure_ranges:
            continue

        chord_parts: list[str] = []
        lyric_parts: list[str] = []
        row_note_labels: list[str] = []
        for measure_start, measure_end in kept_measure_ranges:
            seg_char_start = measure_start * SLOT_WIDTH
            seg_char_end = measure_end * SLOT_WIDTH
            chord_parts.append(chord_line[seg_char_start:seg_char_end])
            lyric_parts.append(lyric_line[seg_char_start:seg_char_end] if seg_char_end <= len(lyric_line) else "")
            row_note_labels.extend(
                label for slot, label in display_section_events if measure_start <= slot < measure_end
            )

        chord_chunk = "".join(chord_parts).rstrip()
        lyric_chunk = "".join(lyric_parts).rstrip()
        row_note = "  ".join(row_note_labels)
        plain_block: list[str] = []
        if row_note:
            plain_block.append(row_note)
        if chord_chunk:
            plain_block.append(chord_chunk)

        row_lines: list[tuple[str, str]] = []
        for string_index in [5, 4, 3, 2, 1, 0]:
            kept_chunks: list[str] = []
            for measure_start, measure_end in kept_measure_ranges:
                seg_char_start = measure_start * SLOT_WIDTH
                seg_char_end = measure_end * SLOT_WIDTH
                kept_chunks.append("".join(lines[string_index][seg_char_start:seg_char_end]))
            chunk = "".join(kept_chunks)
            row_lines.append((STRING_NAMES[string_index], f"{chunk}|"))
            plain_block.append(f"{STRING_NAMES[string_index]}|{chunk}|")

        if lyric_chunk.strip():
            plain_block.append(lyric_chunk)

        html_rows.append(_build_tab_row_html(chord_chunk, row_lines, row_note=row_note, lyric_chunk=lyric_chunk, row_idx=row_idx))
        plain_rows.append("\n".join(plain_block))

    tab_html = f'<div class="tab-container">{"".join(html_rows)}</div>'
    tab_plain = "\n\n".join(plain_rows)
    accuracy, playability = score_placement(lines, display_melody_events, display_total_slots)
    return tab_html, tab_plain, events.was_truncated, accuracy, playability


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


def capo_play_key_from_suggestion(capo_suggestion: str, mode: str) -> Optional[str]:
    marker = "(play in "
    if marker not in capo_suggestion:
        return None
    start = capo_suggestion.find(marker)
    if start < 0:
        return None
    start += len(marker)
    end = capo_suggestion.find(")", start)
    if end < 0:
        return None
    play_key = capo_suggestion[start:end].strip()
    if not play_key:
        return None
    if mode == "minor":
        tonic = play_key[:-1] if play_key.endswith("m") else play_key
        return f"{tonic} minor"
    tonic = play_key[:-1] if play_key.endswith("m") else play_key
    return f"{tonic} major"


def build_key_options() -> list[str]:
    tonic_order = ["C", "G", "D", "A", "E", "B", "F#", "C#", "Ab", "Eb", "Bb", "F"]
    options: list[str] = []
    for tonic in tonic_order:
        options.append(f"{tonic} major")
        options.append(f"{tonic} minor")
    return options


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


def render_score_to_tab_payload(
    score: stream.Score,
    title: str,
    source_label: str,
    forced_key_name: Optional[str] = None,
    difficulty: TabDifficulty = TabDifficulty.STANDARD,
    style: TabStyle = TabStyle.FINGERSTYLE,
) -> dict[str, str]:
    key_name = forced_key_name or "Unknown"
    if forced_key_name is None:
        try:
            analyzed_key = score.analyze("key")
            key_name = f"{analyzed_key.tonic.name} {analyzed_key.mode}"
        except Exception:
            pass

    capo_suggestion = suggest_capo(key_name)
    if difficulty == TabDifficulty.EASY:
        parsed_key = parse_key_name(key_name)
        if parsed_key is not None:
            _, mode = parsed_key
            play_key = capo_play_key_from_suggestion(capo_suggestion, mode)
            if play_key and play_key != key_name:
                score = transpose_score_between_keys(score, key_name, play_key)
                key_name = play_key

    tab_html, tab_plain, was_truncated, accuracy_score, playability_score = arrange_tab(score, difficulty=difficulty, style=style)
    header = [
        f"# {title}",
        f"# Source file: {source_label}",
        f"# Estimated key: {key_name}",
        f"# Style: {style_label(style)} / {difficulty_label(difficulty)}",
        "",
    ]
    return {
        "key_name": key_name,
        "capo_suggestion": capo_suggestion,
        "difficulty": difficulty.value,
        "style": style.value,
        "tab_text": "\n".join(header) + tab_plain,
        "tab_html": tab_html,
        "accuracy_score": accuracy_score,
        "playability_score": playability_score,
        "truncation_warning": (
            f"This score was truncated for display at {MAX_SLOTS} tab slots. Split into sections for full output."
            if was_truncated
            else ""
        ),
    }


def parse_musicxml_bytes(file_bytes: bytes) -> stream.Score:
    # Compressed MusicXML (.mxl) is a zip container ("PK...") and is more
    # reliable when parsed from a real file path.
    if file_bytes.startswith(b"PK"):
        temp_path: Optional[str] = None
        mxl_parse_error: Optional[Exception] = None
        try:
            with tempfile.NamedTemporaryFile(suffix=".mxl", delete=False) as temp_file:
                temp_file.write(file_bytes)
                temp_path = temp_file.name
            return converter.parse(temp_path)
        except Exception as exc:
            mxl_parse_error = exc
        finally:
            if temp_path:
                try:
                    os.unlink(temp_path)
                except Exception:
                    pass
        if mxl_parse_error is not None:
            raise ScoreParseError(f"Saved .mxl source could not be parsed: {mxl_parse_error}") from mxl_parse_error

    try:
        return converter.parseData(file_bytes)
    except Exception:
        try:
            return converter.parseData(file_bytes.decode("utf-8", errors="ignore"))
        except Exception as exc:
            raise ScoreParseError(f"Saved source file could not be parsed: {exc}") from exc


def _clone_part_template(source_part: stream.Stream, fallback_id: str) -> stream.Part:
    part_id = str(getattr(source_part, "id", "") or fallback_id)
    part_name = str(getattr(source_part, "partName", "") or part_id)
    new_part = stream.Part(id=part_id)
    new_part.partName = part_name
    return new_part


def combine_scores_sequential(scores: list[stream.Score]) -> stream.Score:
    """Concatenate multiple page scores end-to-end (page 1 then page 2...)."""
    if not scores:
        raise ScoreParseError("No scores available to combine.")
    if len(scores) == 1:
        return scores[0]

    normalized_parts: list[list[stream.Stream]] = []
    max_parts = 0
    for score in scores:
        parts = list(score.parts) if list(score.parts) else [score]
        normalized_parts.append(parts)
        max_parts = max(max_parts, len(parts))

    combined = stream.Score()
    if getattr(scores[0], "metadata", None) is not None:
        combined.metadata = copy.deepcopy(scores[0].metadata)

    target_parts: list[stream.Part] = []
    for idx in range(max_parts):
        template_source = normalized_parts[0][idx] if idx < len(normalized_parts[0]) else normalized_parts[0][0]
        target_parts.append(_clone_part_template(template_source, f"Part{idx+1}"))

    for page_parts in normalized_parts:
        for part_idx in range(max_parts):
            if part_idx >= len(page_parts):
                continue
            source_part = page_parts[part_idx]
            target_part = target_parts[part_idx]
            source_measures = list(source_part.getElementsByClass(stream.Measure))
            if source_measures:
                for measure_obj in source_measures:
                    target_part.append(copy.deepcopy(measure_obj))
            else:
                page_offset = float(target_part.highestTime)
                for el in source_part.flatten().notesAndRests:
                    target_part.insert(page_offset + float(el.offset), copy.deepcopy(el))

    for p in target_parts:
        combined.insert(0, p)
    return combined


def score_to_musicxml_bytes(score: stream.Score) -> tuple[bytes, str]:
    """Serialize score to MusicXML bytes, returning (bytes, mime_type)."""
    temp_path: Optional[str] = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".musicxml", delete=False) as temp_file:
            temp_path = temp_file.name
        score.write("musicxml", fp=temp_path)
        return Path(temp_path).read_bytes(), "application/vnd.recordare.musicxml+xml"
    finally:
        if temp_path:
            try:
                os.unlink(temp_path)
            except Exception:
                pass


def iter_transcribed_note_points(score: stream.Score):
    parts = list(score.parts) or [score]
    point_index = 0
    for part_idx, part in enumerate(parts):
        part_name = (part.partName or f"Part {part_idx + 1}").strip()
        for element in part.recurse().notes:
            try:
                offset = float(element.getOffsetInHierarchy(score))
            except Exception:
                continue
            slot = quarter_to_slot(offset)
            measure_ctx = element.getContextByClass(stream.Measure)
            measure_no = int(getattr(measure_ctx, "number", 0) or 0)
            measure_offset = 0.0
            if measure_ctx is not None:
                try:
                    measure_offset = float(measure_ctx.getOffsetInHierarchy(score))
                except Exception:
                    measure_offset = 0.0
            beat_pos = max(0.0, offset - measure_offset)
            beat_label = f"{(beat_pos + 1):.2f}".rstrip("0").rstrip(".")

            if isinstance(element, note.Note):
                point_id = f"n:{part_idx}:{slot}:{int(element.pitch.midi)}:{point_index}"
                point_index += 1
                yield {
                    "id": point_id,
                    "part_index": part_idx,
                    "slot": slot,
                    "measure": measure_no,
                    "beat": beat_label,
                    "part_name": part_name,
                    "element": element,
                    "pitch_name": element.pitch.nameWithOctave,
                    "midi": int(element.pitch.midi),
                    "chord_pitch_index": None,
                }
                continue

            if isinstance(element, chord.Chord):
                for pitch_idx, p in enumerate(element.pitches):
                    point_id = f"c:{part_idx}:{slot}:{int(p.midi)}:{pitch_idx}:{point_index}"
                    point_index += 1
                    yield {
                        "id": point_id,
                        "part_index": part_idx,
                        "slot": slot,
                        "measure": measure_no,
                        "beat": beat_label,
                        "part_name": part_name,
                        "element": element,
                        "pitch_name": p.nameWithOctave,
                        "midi": int(p.midi),
                        "chord_pitch_index": pitch_idx,
                    }


def build_note_catalog(score: stream.Score) -> list[dict[str, str]]:
    catalog: list[dict[str, str]] = []
    for point in iter_transcribed_note_points(score):
        label = (
            f"M{point['measure']:>3}  B{point['beat']:<4}  "
            f"{point['pitch_name']:<5} (midi {point['midi']})  {point['part_name']}"
        )
        catalog.append({"id": point["id"], "label": label})
    return catalog


def remove_notes_by_ids(score: stream.Score, remove_ids: set[str]) -> stream.Score:
    if not remove_ids:
        return score
    working = copy.deepcopy(score)
    note_elements: list[note.Note] = []
    chord_pitches: dict[int, tuple[chord.Chord, set[int]]] = {}

    for point in iter_transcribed_note_points(working):
        if point["id"] not in remove_ids:
            continue
        element = point["element"]
        chord_pitch_index = point["chord_pitch_index"]
        if isinstance(element, note.Note):
            note_elements.append(element)
        elif isinstance(element, chord.Chord) and isinstance(chord_pitch_index, int):
            key = id(element)
            if key not in chord_pitches:
                chord_pitches[key] = (element, set())
            chord_pitches[key][1].add(chord_pitch_index)

    for n in note_elements:
        site = n.activeSite
        if site is not None:
            try:
                site.remove(n, recurse=True)
            except Exception:
                try:
                    site.remove(n)
                except Exception:
                    pass

    for _, (ch, idxs) in chord_pitches.items():
        for idx in sorted(idxs, reverse=True):
            if idx < 0 or idx >= len(ch.pitches):
                continue
            try:
                ch.remove(ch.pitches[idx])
            except Exception:
                pass
        if len(ch.pitches) == 0:
            site = ch.activeSite
            if site is not None:
                try:
                    site.remove(ch, recurse=True)
                except Exception:
                    try:
                        site.remove(ch)
                    except Exception:
                        pass

    return working


def transpose_score_between_keys(score: stream.Score, source_key_name: str, target_key_name: str) -> stream.Score:
    source = parse_key_name(source_key_name)
    target = parse_key_name(target_key_name)
    if source is None or target is None:
        return score
    source_tonic, _ = source
    target_tonic, _ = target
    base = (target_tonic.pitchClass - source_tonic.pitchClass) % 12
    candidates = [base]
    if base != 0:
        candidates.append(base - 12)

    melody_stream: stream.Stream = score.parts[0] if score.parts else score
    melody_midis = [int(n.pitch.midi) for n in melody_stream.flatten().notes if isinstance(n, note.Note)]

    def score_candidate(semitones: int) -> tuple[int, int, int]:
        out_of_range = 0
        out_distance = 0
        if melody_midis:
            for midi_value in melody_midis:
                moved = midi_value + semitones
                if moved < GUITAR_MIN_MIDI:
                    out_of_range += 1
                    out_distance += GUITAR_MIN_MIDI - moved
                elif moved > GUITAR_MAX_MIDI:
                    out_of_range += 1
                    out_distance += moved - GUITAR_MAX_MIDI
        return (out_of_range, out_distance, abs(semitones))

    best = min(candidates, key=score_candidate)
    return score.transpose(best, inPlace=False)


def parse_sheet_to_tab(
    saved_path: Path,
    safe_name: str,
    work_dir: Path,
    difficulty: TabDifficulty = TabDifficulty.STANDARD,
    style: TabStyle = TabStyle.FINGERSTYLE,
) -> dict[str, Any]:
    parse_paths: list[Path] = [saved_path]
    source_label = safe_name
    multi_page_warning = ""
    if omr_input_needs_conversion(safe_name):
        parse_paths = convert_sheet_to_musicxml(saved_path, work_dir)
        if len(parse_paths) == 1:
            source_label = f"{safe_name} (via OMR: {parse_paths[0].name})"
        else:
            source_label = f"{safe_name} (via OMR: {len(parse_paths)} exported files)"
            multi_page_warning = f"Detected {len(parse_paths)} OMR-exported files. Combining pages sequentially."

    parsed_scores: list[stream.Score] = []
    parse_errors: list[str] = []
    for p in parse_paths:
        try:
            parsed_scores.append(converter.parse(str(p)))
        except Exception as exc:
            parse_errors.append(f"{p.name}: {exc}")

    if not parsed_scores:
        raise ScoreParseError(
            "MusicXML parsing failed for all OMR exports: " + "; ".join(parse_errors[:5])
        )

    try:
        score = combine_scores_sequential(parsed_scores)
    except Exception as exc:
        raise ScoreParseError(f"Could not combine multi-page MusicXML exports: {exc}") from exc

    title = safe_name
    if score.metadata and score.metadata.title:
        title = str(score.metadata.title)

    rendered = render_score_to_tab_payload(score, title, source_label, difficulty=difficulty, style=style)

    musicxml_bytes, output_mime_type = score_to_musicxml_bytes(score)

    if parse_errors:
        extra = (
            f" Parsed {len(parsed_scores)} of {len(parse_paths)} exported files; "
            f"some pages could not be read ({'; '.join(parse_errors[:3])})."
        )
        multi_page_warning = f"{multi_page_warning} {extra}".strip()

    return {
        "song_title": title,
        "key_name": rendered["key_name"],
        "capo_suggestion": rendered["capo_suggestion"],
        "musicxml_bytes": musicxml_bytes,
        "musicxml_mime_type": output_mime_type,
        "truncation_warning": rendered["truncation_warning"],
        "multi_page_warning": multi_page_warning,
        "tab_text": rendered["tab_text"],
        "tab_html": rendered["tab_html"],
        "difficulty": rendered["difficulty"],
        "style": rendered["style"],
    }


def reprocess_uploaded_bytes_to_tab(
    file_bytes: bytes,
    filename: str,
    difficulty: TabDifficulty = TabDifficulty.STANDARD,
    style: TabStyle = TabStyle.FINGERSTYLE,
) -> dict[str, Any]:
    request_id = uuid4().hex[:10]
    request_dir = UPLOAD_DIR / "requests" / f"reprocess-{request_id}"
    request_dir.mkdir(parents=True, exist_ok=True)
    safe_name = secure_filename(filename) or "reprocess-input.musicxml"
    saved_path = request_dir / safe_name
    saved_path.write_bytes(file_bytes)
    try:
        return parse_sheet_to_tab(saved_path, safe_name, request_dir, difficulty=difficulty, style=style)
    finally:
        shutil.rmtree(request_dir, ignore_errors=True)


@app.errorhandler(413)
def request_entity_too_large(_error):
    message = f"File is too large. Please upload a file up to {MAX_UPLOAD_MB} MB."
    return render_page(
        HOME_BODY,
        error=message,
        result=None,
        omr_warning=None if is_omr_available() else "OMR is currently unavailable on this server. PDF/image uploads may fail; MusicXML uploads still work.",
        style_options=style_options(),
        selected_style=TabStyle.FINGERSTYLE.value,
        difficulty_options=difficulty_options(),
        selected_difficulty=TabDifficulty.STANDARD.value,
        max_upload_mb=MAX_UPLOAD_MB,
    ), 413


@app.route("/", methods=["GET", "POST"])
def index():
    error = None
    result = None
    selected_difficulty = TabDifficulty.STANDARD
    selected_style = TabStyle.FINGERSTYLE
    omr_warning = None
    if not is_omr_available():
        omr_warning = "OMR is currently unavailable on this server. PDF/image uploads may fail; MusicXML uploads still work."

    if request.method == "POST":
        upload = request.files.get("music_file")
        selected_difficulty = parse_difficulty(request.form.get("difficulty"))
        selected_style = parse_style(request.form.get("style"))

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
            original_bytes = saved_path.read_bytes()
            original_mime = ext_to_mime(file_extension(safe_name))

            try:
                parsed = parse_sheet_to_tab(saved_path, safe_name, request_dir, difficulty=selected_difficulty, style=selected_style)
                file_bytes = parsed["musicxml_bytes"]
                song = Song(
                    title=parsed["song_title"],
                    original_filename=safe_name,
                    mime_type=parsed["musicxml_mime_type"],
                    file_data=file_bytes,
                    original_file_data=original_bytes,
                    original_file_mime_type=original_mime,
                )
                db.session.add(song)
                db.session.flush()

                arrangement = Arrangement(
                    song_id=song.id,
                    key_name=parsed["key_name"],
                    difficulty=parsed["difficulty"],
                    style=parsed["style"],
                    capo_suggestion=parsed["capo_suggestion"],
                    tab_text=parsed["tab_text"],
                    tab_html=parsed["tab_html"],
                    accuracy_score=parsed.get("accuracy_score"),
                    playability_score=parsed.get("playability_score"),
                )
                db.session.add(arrangement)
                db.session.commit()

                result = {
                    "title": f"{style_label(selected_style)} Tab (Saved)",
                    "filename": safe_name,
                    "key_name": parsed["key_name"],
                    "style_label": style_label(selected_style),
                    "style_goal": style_goal(selected_style),
                    "difficulty_label": difficulty_label(selected_difficulty),
                    "capo_suggestion": parsed["capo_suggestion"],
                    "truncation_warning": parsed["truncation_warning"],
                    "multi_page_warning": parsed["multi_page_warning"],
                    "tab_html": parsed["tab_html"],
                    "arrangement_id": arrangement.id,
                    "arrangement_url": url_for("view_arrangement", arrangement_id=arrangement.id),
                    "download_url": url_for("download_arrangement", arrangement_id=arrangement.id),
                    "download_original_url": url_for("download_original", arrangement_id=arrangement.id),
                    "accuracy_score": parsed.get("accuracy_score"),
                    "playability_score": parsed.get("playability_score"),
                    "selected_style": selected_style.value,
                    "selected_difficulty": selected_difficulty.value,
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

    return render_page(
        HOME_BODY,
        error=error,
        result=result,
        omr_warning=omr_warning,
        style_options=style_options(),
        selected_style=selected_style.value,
        difficulty_options=difficulty_options(),
        selected_difficulty=selected_difficulty.value,
        max_upload_mb=MAX_UPLOAD_MB,
    )


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
            Arrangement.song_id.label("song_id"),
            Arrangement.tab_text.label("tab_text"),
            Arrangement.tab_html.label("tab_html"),
            Arrangement.key_name.label("key_name"),
            Arrangement.difficulty.label("difficulty"),
            Arrangement.style.label("style"),
            Arrangement.capo_suggestion.label("capo_suggestion"),
            Arrangement.created_at.label("created_at"),
            Song.title.label("song_title"),
            Song.original_filename.label("original_filename"),
            (Song.original_file_data != None).label("has_original"),  # noqa: E711
            Arrangement.accuracy_score.label("accuracy_score"),
            Arrangement.playability_score.label("playability_score"),
        )
        .join(Song, Arrangement.song_id == Song.id)
        .filter(Arrangement.id == arrangement_id)
        .first()
    )

    if row is None:
        abort(404)

    key_options = build_key_options()
    selected_key = row.key_name if row.key_name in key_options else key_options[0]
    selected_difficulty = parse_difficulty(row.difficulty)
    selected_style = parse_style(getattr(row, "style", None))
    transpose_error = None
    transpose_note = None
    note_edit_error = None
    note_edit_note = None
    selected_remove_ids: set[str] = set()
    note_catalog: list[dict[str, str]] = []
    created_label = row.created_at.strftime("%Y-%m-%d %H:%M") if row.created_at else ""
    display = {
        "id": row.id,
        "song_id": row.song_id,
        "tab_text": row.tab_text,
        "tab_html": row.tab_html,
        "key_name": row.key_name,
        "difficulty": selected_difficulty.value,
        "difficulty_label": difficulty_label(selected_difficulty),
        "style": selected_style.value,
        "style_label": style_label(selected_style),
        "style_goal": style_goal(selected_style),
        "capo_suggestion": row.capo_suggestion,
        "created_at": created_label,
        "song_title": row.song_title,
        "original_filename": row.original_filename,
        "has_original": bool(row.has_original),
        "accuracy_score": row.accuracy_score,
        "playability_score": row.playability_score,
    }

    def load_score_for_arrangement() -> stream.Score:
        file_data = (
            db.session.query(Song.file_data)
            .join(Arrangement, Arrangement.song_id == Song.id)
            .filter(Arrangement.id == arrangement_id)
            .scalar()
        )
        if not file_data:
            raise ScoreParseError("Stored source file not found for this arrangement.")
        return parse_musicxml_bytes(file_data)

    try:
        base_score = load_score_for_arrangement()
        note_catalog = build_note_catalog(base_score)
    except Exception:
        note_catalog = []

    if request.method == "POST":
        selected_key = (request.form.get("target_key") or "").strip()
        selected_difficulty = parse_difficulty(request.form.get("target_difficulty"))
        selected_style = parse_style(request.form.get("target_style"))
        transpose_action = (request.form.get("transpose_action") or "preview").strip().lower()
        source_action = (request.form.get("source_action") or "").strip().lower()
        note_edit_action = (request.form.get("note_edit_action") or "").strip().lower()
        selected_remove_ids = set(request.form.getlist("remove_note_ids"))

        if selected_key not in key_options:
            if note_edit_action:
                note_edit_error = "Please choose a valid target key."
            elif source_action:
                transpose_error = "Please choose a valid target key."
            else:
                transpose_error = "Please choose a valid target key."
        else:
            try:
                if source_action in {"reprocess_preview", "reprocess_save"}:
                    source_row = (
                        db.session.query(
                            Song.original_file_data.label("original_file_data"),
                            Song.original_filename.label("original_filename"),
                            Song.file_data.label("file_data"),
                        )
                        .join(Arrangement, Arrangement.song_id == Song.id)
                        .filter(Arrangement.id == arrangement_id)
                        .first()
                    )
                    if source_row is None:
                        raise ScoreParseError("Stored source file not found for this arrangement.")
                    source_bytes = source_row.original_file_data or source_row.file_data
                    source_name = source_row.original_filename or row.original_filename
                    if not source_bytes:
                        raise ScoreParseError("Stored source file not found for this arrangement.")

                    parsed = reprocess_uploaded_bytes_to_tab(
                        source_bytes,
                        source_name,
                        difficulty=selected_difficulty,
                        style=selected_style,
                    )

                    # Allow optional key change after fresh reprocess.
                    if selected_key != parsed["key_name"]:
                        reprocessed_score = parse_musicxml_bytes(parsed["musicxml_bytes"])
                        transposed = transpose_score_between_keys(reprocessed_score, parsed["key_name"], selected_key)
                        note_catalog = build_note_catalog(transposed)
                        source_label = f"{source_name} (reprocessed, transposed to {selected_key})"
                        rendered = render_score_to_tab_payload(
                            transposed,
                            row.song_title,
                            source_label,
                            forced_key_name=selected_key,
                            difficulty=selected_difficulty,
                            style=selected_style,
                        )
                    else:
                        preview_score = parse_musicxml_bytes(parsed["musicxml_bytes"])
                        note_catalog = build_note_catalog(preview_score)
                        rendered = {
                            "key_name": parsed["key_name"],
                            "capo_suggestion": parsed["capo_suggestion"],
                            "difficulty": parsed["difficulty"],
                            "style": parsed["style"],
                            "tab_text": parsed["tab_text"],
                            "tab_html": parsed["tab_html"],
                            "truncation_warning": parsed["truncation_warning"],
                        }

                    if source_action == "reprocess_save":
                        new_song = Song(
                            title=parsed["song_title"],
                            original_filename=source_name,
                            mime_type=parsed["musicxml_mime_type"],
                            file_data=parsed["musicxml_bytes"],
                            original_file_data=source_row.original_file_data or source_bytes,
                            original_file_mime_type=None,
                        )
                        db.session.add(new_song)
                        db.session.flush()
                        new_arrangement = Arrangement(
                            song_id=new_song.id,
                            key_name=rendered["key_name"],
                            difficulty=rendered["difficulty"],
                            style=rendered["style"],
                            capo_suggestion=rendered["capo_suggestion"],
                            tab_text=rendered["tab_text"],
                            tab_html=rendered["tab_html"],
                            accuracy_score=rendered.get("accuracy_score"),
                            playability_score=rendered.get("playability_score"),
                        )
                        db.session.add(new_arrangement)
                        db.session.commit()
                        return redirect(url_for("view_arrangement", arrangement_id=new_arrangement.id))

                    transpose_note = (
                        f"Showing reprocessed preview: {style_label(selected_style)}, {selected_key}, "
                        f"{difficulty_label(selected_difficulty)}. Saved arrangement remains unchanged."
                    )
                    if rendered["truncation_warning"]:
                        transpose_note = f"{transpose_note} {rendered['truncation_warning']}"
                else:
                    score = load_score_for_arrangement()
                    transposed = transpose_score_between_keys(score, row.key_name, selected_key)

                    # Rebuild editor catalog in the current key preview context.
                    note_catalog = build_note_catalog(transposed)
                    if selected_remove_ids:
                        transposed = remove_notes_by_ids(transposed, selected_remove_ids)

                    source_label = f"{row.original_filename} (transposed to {selected_key})"
                    rendered = render_score_to_tab_payload(
                        transposed,
                        row.song_title,
                        source_label,
                        forced_key_name=selected_key,
                        difficulty=selected_difficulty,
                        style=selected_style,
                    )
                display["tab_text"] = rendered["tab_text"]
                display["tab_html"] = rendered["tab_html"]
                display["key_name"] = rendered["key_name"]
                display["difficulty"] = rendered["difficulty"]
                display["difficulty_label"] = difficulty_label(selected_difficulty)
                display["style"] = rendered["style"]
                display["style_label"] = style_label(selected_style)
                display["style_goal"] = style_goal(selected_style)
                display["capo_suggestion"] = rendered["capo_suggestion"]
                display["accuracy_score"] = rendered.get("accuracy_score")
                display["playability_score"] = rendered.get("playability_score")

                save_requested = (transpose_action == "save") or (note_edit_action == "save")
                if save_requested:
                    new_arrangement = Arrangement(
                        song_id=row.song_id,
                        key_name=rendered["key_name"],
                        difficulty=rendered["difficulty"],
                        style=rendered["style"],
                        capo_suggestion=rendered["capo_suggestion"],
                        tab_text=rendered["tab_text"],
                        tab_html=rendered["tab_html"],
                        accuracy_score=rendered.get("accuracy_score"),
                        playability_score=rendered.get("playability_score"),
                    )
                    db.session.add(new_arrangement)
                    db.session.commit()
                    return redirect(url_for("view_arrangement", arrangement_id=new_arrangement.id))

                if source_action:
                    # message already set above
                    pass
                elif note_edit_action:
                    note_edit_note = (
                        f"Previewing edited notes: removed {len(selected_remove_ids)} note(s), "
                        f"{style_label(selected_style)}, {selected_key}, {difficulty_label(selected_difficulty)}. "
                        "Saved arrangement remains unchanged."
                    )
                    if rendered["truncation_warning"]:
                        note_edit_note = f"{note_edit_note} {rendered['truncation_warning']}"
                else:
                    transpose_note = (
                        f"Showing preview: {style_label(selected_style)}, {selected_key}, {difficulty_label(selected_difficulty)}. "
                        "Saved arrangement remains unchanged."
                    )
                    if rendered["truncation_warning"]:
                        transpose_note = f"{transpose_note} {rendered['truncation_warning']}"
            except Exception as exc:
                db.session.rollback()
                if note_edit_action:
                    note_edit_error = f"Could not update note edits: {exc}"
                elif source_action:
                    transpose_error = f"Could not reprocess this arrangement: {exc}"
                else:
                    transpose_error = f"Could not update this arrangement: {exc}"

    return render_page(
        ARRANGEMENT_BODY,
        page_title=f"{display['song_title']} | GuitarTabber",
        row=SimpleNamespace(**display),
        key_options=key_options,
        selected_key=selected_key,
        style_options=style_options(),
        selected_style=selected_style.value,
        difficulty_options=difficulty_options(),
        selected_difficulty=selected_difficulty.value,
        transpose_error=transpose_error,
        transpose_note=transpose_note,
        note_edit_error=note_edit_error,
        note_edit_note=note_edit_note,
        note_catalog=note_catalog,
        selected_remove_ids=selected_remove_ids,
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


@app.route("/arrangement/<int:arrangement_id>/download/original", methods=["GET"])
def download_original(arrangement_id: int):
    row = (
        db.session.query(
            Song.original_filename.label("original_filename"),
            Song.original_file_data.label("original_file_data"),
            Song.original_file_mime_type.label("original_file_mime_type"),
        )
        .join(Arrangement, Arrangement.song_id == Song.id)
        .filter(Arrangement.id == arrangement_id)
        .first()
    )
    if row is None or not row.original_file_data:
        abort(404)

    return Response(
        row.original_file_data,
        mimetype=row.original_file_mime_type or "application/octet-stream",
        headers={"Content-Disposition": f'attachment; filename="{row.original_filename}"'},
    )


if __name__ == "__main__":
    port = int(os.getenv("PORT", "8080"))
    app.run(host="0.0.0.0", port=port, debug=False)
