"""
DocSearch — loads .docx files, finds H1/H2/H3/H4 sections,
and copies them by having Word select+copy the real content natively.
Zero RTF conversion — Word does the copy, so all formatting is preserved.

Run:  python3 docsearch.py
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import zipfile
import xml.etree.ElementTree as ET
import subprocess
import os
import threading
import concurrent.futures
import sys
import time

# ── Word XML namespace ────────────────────────────────────────────────────
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def wtag(n): return f'{{{W}}}{n}'

# ── Style → type mapping ──────────────────────────────────────────────────
STYLE_TYPE = {'Heading1': 'pocket', 'Heading2': 'hat', 'Heading3': 'block', 'Heading4': 'tag'}
# Level number for each type (used to know when a section ends)
TYPE_LEVEL = {'pocket': 1, 'hat': 2, 'block': 3, 'tag': 4}
# Type label colors (pocket adjusted to dark green)
TYPE_COLOR = {'pocket': '#0b6b3a', 'hat': '#ff9c47', 'block': '#c47fff', 'tag': '#ff6b6b'}

# 7 blue shades that repeat for each visible card row
BLUE_SHADES = ['#0b2a4a', '#0f365c', '#12426f', '#165089', '#1b61a3', '#2a78b6', '#3a90c8']
BLUE_BORDERS = ['#07304c', '#0b3460', '#0f3b74', '#134681', '#1a5698', '#256db4', '#2f82c8']
# Sequence pattern requested: 1..7..1..4 (repeat)
BLUE_SHADE_PATTERN = [0, 1, 2, 3, 4, 5, 6, 5, 4, 3, 2, 1, 0, 1, 2, 3]

# Debounce search (ms) so we don't refresh on every keystroke
SEARCH_DEBOUNCE_MS = 150
# Virtualized list: only this many card widgets exist; they're reused as you scroll
ROW_HEIGHT = 44
N_VISIBLE_SLOTS = 80
# Folder to auto-load .docx from on startup (None to disable)
AUTO_LOAD_FOLDER = os.path.join(os.path.expanduser('~'), 'debate')
# Only auto/load .docx from these subfolders within `AUTO_LOAD_FOLDER`
DEBATE_TOPIC_SPECIFIC_SUBDIRS = ['topic specific', 'topic-specific', 'topic_specific']
DEBATE_GENERAL_SUBDIRS = ['general']
# VBA macro to invoke for the "send" action.
SEND_VBA_MACRO_NAME = 'SendToSpeechCursor'

# ── Parse .docx → sections (lightweight — just headings + preview text) ───
def parse_docx(path):
    sections = []
    with zipfile.ZipFile(path) as z:
        doc_xml = z.read('word/document.xml')

    root = ET.fromstring(doc_xml)
    body = root.find(f'.//{wtag("body")}')
    paragraphs = body.findall(wtag('p'))

    current_pocket = current_hat = current_block = None
    para_idx = 0  # 1-based; incremented for every <w:p> including blanks

    for para in paragraphs:
        pPr = para.find(wtag('pPr'))
        style_el = pPr.find(wtag('pStyle')) if pPr is not None else None
        style = style_el.get(f'{{{W}}}val') if style_el is not None else None
        sec_type = STYLE_TYPE.get(style)
        para_text = ''.join(t.text or '' for t in para.iter(wtag('t')))
        text = para_text.strip()
        para_idx += 1  # every paragraph counts, blank or not

        if not text:
            continue

        if sec_type == 'pocket':
            current_pocket = text; current_hat = None; current_block = None
            sections.append({'type': 'pocket', 'heading': text, 'para_idx': para_idx,
                             'parents': {}, 'preview': []})
        elif sec_type == 'hat':
            current_hat = text; current_block = None
            sections.append({'type': 'hat', 'heading': text, 'para_idx': para_idx,
                             'parents': {'pocket': current_pocket}, 'preview': []})
        elif sec_type == 'block':
            current_block = text
            sections.append({'type': 'block', 'heading': text, 'para_idx': para_idx,
                             'parents': {'pocket': current_pocket, 'hat': current_hat}, 'preview': []})
        elif sec_type == 'tag':
            sections.append({'type': 'tag', 'heading': text, 'para_idx': para_idx,
                             'parents': {'pocket': current_pocket, 'hat': current_hat,
                                         'block': current_block}, 'preview': []})
        else:
            # Accumulate preview text into the last section
            if sections and text:
                sections[-1]['preview'].append(text)

    total_paras = para_idx  # total paragraph count in document

    # Assign end_para_idx: last paragraph belonging to this section (inclusive).
    # That is one before the next heading at the same or higher level, or the
    # last paragraph in the document if this section runs to the end.
    for i, sec in enumerate(sections):
        my_level = TYPE_LEVEL[sec['type']]
        end_idx = None
        for j in range(i + 1, len(sections)):
            next_level = TYPE_LEVEL[sections[j]['type']]
            if next_level <= my_level:
                end_idx = sections[j]['para_idx'] - 1  # stop before the next heading
                break
        sec['end_para_idx'] = end_idx if end_idx is not None else total_paras

    return sections


# ── Copy via Word AppleScript ─────────────────────────────────────────────
def copy_via_word(doc_path, para_start, para_end):
    abs_path = os.path.abspath(doc_path)
    escaped_path = abs_path.replace('\\', '\\\\').replace('"', '\\"')

    script = (
        'tell application "Microsoft Word"\n'
        '    open (POSIX file "%s") with read only\n'
        '    delay 0.15\n'
        '    set theDoc to active document\n'
        '    set startPos to start of content of text object of paragraph %d of theDoc\n'
        '    set endPos to end of content of text object of paragraph %d of theDoc\n'
        '    set theRange to create range theDoc start startPos end endPos\n'
        '    select theRange\n'
        '    activate\n'
        '    delay 0.1\n'
        'end tell\n'
        'tell application "System Events"\n'
        '    tell process "Microsoft Word"\n'
        '        keystroke "c" using command down\n'
        '    end tell\n'
        '    delay 0.05\n'
        'end tell\n'
        'tell application "Microsoft Word"\n'
        '    close active document saving no\n'
        'end tell\n'
        'tell application "System Events"\n'
        '    set frontApp to first application process whose frontmost is true\n'
        '    tell process "Python"\n'
        '        set frontmost to true\n'
        '    end tell\n'
        'end tell\n'
    ) % (escaped_path, para_start, para_end)

    result = subprocess.run(
        ['osascript', '-'],
        input=script,
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip())


def send_tilde_via_word(doc_path, para_idx):
    abs_path = os.path.abspath(doc_path)
    escaped_path = abs_path.replace('\\', '\\\\').replace('"', '\\"')
    para_idx = int(para_idx)
    script = (
        'tell application "Microsoft Word"\n'
        '    open (POSIX file "%s") with read only\n'
        '    delay 0.15\n'
        '    set theDoc to active document\n'
        '    set startPos to start of content of text object of paragraph %d of theDoc\n'
        '    set hdrRange to create range theDoc start startPos end startPos\n'
        '    select hdrRange\n'
        '    activate\n'
        'end tell\n'
        'tell application "System Events"\n'
        '    tell process "Microsoft Word"\n'
        '        -- Open Macro dialog from menu, fallback to Option+F8.\n'
        '        try\n'
        '            click menu item "Macros..." of menu 1 of menu item "Macro" of menu 1 of menu bar item "Tools" of menu bar 1\n'
        '        on error\n'
        '            key code 100 using option down\n'
        '        end try\n'
        '        delay 0.3\n'
        '        keystroke "%s"\n'
        '        delay 0.1\n'
        '        key code 36\n'
        '    end tell\n'
        'end tell\n'
        'tell application "Microsoft Word"\n'
        '    delay 0.2\n'
        '    close active document saving no\n'
        'end tell\n'
    ) % (escaped_path, para_idx, SEND_VBA_MACRO_NAME.replace('\\', '\\\\').replace('"', '\\"'))

    result = subprocess.run(
        ['osascript', '-'],
        input=script,
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip())


# ── GUI ───────────────────────────────────────────────────────────────────
class DocSearchApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('DocSearch')
        self.geometry('900x650')
        self.configure(bg='#0e0e0e')
        self.minsize(700, 500)
        self.docs = []
        self.filtered = []
        self._all_sections_cache = []  # flat list; rebuilt only when docs change
        self._search_job = None  # for debounce
        self._load_failure_note = ''  # appended to results meta (no popups)
        self._doc_display_cache = {}  # doc_path -> display string
        # Used by a small on-screen debug label during wheel scrolling.
        self.scroll_debug_var = tk.StringVar(value='wheel debug ready (scroll over results)')
        self._build_ui()
        self._wheel_accum = 0  # accumulates tiny wheel deltas
        self._scroll_velocity = 0
        self._scroll_animating = False
        self._last_scroll_time = 0
        self._last_wheel_debug = 0
        self._last_wheel_debug_move = 0
        self._scroll_debug_enabled = True
        self._scroll_debug_count = 0
        self._smooth_scroll_after_id = None

    def _build_ui(self):
        BG = '#0e0e0e'; SURF = '#161616'; ACC = '#e8ff47'; DIM = '#777'; TEXT = '#e2e2e2'
        self.bind_all('<MouseWheel>', self._on_mousewheel)
        self.bind_all('<Button-4>', self._on_mousewheel)  # linux up
        self.bind_all('<Button-5>', self._on_mousewheel)  # linux down
        self.bind_all('<ButtonPress-4>', self._on_mousewheel)
        self.bind_all('<ButtonPress-5>', self._on_mousewheel)
        self.bind_all('<ButtonRelease-4>', self._on_mousewheel)
        self.bind_all('<ButtonRelease-5>', self._on_mousewheel)

        # Keyboard scroll fallback (trackpad wheel not firing reliably for you).
        self.bind_all('<Up>', lambda e: self._on_scroll_key(e, -1))
        self.bind_all('<Down>', lambda e: self._on_scroll_key(e, 1))
        self.bind_all('<Prior>', lambda e: self._on_scroll_key(e, -1))   # PageUp
        self.bind_all('<Next>', lambda e: self._on_scroll_key(e, 1))      # PageDown

        # Header
        hdr = tk.Frame(self, bg=SURF, height=52)
        hdr.pack(fill='x'); hdr.pack_propagate(False)
        tk.Label(hdr, text='DOCSEARCH', font=('Courier', 13, 'bold'),
                 fg=ACC, bg=SURF).pack(side='left', padx=20, pady=14)
        self.pills_frame = tk.Frame(hdr, bg=SURF)
        self.pills_frame.pack(side='right', padx=12, pady=10)

        # Toolbar
        toolbar = tk.Frame(self, bg=BG, pady=10)
        toolbar.pack(fill='x', padx=20)
        tk.Button(toolbar, text='+ Add .docx files', font=('Courier', 11),
                  fg=ACC, bg=SURF, activeforeground=ACC, activebackground='#222',
                  bd=0, padx=12, pady=6, cursor='hand2',
                  command=self._add_files).pack(side='left')

        self.filter_vars = {}
        filter_frame = tk.Frame(toolbar, bg=BG)
        filter_frame.pack(side='left', padx=20)
        tk.Label(filter_frame, text='SHOW:', font=('Courier', 9), fg=DIM, bg=BG).pack(side='left', padx=(0,6))
        for t, color in TYPE_COLOR.items():
            var = tk.BooleanVar(value=True)
            self.filter_vars[t] = var
            tk.Checkbutton(filter_frame, text=t.upper(), font=('Courier', 9, 'bold'),
                           fg=color, bg=BG, selectcolor='#1a1a1a',
                           activeforeground=color, activebackground=BG,
                           variable=var, command=self._refresh_results,
                           bd=0, highlightthickness=0).pack(side='left', padx=4)

        # Search
        search_frame = tk.Frame(self, bg=BG)
        search_frame.pack(fill='x', padx=20, pady=(0,10))
        self.search_var = tk.StringVar()
        self.search_var.trace_add('write', lambda *_: self._schedule_refresh())
        tk.Entry(search_frame, textvariable=self.search_var,
                 font=('Courier', 13), fg=TEXT, bg=SURF,
                 insertbackground=TEXT, bd=0,
                 highlightthickness=1, highlightcolor=ACC,
                 highlightbackground='#2a2a2a').pack(fill='x', ipady=8, padx=1)

        self.meta_var = tk.StringVar()
        tk.Label(self, textvariable=self.meta_var, font=('Courier', 9),
                 fg=DIM, bg=BG, anchor='w').pack(fill='x', padx=22)

        # If wheel events don't reach stdout, this still shows activity.
        tk.Label(self, textvariable=self.scroll_debug_var, font=('Courier', 8),
                 fg='#666', bg=BG, anchor='w').pack(fill='x', padx=22)

        # Results canvas
        results_frame = tk.Frame(self, bg=BG)
        self.results_frame = results_frame
        results_frame.pack(fill='both', expand=True, padx=20, pady=(6,20))
        self.scrollbar = tk.Scrollbar(results_frame, bg=SURF, troughcolor=BG, bd=0, width=8)
        self.scrollbar.pack(side='right', fill='y')
        self.canvas = tk.Canvas(results_frame, bg=BG, bd=0, highlightthickness=0,
                                yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side='left', fill='both', expand=True)
        self.scrollbar.config(command=self._scroll_command)
        self.results_inner = tk.Frame(self.canvas, bg=BG)
        self.canvas_window = self.canvas.create_window((0,0), window=self.results_inner, anchor='nw')
        self.canvas.bind('<Configure>', lambda e: (self.canvas.itemconfig(self.canvas_window, width=e.width), self._on_canvas_configure()))
        self.canvas.bind('<MouseWheel>', self._on_mousewheel)
        self.results_inner.bind('<MouseWheel>', self._on_mousewheel)
        self._slot_cards = []
        self._show_empty('Add some .docx files to get started.')
        self.after(100, self._auto_load_folder)

    def _sec_full_path(self, sec):
        """Return a readable hierarchy path for a section."""
        parents = sec.get('parents') or {}
        parts = []
        for key in ('pocket', 'hat', 'block'):
            v = parents.get(key) or ''
            if v:
                parts.append(v)
        heading = sec.get('heading') or ''
        if heading:
            parts.append(heading)
        return ' -> '.join(parts)

    def _sec_parents_path(self, sec):
        """Return pocket -> hat -> block for the current section (heading omitted)."""
        parents = sec.get('parents') or {}
        parts = []
        for key in ('pocket', 'hat', 'block'):
            v = parents.get(key) or ''
            if v:
                parts.append(v)
        return ' -> '.join(parts)

    def _sec_truncate(self, text, max_chars):
        if not text:
            return ''
        if len(text) <= max_chars:
            return text
        return text[: max_chars - 1] + '…'

    def _doc_display_name(self, doc_path):
        """Show path like 'topic specific/<subdirs>/<file>' or 'general/<...>/<file>'."""
        if not doc_path:
            return ''
        doc_path = os.path.abspath(doc_path)
        cached = self._doc_display_cache.get(doc_path)
        if cached is not None:
            return cached

        debate_base = os.path.abspath(os.path.expanduser(AUTO_LOAD_FOLDER))
        try:
            rel = os.path.relpath(doc_path, debate_base)
        except Exception:
            rel = os.path.basename(doc_path)

        parts = rel.split(os.sep) if rel else []
        group_label = None
        start_idx = None

        topic_labels = {'topic specific', 'topic-specific', 'topic_specific'}
        general_labels = {'general'}

        for i, part in enumerate(parts):
            if part in topic_labels:
                group_label = 'topic specific'
                start_idx = i
                break
            if part in general_labels:
                group_label = 'general'
                start_idx = i
                break

        if group_label is None or start_idx is None:
            display = os.path.basename(doc_path)
        else:
            rest = parts[start_idx + 1:]
            display = os.path.join(group_label, *rest) if rest else group_label
            filename = os.path.basename(doc_path)
            if rest and rest[-1] != filename:
                display = os.path.join(display, filename)

        if len(display) > 70:
            display = self._sec_truncate(display, 70)
        self._doc_display_cache[doc_path] = display
        return display

    def _scroll_command(self, *args):
        self.canvas.yview(*args)
        self.after_idle(self._update_visible_cards)

    def _on_canvas_configure(self):
        self.canvas.itemconfig(self.canvas_window, width=self.canvas.winfo_width())
        self.after_idle(self._update_visible_cards)

    def _on_mousewheel(self, event):
        self._scroll_debug_count += 1
        inside = self._pointer_in_results()
        widget = getattr(self, 'results_frame', None) or self.canvas
        try:
            x = self.winfo_pointerx()
            y = self.winfo_pointery()
            x0 = widget.winfo_rootx()
            y0 = widget.winfo_rooty()
            x1 = x0 + widget.winfo_width()
            y1 = y0 + widget.winfo_height()
        except Exception:
            x = y = x0 = y0 = x1 = y1 = None
        delta = getattr(event, 'delta', None)
        num = getattr(event, 'num', None)
        self.scroll_debug_var.set(
            f'[wheel] called={self._scroll_debug_count} inside={inside} '
            f'delta={delta} num={num} accum={self._wheel_accum} '
            f'ptr=({x},{y}) rect=({x0},{y0})-({x1},{y1})'
        )

        now = time.time()

        if not inside:
            return

        if hasattr(event, "delta") and getattr(event, "delta", 0) != 0:
            raw = -event.delta
        elif hasattr(event, "num"):
            raw = -120 if event.num == 4 else 120
        else:
            return

        self._wheel_accum += raw

        threshold = 30
        move_units = int(self._wheel_accum / threshold)

        if move_units == 0 and abs(self._wheel_accum) >= 1:
            move_units = 1 if self._wheel_accum > 0 else -1
            self._wheel_accum -= move_units

        if move_units != 0:
            self.canvas.yview_scroll(move_units, "units")
            self._wheel_accum -= move_units * threshold

        dt = now - self._last_scroll_time if self._last_scroll_time else 0.016
        self._last_scroll_time = now

        if dt > 0:
            self._scroll_velocity = raw / dt

        if not self._scroll_animating:
            self._scroll_animating = True
            self._animate_scroll()

        self.after_idle(self._update_visible_cards)
        return "break"

    def _animate_scroll(self):
        friction = 0.90
        min_velocity = 10

        if abs(self._scroll_velocity) < min_velocity:
            self._scroll_animating = False
            return

        move = self._scroll_velocity * 0.016
        move_units = int(move / 120)

        if move_units != 0:
            self.canvas.yview_scroll(move_units, "units")

        self._scroll_velocity *= friction
        self.after_idle(self._update_visible_cards)
        self.after(16, self._animate_scroll)

    def _pointer_in_results(self):
        """Return True if the mouse pointer is currently over the results area."""
        try:
            x = self.winfo_pointerx()
            y = self.winfo_pointery()
            widget = getattr(self, 'results_frame', None) or self.canvas
            x0 = widget.winfo_rootx()
            y0 = widget.winfo_rooty()
            x1 = x0 + widget.winfo_width()
            y1 = y0 + widget.winfo_height()
            return (x0 <= x < x1) and (y0 <= y < y1)
        except Exception:
            return False

    def _get_scrollable_range_px(self):
        total_h = getattr(self, '_scrollregion_height', None)
        if not total_h:
            return None
        viewport_h = self.canvas.winfo_height()
        return max(1, total_h - viewport_h)

    def _smooth_scroll_to_frac(self, target_frac, duration_ms=180):
        scrollable = self._get_scrollable_range_px()
        if scrollable is None:
            return

        if self._smooth_scroll_after_id is not None:
            try:
                self.after_cancel(self._smooth_scroll_after_id)
            except Exception:
                pass
            self._smooth_scroll_after_id = None

        target_frac = max(0.0, min(1.0, float(target_frac)))
        start_frac = self.canvas.yview()[0]
        start_time = time.time()
        dur_s = max(0.05, duration_ms / 1000.0)

        def step():
            t = (time.time() - start_time) / dur_s
            if t >= 1.0:
                self.canvas.yview_moveto(target_frac)
                self._update_visible_cards()
                self._smooth_scroll_after_id = None
                return
            ease = 1 - (1 - t) ** 3
            frac = start_frac + (target_frac - start_frac) * ease
            self.canvas.yview_moveto(frac)
            self._update_visible_cards()
            self._smooth_scroll_after_id = self.after(16, step)

        self._smooth_scroll_after_id = self.after(16, step)

    def _smooth_scroll_by_pixels(self, delta_px, duration_ms=180):
        scrollable = self._get_scrollable_range_px()
        if scrollable is None:
            return

        cur_frac = self.canvas.yview()[0]
        cur_px = cur_frac * scrollable
        target_px = cur_px + float(delta_px)
        target_px = max(0.0, min(float(scrollable), target_px))
        target_frac = target_px / scrollable
        self._smooth_scroll_to_frac(target_frac, duration_ms=duration_ms)

    def _scroll_by_units(self, units):
        if not units:
            return
        delta_px = float(units) * ROW_HEIGHT
        duration_ms = int(max(120, min(300, (abs(delta_px) / max(1.0, ROW_HEIGHT)) * 60)))
        self._smooth_scroll_by_pixels(delta_px, duration_ms=duration_ms)

    def _on_scroll_key(self, event, direction):
        if not self._pointer_in_results():
            return
        self._scroll_animating = False
        self._scroll_velocity = 0
        if getattr(event, 'keysym', '') in ('Prior', 'Next'):
            page_rows = max(1, int(self.canvas.winfo_height() / ROW_HEIGHT))
            step_rows = int(max(1, page_rows * 0.85)) * direction
        else:
            step_rows = 2 * direction
        self._scroll_by_units(step_rows)
        return 'break'

    def _get_debate_roots(self):
        base = os.path.abspath(os.path.expanduser(AUTO_LOAD_FOLDER))
        roots = []
        for sub in DEBATE_TOPIC_SPECIFIC_SUBDIRS:
            p = os.path.join(base, sub)
            if os.path.isdir(p):
                roots.append(os.path.abspath(p))
        for sub in DEBATE_GENERAL_SUBDIRS:
            p = os.path.join(base, sub)
            if os.path.isdir(p):
                roots.append(os.path.abspath(p))
        return roots

    def _get_debate_root_groups(self):
        """Return (topic_specific_roots, general_roots)."""
        base = os.path.abspath(os.path.expanduser(AUTO_LOAD_FOLDER))
        topic_roots = []
        for sub in DEBATE_TOPIC_SPECIFIC_SUBDIRS:
            p = os.path.join(base, sub)
            if os.path.isdir(p):
                topic_roots.append(os.path.abspath(p))
        general_roots = []
        for sub in DEBATE_GENERAL_SUBDIRS:
            p = os.path.join(base, sub)
            if os.path.isdir(p):
                general_roots.append(os.path.abspath(p))
        return topic_roots, general_roots

    def _sort_docs_by_debate_group(self):
        topic_roots, general_roots = self._get_debate_root_groups()

        def doc_priority(doc):
            p = doc.get('path') or ''
            abs_p = os.path.abspath(p)
            for root in topic_roots:
                try:
                    if os.path.commonpath([abs_p, root]) == root:
                        return 0
                except Exception:
                    continue
            for root in general_roots:
                try:
                    if os.path.commonpath([abs_p, root]) == root:
                        return 1
                except Exception:
                    continue
            return 99

        self.docs.sort(key=doc_priority)

    def _path_allowed(self, path, roots=None):
        abs_path = os.path.abspath(path)
        if roots is None:
            roots = self._get_debate_roots()
        for root in roots:
            try:
                if os.path.commonpath([abs_path, root]) == root:
                    return True
            except Exception:
                continue
        return False

    def _auto_load_folder(self):
        """Load .docx from `debate/topic specific` and `debate/general` on startup."""
        if not AUTO_LOAD_FOLDER:
            return
        folder = os.path.abspath(os.path.expanduser(AUTO_LOAD_FOLDER))
        if not os.path.isdir(folder):
            self.meta_var.set(f'Auto-load folder not found: {folder}')
            return

        roots = self._get_debate_roots()
        if not roots:
            self.meta_var.set(f'No allowed subfolders found under: {folder}')
            return

        paths = []
        for allowed_root in roots:
            for dirpath, _dirnames, filenames in os.walk(allowed_root):
                for f in filenames:
                    if f.lower().endswith('.docx'):
                        paths.append(os.path.join(dirpath, f))
        if not paths:
            self.meta_var.set('No .docx files found in allowed folders.')
            return
        self._load_paths(paths)

    def _load_paths(self, paths):
        """Load a list of .docx paths (used by _add_files and _auto_load_folder)."""
        self._load_failure_note = ''

        roots = self._get_debate_roots()
        allowed = [p for p in paths if self._path_allowed(p, roots=roots)]
        allowed_set = set(allowed)
        skipped = [p for p in paths if p not in allowed_set]
        if skipped:
            print(f'Skipping {len(skipped)} file(s) outside allowed debate folders.')
            paths = allowed
        if not paths:
            return

        topic_roots, general_roots = self._get_debate_root_groups()

        def sort_key(p):
            abs_p = os.path.abspath(p)
            for root in topic_roots:
                try:
                    if os.path.commonpath([abs_p, root]) == root:
                        return (0, abs_p.lower())
                except Exception:
                    continue
            for root in general_roots:
                try:
                    if os.path.commonpath([abs_p, root]) == root:
                        return (1, abs_p.lower())
                except Exception:
                    continue
            return (99, abs_p.lower())

        paths = sorted(paths, key=sort_key)

        if len(paths) > 10:
            self.meta_var.set('Loading…')
            self.update_idletasks()

            def load_all():
                loaded_by_idx = {}
                errors_by_idx = {}
                with concurrent.futures.ThreadPoolExecutor(max_workers=8) as ex:
                    future_to_idx = {}
                    for idx, p in enumerate(paths):
                        future = ex.submit(parse_docx, p)
                        future_to_idx[future] = idx
                    for future in concurrent.futures.as_completed(future_to_idx):
                        idx = future_to_idx[future]
                        p = paths[idx]
                        name = os.path.basename(p)
                        try:
                            sections = future.result()
                            loaded_by_idx[idx] = (name, p, sections)
                        except Exception as e:
                            errors_by_idx[idx] = (name, str(e))

                loaded = [loaded_by_idx[i] for i in range(len(paths)) if i in loaded_by_idx]
                errors = [errors_by_idx[i] for i in range(len(paths)) if i in errors_by_idx]
                return loaded, errors

            def on_done():
                loaded, errors = self._add_files_result
                for name, path, sections in loaded:
                    entry = {'name': name, 'path': path, 'sections': sections}
                    existing = next((i for i, d in enumerate(self.docs) if d['name'] == name), None)
                    if existing is not None:
                        self.docs[existing] = entry
                    else:
                        self.docs.append(entry)
                if errors:
                    self._load_failure_note = f' • {len(errors)} failed to load'
                    for name, err in errors[:5]:
                        print(f'Could not load {name}: {err}')
                self._sort_docs_by_debate_group()
                self._rebuild_sections_cache()
                self._render_pills()
                self._refresh_results()

            def run():
                self._add_files_result = load_all()
                self.after(0, on_done)

            threading.Thread(target=run, daemon=True).start()
        else:
            errors = []
            for path in paths:
                name = os.path.basename(path)
                try:
                    sections = parse_docx(path)
                    entry = {'name': name, 'path': path, 'sections': sections}
                    existing = next((i for i, d in enumerate(self.docs) if d['name'] == name), None)
                    if existing is not None:
                        self.docs[existing] = entry
                    else:
                        self.docs.append(entry)
                except Exception as e:
                    errors.append((name, str(e)))
                    print(f'Could not load {name}: {e}')
            if errors:
                self._load_failure_note = f' • {len(errors)} failed to load'
            self._sort_docs_by_debate_group()
            self._rebuild_sections_cache()
            self._render_pills()
            self._refresh_results()

    def _add_files(self):
        allowed_roots = self._get_debate_roots()
        initialdir = allowed_roots[0] if allowed_roots else os.path.abspath(os.path.expanduser(AUTO_LOAD_FOLDER))
        paths = filedialog.askopenfilenames(
            initialdir=initialdir,
            filetypes=[('Word Documents', '*.docx')],
        )
        if not paths:
            return
        self._doc_display_cache = {}
        self._load_paths(paths)

    def _rebuild_sections_cache(self):
        """Rebuild flat section list only when docs change (avoids rebuilding on every keystroke)."""
        self._all_sections_cache = [
            {**s, 'docName': self._doc_display_name(doc['path']), 'docPath': doc['path']}
            for doc in self.docs for s in doc['sections']
        ]

    def _schedule_refresh(self):
        """Debounce search: refresh after SEARCH_DEBOUNCE_MS of no typing."""
        if self._search_job is not None:
            self.after_cancel(self._search_job)
        self._search_job = self.after(SEARCH_DEBOUNCE_MS, self._do_refresh)

    def _do_refresh(self):
        self._search_job = None
        self._refresh_results()

    def _render_pills(self):
        for w in self.pills_frame.winfo_children(): w.destroy()
        for i, doc in enumerate(self.docs):
            pill = tk.Frame(self.pills_frame, bg='#1a1a1a')
            pill.pack(side='left', padx=3)
            tk.Label(pill, text=doc['name'], font=('Courier', 9), fg='#777', bg='#1a1a1a', padx=6, pady=3).pack(side='left')
            tk.Button(pill, text='×', font=('Courier', 9), fg='#ff4747', bg='#1a1a1a',
                      activeforeground='#ff4747', activebackground='#222', bd=0, padx=4, cursor='hand2',
                      command=lambda idx=i: self._remove_doc(idx)).pack(side='left')

    def _remove_doc(self, idx):
        self.docs.pop(idx)
        self._doc_display_cache = {}
        self._rebuild_sections_cache()
        self._render_pills()
        self._refresh_results()

    def _refresh_results(self):
        query = self.search_var.get().strip().lower()
        active = {t for t, var in self.filter_vars.items() if var.get()}
        type_filtered = [s for s in self._all_sections_cache if s['type'] in active]
        if query:
            filtered = [s for s in type_filtered if
                query in s['heading'].lower() or
                query in s['docName'].lower() or
                any(query in (s['parents'].get(k) or '').lower() for k in s['parents']) or
                any(query in line.lower() for line in s['preview'])]
        else:
            filtered = type_filtered
        self.filtered = filtered
        all_count = len(self._all_sections_cache)
        if all_count:
            self.meta_var.set(f'{len(filtered)} / {all_count} sections{self._load_failure_note}')
        else:
            self.meta_var.set('')
        self._render_results()

    def _render_results(self):
        for w in self.results_inner.winfo_children():
            w.destroy()
        self._slot_cards = []
        if not self.docs:
            self._show_empty('Add some .docx files to get started.')
            return
        if not self._all_sections_cache:
            self._show_empty('No sections found.')
            return
        if not self.filtered:
            self._show_empty('No matches.')
            return
        n = len(self.filtered)
        self.results_inner.config(height=n * ROW_HEIGHT)
        self._ensure_slot_cards()
        self.canvas.configure(scrollregion=(0, 0, 0, n * ROW_HEIGHT))
        self._scrollregion_height = n * ROW_HEIGHT
        self.canvas.yview_moveto(0)
        self._update_visible_cards()

    def _show_empty(self, msg):
        for w in self.results_inner.winfo_children():
            w.destroy()
        self._slot_cards = []
        tk.Label(self.results_inner, text=msg, font=('Courier', 11),
                 fg='#555', bg='#0e0e0e', pady=40).pack()

    def _ensure_slot_cards(self):
        """Create N_VISIBLE_SLOTS compact card widgets if we don't have them."""
        SURF = '#161616'
        while len(self._slot_cards) < N_VISIBLE_SLOTS:
            card = tk.Frame(self.results_inner, bg=SURF, height=ROW_HEIGHT,
                            highlightthickness=1, highlightbackground='#2a2a2a')
            card.pack_propagate(False)
            type_lbl = tk.Label(card, text='', font=('Courier', 9, 'bold'),
                               fg=TYPE_COLOR['pocket'], bg=SURF, padx=6, pady=2)
            type_lbl.pack(side='left', padx=(0, 6))

            mid_frame = tk.Frame(card, bg=SURF)
            mid_frame.pack(side='left', fill='x', expand=True, padx=(0, 8))
            row_frame = tk.Frame(mid_frame, bg=SURF)
            row_frame.pack(side='top', fill='x')

            main_lbl = tk.Label(
                row_frame, text='', font=('Courier', 11, 'bold'),
                fg='#ffffff', bg=SURF, anchor='w'
            )
            main_lbl.pack(side='left', fill='x', expand=True)

            side_lbl = tk.Label(
                row_frame, text='', font=('Courier', 9),
                fg='#ffffff', bg=SURF, anchor='w'
            )
            side_lbl.pack(side='right')

            doc_lbl = tk.Label(card, text='', font=('Courier', 9), fg='#ffffff', bg=SURF)
            doc_lbl.pack(side='right', padx=(0, 8))
            send_btn = tk.Button(card, text='send', font=('Courier', 9, 'bold'),
                                 fg='#000', bg='#222', activeforeground='#000',
                                 activebackground='#2a2a2a', bd=0, padx=8, pady=3, cursor='hand2')
            send_btn.pack(side='right', padx=4)
            copy_btn = tk.Button(card, text='copy', font=('Courier', 9, 'bold'),
                                 fg='#000', bg='#222', activeforeground='#000',
                                 activebackground='#2a2a2a', bd=0, padx=8, pady=3, cursor='hand2')
            copy_btn.pack(side='right', padx=4)
            for w in (card, type_lbl, main_lbl, side_lbl, doc_lbl, send_btn, copy_btn):
                w.bind('<MouseWheel>', self._on_mousewheel)
                w.bind('<Button-4>', self._on_mousewheel)
                w.bind('<Button-5>', self._on_mousewheel)
            self._slot_cards.append({
                'frame': card, 'mid_frame': mid_frame, 'row_frame': row_frame,
                'type_lbl': type_lbl, 'main_lbl': main_lbl, 'side_lbl': side_lbl,
                'doc_lbl': doc_lbl, 'send_btn': send_btn, 'copy_btn': copy_btn
            })

    def _update_visible_cards(self):
        """Reposition and update slot cards to show the visible slice of self.filtered."""
        if not self._slot_cards or not self.filtered:
            return
        try:
            y0, y1 = self.canvas.yview()
        except Exception:
            return
        n = len(self.filtered)
        start = int(y0 * n)
        start = max(0, min(start, n - 1))
        end = min(start + N_VISIBLE_SLOTS, n)
        for i, slot in enumerate(self._slot_cards):
            idx = start + i
            if idx < end:
                sec = self.filtered[idx]
                w = self.canvas.winfo_width()
                if w <= 1:
                    w = 800
                slot['frame'].place(x=0, y=idx * ROW_HEIGHT, width=w, height=ROW_HEIGHT)
                self._update_slot_content(slot, sec, idx)
                slot['frame'].tkraise()
            else:
                slot['frame'].place_forget()

    def _update_slot_content(self, slot, sec, idx):
        shade_idx = BLUE_SHADE_PATTERN[idx % len(BLUE_SHADE_PATTERN)]
        shade = BLUE_SHADES[shade_idx]
        border = BLUE_BORDERS[shade_idx]

        slot['frame'].config(bg=shade, highlightbackground=border)
        slot['mid_frame'].config(bg=shade)
        slot['row_frame'].config(bg=shade)
        slot['type_lbl'].config(bg=shade)
        slot['main_lbl'].config(bg=shade)
        slot['side_lbl'].config(bg=shade)
        slot['doc_lbl'].config(bg=shade)
        slot['send_btn'].config(bg=border, activebackground=shade, fg='#000', activeforeground='#000')
        slot['copy_btn'].config(bg=border, activebackground=shade, fg='#000', activeforeground='#000')

        color = TYPE_COLOR[sec['type']]
        slot['type_lbl'].config(text=sec['type'].upper(), fg=color)
        slot['main_lbl'].config(text=self._sec_truncate(sec.get('heading') or '', 85))
        side = self._sec_parents_path(sec)
        slot['side_lbl'].config(text=self._sec_truncate(side, 65), fg='#ffffff')
        slot['doc_lbl'].config(text=sec['docName'], fg='#ffffff')
        slot['send_btn'].config(command=lambda s=sec, b=slot['send_btn']: self._send_section(s, b))
        slot['copy_btn'].config(command=lambda s=sec, b=slot['copy_btn']: self._copy_section(s, b))

    def _render_card(self, sec, idx):
        SURF = '#161616'; TEXT = '#e2e2e2'; DIM = '#666'
        color = TYPE_COLOR[sec['type']]
        crumbs = self._sec_parents_path(sec)

        card = tk.Frame(self.results_inner, bg=SURF, bd=0,
                        highlightthickness=1, highlightbackground='#2a2a2a')
        card.pack(fill='x', pady=4)

        hdr_row = tk.Frame(card, bg=SURF)
        hdr_row.pack(fill='x', padx=12, pady=8)

        tk.Label(hdr_row, text=sec['type'].upper(), font=('Courier', 9, 'bold'),
                 fg=color, bg=SURF, padx=6, pady=2).pack(side='left', padx=(0,8))
        mid_frame = tk.Frame(hdr_row, bg=SURF)
        mid_frame.pack(side='left', fill='x', expand=True)
        row_frame = tk.Frame(mid_frame, bg=SURF)
        row_frame.pack(side='top', fill='x')
        tk.Label(row_frame, text=sec['heading'], font=('Courier', 11, 'bold'),
                 fg='#ffffff', bg=SURF, anchor='w').pack(side='left', fill='x', expand=True)
        if crumbs:
            tk.Label(row_frame, text=crumbs, font=('Courier', 9),
                     fg=DIM, bg=SURF, anchor='w').pack(side='right')
        tk.Label(hdr_row, text=sec['docName'], font=('Courier', 9),
                 fg=DIM, bg=SURF).pack(side='right', padx=(0,8))

        copy_btn = tk.Button(hdr_row, text='copy', font=('Courier', 9, 'bold'),
                             fg='#888', bg='#222', activeforeground=color,
                             activebackground='#2a2a2a', bd=0, padx=8, pady=3, cursor='hand2')
        copy_btn.pack(side='right', padx=4)
        copy_btn.config(command=lambda s=sec, b=copy_btn: self._copy_section(s, b))
        send_btn = tk.Button(hdr_row, text='send', font=('Courier', 9, 'bold'),
                             fg='#888', bg='#222', activeforeground=color,
                             activebackground='#2a2a2a', bd=0, padx=8, pady=3, cursor='hand2')
        send_btn.pack(side='right', padx=4)
        send_btn.config(command=lambda s=sec, b=send_btn: self._send_section(s, b))

        preview = ' '.join(sec['preview'][:3])[:300]
        if preview:
            tk.Label(card, text=preview, font=('Courier', 10),
                     fg='#666', bg=SURF, anchor='w', justify='left',
                     wraplength=820, padx=12, pady=6).pack(fill='x', anchor='w')

        card.bind('<Enter>', lambda e, c=card: c.config(highlightbackground='#444'))
        card.bind('<Leave>', lambda e, c=card: c.config(highlightbackground='#2a2a2a'))

    def _copy_section(self, sec, btn):
        btn.config(text='…', fg='#aaa')
        self.update_idletasks()

        def do_copy():
            err = None
            try:
                # para_idx and end_para_idx are 1-based paragraph numbers matching
                # Word's own `paragraph N of active document` count.
                copy_via_word(sec['docPath'], sec['para_idx'], sec['end_para_idx'])
            except Exception as e:
                err = e

            def on_done():
                if err:
                    btn.config(text='copy', fg='#888')
                    messagebox.showerror('Copy failed',
                        f'Could not copy via Word.\n\nMake sure Microsoft Word is installed.\n\nError: {err}')
                else:
                    btn.config(text='✓ copied', fg='#47ffb8')
                    self.after(2000, lambda: btn.config(text='copy', fg='#888'))

            self.after(0, on_done)

        threading.Thread(target=do_copy, daemon=True).start()

    def _send_section(self, sec, btn):
        btn.config(text='…', fg='#aaa')
        self.update_idletasks()

        def do_send():
            err = None
            try:
                print(
                    f"[send_tilde_via_word] {sec['type']} {sec['heading']!r} "
                    f"para {sec['para_idx']} doc={sec['docPath']}"
                )
                send_tilde_via_word(sec['docPath'], sec['para_idx'])
            except Exception as e:
                err = e

            def on_done():
                if err:
                    btn.config(text='send', fg='#888')
                    messagebox.showerror('Send failed',
                        f'Could not run VBA send macro via Word.\n\nMake sure Microsoft Word is installed and macro "{SEND_VBA_MACRO_NAME}" exists.\n\nError: {err}')
                else:
                    btn.config(text='✓ sent', fg='#47ffb8')
                    self.after(2000, lambda: btn.config(text='send', fg='#888'))

            self.after(0, on_done)

        threading.Thread(target=do_send, daemon=True).start()


if __name__ == '__main__':
    app = DocSearchApp()
    app.mainloop()