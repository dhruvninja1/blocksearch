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

# ── Word XML namespace ────────────────────────────────────────────────────
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def wtag(n): return f'{{{W}}}{n}'

# ── Style → type mapping ──────────────────────────────────────────────────
STYLE_TYPE = {'Heading1': 'pocket', 'Heading2': 'hat', 'Heading3': 'block', 'Heading4': 'tag'}
# Level number for each type (used to know when a section ends)
TYPE_LEVEL = {'pocket': 1, 'hat': 2, 'block': 3, 'tag': 4}
TYPE_COLOR = {'pocket': '#47b8ff', 'hat': '#ff9c47', 'block': '#c47fff', 'tag': '#ff6b6b'}

# Debounce search (ms) so we don't refresh on every keystroke
SEARCH_DEBOUNCE_MS = 150
# Virtualized list: only this many card widgets exist; they're reused as you scroll
ROW_HEIGHT = 44
N_VISIBLE_SLOTS = 80
# Folder to auto-load .docx from on startup (None to disable)
AUTO_LOAD_FOLDER = os.path.join(os.path.expanduser('~'), 'debate')

# ── Parse .docx → sections (lightweight — just headings + preview text) ───
def parse_docx(path):
    sections = []
    with zipfile.ZipFile(path) as z:
        doc_xml = z.read('word/document.xml')

    root = ET.fromstring(doc_xml)
    body = root.find(f'.//{wtag("body")}')
    paragraphs = body.findall(wtag('p'))

    current_pocket = current_hat = current_block = None
    para_idx = 0
    char_pos = 0
    para_start_positions = {}  # 1-based para_idx -> char position at start (for copy range)

    for para in paragraphs:
        pPr = para.find(wtag('pPr'))
        style_el = pPr.find(wtag('pStyle')) if pPr is not None else None
        style = style_el.get(f'{{{W}}}val') if style_el is not None else None
        sec_type = STYLE_TYPE.get(style)
        para_text = ''.join(t.text or '' for t in para.iter(wtag('t')))
        text = para_text.strip()
        para_idx += 1
        para_start_positions[para_idx] = char_pos
        char_pos += len(para_text) + 1  # +1 for paragraph mark (same as get_char_range)

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

    # Assign end_para_idx and char_start/char_end for each section
    for i, sec in enumerate(sections):
        my_level = TYPE_LEVEL[sec['type']]
        end_idx = None
        for j in range(i + 1, len(sections)):
            next_level = TYPE_LEVEL[sections[j]['type']]
            if next_level <= my_level:
                end_idx = sections[j]['para_idx']
                break
        sec['end_para_idx'] = end_idx  # None = goes to end of doc
        sec['char_start'] = para_start_positions[sec['para_idx']]
        sec['char_end'] = para_start_positions[end_idx] if end_idx is not None else char_pos

    return sections


# ── Copy via Word AppleScript ─────────────────────────────────────────────
def copy_via_word(doc_path, char_start, char_end):
    """
    Opens the .docx in Word, selects the character range, copies via Word, closes.
    char_start/char_end are precomputed when the doc is loaded (no re-parse on copy).
    """
    abs_path = os.path.abspath(doc_path)
    escaped_path = abs_path.replace('\\', '\\\\').replace('"', '\\"')

    # No activate (Word opens in background). Minimal delay so doc is ready.
    script = (
        'tell application "Microsoft Word"\n'
        '    open (POSIX file "%s") with read only\n'
        '    delay 0.05\n'
        '    set theRange to create range active document start %d end %d\n'
        '    select theRange\n'
        '    copy object selection\n'
        '    close active document saving no\n'
        'end tell\n'
    ) % (escaped_path, char_start, char_end)

    result = subprocess.run(
        ['osascript', '-'],
        input=script,
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip())


def get_char_range(doc_path, heading_text, sec_type):
    """
    Parse the docx XML to get the character start/end offsets
    for the section starting at heading_text.
    Word counts characters the same way we do — one per character,
    plus 1 for each paragraph mark.
    """
    with zipfile.ZipFile(doc_path) as z:
        doc_xml = z.read('word/document.xml')

    root = ET.fromstring(doc_xml)
    body = root.find(f'.//{wtag("body")}')
    paragraphs = body.findall(wtag('p'))

    my_level = TYPE_LEVEL[sec_type]

    # Walk paragraphs, counting characters (text + 1 for paragraph mark)
    char_pos = 0
    section_start = None
    section_end = None

    for para in paragraphs:
        pPr = para.find(wtag('pPr'))
        style_el = pPr.find(wtag('pStyle')) if pPr is not None else None
        style = style_el.get(f'{{{W}}}val') if style_el is not None else None
        para_type = STYLE_TYPE.get(style)
        para_level = TYPE_LEVEL.get(para_type, 99)

        para_text = ''.join(t.text or '' for t in para.iter(wtag('t')))
        para_len = len(para_text) + 1  # +1 for paragraph mark

        if section_start is None:
            # Looking for our heading
            if para_text.strip() == heading_text.strip():
                section_start = char_pos
        else:
            # We found our heading — now look for the end
            if para_level <= my_level:
                section_end = char_pos
                break

        char_pos += para_len

    if section_start is None:
        raise RuntimeError(f'Heading not found: {heading_text!r}')

    if section_end is None:
        section_end = char_pos  # goes to end of doc

    return section_start, section_end


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
        self._build_ui()

    def _build_ui(self):
        BG = '#0e0e0e'; SURF = '#161616'; ACC = '#e8ff47'; DIM = '#777'; TEXT = '#e2e2e2'

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

        # Results canvas
        results_frame = tk.Frame(self, bg=BG)
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

    def _scroll_command(self, *args):
        self.canvas.yview(*args)
        self.after_idle(self._update_visible_cards)

    def _on_canvas_configure(self):
        self.canvas.itemconfig(self.canvas_window, width=self.canvas.winfo_width())
        self.after_idle(self._update_visible_cards)

    def _on_mousewheel(self, e):
        self.canvas.yview_scroll(int(-1 * (e.delta / 120)), 'units')
        self.after_idle(self._update_visible_cards)

    def _auto_load_folder(self):
        """Load all .docx from AUTO_LOAD_FOLDER and all subfolders on startup."""
        if not AUTO_LOAD_FOLDER:
            return
        folder = os.path.abspath(os.path.expanduser(AUTO_LOAD_FOLDER))
        if not os.path.isdir(folder):
            self.meta_var.set(f'Auto-load folder not found: {folder}')
            return
        paths = []
        for dirpath, _dirnames, filenames in os.walk(folder):
            for f in filenames:
                if f.lower().endswith('.docx'):
                    paths.append(os.path.join(dirpath, f))
        paths.sort()
        if not paths:
            self.meta_var.set(f'No .docx files in {folder}')
            return
        self._load_paths(paths)

    def _load_paths(self, paths):
        """Load a list of .docx paths (used by _add_files and _auto_load_folder)."""
        if len(paths) > 10:
            self.meta_var.set('Loading…')
            self.update_idletasks()

            def load_all():
                loaded, errors = [], []
                with concurrent.futures.ThreadPoolExecutor(max_workers=8) as ex:
                    future_to_path = {ex.submit(parse_docx, p): p for p in paths}
                    for future in concurrent.futures.as_completed(future_to_path):
                        path = future_to_path[future]
                        name = os.path.basename(path)
                        try:
                            sections = future.result()
                            loaded.append((name, path, sections))
                        except Exception as e:
                            errors.append((name, str(e)))
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
                for name, err in errors:
                    messagebox.showerror('Error', f'Could not load {name}:\n{err}')
                self._rebuild_sections_cache()
                self._render_pills()
                self._refresh_results()

            def run():
                self._add_files_result = load_all()
                self.after(0, on_done)

            threading.Thread(target=run, daemon=True).start()
        else:
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
                    messagebox.showerror('Error', f'Could not load {name}:\n{e}')
            self._rebuild_sections_cache()
            self._render_pills()
            self._refresh_results()

    def _add_files(self):
        paths = filedialog.askopenfilenames(filetypes=[('Word Documents', '*.docx')])
        if not paths:
            return
        self._load_paths(paths)

    def _rebuild_sections_cache(self):
        """Rebuild flat section list only when docs change (avoids rebuilding on every keystroke)."""
        self._all_sections_cache = [
            {**s, 'docName': doc['name'], 'docPath': doc['path']}
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
            self.meta_var.set(f'{len(filtered)} / {all_count} sections')
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
        TEXT = '#e2e2e2'
        DIM = '#666'
        while len(self._slot_cards) < N_VISIBLE_SLOTS:
            card = tk.Frame(self.results_inner, bg=SURF, height=ROW_HEIGHT,
                            highlightthickness=1, highlightbackground='#2a2a2a')
            card.pack_propagate(False)
            type_lbl = tk.Label(card, text='', font=('Courier', 9, 'bold'),
                               fg=TYPE_COLOR['pocket'], bg=SURF, padx=6, pady=2)
            type_lbl.pack(side='left', padx=(0, 6))
            heading_lbl = tk.Label(card, text='', font=('Courier', 10, 'bold'),
                                   fg=TEXT, bg=SURF, anchor='w')
            heading_lbl.pack(side='left', fill='x', expand=True, padx=(0, 8))
            doc_lbl = tk.Label(card, text='', font=('Courier', 9), fg=DIM, bg=SURF)
            doc_lbl.pack(side='right', padx=(0, 8))
            copy_btn = tk.Button(card, text='copy', font=('Courier', 9, 'bold'),
                                 fg='#888', bg='#222', activeforeground=TYPE_COLOR['pocket'],
                                 activebackground='#2a2a2a', bd=0, padx=8, pady=3, cursor='hand2')
            copy_btn.pack(side='right', padx=4)
            self._slot_cards.append({
                'frame': card, 'type_lbl': type_lbl, 'heading_lbl': heading_lbl,
                'doc_lbl': doc_lbl, 'copy_btn': copy_btn
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
                self._update_slot_content(slot, sec)
                slot['frame'].tkraise()
            else:
                slot['frame'].place_forget()

    def _update_slot_content(self, slot, sec):
        color = TYPE_COLOR[sec['type']]
        slot['type_lbl'].config(text=sec['type'].upper(), fg=color)
        heading = (sec['heading'] or '')[:90]
        slot['heading_lbl'].config(text=heading + ('…' if len(sec['heading'] or '') > 90 else ''))
        slot['doc_lbl'].config(text=sec['docName'])
        slot['copy_btn'].config(command=lambda s=sec, b=slot['copy_btn']: self._copy_section(s, b))

    def _render_card(self, sec, idx):
        SURF = '#161616'; TEXT = '#e2e2e2'; DIM = '#666'
        color = TYPE_COLOR[sec['type']]
        crumbs = ' › '.join(v for v in sec['parents'].values() if v)

        card = tk.Frame(self.results_inner, bg=SURF, bd=0,
                        highlightthickness=1, highlightbackground='#2a2a2a')
        card.pack(fill='x', pady=4)

        hdr_row = tk.Frame(card, bg=SURF)
        hdr_row.pack(fill='x', padx=12, pady=8)

        tk.Label(hdr_row, text=sec['type'].upper(), font=('Courier', 9, 'bold'),
                 fg=color, bg=SURF, padx=6, pady=2).pack(side='left', padx=(0,8))
        tk.Label(hdr_row, text=sec['heading'], font=('Courier', 11, 'bold'),
                 fg=TEXT, bg=SURF, anchor='w').pack(side='left')
        if crumbs:
            tk.Label(hdr_row, text=crumbs, font=('Courier', 9),
                     fg=DIM, bg=SURF).pack(side='left', padx=10)
        tk.Label(hdr_row, text=sec['docName'], font=('Courier', 9),
                 fg=DIM, bg=SURF).pack(side='right', padx=(0,8))

        copy_btn = tk.Button(hdr_row, text='copy', font=('Courier', 9, 'bold'),
                             fg='#888', bg='#222', activeforeground=color,
                             activebackground='#2a2a2a', bd=0, padx=8, pady=3, cursor='hand2')
        copy_btn.pack(side='right', padx=4)
        copy_btn.config(command=lambda s=sec, b=copy_btn: self._copy_section(s, b))

        preview = ' '.join(sec['preview'][:3])[:300]
        if preview:
            tk.Label(card, text=preview, font=('Courier', 10),
                     fg='#666', bg=SURF, anchor='w', justify='left',
                     wraplength=820, padx=12, pady=6).pack(fill='x', anchor='w')

        # Single Enter/Leave on card only (fewer bindings = less lag with many cards)
        card.bind('<Enter>', lambda e, c=card: c.config(highlightbackground='#444'))
        card.bind('<Leave>', lambda e, c=card: c.config(highlightbackground='#2a2a2a'))

    def _copy_section(self, sec, btn):
        btn.config(text='…', fg='#aaa')
        self.update_idletasks()

        def do_copy():
            err = None
            try:
                copy_via_word(sec['docPath'], sec['char_start'], sec['char_end'])
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


if __name__ == '__main__':
    app = DocSearchApp()
    app.mainloop()
