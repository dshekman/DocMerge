"""
DocMerge — Tracked Changes Merger  (python-docx / lxml edition)
Merges tracked changes + comments from a folder of colleague docs into your base document.
Reads OOXML (docx) XML directly — no Microsoft Word required.
Requires: python-docx  (pip install python-docx)
"""

import os
import threading
import tkinter as tk
from tkinter import filedialog, ttk
from pathlib import Path
import zipfile
import shutil
import copy
import tempfile

from lxml import etree

# ── Palette ──────────────────────────────────────────────────────────────────
BG         = "#1a1a1a"
SURFACE    = "#242424"
SURFACE2   = "#2e2e2e"
BORDER     = "#3a3a3a"
ACCENT     = "#c8a96e"
ACCENT_DIM = "#8a7048"
TEXT       = "#e8e4dc"
TEXT_DIM   = "#8a8680"
SUCCESS    = "#6ea87e"
ERROR      = "#c86e6e"
WARN       = "#c8a96e"
FONT_UI    = ("Segoe UI", 10)
FONT_MONO  = ("Consolas", 9)
FONT_TITLE = ("Segoe UI Light", 18)


# ── OOXML constants ───────────────────────────────────────────────────────────

W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"

REL_TYPE_COMMENTS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/"
    "relationships/comments"
)
CT_COMMENTS = (
    "application/vnd.openxmlformats-officedocument"
    ".wordprocessingml.comments+xml"
)
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def _q(local: str) -> str:
    return f"{{{W}}}{local}"


_REV_TAGS     = {_q(t) for t in ("ins", "del", "rPrChange", "pPrChange")}
_CMT_REF_TAGS = {_q(t) for t in ("commentRangeStart", "commentRangeEnd", "commentReference")}


# ── Low-level XML helpers ─────────────────────────────────────────────────────

def _read_xml(zf: zipfile.ZipFile, path: str):
    try:
        return etree.fromstring(zf.read(path))
    except KeyError:
        return None


def _all_paragraphs(root) -> list:
    return root.findall(f".//{_q('p')}")


def _has_tracked_changes(para, include_format: bool) -> bool:
    tags = _REV_TAGS if include_format else {_q("ins"), _q("del")}
    return any(para.find(f".//{t}") is not None for t in tags)


def _max_id(root, tags) -> int:
    hi = 0
    for tag in tags:
        for elem in root.iter(tag):
            raw = elem.get(_q("id"))
            if raw is not None:
                try:
                    hi = max(hi, int(raw))
                except ValueError:
                    pass
    return hi


def _offset_ids(root, tags, offset: int):
    for tag in tags:
        for elem in root.iter(tag):
            raw = elem.get(_q("id"))
            if raw is not None:
                try:
                    elem.set(_q("id"), str(int(raw) + offset))
                except ValueError:
                    pass


def _set_author(root, author: str):
    for tag in (_q("ins"), _q("del"), _q("rPrChange"), _q("pPrChange")):
        for elem in root.iter(tag):
            if elem.get(_q("author")) is not None:
                elem.set(_q("author"), author)


def _serialize(root) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _extract_authors(zf: zipfile.ZipFile) -> set:
    authors = set()
    doc = _read_xml(zf, "word/document.xml")
    if doc is None:
        return authors
    for tag in (_q("ins"), _q("del"), _q("rPrChange"), _q("pPrChange")):
        for elem in doc.iter(tag):
            a = elem.get(_q("author"))
            if a:
                authors.add(a)
    return authors


# ── Paragraph-level conflict merge ────────────────────────────────────────────

def _run_text(run_elem) -> str:
    return "".join(t.text or "" for t in run_elem.iter(_q("t")))


def _del_text(del_elem) -> str:
    return "".join(t.text or "" for t in del_elem.iter(_q("delText")))


def _merge_b_into_a(para_a, para_b):
    a_anchors = []
    for child in list(para_a):
        if child.tag == _q("r"):
            a_anchors.append((child, _run_text(child)))
        elif child.tag == _q("del"):
            a_anchors.append((child, _del_text(child)))

    last_a_anchor = None
    a_search_from = 0
    insertions = []

    for b_child in list(para_b):
        if b_child.tag == _q("pPr"):
            continue

        if b_child.tag == _q("r"):
            bt = _run_text(b_child)
            for i in range(a_search_from, len(a_anchors)):
                if a_anchors[i][1] == bt:
                    last_a_anchor = a_anchors[i][0]
                    a_search_from = i + 1
                    break

        elif b_child.tag == _q("del"):
            dt = _del_text(b_child)
            matched_idx = None
            for i in range(a_search_from, len(a_anchors)):
                if a_anchors[i][1] == dt:
                    matched_idx = i
                    last_a_anchor = a_anchors[i][0]
                    a_search_from = i + 1
                    break
            if matched_idx is not None and a_anchors[matched_idx][0].tag == _q("r"):
                insertions.append((last_a_anchor, copy.deepcopy(b_child)))

        elif b_child.tag == _q("ins"):
            insertions.append((last_a_anchor, copy.deepcopy(b_child)))

    for after_elem, new_elem in insertions:
        children = list(para_a)
        if after_elem is not None and after_elem in children:
            pos = children.index(after_elem) + 1
        else:
            pos = 0
            for i, ch in enumerate(children):
                if ch.tag == _q("pPr"):
                    pos = i + 1
                    break
        para_a.insert(pos, new_elem)


# ── Core merge engine ─────────────────────────────────────────────────────────

def merge_documents(base_path: str, colleague_paths: list, output_path: str,
                    detect_format: bool, author_map: dict, log_fn):
    """
    author_map: {filepath -> author_name_string}
                Empty string means keep the original author from the file.
    """
    base_path       = str(Path(base_path).resolve())
    output_path     = str(Path(output_path).resolve())
    colleague_paths = [str(Path(p).resolve()) for p in colleague_paths]

    log_fn("Copying base document...", "info")
    shutil.copy2(base_path, output_path)

    with zipfile.ZipFile(output_path, "r") as zf:
        doc_xml      = etree.fromstring(zf.read("word/document.xml"))
        comments_xml = _read_xml(zf, "word/comments.xml")
        zip_names    = set(zf.namelist())

    body      = doc_xml.find(f".//{_q('body')}")
    out_paras = _all_paragraphs(body)

    base_had_comments = comments_xml is not None
    max_rev_id = _max_id(doc_xml, _REV_TAGS)
    max_cmt_id = _max_id(comments_xml, {_q("comment")}) if comments_xml else 0

    touched: dict = {}

    for ci, cpath in enumerate(colleague_paths, 1):
        cname  = Path(cpath).name
        author = author_map.get(cpath, "").strip()
        label  = f"{cname}" + (f"  [{author}]" if author else "")
        log_fn(f"Merging [{ci}/{len(colleague_paths)}]: {label}", "info")
        try:
            with zipfile.ZipFile(cpath, "r") as czf:
                c_doc      = etree.fromstring(czf.read("word/document.xml"))
                c_comments = _read_xml(czf, "word/comments.xml")

            c_body  = c_doc.find(f".//{_q('body')}")
            c_paras = _all_paragraphs(c_body)

            _offset_ids(c_body, _REV_TAGS, max_rev_id + 1)

            if author:
                _set_author(c_body, author)

            cmt_id_map: dict = {}
            if c_comments is not None:
                cmt_offset = max_cmt_id + 1
                for cmt in c_comments.findall(_q("comment")):
                    old_id = cmt.get(_q("id"), "")
                    new_id = str(int(old_id) + cmt_offset)
                    cmt_id_map[old_id] = new_id
                    cmt.set(_q("id"), new_id)
                    if author:
                        cmt.set(_q("author"), author)

                for tag in _CMT_REF_TAGS:
                    for elem in c_body.iter(tag):
                        old_id = elem.get(_q("id"), "")
                        if old_id in cmt_id_map:
                            elem.set(_q("id"), cmt_id_map[old_id])

                if comments_xml is None:
                    comments_xml = etree.Element(_q("comments"), nsmap={"w": W})
                for cmt in c_comments.findall(_q("comment")):
                    comments_xml.append(copy.deepcopy(cmt))
                if cmt_id_map:
                    max_cmt_id = max(int(v) for v in cmt_id_map.values())

            changed = spliced = 0
            for i, c_para in enumerate(c_paras):
                if not _has_tracked_changes(c_para, detect_format):
                    continue
                if i >= len(out_paras):
                    log_fn(f"  Para {i} beyond output document -- skipped", "warn")
                    continue

                if i not in touched:
                    out_para = out_paras[i]
                    parent   = out_para.getparent()
                    pos      = list(parent).index(out_para)
                    new_para = copy.deepcopy(c_para)
                    parent.remove(out_para)
                    parent.insert(pos, new_para)
                    out_paras[i] = new_para
                    touched[i]   = cname
                    changed += 1
                else:
                    log_fn(
                        f"  <-> Para {i} also edited by {touched[i]} -- splicing both authors",
                        "info",
                    )
                    _merge_b_into_a(out_paras[i], c_para)
                    spliced += 1

            summary = f"  + {changed} paragraph(s) replaced"
            if spliced:
                summary += f", {spliced} paragraph(s) spliced (multi-author)"
            log_fn(summary, "success")

            max_rev_id = _max_id(body, _REV_TAGS)

        except Exception as exc:
            import traceback
            log_fn(f"  ERROR {cname}: {exc}", "error")
            log_fn(traceback.format_exc(), "dim")

    log_fn("Writing output...", "info")
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".docx")
    os.close(tmp_fd)

    inject_comments = (not base_had_comments) and (comments_xml is not None)

    try:
        with zipfile.ZipFile(output_path, "r") as zf_in:
            with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zf_out:
                for item in zf_in.namelist():
                    if item == "word/document.xml":
                        zf_out.writestr(item, _serialize(doc_xml))
                    elif item == "word/comments.xml":
                        data = _serialize(comments_xml) if comments_xml else zf_in.read(item)
                        zf_out.writestr(item, data)
                    elif item == "word/_rels/document.xml.rels" and inject_comments:
                        rels = etree.fromstring(zf_in.read(item))
                        existing_ids = {r.get("Id", "") for r in rels}
                        new_rid = "rIdComments"
                        while new_rid in existing_ids:
                            new_rid += "x"
                        etree.SubElement(
                            rels, f"{{{PKG}}}Relationship",
                            {"Id": new_rid, "Type": REL_TYPE_COMMENTS, "Target": "comments.xml"},
                        )
                        zf_out.writestr(item, _serialize(rels))
                    elif item == "[Content_Types].xml" and inject_comments:
                        ct = etree.fromstring(zf_in.read(item))
                        existing = {e.get("PartName") for e in ct}
                        if "/word/comments.xml" not in existing:
                            etree.SubElement(
                                ct, f"{{{CT_NS}}}Override",
                                {"PartName": "/word/comments.xml", "ContentType": CT_COMMENTS},
                            )
                        zf_out.writestr(item, _serialize(ct))
                    else:
                        zf_out.writestr(item, zf_in.read(item))

                if inject_comments and "word/comments.xml" not in zip_names:
                    zf_out.writestr("word/comments.xml", _serialize(comments_xml))

        shutil.move(tmp_path, output_path)
        log_fn(f"Done -> {Path(output_path).name}", "success")
        return True

    except Exception as exc:
        import traceback
        log_fn(f"Fatal error writing output: {exc}", "error")
        log_fn(traceback.format_exc(), "dim")
        try:
            os.unlink(tmp_path)
        except OSError:
            pass
        return False


# ── UI helpers ────────────────────────────────────────────────────────────────

def styled_button(parent, text, command, accent=False):
    bg  = ACCENT      if accent else SURFACE2
    fg  = "#1a1a1a"   if accent else TEXT
    abg = ACCENT_DIM  if accent else BORDER
    return tk.Button(
        parent, text=text, command=command,
        bg=bg, fg=fg, activebackground=abg, activeforeground=fg,
        relief="flat", bd=0, padx=14, pady=6,
        font=FONT_UI, cursor="hand2",
    )


def dark_entry(parent, textvariable=None, width=None):
    kw = {}
    if textvariable:
        kw["textvariable"] = textvariable
    if width:
        kw["width"] = width
    return tk.Entry(
        parent, bg=SURFACE2, fg=TEXT, insertbackground=ACCENT,
        relief="flat", bd=0, font=FONT_UI,
        highlightthickness=1, highlightbackground=BORDER,
        highlightcolor=ACCENT, **kw,
    )


# ── Per-file author table ─────────────────────────────────────────────────────

class FileAuthorTable(tk.Frame):
    """Scrollable table: one row per colleague file with an editable Author field."""

    ROW_H   = 30
    VISIBLE = 5

    def __init__(self, parent, **kw):
        super().__init__(parent, bg=SURFACE,
                         highlightthickness=1, highlightbackground=BORDER, **kw)
        self._rows: list[dict] = []
        self._build()

    def _build(self):
        hdr = tk.Frame(self, bg=SURFACE2)
        hdr.pack(fill="x")
        tk.Label(hdr, text="  FILE", bg=SURFACE2, fg=TEXT_DIM,
                 font=("Segoe UI", 8), anchor="w", width=36
                 ).pack(side="left", padx=(4, 0), pady=5)
        tk.Label(hdr, text="AUTHOR NAME", bg=SURFACE2, fg=TEXT_DIM,
                 font=("Segoe UI", 8), anchor="w"
                 ).pack(side="left", padx=(0, 4), pady=5)

        body_frame = tk.Frame(self, bg=SURFACE)
        body_frame.pack(fill="both", expand=True)

        self._canvas = tk.Canvas(body_frame, bg=SURFACE, bd=0,
                                  highlightthickness=0,
                                  height=self.ROW_H * self.VISIBLE)
        self._sb = tk.Scrollbar(body_frame, orient="vertical",
                                 command=self._canvas.yview,
                                 bg=SURFACE2, troughcolor=SURFACE, width=10)
        self._canvas.configure(yscrollcommand=self._sb.set)
        self._canvas.pack(side="left", fill="both", expand=True)
        self._sb.pack(side="right", fill="y")

        self._inner = tk.Frame(self._canvas, bg=SURFACE)
        self._win_id = self._canvas.create_window((0, 0), window=self._inner, anchor="nw")
        self._inner.bind("<Configure>", lambda e: self._canvas.configure(
            scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<Configure>", lambda e: self._canvas.itemconfig(
            self._win_id, width=e.width))

    def clear(self):
        for w in self._inner.winfo_children():
            w.destroy()
        self._rows.clear()

    def add_file(self, path: str, detected_author: str = ""):
        var = tk.StringVar(value=detected_author)
        row = tk.Frame(self._inner, bg=SURFACE)
        row.pack(fill="x")

        tk.Label(row, text=f"  {Path(path).name}", bg=SURFACE, fg=TEXT,
                 font=FONT_MONO, anchor="w", width=34
                 ).pack(side="left", padx=(0, 6), ipady=5)

        dark_entry(row, textvariable=var).pack(
            side="left", fill="x", expand=True, padx=(0, 8), ipady=4)

        tk.Frame(self._inner, bg=BORDER, height=1).pack(fill="x")
        self._rows.append({"path": path, "var": var})

    def set_author(self, path: str, author: str):
        for r in self._rows:
            if r["path"] == path:
                r["var"].set(author)
                break

    def get_author_map(self) -> dict:
        return {r["path"]: r["var"].get().strip() for r in self._rows}

    def paths(self) -> list:
        return [r["path"] for r in self._rows]


# ── Main application ──────────────────────────────────────────────────────────

class DocMergeApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DocMerge")
        self.configure(bg=BG)
        self.resizable(True, True)
        self.minsize(680, 680)

        self.base_var   = tk.StringVar()
        self.folder_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.format_var = tk.BooleanVar(value=False)
        self.running    = False

        self._build_ui()
        self._center()

    def _center(self):
        self.update_idletasks()
        w, h = 740, 740
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build_ui(self):
        header = tk.Frame(self, bg=SURFACE, pady=16)
        header.pack(fill="x")
        tk.Label(header, text="DocMerge", font=FONT_TITLE,
                 bg=SURFACE, fg=ACCENT).pack(side="left", padx=24)
        tk.Label(header, text="Tracked Changes Merger",
                 font=("Segoe UI Light", 10), bg=SURFACE, fg=TEXT_DIM
                 ).pack(side="left", pady=(6, 0))
        tk.Frame(self, bg=ACCENT, height=2).pack(fill="x")

        content = tk.Frame(self, bg=BG, padx=28, pady=24)
        content.pack(fill="both", expand=True)
        content.columnconfigure(0, weight=1)

        row = 0

        # Base document
        tk.Label(content, text="BASE DOCUMENT  (your draft)",
                 bg=BG, fg=TEXT_DIM, font=("Segoe UI", 9), anchor="w"
                 ).grid(row=row, column=0, sticky="w", pady=(0, 3))
        row += 1
        base_row = tk.Frame(content, bg=BG)
        base_row.grid(row=row, column=0, sticky="ew", pady=(0, 18))
        base_row.columnconfigure(0, weight=1)
        dark_entry(base_row, textvariable=self.base_var).grid(
            row=0, column=0, sticky="ew", ipady=7, padx=(0, 8))
        styled_button(base_row, "Browse...", self._pick_base).grid(row=0, column=1)
        row += 1

        # Colleague folder
        tk.Label(content, text="COLLEAGUE DOCS FOLDER",
                 bg=BG, fg=TEXT_DIM, font=("Segoe UI", 9), anchor="w"
                 ).grid(row=row, column=0, sticky="w", pady=(0, 3))
        row += 1
        folder_row = tk.Frame(content, bg=BG)
        folder_row.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        folder_row.columnconfigure(0, weight=1)
        dark_entry(folder_row, textvariable=self.folder_var).grid(
            row=0, column=0, sticky="ew", ipady=7, padx=(0, 8))
        styled_button(folder_row, "Browse...", self._pick_folder).grid(row=0, column=1)
        row += 1

        # File + author table
        tk.Label(
            content,
            text="FILES & AUTHORS  (pre-filled from file; edit to override the author attribution)",
            bg=BG, fg=TEXT_DIM, font=("Segoe UI", 9), anchor="w",
        ).grid(row=row, column=0, sticky="w", pady=(0, 3))
        row += 1
        self.file_table = FileAuthorTable(content)
        self.file_table.grid(row=row, column=0, sticky="ew", pady=(0, 4))
        row += 1
        self.file_count_label = tk.Label(
            content, text="No folder selected",
            bg=BG, fg=TEXT_DIM, font=("Segoe UI", 9), anchor="w",
        )
        self.file_count_label.grid(row=row, column=0, sticky="w", pady=(0, 18))
        row += 1

        # Output
        tk.Label(content, text="OUTPUT FILE",
                 bg=BG, fg=TEXT_DIM, font=("Segoe UI", 9), anchor="w"
                 ).grid(row=row, column=0, sticky="w", pady=(0, 3))
        row += 1
        out_row = tk.Frame(content, bg=BG)
        out_row.grid(row=row, column=0, sticky="ew", pady=(0, 16))
        out_row.columnconfigure(0, weight=1)
        dark_entry(out_row, textvariable=self.output_var).grid(
            row=0, column=0, sticky="ew", ipady=7, padx=(0, 8))
        styled_button(out_row, "Browse...", self._pick_output).grid(row=0, column=1)
        row += 1

        # Options
        opts = tk.Frame(content, bg=BG)
        opts.grid(row=row, column=0, sticky="w", pady=(0, 20))
        tk.Checkbutton(
            opts,
            text="Include formatting changes  (off = content changes only, cleaner for legal docs)",
            variable=self.format_var, bg=BG, fg=TEXT_DIM,
            activebackground=BG, activeforeground=TEXT,
            selectcolor=SURFACE2, font=("Segoe UI", 9), highlightthickness=0,
        ).pack(side="left")
        row += 1

        self.run_btn = styled_button(content, "  Run Merge  ", self._run, accent=True)
        self.run_btn.grid(row=row, column=0, sticky="w")

        # Log
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")
        log_header = tk.Frame(self, bg=SURFACE, pady=6)
        log_header.pack(fill="x")
        tk.Label(log_header, text="LOG", bg=SURFACE, fg=TEXT_DIM,
                 font=("Segoe UI Semibold", 8), padx=28).pack(side="left")
        tk.Button(log_header, text="Clear", bg=SURFACE, fg=TEXT_DIM,
                  relief="flat", bd=0, font=("Segoe UI", 8), cursor="hand2",
                  command=self._clear_log, padx=8).pack(side="right", padx=12)

        log_frame = tk.Frame(self, bg="#111111")
        log_frame.pack(fill="both", expand=False)
        self.log_box = tk.Text(
            log_frame, bg="#111111", fg=TEXT_DIM, font=FONT_MONO,
            relief="flat", bd=0, height=7, state="disabled",
            highlightthickness=0, padx=12, pady=8, wrap="word",
        )
        log_sb = tk.Scrollbar(log_frame, orient="vertical", command=self.log_box.yview,
                               bg="#111111", troughcolor="#111111", width=10)
        self.log_box.configure(yscrollcommand=log_sb.set)
        self.log_box.pack(side="left", fill="both", expand=True)
        log_sb.pack(side="right", fill="y")

        self.log_box.tag_configure("info",    foreground=TEXT)
        self.log_box.tag_configure("dim",     foreground=TEXT_DIM)
        self.log_box.tag_configure("success", foreground=SUCCESS)
        self.log_box.tag_configure("error",   foreground=ERROR)
        self.log_box.tag_configure("warn",    foreground=WARN)

    def _pick_base(self):
        p = filedialog.askopenfilename(
            title="Select your base document",
            filetypes=[("Word Documents", "*.docx"), ("All files", "*.*")],
        )
        if p:
            self.base_var.set(p)
            if not self.output_var.get():
                base = Path(p)
                self.output_var.set(str(base.parent / (base.stem + "_merged.docx")))

    def _pick_folder(self):
        d = filedialog.askdirectory(title="Select folder of colleague documents")
        if d:
            self.folder_var.set(d)
            self._refresh_file_list(d)

    def _pick_output(self):
        p = filedialog.asksaveasfilename(
            title="Save merged document as...",
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx")],
        )
        if p:
            self.output_var.set(p)

    def _refresh_file_list(self, folder: str):
        self.file_table.clear()
        files   = sorted(Path(folder).glob("*.docx"))
        base    = Path(self.base_var.get()).resolve() if self.base_var.get() else None
        col_files = []
        for f in files:
            if base and f.resolve() == base:
                continue
            self.file_table.add_file(str(f))
            col_files.append(str(f))

        n = len(col_files)
        self.file_count_label.config(
            text=f"{n} colleague document{'s' if n != 1 else ''} found  --  author names auto-detected, edit to override",
            fg=ACCENT if n > 0 else ERROR,
        )
        threading.Thread(
            target=self._scan_authors, args=(list(col_files),), daemon=True
        ).start()

    def _scan_authors(self, paths: list):
        for p in paths:
            try:
                with zipfile.ZipFile(p, "r") as zf:
                    authors = _extract_authors(zf)
                if len(authors) == 1:
                    author = next(iter(authors))
                elif len(authors) > 1:
                    author = ", ".join(sorted(authors))
                else:
                    author = ""
                self.after(0, self.file_table.set_author, p, author)
            except Exception:
                pass

    def _log(self, msg: str, tag: str = "info"):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", msg + "\n", tag)
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _clear_log(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    def _run(self):
        if self.running:
            return

        base   = self.base_var.get().strip()
        folder = self.folder_var.get().strip()
        output = self.output_var.get().strip()

        colleague_paths = self.file_table.paths()
        author_map      = self.file_table.get_author_map()

        errors = []
        if not base or not Path(base).exists():
            errors.append("Base document path is invalid or missing.")
        if not colleague_paths:
            if folder:
                self._refresh_file_list(folder)
                colleague_paths = self.file_table.paths()
                author_map      = self.file_table.get_author_map()
            if not colleague_paths:
                errors.append("No colleague .docx files found in the selected folder.")
        if not output:
            errors.append("Please specify an output file path.")
        if errors:
            for e in errors:
                self._log(f"  {e}", "warn")
            return

        colleague_paths = [
            p for p in colleague_paths
            if Path(p).resolve() != Path(base).resolve()
        ]

        self.running = True
        self.run_btn.config(state="disabled", text="  Running...  ")
        self._log("─" * 52, "dim")
        self._log(f"Starting merge: {len(colleague_paths)} file(s) -> {Path(output).name}", "info")

        def worker():
            success = merge_documents(
                base_path=base,
                colleague_paths=colleague_paths,
                output_path=output,
                detect_format=self.format_var.get(),
                author_map=author_map,
                log_fn=lambda msg, tag="info": self.after(0, self._log, msg, tag),
            )
            def done():
                self.running = False
                self.run_btn.config(state="normal", text="  Run Merge  ")
                if success:
                    self._log("─" * 52, "dim")
                    self._log("Merge complete. File saved:", "success")
                    self._log(f"   {output}", "dim")
            self.after(0, done)

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    app = DocMergeApp()
    app.mainloop()
