# IVI_pdf_packer.py

#region IVI_pdf_packer.py PDF folder packer with TOC + header/footer + filename prefix + file list + help
# Version 0.1.3 (2026/02/13)
# Added GUI fields for output filename prefix (default set) and updated default header text.

import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from pypdf import PdfReader, PdfWriter
except Exception as e:
    raise SystemExit(
        "Missing dependency: pypdf\n"
        "Install with: pip install pypdf\n\n"
        f"Original error: {e}"
    )

try:
    from reportlab.pdfgen import canvas
except Exception as e:
    raise SystemExit(
        "Missing dependency: reportlab\n"
        "Install with: pip install reportlab\n\n"
        f"Original error: {e}"
    )

_HAS_DND = False
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore
    _HAS_DND = True
except Exception:
    _HAS_DND = False


DEFAULT_PREFIX = "566-02-48305_Grievor_Book"
DEFAULT_HEADER = "566-02-48305"


def _is_pdf(path: str) -> bool:
    return path.lower().endswith(".pdf")


def _list_pdfs_in_folder(folder: str) -> list[str]:
    try:
        names = os.listdir(folder)
    except Exception:
        return []
    files = [os.path.join(folder, n) for n in names if _is_pdf(n)]
    files = [f for f in files if os.path.isfile(f)]
    files.sort(key=lambda p: os.path.basename(p).lower())
    return files


def _safe_basename(path: str) -> str:
    return os.path.basename(path)


def _count_pages(pdf_path: str) -> int:
    r = PdfReader(pdf_path)
    return len(r.pages)


def _build_toc_pdf_bytes(
    entries: list[tuple[str, int]],
    title: str = "Table of Contents",
    pagesize=(612.0, 792.0),  # letter in points
    left_margin=54,
    top_margin=54,
    bottom_margin=54,
    font_name="Helvetica",
    font_size=11,
    line_gap=3,
) -> tuple[bytes, int]:
    import io

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=pagesize)
    w, h = pagesize

    usable_h = h - top_margin - bottom_margin
    line_h = font_size + line_gap

    title_font = 16
    title_gap = 18
    header_space = title_gap + 2 * line_h
    lines_per_page = max(1, int((usable_h - header_space) // line_h))

    toc_pages = (len(entries) + (lines_per_page - 1)) // lines_per_page
    toc_pages = max(1, toc_pages)

    def draw_page(page_idx: int):
        c.setFont(font_name, title_font)
        c.drawString(left_margin, h - top_margin, title)

        y = h - top_margin - title_gap
        c.setFont(font_name, 9)
        c.drawString(left_margin, y, "Merged PDF contents (file start pages):")
        y -= (line_h + 6)

        c.setFont(font_name, font_size)

        start = page_idx * lines_per_page
        end = min(len(entries), start + lines_per_page)

        for name, pg in entries[start:end]:
            max_chars = 95
            disp = name if len(name) <= max_chars else (name[: max_chars - 3] + "...")
            c.drawString(left_margin, y, disp)

            dots_start_x = left_margin + 320
            dots_end_x = w - left_margin - 40
            if dots_end_x > dots_start_x:
                c.drawString(dots_start_x, y, "." * 60)

            c.drawRightString(w - left_margin, y, str(pg))
            y -= line_h

        c.showPage()

    for i in range(toc_pages):
        draw_page(i)

    c.save()
    buf.seek(0)
    return buf.read(), toc_pages


def _overlay_header_footer_pdf_bytes(
    page_w: float,
    page_h: float,
    header_text: str,
    page_num_1based: int,
    total_pages: int,
    left_margin=36,
    top_margin=24,
    bottom_margin=24,
    font_name="Helvetica",
    header_font_size=10,
    footer_font_size=9,
) -> bytes:
    import io

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(page_w, page_h))

    if header_text.strip():
        c.setFont(font_name, header_font_size)
        c.drawString(left_margin, page_h - top_margin, header_text.strip())

    c.setFont(font_name, footer_font_size)
    c.drawRightString(page_w - left_margin, bottom_margin, f"Page {page_num_1based} of {total_pages}")

    c.save()
    buf.seek(0)
    return buf.read()


def _stamp_header_footer_on_writer(writer: PdfWriter, header_text: str):
    import io

    total = len(writer.pages)
    for idx, page in enumerate(writer.pages):
        w = float(page.mediabox.width)
        h = float(page.mediabox.height)
        overlay_bytes = _overlay_header_footer_pdf_bytes(
            page_w=w,
            page_h=h,
            header_text=header_text,
            page_num_1based=idx + 1,
            total_pages=total,
        )
        overlay_reader = PdfReader(io.BytesIO(overlay_bytes))
        overlay_page = overlay_reader.pages[0]
        try:
            page.merge_page(overlay_page)
        except Exception:
            pass


def merge_pdfs_with_toc(
    folder: str,
    output_path: str,
    header_text: str,
    progress_cb=None,
    log_cb=None,
) -> dict:
    files = _list_pdfs_in_folder(folder)
    if not files:
        raise ValueError("No PDF files found in the selected folder.")

    page_counts = []
    for i, f in enumerate(files, start=1):
        if progress_cb:
            progress_cb(i, len(files), f"Counting pages: {_safe_basename(f)}")
        try:
            n = _count_pages(f)
        except Exception as e:
            raise RuntimeError(f"Failed to read PDF: {f}\n{e}")
        page_counts.append(n)

    placeholder_entries = [(_safe_basename(f), 1) for f in files]
    _, toc_pages = _build_toc_pdf_bytes(placeholder_entries)

    starts = []
    current = toc_pages + 1
    for n in page_counts:
        starts.append(current)
        current += n

    toc_entries = [(_safe_basename(f), s) for f, s in zip(files, starts)]
    toc_bytes, toc_pages_check = _build_toc_pdf_bytes(toc_entries)
    if toc_pages_check != toc_pages:
        toc_pages = toc_pages_check
        starts = []
        current = toc_pages + 1
        for n in page_counts:
            starts.append(current)
            current += n
        toc_entries = [(_safe_basename(f), s) for f, s in zip(files, starts)]
        toc_bytes, toc_pages = _build_toc_pdf_bytes(toc_entries)

    if log_cb:
        log_cb(f"TOC pages: {toc_pages}")
        for (name, s) in toc_entries:
            log_cb(f"  {name} -> page {s}")

    writer = PdfWriter()

    import io
    toc_reader = PdfReader(io.BytesIO(toc_bytes))
    for p in toc_reader.pages:
        writer.add_page(p)

    for i, f in enumerate(files, start=1):
        if progress_cb:
            progress_cb(i, len(files), f"Merging: {_safe_basename(f)}")
        r = PdfReader(f)
        file_start_index = starts[i - 1] - 1
        for p in r.pages:
            writer.add_page(p)
        try:
            writer.add_outline_item(title=_safe_basename(f), page_number=file_start_index)
        except Exception:
            pass

    if progress_cb:
        progress_cb(1, 1, "Stamping header/footer...")
    _stamp_header_footer_on_writer(writer, header_text=header_text)

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with open(output_path, "wb") as out:
        writer.write(out)

    return {
        "folder": folder,
        "output_path": output_path,
        "header_text": header_text,
        "files": files,
        "page_counts": page_counts,
        "toc_pages": toc_pages,
        "start_pages": starts,
        "total_pages": len(writer.pages),
    }


class AppBase:
    APP_NAME = "IVI_pdf_packer"

    HELP_TEXT = (
        "IVI_pdf_packer\n\n"
        "Steps:\n"
        "  1) Drop a folder (or Browse) containing PDFs.\n"
        "  2) (Optional) Change the filename prefix and/or header text.\n"
        "  3) Click 'Generate merged PDF'.\n\n"
        "Output:\n"
        "  • A single merged PDF with TOC, header, and 'Page X of Y' footer.\n"
        "  • A copy/paste file list in the output window.\n\n"
        "Notes:\n"
        "  • PDFs are taken from the selected folder (non-recursive) and sorted by filename.\n"
        "  • Default output location is the PARENT folder of the selected folder.\n\n"
        "Dependencies:\n"
        "  pip install pypdf reportlab\n"
        "  pip install tkinterdnd2   (optional drag & drop)\n\n"
        "For more IVI tools visit: IVIM.ca\n"
    )

    def __init__(self, root):
        self.root = root
        self.root.title(self.APP_NAME)
        self.root.geometry("900x760")

        self.varFolder = tk.StringVar(value="")
        self.varPrefix = tk.StringVar(value=DEFAULT_PREFIX)
        self.varOut = tk.StringVar(value="")
        self.varHeader = tk.StringVar(value=DEFAULT_HEADER)
        self.varStatus = tk.StringVar(value="Drop a folder, or click Browse.")

        self._build_ui()
        self._set_default_output()

    def _build_ui(self):
        frmTop = ttk.Frame(self.root, padding=12)
        frmTop.pack(fill="x")

        ttk.Label(frmTop, text="Folder:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frmTop, textvariable=self.varFolder).grid(row=0, column=1, sticky="ew", padx=(8, 8))
        ttk.Button(frmTop, text="Browse...", command=self.on_browse_folder).grid(row=0, column=2, sticky="e")
        ttk.Button(frmTop, text="Help", command=self.on_help).grid(row=0, column=3, sticky="e", padx=(8, 0))

        ttk.Label(frmTop, text="Filename prefix:").grid(row=1, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(frmTop, textvariable=self.varPrefix).grid(row=1, column=1, sticky="ew", padx=(8, 8), pady=(10, 0))
        ttk.Label(frmTop, text="(used for default output name)").grid(row=1, column=2, sticky="w", pady=(10, 0))

        ttk.Label(frmTop, text="Output PDF:").grid(row=2, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(frmTop, textvariable=self.varOut).grid(row=2, column=1, sticky="ew", padx=(8, 8), pady=(10, 0))
        ttk.Button(frmTop, text="Save as...", command=self.on_choose_output).grid(row=2, column=2, sticky="e", pady=(10, 0))

        ttk.Label(frmTop, text="Header text:").grid(row=3, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(frmTop, textvariable=self.varHeader).grid(row=3, column=1, sticky="ew", padx=(8, 8), pady=(10, 0))
        ttk.Label(frmTop, text="(stamped on each page)").grid(row=3, column=2, sticky="w", pady=(10, 0))

        frmTop.columnconfigure(1, weight=1)

        frmMid = ttk.Frame(self.root, padding=(12, 0, 12, 12))
        frmMid.pack(fill="x")

        self.btnRun = ttk.Button(frmMid, text="Generate merged PDF", command=self.on_run)
        self.btnRun.pack(side="left")

        self.prog = ttk.Progressbar(frmMid, length=280, mode="determinate")
        self.prog.pack(side="left", padx=12)

        ttk.Label(frmMid, textvariable=self.varStatus).pack(side="left", fill="x", expand=True)

        frmDrop = ttk.Frame(self.root, padding=(12, 0, 12, 12))
        frmDrop.pack(fill="x")
        self.lblDrop = ttk.Label(
            frmDrop,
            text=("Drag & drop a folder here" if _HAS_DND else "Drag & drop not available (install tkinterdnd2). Use Browse."),
            anchor="center",
            relief="groove",
            padding=12,
        )
        self.lblDrop.pack(fill="x")

        frmBottom = ttk.Frame(self.root, padding=(12, 0, 12, 12))
        frmBottom.pack(fill="both", expand=True)

        ttk.Label(frmBottom, text="Output / file list (copy from here):").pack(anchor="w")
        self.txt = tk.Text(frmBottom, wrap="none", height=20)
        self.txt.pack(fill="both", expand=True, pady=(6, 6))

        frmBtns = ttk.Frame(frmBottom)
        frmBtns.pack(fill="x")

        ttk.Button(frmBtns, text="Copy all", command=self.on_copy_all).pack(side="left")
        ttk.Button(frmBtns, text="Clear", command=self.on_clear).pack(side="left", padx=8)

        self._log(f"{self.APP_NAME} ready.")
        self._log(f"Drag & drop enabled: {_HAS_DND}")
        self._log("Default output is in the PARENT folder of the selected folder.")

        self.varPrefix.trace_add("write", lambda *_: self._set_default_output_if_not_custom())

    def on_help(self):
        messagebox.showinfo("Help / ReadMe", self.HELP_TEXT)

    def _default_output_for_folder_prefix(self, folder: str, prefix: str) -> str:
        folder = os.path.abspath(folder)
        parent = os.path.dirname(folder.rstrip(os.sep)) or folder
        safe_prefix = (prefix.strip() or "packed").replace(os.sep, "_")
        return os.path.join(parent, f"{safe_prefix}.pdf")

    def _set_default_output(self):
        folder = self.varFolder.get().strip()
        prefix = self.varPrefix.get().strip()
        if folder and os.path.isdir(folder):
            out = self._default_output_for_folder_prefix(folder, prefix)
        else:
            out = os.path.abspath(f"{(prefix or 'packed')}.pdf")
        self.varOut.set(out)
        self._outIsDefault = True

    def _set_default_output_if_not_custom(self):
        if getattr(self, "_outIsDefault", True):
            self._set_default_output()

    def _mark_output_custom(self):
        self._outIsDefault = False

    def _log(self, s: str):
        self.txt.insert("end", s + "\n")
        self.txt.see("end")

    def _clear_log(self):
        self.txt.delete("1.0", "end")

    def on_clear(self):
        self._clear_log()

    def on_copy_all(self):
        data = self.txt.get("1.0", "end-1c")
        self.root.clipboard_clear()
        self.root.clipboard_append(data)
        self.varStatus.set("Copied to clipboard.")

    def on_browse_folder(self):
        folder = filedialog.askdirectory(title="Select folder containing PDFs")
        if folder:
            self.varFolder.set(folder)
            self.varOut.set(self._default_output_for_folder_prefix(folder, self.varPrefix.get()))
            self._outIsDefault = True
            self._preview_files(folder)

    def on_choose_output(self):
        self._mark_output_custom()
        path = filedialog.asksaveasfilename(
            title="Save merged PDF as",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile=os.path.basename(self.varOut.get() or f"{self.varPrefix.get().strip() or 'packed'}.pdf"),
        )
        if path:
            self.varOut.set(path)

    def _preview_files(self, folder: str):
        files = _list_pdfs_in_folder(folder)
        self._log(f"Folder: {folder}")
        self._log(f"PDF files found: {len(files)}")
        for f in files:
            self._log(f"  {os.path.basename(f)}")

    def _set_running(self, running: bool):
        self.btnRun.configure(state=("disabled" if running else "normal"))

    def _progress(self, i, n, msg):
        self.prog.configure(maximum=max(1, n), value=min(i, n))
        self.varStatus.set(msg)

    def on_run(self):
        folder = self.varFolder.get().strip()
        out = self.varOut.get().strip()
        header = self.varHeader.get().strip()

        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Missing folder", "Please choose a valid folder containing PDFs.")
            return
        if not out or not out.lower().endswith(".pdf"):
            messagebox.showerror("Invalid output", "Please choose an output .pdf file path.")
            return

        self._set_running(True)
        self.prog.configure(value=0)
        self.varStatus.set("Starting...")

        def worker():
            try:
                self._log("")
                self._log("Running pack...")
                res = merge_pdfs_with_toc(
                    folder=folder,
                    output_path=out,
                    header_text=header,
                    progress_cb=lambda i, n, msg: self.root.after(0, self._progress, i, n, msg),
                    log_cb=lambda s: self.root.after(0, self._log, s),
                )

                def done_ok():
                    self._log("")
                    self._log("Done.")
                    self._log(f"Output: {res['output_path']}")
                    self._log(f"Header: {res['header_text'] or '(none)'}")
                    self._log(f"TOC pages: {res['toc_pages']}")
                    self._log(f"Total pages: {res['total_pages']}")
                    self._log("")
                    self._log("File list (copy/paste):")
                    for f in res["files"]:
                        self._log(f)
                    self.varStatus.set("Completed.")
                    self._set_running(False)

                self.root.after(0, done_ok)

            except Exception as e:
                def done_err():
                    self._set_running(False)
                    self.varStatus.set("Error.")
                    messagebox.showerror("Error", str(e))
                    self._log("")
                    self._log("Error:")
                    self._log(str(e))

                self.root.after(0, done_err)

        threading.Thread(target=worker, daemon=True).start()


def _normalize_drop_path(s: str) -> str:
    s = s.strip()
    if not s:
        return ""
    if s.startswith("{") and s.endswith("}"):
        s = s[1:-1]
    return s


def main():
    if _HAS_DND:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()

    app = AppBase(root)

    if _HAS_DND:
        def on_drop(event):
            path = _normalize_drop_path(event.data)
            if path and os.path.isdir(path):
                app.varFolder.set(path)
                app.varOut.set(app._default_output_for_folder_prefix(path, app.varPrefix.get()))
                app._outIsDefault = True
                app._preview_files(path)
                app.varStatus.set("Folder dropped.")
            else:
                app.varStatus.set("Drop a folder (not a file).")

        app.lblDrop.drop_target_register(DND_FILES)
        app.lblDrop.dnd_bind("<<Drop>>", on_drop)

    root.mainloop()


if __name__ == "__main__":
    main()

#endregion
