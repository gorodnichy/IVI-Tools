# docx_pdf_booklet_4up_ivim_app.py

#region docx_pdf_booklet_4up_ivim_app.py Professional ttkbootstrap UI + drag&drop: DOCX/PDF -> 4-up (cut-in-middle) print-ready PDF + clickable IVIM link
# Version 0.5.0 (2025/12/18)
# Added drag-and-drop file support (tkinterdnd2); upgraded UI polish and added clickable www.IVIM.ca link; kept same 4-up ordering logic

import os
import sys
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox

try:
    import ttkbootstrap as tb
except Exception:
    tb = None

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES  # pip install tkinterdnd2
except Exception:
    TkinterDnD = None
    DND_FILES = None

try:
    from docx2pdf import convert as docx2pdf_convert
except Exception:
    docx2pdf_convert = None

try:
    from pypdf import PdfReader, PdfWriter, Transformation
    from pypdf._page import PageObject
except Exception:
    PdfReader = PdfWriter = Transformation = PageObject = None


# ======================== =
# Helpers -----
# ======================== =

def resource_path(rel_path: str) -> str:
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, rel_path)
    return os.path.join(os.path.abspath("."), rel_path)

def _next_multiple(intN: int, intK: int) -> int:
    return ((intN + intK - 1) // intK) * intK

def _as_real_or_blank(intP: int | None, intRealPages: int) -> int | None:
    if intP is None:
        return None
    return intP if 1 <= intP <= intRealPages else None

def _blank_like(page) -> PageObject:
    fltW = float(page.mediabox.width)
    fltH = float(page.mediabox.height)
    return PageObject.create_blank_page(width=fltW, height=fltH)

def _get_page_or_blank(pages, idx_1_based: int | None, blank_template: PageObject) -> PageObject:
    if idx_1_based is None:
        return _blank_like(blank_template)
    return pages[idx_1_based - 1]


# ======================== =
# Ordering logic (pads to multiple of 8) -----
# ======================== =

def _sides_4up_cut_middle_user_scheme(intRealPages: int) -> list[tuple[int | None, int | None, int | None, int | None]]:
    """
    Each OUTPUT PDF page corresponds to one printed SIDE and has 4 slots:
      (top-left, top-right, bottom-left, bottom-right)

    Pads to nearest multiple of 8 pages (by blanks), then:
      Page 1: (last,1) and (last-2,3)
      Page 2: (2,last-1) and (4,last-3)
      then repeats with 4-page stride on the small numbers, 4-page stride down on the big numbers.
    """
    intTotal = _next_multiple(intRealPages, 8)
    intBlocks = intTotal // 8
    vecSides: list[tuple[int | None, int | None, int | None, int | None]] = []

    for i in range(intBlocks):
        a = intTotal - 4 * i
        front = (
            _as_real_or_blank(a, intRealPages),
            _as_real_or_blank(1 + 4 * i, intRealPages),
            _as_real_or_blank(a - 2, intRealPages),
            _as_real_or_blank(3 + 4 * i, intRealPages),
        )
        back = (
            _as_real_or_blank(2 + 4 * i, intRealPages),
            _as_real_or_blank(a - 1, intRealPages),
            _as_real_or_blank(4 + 4 * i, intRealPages),
            _as_real_or_blank(a - 3, intRealPages),
        )
        vecSides.append(front)
        vecSides.append(back)

    return vecSides


# ======================== =
# PDF conversion / imposition -----
# ======================== =

def docx_to_pdf(strDocx: str, strPdf: str) -> None:
    if docx2pdf_convert is None:
        raise RuntimeError("DOCX support requires docx2pdf (and typically Microsoft Word on Windows). Install: pip install docx2pdf")
    docx2pdf_convert(strDocx, strPdf)

def impose_booklet_4up_cut_middle(strPdfIn: str, strPdfOut: str) -> None:
    if PdfReader is None or PdfWriter is None or Transformation is None or PageObject is None:
        raise RuntimeError("Missing dependency: pypdf. Install: pip install pypdf")

    reader = PdfReader(strPdfIn)
    pages = reader.pages
    if len(pages) == 0:
        raise RuntimeError("Input PDF has 0 pages.")

    template = pages[0]
    fltW = float(template.mediabox.width)
    fltH = float(template.mediabox.height)

    vecSides = _sides_4up_cut_middle_user_scheme(len(pages))
    writer = PdfWriter()
    s = Transformation().scale(0.5, 0.5)

    for (tl, tr, bl, br) in vecSides:
        pTL = _get_page_or_blank(pages, tl, template)
        pTR = _get_page_or_blank(pages, tr, template)
        pBL = _get_page_or_blank(pages, bl, template)
        pBR = _get_page_or_blank(pages, br, template)

        out = PageObject.create_blank_page(width=fltW, height=fltH)
        out.merge_transformed_page(pTL, s.translate(0, fltH / 2))
        out.merge_transformed_page(pTR, s.translate(fltW / 2, fltH / 2))
        out.merge_transformed_page(pBL, s.translate(0, 0))
        out.merge_transformed_page(pBR, s.translate(fltW / 2, 0))
        writer.add_page(out)

    with open(strPdfOut, "wb") as f:
        writer.write(f)


# ======================== =
# UI text -----
# ======================== =

ABOUT_TEXT = (
    "IVIM Booklet Print-Ready (4 pages per sheet)\n\n"
    "Drag & drop or browse a DOCX/PDF.\n"
    "The app creates a print-ready 4-up PDF (cut-in-middle).\n\n"
    "Ordering (with blanks padded to nearest multiple of 8):\n"
    "• Side 1: (last,1) and (last-2,3)\n"
    "• Side 2: (2,last-1) and (4,last-3)\n"
    "…and so on.\n\n"
    "Printing:\n"
    "1) Print duplex (flip on long edge is usually correct)\n"
    "2) Cut each sheet in half horizontally\n"
    "3) Stack halves in order, fold, staple\n\n"
    "Made by IVIM. More tools: www.ivim.ca"
)


# ======================== =
# App -----
# ======================== =

class App:
    def __init__(self):
        if tb is None:
            raise RuntimeError("Missing dependency: ttkbootstrap. Install: pip install ttkbootstrap")

        # Root with optional drag&drop
        if TkinterDnD is not None:
            self.root = TkinterDnD.Tk()
            self.hasDnd = True
        else:
            self.root = tk.Tk()
            self.hasDnd = False

        tb.Style(theme="flatly")  # apply theme to current root

        self.root.title("IVIM Booklet Print-Ready (4-up)")
        self.root.geometry("860x360")
        self.root.resizable(False, False)

        try:
            self.root.iconbitmap(resource_path("ivim.ico"))
        except Exception:
            pass

        self.strInPath = tk.StringVar(value="")
        self.strStatus = tk.StringVar(value="Drop a DOCX/PDF or click Browse. Output: *-printready-4-per-sheet.pdf")

        self._build_ui()
        if self.hasDnd:
            self._enable_dnd()

    def _build_ui(self):
        frm = tb.Frame(self.root, padding=16)
        frm.pack(fill="both", expand=True)

        tb.Label(frm, text="Booklet Print-Ready", font=("Segoe UI", 16, "bold")).pack(anchor="w")
        tb.Label(frm, text="DOCX/PDF → 4 pages per sheet → cut in middle → fold & staple", bootstyle="secondary").pack(anchor="w", pady=(2, 14))

        frmPick = tb.Frame(frm)
        frmPick.pack(fill="x", pady=(0, 10))

        self.entPath = tb.Entry(frmPick, textvariable=self.strInPath)
        self.entPath.pack(side="left", fill="x", expand=True)

        tb.Button(frmPick, text="Browse...", command=self.pick_input, bootstyle="secondary").pack(side="left", padx=(10, 0))

        frmBtns = tb.Frame(frm)
        frmBtns.pack(fill="x", pady=(6, 0))

        self.btnRun = tb.Button(frmBtns, text="Generate 4-up PDF", command=self.run, bootstyle="primary")
        self.btnRun.pack(side="left")

        tb.Button(frmBtns, text="About", command=self.show_about, bootstyle="info-outline").pack(side="left", padx=(10, 0))
        tb.Button(frmBtns, text="Quit", command=self.root.destroy, bootstyle="secondary").pack(side="left", padx=(10, 0))

        frmProg = tb.Frame(frm)
        frmProg.pack(fill="x", pady=(14, 0))

        self.pb = tb.Progressbar(frmProg, mode="indeterminate")
        self.pb.pack(side="left", fill="x", expand=True)

        tb.Label(frm, textvariable=self.strStatus, bootstyle="secondary").pack(anchor="w", pady=(10, 0))

        frmFooter = tb.Frame(frm)
        frmFooter.pack(fill="x", pady=(14, 0))

        tb.Label(frmFooter, text="Made by IVIM. More tools:", bootstyle="secondary").pack(side="left")

        lblLink = tb.Label(frmFooter, text="www.IVIM.ca", bootstyle="primary")
        lblLink.pack(side="left", padx=(6, 0))

        # make it look and act like a link
        lblLink.configure(cursor="hand2")
        try:
            f = ("Segoe UI", 10, "underline")
            lblLink.configure(font=f)
        except Exception:
            pass
        lblLink.bind("<Button-1>", lambda _e: webbrowser.open("https://www.ivim.ca"))

        if not self.hasDnd:
            tb.Label(frm, text="Tip: install drag & drop support with: pip install tkinterdnd2", bootstyle="secondary").pack(anchor="w", pady=(10, 0))

    def _enable_dnd(self):
        # Accept drops on the whole window and on the entry
        try:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<Drop>>", self._on_drop)

            self.entPath.drop_target_register(DND_FILES)
            self.entPath.dnd_bind("<<Drop>>", self._on_drop)
        except Exception:
            self.hasDnd = False

    def _on_drop(self, event):
        # event.data can be '{C:\path with spaces\file.pdf}' or 'C:\a.pdf C:\b.pdf'
        strData = (event.data or "").strip()
        if not strData:
            return

        if strData.startswith("{") and strData.endswith("}"):
            strData = strData[1:-1]

        # if multiple files, take the first token (handles basic cases)
        if " " in strData and os.path.exists(strData) is False:
            # fallback parse for multiple paths; handle braces already stripped
            parts = strData.split()
            for p in parts:
                p = p.strip().strip("{").strip("}")
                if os.path.exists(p):
                    strData = p
                    break

        self.strInPath.set(strData)

    def pick_input(self):
        p = filedialog.askopenfilename(filetypes=[("DOCX or PDF", "*.docx *.pdf"), ("Word Document", "*.docx"), ("PDF", "*.pdf")])
        if p:
            self.strInPath.set(p)

    def show_about(self):
        messagebox.showinfo("About", ABOUT_TEXT)

    def _set_busy(self, isBusy: bool, strMsg: str | None = None):
        self.btnRun.configure(state=("disabled" if isBusy else "normal"))
        if strMsg is not None:
            self.strStatus.set(strMsg)
        if isBusy:
            self.pb.start(10)
        else:
            self.pb.stop()
        self.root.update_idletasks()

    def run(self):
        strIn = self.strInPath.get().strip()
        if not strIn or not os.path.isfile(strIn):
            messagebox.showerror("Error", "Please select a valid .docx or .pdf file.")
            return

        if PdfReader is None:
            messagebox.showerror("Error", "Missing pypdf. Install: pip install pypdf")
            return

        strExt = os.path.splitext(strIn)[1].lower()
        if strExt not in (".docx", ".pdf"):
            messagebox.showerror("Error", "Input must be .docx or .pdf")
            return

        if strExt == ".docx" and docx2pdf_convert is None:
            messagebox.showerror("Error", "DOCX support requires docx2pdf. Install: pip install docx2pdf\n(Windows typically needs Microsoft Word installed.)")
            return

        strBase = os.path.splitext(os.path.basename(strIn))[0]
        strOutDir = os.path.dirname(strIn)
        strTmpPdf = os.path.join(strOutDir, f"{strBase}__tmp.pdf")
        strPdfSrc = strIn if strExt == ".pdf" else strTmpPdf
        strPdfOut = os.path.join(strOutDir, f"{strBase}-printready-4-per-sheet.pdf")

        try:
            if strExt == ".docx":
                self._set_busy(True, "Converting DOCX → PDF...")
                if os.path.exists(strTmpPdf):
                    os.remove(strTmpPdf)
                docx_to_pdf(strIn, strTmpPdf)

            self._set_busy(True, "Generating booklet 4-up PDF...")
            impose_booklet_4up_cut_middle(strPdfSrc, strPdfOut)

            if strExt == ".docx":
                try:
                    os.remove(strTmpPdf)
                except Exception:
                    pass

            self._set_busy(False, f"Done: {strPdfOut}")
            messagebox.showinfo("Done", f"Created:\n{strPdfOut}")
        except Exception as e:
            self._set_busy(False, "Error.")
            messagebox.showerror("Error", str(e))

    def mainloop(self):
        self.root.mainloop()


if __name__ == "__main__":
    App().mainloop()

if (False):
    # Install:
    # pip install ttkbootstrap pypdf docx2pdf tkinterdnd2
    #
    # Build (PowerShell):
    # pyinstaller --onefile --windowed --icon ivim.ico --add-data "ivim.ico;." --name IVIM_Booklet4Up docx_pdf_booklet_4up_ivim_app.py
    pass
#endregion
