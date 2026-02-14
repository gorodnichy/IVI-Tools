# IVI_a2v.py
# Version 0.4.0 (2026/01/11)
# Add user-selectable aspect layouts: Horizontal (16:9), Square (1:1), Vertical (9:16) for YouTube/Shorts

#region IVI_a2v mp3+image -> mp4

from __future__ import annotations

import os
import shutil
import subprocess
import threading
import traceback
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Tuple

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # pip install tkinterdnd2
    _HAS_DND = True
except Exception:
    TkinterDnD = None  # type: ignore
    DND_FILES = None   # type: ignore
    _HAS_DND = False


APP_NAME = "IVI_a2v"
APP_LINK = "www.IVIM.ca"
ICON_NAME = "IVI_a2v.ico"

# Speed controls
VIDEO_FPS = 1
X264_PRESET = "ultrafast"

# Audio defaults (controlled via UI)
AUDIO_MODE_DEFAULT = "auto"         # "auto" | "copy" | "aac"
AUDIO_BITRATE_DEFAULT = "320k"

# Layout defaults (controlled via UI)
LAYOUT_DEFAULT = "horizontal"       # "horizontal" | "square" | "vertical"


@dataclass
class Inputs:
    pAudio: Optional[Path] = None
    pImage: Optional[Path] = None
    pOut: Optional[Path] = None


def _app_dir() -> Path:
    return Path(__file__).resolve().parent


def _icon_path() -> Path:
    return _app_dir() / ICON_NAME


def _find_ffmpeg() -> Optional[str]:
    exe = "ffmpeg.exe" if os.name == "nt" else "ffmpeg"
    pLocal = _app_dir() / exe
    if pLocal.exists():
        return str(pLocal)
    return shutil.which("ffmpeg")


def _is_audio(p: Path) -> bool:
    return p.suffix.lower() in {".mp3", ".wav", ".m4a", ".aac", ".flac", ".ogg"}


def _is_image(p: Path) -> bool:
    return p.suffix.lower() in {".jpg", ".jpeg", ".png", ".bmp", ".webp"}


def _safe_out_path(pAudio: Path, pImage: Path) -> Path:
    stem = f"{pAudio.stem}__{pImage.stem}".replace(" ", "_")
    return pAudio.with_name(stem).with_suffix(".mp4")


def _parse_dnd_paths(raw: str) -> list[str]:
    raw = (raw or "").strip()
    if not raw:
        return []
    paths: list[str] = []
    cur, in_brace = "", False
    for ch in raw:
        if ch == "{":
            in_brace, cur = True, ""
        elif ch == "}":
            in_brace = False
            if cur:
                paths.append(cur)
                cur = ""
        elif ch.isspace() and not in_brace:
            if cur:
                paths.append(cur)
                cur = ""
        else:
            cur += ch
    if cur:
        paths.append(cur)
    return paths


def _win_font_arial() -> Optional[str]:
    if os.name != "nt":
        return None
    p = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts" / "arial.ttf"
    return str(p) if p.exists() else None


def _ff_escape_drawtext(s: str) -> str:
    return (
        s.replace("\\", "\\\\")
         .replace(":", "\\:")
         .replace("'", "\\'")
         .replace("%", "\\%")
    )


def _layout_dims(layout: str) -> Tuple[int, int]:
    layout = (layout or "").strip().lower()
    if layout == "vertical":
        return (1080, 1920)   # YouTube Shorts / Reels
    if layout == "square":
        return (1080, 1080)   # IG square / generic
    return (1920, 1080)       # default horizontal YouTube


class AppUI:
    def __init__(self) -> None:
        self.inputs = Inputs()
        self.ffmpeg = _find_ffmpeg()
        self.fontfile = _win_font_arial()

        self.root = self._make_root_safely()
        self._apply_icon()
        self._build()
        self._wire_dnd()
        self.root.after(50, self._bring_to_front)

    def _make_root_safely(self):
        if _HAS_DND and TkinterDnD is not None:
            try:
                root = TkinterDnD.Tk()
                root.title(f"{APP_NAME}  |  MP3 + Image → MP4")
                root.minsize(900, 760)
                return root
            except Exception:
                pass
        root = tk.Tk()
        root.title(f"{APP_NAME}  |  MP3 + Image → MP4")
        root.minsize(900, 760)
        return root

    def _bring_to_front(self) -> None:
        try:
            self.root.deiconify()
            self.root.lift()
            self.root.attributes("-topmost", True)
            self.root.after(150, lambda: self.root.attributes("-topmost", False))
            self.root.focus_force()
        except Exception:
            pass

    def _apply_icon(self) -> None:
        pIco = _icon_path()
        if pIco.exists():
            try:
                self.root.iconbitmap(str(pIco))
            except Exception:
                pass

    def _build(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(5, weight=1)

        frmTop = ttk.Frame(self.root, padding=14)
        frmTop.grid(row=0, column=0, sticky="ew")
        frmTop.columnconfigure(1, weight=1)

        ttk.Label(frmTop, text="MP3 + Image → MP4", font=("Segoe UI", 16, "bold")).grid(row=0, column=0, sticky="w")

        frmBtns = ttk.Frame(frmTop)
        frmBtns.grid(row=0, column=1, sticky="e")
        ttk.Button(frmBtns, text="About", command=self.on_about).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(frmBtns, text="Open IVIM.ca", command=self.on_open_site).grid(row=0, column=1)

        frmInputs = ttk.LabelFrame(self.root, text="Inputs", padding=14)
        frmInputs.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 10))
        frmInputs.columnconfigure(1, weight=1)

        self.varAudio = tk.StringVar(value="")
        self.varImage = tk.StringVar(value="")
        self.varOut = tk.StringVar(value="")

        ttk.Label(frmInputs, text="Audio:").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(frmInputs, textvariable=self.varAudio).grid(row=0, column=1, sticky="ew", padx=10, pady=4)
        ttk.Button(frmInputs, text="Browse…", command=self.on_browse_audio).grid(row=0, column=2, pady=4)

        ttk.Label(frmInputs, text="Image:").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(frmInputs, textvariable=self.varImage).grid(row=1, column=1, sticky="ew", padx=10, pady=4)
        ttk.Button(frmInputs, text="Browse…", command=self.on_browse_image).grid(row=1, column=2, pady=4)

        ttk.Label(frmInputs, text="Output (MP4):").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(frmInputs, textvariable=self.varOut).grid(row=2, column=1, sticky="ew", padx=10, pady=4)
        ttk.Button(frmInputs, text="Save as…", command=self.on_save_as).grid(row=2, column=2, pady=4)

        frmCover = ttk.LabelFrame(self.root, text="Song cover text (optional)", padding=14)
        frmCover.grid(row=2, column=0, sticky="ew", padx=14, pady=(0, 10))
        frmCover.columnconfigure(1, weight=1)

        self.varTitle = tk.StringVar(value="")
        self.varSubtitle = tk.StringVar(value="")

        ttk.Label(frmCover, text="Title:").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(frmCover, textvariable=self.varTitle).grid(row=0, column=1, sticky="ew", padx=10, pady=4)

        ttk.Label(frmCover, text="Subtitle (authors):").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(frmCover, textvariable=self.varSubtitle).grid(row=1, column=1, sticky="ew", padx=10, pady=4)

        frmLayout = ttk.LabelFrame(self.root, text="Layout / aspect ratio", padding=14)
        frmLayout.grid(row=3, column=0, sticky="ew", padx=14, pady=(0, 10))
        frmLayout.columnconfigure(3, weight=1)

        self.varLayout = tk.StringVar(value=LAYOUT_DEFAULT)

        ttk.Radiobutton(frmLayout, text="Horizontal (16:9) — YouTube", value="horizontal", variable=self.varLayout, command=self.on_layout_change).grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(frmLayout, text="Square (1:1)", value="square", variable=self.varLayout, command=self.on_layout_change).grid(row=0, column=1, sticky="w", padx=(18, 0))
        ttk.Radiobutton(frmLayout, text="Vertical (9:16) — Shorts", value="vertical", variable=self.varLayout, command=self.on_layout_change).grid(row=0, column=2, sticky="w", padx=(18, 0))

        self.varLayoutInfo = tk.StringVar(value="")
        ttk.Label(frmLayout, textvariable=self.varLayoutInfo).grid(row=0, column=3, sticky="e")
        self.on_layout_change()

        frmAudio = ttk.LabelFrame(self.root, text="Audio mode", padding=14)
        frmAudio.grid(row=4, column=0, sticky="ew", padx=14, pady=(0, 10))
        frmAudio.columnconfigure(4, weight=1)

        self.varAudioMode = tk.StringVar(value=AUDIO_MODE_DEFAULT)
        self.varAudioBitrate = tk.StringVar(value=AUDIO_BITRATE_DEFAULT)

        ttk.Radiobutton(frmAudio, text="AUTO (recommended)", value="auto", variable=self.varAudioMode, command=self.on_audio_mode_change).grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(frmAudio, text="COPY (best quality, may be incompatible)", value="copy", variable=self.varAudioMode, command=self.on_audio_mode_change).grid(row=0, column=1, sticky="w", padx=(18, 0))
        ttk.Radiobutton(frmAudio, text="AAC (most compatible)", value="aac", variable=self.varAudioMode, command=self.on_audio_mode_change).grid(row=0, column=2, sticky="w", padx=(18, 0))

        ttk.Label(frmAudio, text="AAC bitrate:").grid(row=0, column=3, sticky="e", padx=(18, 6))
        self.cmbBitrate = ttk.Combobox(frmAudio, width=8, textvariable=self.varAudioBitrate, values=["128k", "160k", "192k", "256k", "320k"], state="readonly")
        self.cmbBitrate.grid(row=0, column=4, sticky="w")
        self.on_audio_mode_change()

        hint = "Tip: drag & drop an audio file and an image onto this window." if _HAS_DND else "Tip: optional drag&drop: pip install tkinterdnd2"
        ttk.Label(self.root, text=hint, padding=(14, 0)).grid(row=1, column=0, sticky="e", padx=14, pady=(0, 6))

        frmMid = ttk.Frame(self.root, padding=(14, 0, 14, 14))
        frmMid.grid(row=5, column=0, sticky="nsew")
        frmMid.columnconfigure(0, weight=1)
        frmMid.rowconfigure(1, weight=1)

        frmActions = ttk.Frame(frmMid)
        frmActions.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        frmActions.columnconfigure(0, weight=1)

        self.varStatus = tk.StringVar(value=("Status: ready" if self.ffmpeg else "Status: ffmpeg not found"))
        ttk.Label(frmActions, textvariable=self.varStatus).grid(row=0, column=0, sticky="w")

        self.btnMake = ttk.Button(frmActions, text="Create MP4", command=self.on_make)
        self.btnMake.grid(row=0, column=1, sticky="e", padx=(10, 0))

        self.btnClear = ttk.Button(frmActions, text="Clear", command=self.on_clear)
        self.btnClear.grid(row=0, column=2, sticky="e", padx=(10, 0))

        frmLog = ttk.LabelFrame(frmMid, text="Log", padding=10)
        frmLog.grid(row=1, column=0, sticky="nsew")
        frmLog.columnconfigure(0, weight=1)
        frmLog.rowconfigure(0, weight=1)

        self.txtLog = tk.Text(frmLog, wrap="word")
        self.txtLog.grid(row=0, column=0, sticky="nsew")
        scr = ttk.Scrollbar(frmLog, orient="vertical", command=self.txtLog.yview)
        scr.grid(row=0, column=1, sticky="ns")
        self.txtLog.configure(yscrollcommand=scr.set)

        self._log(f"{APP_NAME} ready.")
        self._log(f"ffmpeg: {self.ffmpeg if self.ffmpeg else 'NOT FOUND'}")
        if not self.ffmpeg:
            self._log("Install ffmpeg or put ffmpeg.exe next to this script.")
        self._log(f"Speed: VIDEO_FPS={VIDEO_FPS}, X264_PRESET={X264_PRESET}")
        self._log(f"Layout: {self.varLayout.get()} ({self.varLayoutInfo.get()})")
        self._log(f"Audio mode: {self.varAudioMode.get()} (AAC bitrate={self.varAudioBitrate.get()})")

    def on_layout_change(self) -> None:
        w, h = _layout_dims(self.varLayout.get())
        self.varLayoutInfo.set(f"{w}×{h}")

    def on_audio_mode_change(self) -> None:
        st = "readonly" if self.varAudioMode.get() == "aac" else "disabled"
        try:
            self.cmbBitrate.configure(state=st)
        except Exception:
            pass

    def _wire_dnd(self) -> None:
        if not (_HAS_DND and hasattr(self.root, "drop_target_register")):
            return
        try:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind("<<Drop>>", self.on_drop_files)
        except Exception:
            pass

    def _log(self, s: str) -> None:
        self.txtLog.insert("end", s.rstrip() + "\n")
        self.txtLog.see("end")

    def _set_inputs(self, p: Path) -> None:
        if _is_audio(p):
            self.inputs.pAudio = p
            self.varAudio.set(str(p))
        elif _is_image(p):
            self.inputs.pImage = p
            self.varImage.set(str(p))

        if self.inputs.pAudio and self.inputs.pImage and not self.inputs.pOut:
            pOut = _safe_out_path(self.inputs.pAudio, self.inputs.pImage)
            self.inputs.pOut = pOut
            self.varOut.set(str(pOut))

    def _validate(self) -> bool:
        if not self.ffmpeg:
            messagebox.showerror(APP_NAME, "ffmpeg was not found.\nInstall ffmpeg or put ffmpeg.exe next to this script.")
            return False

        pAudio = Path(self.varAudio.get().strip()) if self.varAudio.get().strip() else None
        pImage = Path(self.varImage.get().strip()) if self.varImage.get().strip() else None
        pOut = Path(self.varOut.get().strip()) if self.varOut.get().strip() else None

        if not pAudio or not pAudio.exists() or not _is_audio(pAudio):
            messagebox.showerror(APP_NAME, "Select a valid audio file (.mp3, .wav, .m4a, ...).")
            return False
        if not pImage or not pImage.exists() or not _is_image(pImage):
            messagebox.showerror(APP_NAME, "Select a valid image file (.jpg, .png, ...).")
            return False
        if not pOut:
            messagebox.showerror(APP_NAME, "Select an output .mp4 file.")
            return False
        if pOut.suffix.lower() != ".mp4":
            messagebox.showerror(APP_NAME, "Output file must end with .mp4.")
            return False

        self.inputs.pAudio, self.inputs.pImage, self.inputs.pOut = pAudio, pImage, pOut
        return True

    def on_about(self) -> None:
        sF = ("Arial (Windows)" if self.fontfile else "Default/system")
        w, h = _layout_dims(self.varLayout.get())
        messagebox.showinfo(
            f"About {APP_NAME}",
            f"{APP_NAME}\n\n"
            "Create an MP4 from an audio file and a single image.\n"
            "Optional: overlay Title + Subtitle text.\n\n"
            f"Layout: {self.varLayout.get()} ({w}×{h})\n"
            f"ffmpeg: {'OK' if self.ffmpeg else 'NOT FOUND'}\n"
            f"Font: {sF}\n"
            f"Speed: VIDEO_FPS={VIDEO_FPS}, X264_PRESET={X264_PRESET}\n"
            f"Audio mode: {self.varAudioMode.get()} (AAC bitrate={self.varAudioBitrate.get()})\n\n"
            f"Site: {APP_LINK}"
        )

    def on_open_site(self) -> None:
        import webbrowser
        webbrowser.open(f"https://{APP_LINK}")

    def on_browse_audio(self) -> None:
        p = filedialog.askopenfilename(
            title="Select audio file",
            filetypes=[("Audio", "*.mp3 *.wav *.m4a *.aac *.flac *.ogg"), ("All files", "*.*")]
        )
        if p:
            pP = Path(p)
            self._set_inputs(pP)
            self._log(f"Audio: {pP}")

    def on_browse_image(self) -> None:
        p = filedialog.askopenfilename(
            title="Select image file",
            filetypes=[("Images", "*.jpg *.jpeg *.png *.bmp *.webp"), ("All files", "*.*")]
        )
        if p:
            pP = Path(p)
            self._set_inputs(pP)
            self._log(f"Image: {pP}")

    def on_save_as(self) -> None:
        p = filedialog.asksaveasfilename(
            title="Save MP4 as",
            defaultextension=".mp4",
            filetypes=[("MP4 video", "*.mp4")]
        )
        if p:
            pP = Path(p)
            self.inputs.pOut = pP
            self.varOut.set(str(pP))
            self._log(f"Output: {pP}")

    def on_clear(self) -> None:
        self.inputs = Inputs()
        self.varAudio.set("")
        self.varImage.set("")
        self.varOut.set("")
        self.varTitle.set("")
        self.varSubtitle.set("")
        self.varLayout.set(LAYOUT_DEFAULT)
        self.on_layout_change()
        self.varAudioMode.set(AUDIO_MODE_DEFAULT)
        self.varAudioBitrate.set(AUDIO_BITRATE_DEFAULT)
        self.on_audio_mode_change()
        self._log("Cleared.")

    def on_drop_files(self, event) -> None:
        for p in _parse_dnd_paths(getattr(event, "data", "")):
            pP = Path(p)
            if pP.exists():
                self._set_inputs(pP)
                self._log(f"Dropped: {pP}")

    def on_make(self) -> None:
        if not self._validate():
            return
        self.btnMake.configure(state="disabled")
        self.varStatus.set("Status: working…")
        self._log("Starting ffmpeg…")
        threading.Thread(target=self._run_ffmpeg, daemon=True).start()

    def _drawtext_filters(self) -> list[str]:
        title = self.varTitle.get().strip()
        subtitle = self.varSubtitle.get().strip()
        if not title and not subtitle:
            return []

        common = [
            "fontcolor=white",
            "shadowcolor=black",
            "shadowx=2",
            "shadowy=2",
            "box=1",
            "boxcolor=black@0.45",
            "boxborderw=18",
            "x=80",
        ]
        font = [f"fontfile='{_ff_escape_drawtext(self.fontfile)}'"] if self.fontfile else []

        out: list[str] = []
        if title:
            out.append(
                "drawtext=" + ":".join(
                    font + common + [
                        f"text='{_ff_escape_drawtext(title)}'",
                        "fontsize=64",
                        "y=H-220",
                    ]
                )
            )
        if subtitle:
            out.append(
                "drawtext=" + ":".join(
                    font + common + [
                        f"text='{_ff_escape_drawtext(subtitle)}'",
                        "fontsize=38",
                        "y=H-140",
                    ]
                )
            )
        return out

    def _audio_args(self, pAudio: Path) -> list[str]:
        mode = (self.varAudioMode.get() or "auto").strip().lower()

        if mode == "auto":
            ext = pAudio.suffix.lower()
            if ext in {".mp3", ".aac", ".m4a"}:
                return ["-c:a", "copy"]
            br = (self.varAudioBitrate.get() or AUDIO_BITRATE_DEFAULT).strip()
            return ["-c:a", "aac", "-b:a", br]

        if mode == "copy":
            return ["-c:a", "copy"]

        br = (self.varAudioBitrate.get() or AUDIO_BITRATE_DEFAULT).strip()
        return ["-c:a", "aac", "-b:a", br]

    def _ffmpeg_cmd(self, pAudio: Path, pImage: Path, pOut: Path) -> list[str]:
        out_w, out_h = _layout_dims(self.varLayout.get())
        txt = self._drawtext_filters()

        vf_parts = [
            f"[0:v]scale={out_w}:{out_h}:force_original_aspect_ratio=increase,crop={out_w}:{out_h},gblur=sigma=28[bg]",
            f"[0:v]scale={out_w}:{out_h}:force_original_aspect_ratio=decrease[fg]",
            f"[bg][fg]overlay=(W-w)/2:(H-h)/2[v0]",
        ]

        if txt:
            vf_parts.append(f"[v0]{','.join(txt)}[v]")
            map_v = "[v]"
        else:
            map_v = "[v0]"

        filter_complex = ";".join(vf_parts)

        return [
            str(self.ffmpeg),
            "-y",
            "-loop", "1",
            "-i", str(pImage),
            "-i", str(pAudio),
            "-filter_complex", filter_complex,
            "-map", map_v,
            "-map", "1:a",
            "-c:v", "libx264",
            "-preset", str(X264_PRESET),
            "-pix_fmt", "yuv420p",
            "-r", str(VIDEO_FPS),
            *self._audio_args(pAudio),
            "-shortest",
            str(pOut),
        ]

    def _run_ffmpeg(self) -> None:
        assert self.inputs.pAudio and self.inputs.pImage and self.inputs.pOut
        cmd = self._ffmpeg_cmd(self.inputs.pAudio, self.inputs.pImage, self.inputs.pOut)
        self._log("Command:")
        self._log("  " + " ".join(cmd))

        try:
            p = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1)
            assert p.stdout is not None
            for line in p.stdout:
                self._log(line.rstrip())
            rc = p.wait()
        except Exception as e:
            self.root.after(0, lambda: self._done(False, f"Failed to run ffmpeg: {e}"))
            return

        if rc == 0 and self.inputs.pOut.exists():
            self.root.after(0, lambda: self._done(True, f"Created: {self.inputs.pOut}"))
        else:
            self.root.after(0, lambda: self._done(False, f"ffmpeg exited with code {rc}"))

    def _done(self, ok: bool, msg: str) -> None:
        self.btnMake.configure(state="normal")
        self.varStatus.set("Status: done" if ok else "Status: error")
        self._log(msg)
        (messagebox.showinfo if ok else messagebox.showerror)(APP_NAME, msg)

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    AppUI().run()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        tb = traceback.format_exc()
        print(tb)
        try:
            messagebox.showerror(APP_NAME, tb)
        except Exception:
            pass


# Build EXE (PyInstaller)
# pyinstaller --onefile --windowed --icon=IVI_a2v.ico --add-binary "ffmpeg.exe;." IVI_a2v.py

#endregion
