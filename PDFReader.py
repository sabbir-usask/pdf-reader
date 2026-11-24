"""
PDF Reader - reads PDF files aloud on Windows.

Uses the speech voices built into Windows (SAPI), PyMuPDF to pull the text
out of the PDF, and Tkinter for the window.

Run directly:   python PDFReader.py  [optional path to a .pdf]
Or build an .exe with Setup-and-Build.bat
"""

import bisect
import json
import os
import re
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

APP_NAME = "PDF Reader"
SETTINGS_DIR = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "PDFReader")
SETTINGS_FILE = os.path.join(SETTINGS_DIR, "settings.json")

# SAPI flags
SVSF_ASYNC = 1
SVSF_PURGE = 2
SSFM_CREATE_FOR_WRITE = 3


# ----------------------------------------------------------------------------
# Settings
# ----------------------------------------------------------------------------
def load_settings():
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_settings(data):
    try:
        os.makedirs(SETTINGS_DIR, exist_ok=True)
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass


# ----------------------------------------------------------------------------
# PDF text extraction
# ----------------------------------------------------------------------------
def clean_block(text):
    """Join the lines of one paragraph block into flowing text."""
    text = text.replace("­", "")                 # soft hyphens
    text = re.sub(r"(\w)-\n(?=[a-z])", r"\1", text)   # re-join hyphen-split words
    text = re.sub(r"\s*\n\s*", " ", text)
    text = re.sub(r"[ \t ]+", " ", text)
    return text.strip()


def extract_pages(path, password=None):
    """Return a list of pages; each page is a list of paragraph strings."""
    try:
        import fitz  # PyMuPDF
    except ImportError:
        fitz = None

    pages = []
    if fitz is not None:
        doc = fitz.open(path)
        if doc.needs_pass and not doc.authenticate(password or ""):
            raise PermissionError("password")
        for page in doc:
            paras = []
            for b in page.get_text("blocks", sort=True):
                if len(b) > 6 and b[6] != 0:      # skip image blocks
                    continue
                t = clean_block(b[4])
                if t:
                    paras.append(t)
            pages.append(paras)
        doc.close()
        return pages

    # Fallback: pypdf
    from pypdf import PdfReader
    reader = PdfReader(path)
    if reader.is_encrypted and not reader.decrypt(password or ""):
        raise PermissionError("password")
    for page in reader.pages:
        raw = page.extract_text() or ""
        paras = [clean_block(p) for p in re.split(r"\n\s*\n", raw)]
        pages.append([p for p in paras if p])
    return pages


SENT_RE = re.compile(r"(?<=[.!?])\s+|(?<=[.!?][\"'\u201d\u2019)\]])\s+")


ABBREV_RE = re.compile(r"(?:\b(?:Dr|Mr|Mrs|Ms|Prof|St|vs|etc|al|Fig|Figs|No|Vol|pp|approx|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec|e\.g|i\.e)|\b[A-Z]|\d)\.$")


def split_sentences(paragraph):
    pieces = []
    for s in SENT_RE.split(paragraph):
        s = s.strip()
        if not s:
            continue
        # don't break after "Dr.", "et al.", "e.g.", initials or "3." style numbers
        if pieces and ABBREV_RE.search(pieces[-1]):
            pieces[-1] += " " + s
        else:
            pieces.append(s)
    out = []
    for s in pieces:
        # break very long run-ons (tables, lists) so highlighting keeps moving
        while len(s) > 260:
            cut = s.rfind(", ", 0, 240)
            if cut < 80:
                cut = s.rfind(" ", 0, 240)
            if cut < 40:
                cut = 240
            out.append(s[:cut + 1].strip())
            s = s[cut + 1:].strip()
        if s:
            out.append(s)
    return out


# ----------------------------------------------------------------------------
# Speech (Windows SAPI)
# ----------------------------------------------------------------------------
class Speaker:
    def __init__(self):
        self.voice = None
        self.tokens = []
        self.error = None
        try:
            import win32com.client
            self.voice = win32com.client.Dispatch("SAPI.SpVoice")
            toks = self.voice.GetVoices()
            self.tokens = [toks.Item(i) for i in range(toks.Count)]
        except Exception as e:  # not Windows, or speech missing
            self.error = str(e)

    @property
    def ok(self):
        return self.voice is not None

    def voice_names(self):
        return [t.GetDescription() for t in self.tokens]

    def set_voice(self, index):
        if self.ok and 0 <= index < len(self.tokens):
            self.voice.Voice = self.tokens[index]

    def set_rate(self, rate):        # -10 .. 10
        if self.ok:
            self.voice.Rate = int(rate)

    def set_volume(self, vol):       # 0 .. 100
        if self.ok:
            self.voice.Volume = int(vol)

    def speak(self, text):
        if self.ok:
            self.voice.Speak(text, SVSF_ASYNC | SVSF_PURGE)

    def stop(self):
        if self.ok:
            self.voice.Speak("", SVSF_ASYNC | SVSF_PURGE)

    def pause(self):
        if self.ok:
            self.voice.Pause()

    def resume(self):
        if self.ok:
            self.voice.Resume()

    def is_done(self):
        if not self.ok:
            return True
        return bool(self.voice.WaitUntilDone(0))


def save_wav(text, path, voice_index, rate, volume, on_done):
    """Write speech to a WAV file on a background thread."""
    def work():
        err = None
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            try:
                v = win32com.client.Dispatch("SAPI.SpVoice")
                toks = v.GetVoices()
                if 0 <= voice_index < toks.Count:
                    v.Voice = toks.Item(voice_index)
                v.Rate = int(rate)
                v.Volume = int(volume)
                stream = win32com.client.Dispatch("SAPI.SpFileStream")
                stream.Open(path, SSFM_CREATE_FOR_WRITE)
                v.AudioOutputStream = stream
                v.Speak(text)
                stream.Close()
            finally:
                pythoncom.CoUninitialize()
        except Exception as e:
            err = str(e)
        on_done(err)
    threading.Thread(target=work, daemon=True).start()


# ----------------------------------------------------------------------------
# Main window
# ----------------------------------------------------------------------------
class App:
    INK = "#1b1a2e"
    STRIP = "#24215a"
    MARK = "#ffe066"
    PAGE = "#fbfbfe"
    MUTED = "#6a6886"

    def __init__(self, root, start_file=None):
        self.root = root
        self.settings = load_settings()
        self.speaker = Speaker()

        self.path = None
        self.sentences = []     # list of dicts: text, page, start, end (char offsets)
        self.starts = []        # sentence start offsets, for click lookup
        self.page_starts = {}   # page number -> first sentence index
        self.idx = 0
        self.state = "stopped"  # stopped | playing | paused
        self.poll_id = None
        self.font_size = int(self.settings.get("font_size", 13))

        root.title(APP_NAME)
        root.geometry(self.settings.get("geometry", "1000x720"))
        root.minsize(640, 420)
        self._build_ui()
        self._bind_keys()
        root.protocol("WM_DELETE_WINDOW", self.on_close)

        if not self.speaker.ok:
            self.set_status("Speech is not available on this computer.")
            messagebox.showwarning(
                APP_NAME,
                "Windows speech could not be started, so the reader can show text "
                "but not read it aloud.\n\nDetails: " + str(self.speaker.error))

        if start_file:
            root.after(150, lambda: self.open_pdf(start_file))

    # ---------------- UI ----------------
    def _build_ui(self):
        style = ttk.Style()
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("Strip.TFrame", background=self.STRIP)
        style.configure("Strip.TLabel", background=self.STRIP, foreground="white",
                        font=("Segoe UI", 10))
        style.configure("Big.TButton", font=("Segoe UI", 10, "bold"), padding=(10, 6))
        style.configure("TButton", font=("Segoe UI", 10), padding=(8, 6))

        strip = ttk.Frame(self.root, style="Strip.TFrame", padding=(10, 8))
        strip.pack(fill="x")

        row1 = ttk.Frame(strip, style="Strip.TFrame")
        row1.pack(fill="x")
        ttk.Button(row1, text="Open PDF…", command=self.choose_file).pack(side="left", padx=(0, 10))
        self.play_btn = ttk.Button(row1, text="▶  Play", style="Big.TButton", width=10,
                                   command=self.toggle_play)
        self.play_btn.pack(side="left")
        ttk.Button(row1, text="■  Stop", command=self.stop).pack(side="left", padx=4)
        ttk.Button(row1, text="⏮", width=3, command=self.prev_sentence).pack(side="left")
        ttk.Button(row1, text="⏭", width=3, command=self.next_sentence).pack(side="left", padx=(4, 14))

        ttk.Label(row1, text="Page", style="Strip.TLabel").pack(side="left")
        self.page_var = tk.StringVar(value="–")
        self.page_box = ttk.Combobox(row1, textvariable=self.page_var, width=10, state="readonly")
        self.page_box.pack(side="left", padx=(6, 14))
        self.page_box.bind("<<ComboboxSelected>>", self.goto_page)

        ttk.Button(row1, text="A−", width=3, command=lambda: self.zoom(-1)).pack(side="left")
        ttk.Button(row1, text="A+", width=3, command=lambda: self.zoom(1)).pack(side="left", padx=(4, 14))
        ttk.Button(row1, text="Save as audio…", command=self.save_audio).pack(side="left")

        row2 = ttk.Frame(strip, style="Strip.TFrame")
        row2.pack(fill="x", pady=(8, 0))
        ttk.Label(row2, text="Voice", style="Strip.TLabel").pack(side="left")
        names = self.speaker.voice_names() or ["(no voices found)"]
        self.voice_var = tk.StringVar()
        self.voice_box = ttk.Combobox(row2, textvariable=self.voice_var, values=names,
                                      width=42, state="readonly")
        self.voice_box.pack(side="left", padx=(6, 16))
        saved_voice = self.settings.get("voice")
        self.voice_box.current(names.index(saved_voice) if saved_voice in names else 0)
        self.voice_box.bind("<<ComboboxSelected>>", self.on_voice)
        self.speaker.set_voice(self.voice_box.current())

        ttk.Label(row2, text="Speed", style="Strip.TLabel").pack(side="left")
        self.rate_var = tk.IntVar(value=int(self.settings.get("rate", 0)))
        rate = ttk.Scale(row2, from_=-10, to=10, orient="horizontal", length=150,
                         command=self.on_rate)
        rate.set(self.rate_var.get())
        rate.pack(side="left", padx=6)
        self.rate_lbl = ttk.Label(row2, text="", style="Strip.TLabel", width=6)
        self.rate_lbl.pack(side="left", padx=(0, 16))
        self.speaker.set_rate(self.rate_var.get())
        self._show_rate()

        ttk.Label(row2, text="Volume", style="Strip.TLabel").pack(side="left")
        self.vol_var = tk.IntVar(value=int(self.settings.get("volume", 100)))
        vol = ttk.Scale(row2, from_=0, to=100, orient="horizontal", length=110,
                        command=self.on_volume)
        vol.set(self.vol_var.get())
        vol.pack(side="left", padx=6)
        self.speaker.set_volume(self.vol_var.get())

        # Text area
        body = ttk.Frame(self.root)
        body.pack(fill="both", expand=True)
        self.text = tk.Text(body, wrap="word", bg=self.PAGE, fg=self.INK, relief="flat",
                            padx=40, pady=24, spacing2=4, spacing3=10, cursor="hand2",
                            font=("Segoe UI", self.font_size), insertwidth=0,
                            selectbackground="#cfd6ff", selectforeground=self.INK)
        sb = ttk.Scrollbar(body, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.text.pack(side="left", fill="both", expand=True)
        self._style_tags()
        self.text.bind("<Button-1>", self.on_click)
        self.text.bind("<Key>", self._readonly)   # selectable and copyable, not editable

        self.show_welcome()

        self.status = tk.StringVar(value="Open a PDF to begin.")
        bar = ttk.Label(self.root, textvariable=self.status, anchor="w", padding=(10, 4),
                        font=("Segoe UI", 9))
        bar.pack(fill="x", side="bottom")

    def _style_tags(self):
        fs = self.font_size
        self.text.tag_configure("pagehead", font=("Segoe UI", max(9, fs - 3), "bold"),
                                foreground=self.MUTED, spacing1=18, spacing3=6)
        self.text.tag_configure("hl", background=self.MARK)
        self.text.tag_configure("muted", foreground=self.MUTED)
        self.text.tag_configure("title", font=("Segoe UI", fs + 9, "bold"), spacing3=12)
        self.text.tag_raise("sel")

    def _readonly(self, e):
        ctrl = e.state & 0x4
        if ctrl and e.keysym.lower() in ("c", "a"):
            if e.keysym.lower() == "a":
                self.text.tag_add("sel", "1.0", "end")
                return "break"
            return None
        if e.keysym in ("Up", "Down", "Prior", "Next", "Home", "End"):
            return None
        return "break"

    def _bind_keys(self):
        r = self.root
        r.bind("<space>", lambda e: (self.toggle_play(), "break")[1])
        r.bind("<Right>", lambda e: self.next_sentence())
        r.bind("<Left>", lambda e: self.prev_sentence())
        r.bind("<Escape>", lambda e: self.stop())
        r.bind("<Control-o>", lambda e: self.choose_file())
        r.bind("<Control-plus>", lambda e: self.zoom(1))
        r.bind("<Control-equal>", lambda e: self.zoom(1))
        r.bind("<Control-minus>", lambda e: self.zoom(-1))

    def show_welcome(self):
        t = self.text
        t.delete("1.0", "end")
        t.insert("end", "PDF Reader\n", "title")
        t.insert("end",
                 "Click “Open PDF…” (or press Ctrl+O) and choose a file.\n\n"
                 "Then press Play, or click any sentence to start reading from there.\n\n"
                 "Keys:  Space = play / pause    ← → = previous / next sentence    "
                 "Esc = stop    Ctrl + / Ctrl − = text size\n\n"
                 "Tip: you can also right-click a PDF in File Explorer → Open with → PDF Reader, "
                 "or drag a PDF onto the PDF Reader icon.", "muted")

    def set_status(self, msg):
        self.status.set(msg)

    # ---------------- opening ----------------
    def choose_file(self):
        start = self.settings.get("last_dir") or os.path.expanduser("~")
        path = filedialog.askopenfilename(title="Open a PDF", initialdir=start,
                                          filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if path:
            self.open_pdf(path)

    def open_pdf(self, path, password=None):
        self.stop()
        self.set_status("Opening " + os.path.basename(path) + " …")
        self.root.config(cursor="watch")

        def work():
            try:
                pages = extract_pages(path, password)
                self.root.after(0, lambda: self._loaded(path, pages))
            except PermissionError:
                self.root.after(0, lambda: self._ask_password(path))
            except Exception as e:
                msg = str(e)
                self.root.after(0, lambda: self._load_failed(msg))
        threading.Thread(target=work, daemon=True).start()

    def _ask_password(self, path):
        self.root.config(cursor="")
        win = tk.Toplevel(self.root)
        win.title("Password needed")
        win.transient(self.root)
        win.grab_set()
        ttk.Label(win, text="This PDF is password-protected. Enter the password:",
                  padding=12).pack()
        pw = ttk.Entry(win, show="•", width=30)
        pw.pack(padx=12)
        pw.focus_set()

        def ok(_=None):
            p = pw.get()
            win.destroy()
            self.open_pdf(path, p)
        ttk.Button(win, text="Open", command=ok).pack(pady=12)
        win.bind("<Return>", ok)

    def _load_failed(self, msg):
        self.root.config(cursor="")
        self.set_status("Could not open the file.")
        messagebox.showerror(APP_NAME, "Sorry, that PDF could not be opened.\n\n" + msg)

    def _loaded(self, path, pages):
        self.root.config(cursor="")
        self.path = path
        self.settings["last_dir"] = os.path.dirname(path)
        self.root.title(f"{os.path.basename(path)} – {APP_NAME}")

        t = self.text
        t.delete("1.0", "end")
        self.sentences, self.starts, self.page_starts = [], [], {}
        pos = 0

        def put(s, tag=None):
            nonlocal pos
            t.insert("end", s, tag) if tag else t.insert("end", s)
            pos += len(s)

        for pno, paras in enumerate(pages, start=1):
            put(f"Page {pno}\n", "pagehead")
            self.page_starts[pno] = len(self.sentences)
            if not paras:
                put("(no readable text on this page)\n", "muted")
                continue
            for para in paras:
                for i, s in enumerate(split_sentences(para)):
                    if i:
                        put(" ")
                    self.sentences.append({"text": s, "page": pno, "start": pos, "end": pos + len(s)})
                    self.starts.append(pos)
                    put(s)
                put("\n")

        n = len(pages)
        self.page_box["values"] = [f"{i} of {n}" for i in range(1, n + 1)]
        if n:
            self.page_box.current(0)

        if not self.sentences:
            self.set_status(f"{n} pages, but no readable text.")
            messagebox.showinfo(
                APP_NAME,
                "This PDF doesn't contain text the reader can use. It is probably scanned "
                "pages (pictures of text), which would need OCR (text recognition) first.")
            return

        # resume where you left off in this file
        resume = self.settings.get("positions", {}).get(os.path.abspath(path), 0)
        self.idx = resume if 0 < resume < len(self.sentences) else 0
        self.highlight(self.idx, scroll=True)
        extra = "  ·  resumed where you left off" if self.idx else ""
        self.set_status(f"{n} pages, {len(self.sentences)} sentences. Press Play (Space).{extra}")

    # ---------------- reading ----------------
    def toggle_play(self):
        if not self.sentences:
            if not self.path:
                self.choose_file()
            return
        if self.state == "playing":
            self.speaker.pause()
            self.state = "paused"
            self._sync_buttons()
            self.update_status()
        elif self.state == "paused":
            self.speaker.resume()
            self.state = "playing"
            self._sync_buttons()
            self._poll()
        else:
            self.speak_from(self.idx)

    def speak_from(self, i):
        if not self.sentences:
            return
        if i >= len(self.sentences):
            self.stop()
            self.idx = 0
            self.set_status("Finished reading.")
            return
        self.idx = max(0, i)
        if self.state == "paused":
            self.speaker.resume()
        self.state = "playing"
        self._sync_buttons()
        self.highlight(self.idx, scroll=True)
        self.update_status()
        self.speaker.speak(self.sentences[self.idx]["text"])
        if self.poll_id:
            self.root.after_cancel(self.poll_id)
        self.poll_id = self.root.after(120, self._poll)

    def _poll(self):
        if self.poll_id:
            self.root.after_cancel(self.poll_id)
            self.poll_id = None
        if self.state != "playing":
            return
        if self.speaker.is_done():
            self.speak_from(self.idx + 1)
        else:
            self.poll_id = self.root.after(80, self._poll)

    def stop(self):
        if self.poll_id:
            self.root.after_cancel(self.poll_id)
            self.poll_id = None
        if self.state == "paused":
            self.speaker.resume()
        self.speaker.stop()
        self.state = "stopped"
        self._sync_buttons()
        if self.sentences:
            self.update_status()

    def _restart_if_playing(self):
        if self.state in ("playing", "paused"):
            self.speak_from(self.idx)

    def next_sentence(self):
        if not self.sentences:
            return
        i = min(len(self.sentences) - 1, self.idx + 1)
        if self.state == "playing":
            self.speak_from(i)
        else:
            self.stop()
            self.idx = i
            self.highlight(i, scroll=True)
            self.update_status()

    def prev_sentence(self):
        if not self.sentences:
            return
        i = max(0, self.idx - 1)
        if self.state == "playing":
            self.speak_from(i)
        else:
            self.stop()
            self.idx = i
            self.highlight(i, scroll=True)
            self.update_status()

    def goto_page(self, _=None):
        if not self.sentences:
            return
        p = self.page_box.current() + 1
        i = self.page_starts.get(p, 0)
        if i >= len(self.sentences):
            return
        if self.state == "playing":
            self.speak_from(i)
        else:
            self.stop()
            self.idx = i
            self.highlight(i, scroll=True)
            self.update_status()

    def on_click(self, event):
        if not self.sentences:
            return
        if self.text.tag_ranges("sel"):
            self.text.tag_remove("sel", "1.0", "end")
        index = self.text.index(f"@{event.x},{event.y}")
        off = self._chars_to(index)
        i = bisect.bisect_right(self.starts, off) - 1
        if i < 0:
            return
        s = self.sentences[i]
        if off > s["end"] + 1:      # clicked on a page heading or blank
            return
        self.speak_from(i)

    def _chars_to(self, index):
        """Number of characters from the start of the text to index."""
        if not self.text.compare(index, ">", "1.0"):
            return 0
        n = self.text.count("1.0", index, "chars")
        if isinstance(n, tuple):
            n = n[0]
        return int(n or 0)

    # ---------------- display ----------------
    def highlight(self, i, scroll=False):
        t = self.text
        t.tag_remove("hl", "1.0", "end")
        if not (0 <= i < len(self.sentences)):
            return
        s = self.sentences[i]
        a, b = f"1.0+{s['start']}c", f"1.0+{s['end']}c"
        t.tag_add("hl", a, b)
        if scroll:
            info = t.dlineinfo(a)
            h = t.winfo_height()
            if info is None or info[1] > h * 0.75 or info[1] < 10:
                t.see(a)
                info = t.dlineinfo(a)
                if info is not None and info[1] > h * 0.4:
                    t.yview_scroll(int((info[1] - h * 0.3) / max(1, info[3])), "units")

    def update_status(self):
        if not self.sentences:
            return
        s = self.sentences[self.idx]
        state = {"playing": "Reading", "paused": "Paused", "stopped": "Ready"}[self.state]
        self.set_status(f"{state}  ·  sentence {self.idx + 1} of {len(self.sentences)}  ·  "
                        f"page {s['page']} of {len(self.page_box['values'])}")
        self.page_box.current(s["page"] - 1)

    def _sync_buttons(self):
        self.play_btn.configure(text="❚❚  Pause" if self.state == "playing" else "▶  Play")

    def zoom(self, d):
        self.font_size = max(9, min(28, self.font_size + d))
        self.text.configure(font=("Segoe UI", self.font_size))
        self._style_tags()
        self.settings["font_size"] = self.font_size

    # ---------------- settings controls ----------------
    def _show_rate(self):
        r = self.rate_var.get()
        self.rate_lbl.configure(text=("normal" if r == 0 else f"{r:+d}"))

    def on_rate(self, val):
        r = int(round(float(val)))
        if r == self.rate_var.get():
            return
        self.rate_var.set(r)
        self._show_rate()
        self.speaker.set_rate(r)
        self.settings["rate"] = r
        if self.state == "playing":
            if self.poll_id:
                self.root.after_cancel(self.poll_id)
            self.poll_id = self.root.after(250, lambda: self.speak_from(self.idx))

    def on_volume(self, val):
        v = int(round(float(val)))
        self.vol_var.set(v)
        self.speaker.set_volume(v)
        self.settings["volume"] = v

    def on_voice(self, _=None):
        self.speaker.set_voice(self.voice_box.current())
        self.settings["voice"] = self.voice_var.get()
        self._restart_if_playing()

    # ---------------- audio export ----------------
    def save_audio(self):
        if not self.sentences:
            messagebox.showinfo(APP_NAME, "Open a PDF first.")
            return
        if not self.speaker.ok:
            messagebox.showwarning(APP_NAME, "Windows speech is not available on this computer.")
            return
        base = os.path.splitext(os.path.basename(self.path))[0]
        out = filedialog.asksaveasfilename(title="Save the reading as audio", defaultextension=".wav",
                                           initialfile=base + ".wav", filetypes=[("WAV audio", "*.wav")])
        if not out:
            return
        text = "\n".join(s["text"] for s in self.sentences)
        self.set_status("Saving audio… this can take a while for long PDFs.")

        def done(err):
            def ui():
                if err:
                    self.set_status("Saving audio failed.")
                    messagebox.showerror(APP_NAME, "Could not save the audio.\n\n" + err)
                else:
                    self.set_status("Saved audio to " + out)
            self.root.after(0, ui)
        save_wav(text, out, self.voice_box.current(), self.rate_var.get(), self.vol_var.get(), done)

    # ---------------- closing ----------------
    def on_close(self):
        try:
            if self.path and self.sentences:
                self.settings.setdefault("positions", {})[os.path.abspath(self.path)] = self.idx
                # keep the list short
                pos = self.settings["positions"]
                if len(pos) > 50:
                    for k in list(pos)[:-50]:
                        pos.pop(k, None)
            self.settings["geometry"] = self.root.geometry()
            save_settings(self.settings)
            self.speaker.stop()
        finally:
            self.root.destroy()


def main():
    # Sharper text on high-DPI screens
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    start = sys.argv[1] if len(sys.argv) > 1 and os.path.isfile(sys.argv[1]) else None
    root = tk.Tk()
    App(root, start)
    root.mainloop()


if __name__ == "__main__":
    main()
