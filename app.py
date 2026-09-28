"""File Convertor - drag & drop file format converter.

Supported types: PDF, DOCX, PPTX, PNG, JPEG, HEIC.
"""

import os
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk
from tkinterdnd2 import DND_FILES, TkinterDnD

from converter import FORMATS, ConversionError, convert_file, detect_format, soffice_available

APP_TITLE = "File Convertor"

# (label shown in dropdown, internal format key)
TARGET_CHOICES = [
    ("PDF (.pdf)", "pdf"),
    ("Word (.docx)", "docx"),
    ("PowerPoint (.pptx)", "pptx"),
    ("PNG (.png)", "png"),
    ("JPEG (.jpg)", "jpeg"),
    ("HEIC (.heic)", "heic"),
]


class ConvertorApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        # Root is a TkinterDnD.Tk (not ctk.CTk), so set its background
        # explicitly to match the dark theme.
        self.configure(background="#242424")

        self.title(APP_TITLE)
        self.geometry("920x680")
        self.minsize(760, 560)

        self.files: list[Path] = []
        self.target = "pdf"
        self.output_dir = str(Path.home() / "FileConvertor output")
        self._msg_queue: queue.Queue = queue.Queue()
        self._converting = False

        self._build_ui()
        self._poll_queue()

    # ------------------------------------------------------------------ UI

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # Header
        header = ctk.CTkLabel(self, text=APP_TITLE, font=ctk.CTkFont(size=26, weight="bold"))
        header.grid(row=0, column=0, pady=(18, 4), sticky="ew")
        sub = ctk.CTkLabel(
            self,
            text="Drop files below (or choose them), pick a target format, hit Convert.",
            font=ctk.CTkFont(size=13),
            text_color="gray70",
        )
        sub.grid(row=1, column=0, pady=(0, 10), sticky="ew")

        # Drop zone
        self.drop_frame = ctk.CTkFrame(self, height=150, corner_radius=12,
                                       border_width=2, border_color="gray40",
                                       fg_color="transparent")
        self.drop_frame.grid(row=2, column=0, padx=20, pady=6, sticky="nsew")
        self.drop_frame.grid_columnconfigure(0, weight=1)
        self.drop_frame.grid_rowconfigure(0, weight=1)
        self.drop_frame.grid_rowconfigure(3, weight=1)
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind("<<Drop>>", self._on_drop)

        drop_label = ctk.CTkLabel(
            self.drop_frame,
            text="Drag & drop files here\nor",
            font=ctk.CTkFont(size=15),
            text_color="gray70",
            justify="center",
        )
        drop_label.grid(row=1, column=0, pady=(10, 2))

        choose_btn = ctk.CTkButton(self.drop_frame, text="Choose files…",
                                   command=self._choose_files, width=180)
        choose_btn.grid(row=2, column=0, pady=(2, 10))

        # File list
        list_label = ctk.CTkLabel(self, text="Files (0)", font=ctk.CTkFont(size=14, weight="bold"),
                                  anchor="w")
        list_label.grid(row=3, column=0, padx=24, pady=(10, 2), sticky="ew")
        self.list_label = list_label

        self.file_listbox = tk.Listbox(self, height=8, font=("TkDefaultFont", 12),
                                       bg="#2b2b2b", fg="white",
                                       selectbackground="#1f6aa5",
                                       highlightthickness=0, borderwidth=0)
        self.file_listbox.grid(row=4, column=0, padx=20, sticky="nsew")

        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.grid(row=5, column=0, padx=20, pady=6, sticky="ew")
        ctk.CTkButton(btn_row, text="Remove selected", width=150,
                      fg_color="transparent", border_width=1,
                      command=self._remove_selected).pack(side="left", padx=(0, 8))
        ctk.CTkButton(btn_row, text="Clear all", width=110,
                      fg_color="transparent", border_width=1,
                      command=self._clear_files).pack(side="left")

        # Controls: target dropdown + output folder
        controls = ctk.CTkFrame(self, fg_color="transparent")
        controls.grid(row=6, column=0, padx=20, pady=8, sticky="ew")
        controls.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(controls, text="Convert to:", font=ctk.CTkFont(size=14)).grid(
            row=0, column=0, padx=(0, 8), sticky="w")
        self.target_menu = ctk.CTkOptionMenu(
            controls, values=[label for label, _ in TARGET_CHOICES],
            command=self._on_target_change, width=200)
        self.target_menu.set(TARGET_CHOICES[0][0])
        self.target_menu.grid(row=0, column=1, sticky="w")

        ctk.CTkLabel(controls, text="Save to:", font=ctk.CTkFont(size=14)).grid(
            row=1, column=0, padx=(0, 8), pady=(8, 0), sticky="w")
        self.output_entry = ctk.CTkEntry(controls, width=420)
        self.output_entry.insert(0, self.output_dir)
        self.output_entry.grid(row=1, column=1, pady=(8, 0), sticky="ew", padx=(0, 8))
        ctk.CTkButton(controls, text="Browse…", width=100,
                      command=self._choose_output).grid(row=1, column=2, pady=(8, 0))

        # Convert button + progress
        self.convert_btn = ctk.CTkButton(self, text="Convert", height=44,
                                         font=ctk.CTkFont(size=16, weight="bold"),
                                         command=self._start_convert)
        self.convert_btn.grid(row=7, column=0, padx=20, pady=8, sticky="ew")

        self.progress = ctk.CTkProgressBar(self, height=12)
        self.progress.grid(row=8, column=0, padx=20, pady=(0, 4), sticky="ew")
        self.progress.set(0)

        self.status = ctk.CTkLabel(self, text="Ready.", font=ctk.CTkFont(size=12),
                                   text_color="gray70", anchor="w", wraplength=860,
                                   justify="left")
        self.status.grid(row=9, column=0, padx=24, pady=(0, 14), sticky="ew")

        if not soffice_available():
            self._set_status("Note: LibreOffice not found — DOCX/PPTX conversions "
                             "need it installed.", error=False)

    # -------------------------------------------------------------- actions

    def _on_target_change(self, label: str):
        for text, key in TARGET_CHOICES:
            if text == label:
                self.target = key
                break

    def _on_drop(self, event):
        # tkinterdnd2 gives a Tcl list of paths, possibly wrapped in braces
        raw = self.tk.splitlist(event.data)
        self._add_files([Path(p) for p in raw if os.path.isfile(p)])

    def _choose_files(self):
        paths = filedialog.askopenfilenames(title="Choose files to convert")
        self._add_files([Path(p) for p in paths])

    def _add_files(self, paths: list[Path]):
        added = 0
        skipped = 0
        for p in paths:
            if p in self.files:
                continue
            if detect_format(p) is None:
                skipped += 1
                continue
            self.files.append(p)
            added += 1
        self._refresh_list()
        if skipped:
            self._set_status(f"Added {added} file(s); skipped {skipped} unsupported file(s).")
        elif added:
            self._set_status(f"Added {added} file(s).")

    def _remove_selected(self):
        for idx in reversed(self.file_listbox.curselection()):
            del self.files[idx]
        self._refresh_list()

    def _clear_files(self):
        self.files.clear()
        self._refresh_list()

    def _refresh_list(self):
        self.file_listbox.delete(0, tk.END)
        for p in self.files:
            fmt = detect_format(p) or "?"
            self.file_listbox.insert(tk.END, f"{p.name}   [{fmt.upper()}]")
        self.list_label.configure(text=f"Files ({len(self.files)})")

    def _choose_output(self):
        d = filedialog.askdirectory(title="Choose output folder",
                                    initialdir=self.output_entry.get())
        if d:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, d)

    def _set_status(self, text: str, error: bool = False):
        self.status.configure(text=text,
                              text_color="#ff7b72" if error else "gray70")

    # ------------------------------------------------------------ conversion

    def _start_convert(self):
        if self._converting:
            return
        if not self.files:
            self._set_status("Add some files first.", error=True)
            return
        out = self.output_entry.get().strip()
        if not out:
            self._set_status("Choose an output folder.", error=True)
            return
        self._converting = True
        self.convert_btn.configure(state="disabled", text="Converting…")
        self.progress.set(0)
        t = threading.Thread(target=self._convert_worker,
                             args=(list(self.files), self.target, out),
                             daemon=True)
        t.start()

    def _convert_worker(self, files: list[Path], target: str, out_dir: str):
        total = len(files)
        ok, failed = 0, 0
        for i, path in enumerate(files, start=1):
            try:
                created = convert_file(path, target, out_dir)
                names = ", ".join(p.name for p in created)
                self._msg_queue.put(("log", f"✓ {path.name} → {names}"))
                ok += 1
            except ConversionError as e:
                self._msg_queue.put(("log", f"✗ {path.name}: {e}"))
                failed += 1
            except Exception as e:  # never kill the batch on one bad file
                self._msg_queue.put(("log", f"✗ {path.name}: unexpected error: {e}"))
                failed += 1
            self._msg_queue.put(("progress", i / total))
        self._msg_queue.put(("done", (ok, failed, out_dir)))

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self._msg_queue.get_nowait()
                if kind == "log":
                    self._set_status(payload, error=payload.startswith("✗"))
                elif kind == "progress":
                    self.progress.set(payload)
                elif kind == "done":
                    ok, failed, out_dir = payload
                    self._converting = False
                    self.convert_btn.configure(state="normal", text="Convert")
                    if failed == 0:
                        self._set_status(f"Done — {ok} file(s) converted to {out_dir}")
                    else:
                        self._set_status(
                            f"Finished with {failed} error(s), {ok} converted. "
                            f"Output: {out_dir}", error=True)
                    if ok and messagebox.askyesno("Done",
                            f"Converted {ok} file(s) to {out_dir}.\nOpen the folder?"):
                        self._open_folder(out_dir)
        except queue.Empty:
            pass
        self.after(120, self._poll_queue)

    def _open_folder(self, path: str):
        import subprocess, sys
        if sys.platform == "darwin":
            subprocess.run(["open", path])
        elif sys.platform.startswith("win"):
            os.startfile(path)  # noqa: S606
        else:
            subprocess.run(["xdg-open", path])


if __name__ == "__main__":
    app = ConvertorApp()
    app.mainloop()
