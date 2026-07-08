"""
PDF Folder Merger
-----------------
Select a folder, scan it (and all subfolders) for PDF files, merge them
into a single PDF, and save the result in the selected folder under a
user-supplied name.

Requirements:
    pip install pypdf

Run:
    python pdf_merger.py
"""

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    raise SystemExit(
        "pypdf is not installed.\n\nInstall it with:\n    pip install pypdf"
    )


class PDFMergerApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PDF Folder Merger")
        self.root.geometry("640x560")
        self.root.minsize(560, 480)

        self.source_folder: Path | None = None
        self.pdf_files: list[Path] = []
        self.merging = False

        self._build_ui()

    # ------------------------------------------------------------------ UI

    def _build_ui(self):
        pad = {"padx": 10, "pady": 5}

        # --- Folder selection row
        folder_frame = ttk.Frame(self.root)
        folder_frame.pack(fill="x", **pad)

        ttk.Button(
            folder_frame, text="Select Folder…", command=self.select_folder
        ).pack(side="left")

        self.folder_label = ttk.Label(
            folder_frame, text="No folder selected", foreground="grey"
        )
        self.folder_label.pack(side="left", padx=10)

        # --- File list
        list_frame = ttk.LabelFrame(self.root, text="PDF files found (merge order)")
        list_frame.pack(fill="both", expand=True, **pad)

        self.file_list = tk.Listbox(list_frame, activestyle="none")
        scrollbar = ttk.Scrollbar(
            list_frame, orient="vertical", command=self.file_list.yview
        )
        self.file_list.configure(yscrollcommand=scrollbar.set)
        self.file_list.pack(side="left", fill="both", expand=True, padx=(5, 0), pady=5)
        scrollbar.pack(side="right", fill="y", pady=5)

        # --- Output name row
        name_frame = ttk.Frame(self.root)
        name_frame.pack(fill="x", **pad)

        ttk.Label(name_frame, text="Output file name:").pack(side="left")
        self.name_var = tk.StringVar(value="merged")
        self.name_entry = ttk.Entry(name_frame, textvariable=self.name_var)
        self.name_entry.pack(side="left", fill="x", expand=True, padx=(8, 4))
        ttk.Label(name_frame, text=".pdf").pack(side="left")

        # --- Merge button + progress
        action_frame = ttk.Frame(self.root)
        action_frame.pack(fill="x", **pad)

        self.merge_button = ttk.Button(
            action_frame, text="Merge PDFs", command=self.start_merge
        )
        self.merge_button.pack(side="left")

        self.progress = ttk.Progressbar(action_frame, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=10)

        # --- Status log
        log_frame = ttk.LabelFrame(self.root, text="Log")
        log_frame.pack(fill="both", expand=False, **pad)

        self.log_box = tk.Text(log_frame, height=7, state="disabled", wrap="word")
        log_scroll = ttk.Scrollbar(
            log_frame, orient="vertical", command=self.log_box.yview
        )
        self.log_box.configure(yscrollcommand=log_scroll.set)
        self.log_box.pack(side="left", fill="both", expand=True, padx=(5, 0), pady=5)
        log_scroll.pack(side="right", fill="y", pady=5)

    def log(self, message: str):
        """Thread-safe append to the log box."""
        def _append():
            self.log_box.configure(state="normal")
            self.log_box.insert("end", message + "\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
        self.root.after(0, _append)

    # ----------------------------------------------------------- Scanning

    def select_folder(self):
        if self.merging:
            return
        folder = filedialog.askdirectory(title="Select folder containing PDFs")
        if not folder:
            return

        self.source_folder = Path(folder)
        self.folder_label.configure(text=str(self.source_folder), foreground="black")
        self.scan_folder()

    def scan_folder(self):
        """Find all PDFs in the folder and subfolders, sorted by relative path."""
        assert self.source_folder is not None
        self.pdf_files = sorted(
            (
                p for p in self.source_folder.rglob("*")
                if p.is_file() and p.suffix.lower() == ".pdf"
            ),
            key=lambda p: str(p.relative_to(self.source_folder)).lower(),
        )

        self.file_list.delete(0, "end")
        for p in self.pdf_files:
            self.file_list.insert("end", str(p.relative_to(self.source_folder)))

        self.log(f"Found {len(self.pdf_files)} PDF file(s) in {self.source_folder}")

    # ------------------------------------------------------------ Merging

    def start_merge(self):
        if self.merging:
            return
        if not self.source_folder or not self.pdf_files:
            messagebox.showwarning(
                "Nothing to merge", "Select a folder containing PDF files first."
            )
            return

        name = self.name_var.get().strip()
        if not name:
            messagebox.showwarning("No name", "Enter a name for the output file.")
            return

        # Strip .pdf if the user typed it, and block characters Windows rejects.
        if name.lower().endswith(".pdf"):
            name = name[:-4]
        if any(c in name for c in '\\/:*?"<>|'):
            messagebox.showwarning(
                "Invalid name", 'The file name cannot contain \\ / : * ? " < > |'
            )
            return

        output_path = self.source_folder / f"{name}.pdf"

        # Exclude the output file itself (e.g. re-running with the same name).
        sources = [p for p in self.pdf_files if p.resolve() != output_path.resolve()]
        if not sources:
            messagebox.showwarning(
                "Nothing to merge",
                "The only PDF found is the output file from a previous run.",
            )
            return

        if output_path.exists():
            if not messagebox.askyesno(
                "Overwrite?", f"{output_path.name} already exists. Overwrite it?"
            ):
                return

        self.merging = True
        self.merge_button.configure(state="disabled")
        self.progress.configure(maximum=len(sources), value=0)

        thread = threading.Thread(
            target=self._merge_worker, args=(sources, output_path), daemon=True
        )
        thread.start()

    def _merge_worker(self, sources: list[Path], output_path: Path):
        writer = PdfWriter()
        merged, skipped = 0, 0

        for i, pdf_path in enumerate(sources, start=1):
            rel = pdf_path.relative_to(self.source_folder)
            try:
                reader = PdfReader(pdf_path)
                if reader.is_encrypted:
                    # Try an empty password (some PDFs are "encrypted" but open freely)
                    try:
                        reader.decrypt("")
                    except Exception:
                        raise ValueError("password-protected")
                for page in reader.pages:
                    writer.add_page(page)
                merged += 1
                self.log(f"Added: {rel}")
            except Exception as exc:
                skipped += 1
                self.log(f"SKIPPED {rel} ({exc})")

            self.root.after(0, lambda v=i: self.progress.configure(value=v))

        if merged == 0:
            self.log("No PDFs could be merged — nothing was saved.")
            self.root.after(0, self._merge_done, None, merged, skipped)
            return

        try:
            with open(output_path, "wb") as f:
                writer.write(f)
            self.log(f"Saved: {output_path}")
        except Exception as exc:
            self.log(f"ERROR saving output: {exc}")
            output_path = None

        self.root.after(0, self._merge_done, output_path, merged, skipped)

    def _merge_done(self, output_path, merged: int, skipped: int):
        self.merging = False
        self.merge_button.configure(state="normal")

        if output_path:
            msg = f"Merged {merged} file(s) into:\n{output_path}"
            if skipped:
                msg += f"\n\n{skipped} file(s) were skipped — see the log."
            messagebox.showinfo("Done", msg)
        else:
            messagebox.showerror(
                "Merge failed",
                "No output file was created. Check the log for details.",
            )


def main():
    root = tk.Tk()
    # Use a slightly nicer theme where available
    try:
        ttk.Style().theme_use("vista")
    except tk.TclError:
        pass
    PDFMergerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
