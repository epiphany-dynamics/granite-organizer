"""
Granite Net File Organizer
Sorts inspection PDFs and videos into segment folders by matching
manhole IDs in filenames to an Excel prep sheet.
"""

import json
import os
import re
import shutil
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import openpyxl

MH_PATTERN = re.compile(r"(\d{3}-\d{2}-\d{3})")
SEGMENT_PATTERN = re.compile(r"^[A-Z]-\d+$")
PDF_EXTENSIONS = {".pdf"}
CONFIG_FILE = os.path.join(
    os.path.dirname(os.path.abspath(sys.argv[0])), "organizer_config.json"
)


def load_config():
    try:
        with open(CONFIG_FILE) as f:
            return json.load(f)
    except Exception:
        return {}


def save_config(data):
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump(data, f)
    except Exception:
        pass


def build_lookup(excel_path):
    """Scan all sheets for rows with a segment ID and at least 2 manhole IDs.
    Returns (lookup_dict, set_of_segment_ids)."""
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    lookup = {}
    segments = set()

    for ws in wb.worksheets:
        for row in ws.iter_rows(min_row=1, values_only=True):
            cells = [str(c).strip() if c is not None else "" for c in row]
            segment_id = None
            mh_ids = []

            for cell in cells:
                if not segment_id and SEGMENT_PATTERN.match(cell):
                    segment_id = cell
                mh_ids.extend(MH_PATTERN.findall(cell))

            if segment_id and len(mh_ids) >= 2:
                unique = list(dict.fromkeys(mh_ids))
                if len(unique) >= 2:
                    lookup[(unique[0], unique[1])] = segment_id
                    lookup[(unique[1], unique[0])] = segment_id
                    segments.add(segment_id)

    wb.close()
    return lookup, segments


def match_file(filename, lookup):
    """Try to match a filename to a segment via its manhole IDs."""
    mh_ids = MH_PATTERN.findall(filename)
    if len(mh_ids) < 2:
        return None, mh_ids
    for i in range(len(mh_ids)):
        for j in range(i + 1, len(mh_ids)):
            seg = lookup.get((mh_ids[i], mh_ids[j])) or lookup.get(
                (mh_ids[j], mh_ids[i])
            )
            if seg:
                return seg, mh_ids
    return None, mh_ids


def organize_files(excel_path, pdf_folder, video_folder, output_folder, move, log):
    log("=" * 55)
    log("  Granite Net File Organizer")
    log("=" * 55)

    log(f"\nReading: {os.path.basename(excel_path)}")
    try:
        lookup, segments = build_lookup(excel_path)
    except Exception as e:
        log(f"ERROR reading spreadsheet: {e}")
        return
    log(f"Found {len(segments)} segments with MH mappings\n")

    action_word = "Moving" if move else "Copying"
    file_op = shutil.move if move else shutil.copy2
    matched = 0
    unmatched = []

    sources = []
    if pdf_folder and os.path.isdir(pdf_folder):
        sources.append(("PDF", pdf_folder, True))
    if video_folder and os.path.isdir(video_folder):
        sources.append(("Video", video_folder, False))

    for label, folder, pdf_only in sources:
        files = [
            f
            for f in os.listdir(folder)
            if os.path.isfile(os.path.join(folder, f))
        ]
        if pdf_only:
            files = [f for f in files if os.path.splitext(f)[1].lower() in PDF_EXTENSIONS]
        else:
            files = [f for f in files if os.path.splitext(f)[1].lower() not in {".xlsx", ".xls"}]

        log(f"--- {label}s: {len(files)} files in {folder} ---")
        for idx, filename in enumerate(files, 1):
            seg, mh_ids = match_file(filename, lookup)
            if seg:
                dest_dir = os.path.join(output_folder, seg)
                os.makedirs(dest_dir, exist_ok=True)
                dest = os.path.join(dest_dir, filename)
                if os.path.exists(dest):
                    base, ext = os.path.splitext(filename)
                    n = 1
                    while os.path.exists(dest):
                        dest = os.path.join(dest_dir, f"{base}_{n}{ext}")
                        n += 1
                try:
                    file_op(os.path.join(folder, filename), dest)
                    log(f"  {idx}/{len(files)}  {seg}/  <-  {filename}")
                    matched += 1
                except Exception as e:
                    log(f"  ERROR on {filename}: {e}")
            else:
                reason = f"MH IDs {mh_ids}" if mh_ids else "no MH IDs found"
                unmatched.append((label, filename, reason))
                log(f"  {idx}/{len(files)}  NO MATCH  {filename}  ({reason})")
        log("")

    log("=" * 55)
    log(f"  DONE  —  {matched} files organized, {len(unmatched)} unmatched")
    log("=" * 55)
    if unmatched:
        log("\nUnmatched files:")
        for label, fn, reason in unmatched:
            log(f"  [{label}] {fn}  ({reason})")


class App:
    def __init__(self, root):
        self.root = root
        root.title("Granite Net File Organizer")
        root.geometry("750x580")
        root.resizable(True, True)

        cfg = load_config()
        self.excel_var = tk.StringVar(value=cfg.get("excel", ""))
        self.pdf_var = tk.StringVar(value=cfg.get("pdf_folder", ""))
        self.video_var = tk.StringVar(value=cfg.get("video_folder", ""))
        self.output_var = tk.StringVar(value=cfg.get("output_folder", ""))
        self.move_var = tk.BooleanVar(value=cfg.get("move", False))

        self._build()

    def _build(self):
        m = ttk.Frame(self.root, padding=12)
        m.pack(fill=tk.BOTH, expand=True)

        ttk.Label(m, text="Granite Net File Organizer", font=("Segoe UI", 14, "bold")).pack(pady=(0, 8))
        ttk.Label(m, text="Match inspection PDFs & videos to segment folders using your prep sheet.").pack(pady=(0, 12))

        self._row(m, "Excel Spreadsheet:", self.excel_var, file=True)
        self._row(m, "PDF Folder:", self.pdf_var)
        self._row(m, "Video Folder:", self.video_var)
        self._row(m, "Output Folder:", self.output_var)

        opts = ttk.Frame(m)
        opts.pack(fill=tk.X, pady=6)
        ttk.Checkbutton(opts, text="Move files instead of copy", variable=self.move_var).pack(side=tk.LEFT)

        self.btn = ttk.Button(m, text="Organize Files", command=self._run)
        self.btn.pack(pady=8)

        self.log = tk.Text(m, height=16, font=("Consolas", 9), wrap=tk.WORD)
        self.log.pack(fill=tk.BOTH, expand=True)
        sb = ttk.Scrollbar(m, command=self.log.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.log.configure(yscrollcommand=sb.set)

    def _row(self, parent, label, var, file=False):
        f = ttk.Frame(parent)
        f.pack(fill=tk.X, pady=2)
        ttk.Label(f, text=label, width=20, anchor=tk.W).pack(side=tk.LEFT)
        ttk.Entry(f, textvariable=var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
        if file:
            cmd = lambda: var.set(
                filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")])
                or var.get()
            )
        else:
            cmd = lambda: var.set(filedialog.askdirectory() or var.get())
        ttk.Button(f, text="Browse", command=cmd, width=8).pack(side=tk.RIGHT)

    def _log_msg(self, msg):
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)
        self.root.update_idletasks()

    def _run(self):
        excel = self.excel_var.get().strip()
        pdfs = self.pdf_var.get().strip()
        videos = self.video_var.get().strip()
        output = self.output_var.get().strip()

        if not excel or not os.path.isfile(excel):
            messagebox.showerror("Error", "Select a valid Excel spreadsheet.")
            return
        if not output:
            messagebox.showerror("Error", "Select an output folder.")
            return
        if not pdfs and not videos:
            messagebox.showerror("Error", "Select at least one source folder (PDFs or Videos).")
            return

        save_config({
            "excel": excel,
            "pdf_folder": pdfs,
            "video_folder": videos,
            "output_folder": output,
            "move": self.move_var.get(),
        })

        self.log.delete("1.0", tk.END)
        self.btn.configure(state=tk.DISABLED)

        def worker():
            try:
                organize_files(excel, pdfs, videos, output, self.move_var.get(), self._log_msg)
            except Exception as e:
                self._log_msg(f"\nFATAL ERROR: {e}")
            finally:
                self.btn.configure(state=tk.NORMAL)

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
