"""Desktop GUI: pick timesheet PDFs and a month, process them, export the report."""

import queue
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from mapper import MONTH_NAMES
from pipeline import build_report, report_to_xlsx_bytes


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Otis timesheet mapper")
        self.geometry("900x640")
        self.minsize(760, 520)

        self.files = []      # paths of the selected PDFs
        self.report = None   # compiled report of the last run
        self.report_month = None
        self.results_queue = queue.Queue()

        self._build()

    def _build(self):
        pad = {"padx": 10, "pady": 6}

        # --- 1. Files ---
        files_box = ttk.LabelFrame(self, text="1. Timesheet PDFs")
        files_box.pack(fill="x", **pad)
        self.file_list = tk.Listbox(files_box, height=6, selectmode="extended")
        self.file_list.grid(row=0, column=0, rowspan=3, sticky="nsew", padx=6, pady=6)
        scroll = ttk.Scrollbar(files_box, command=self.file_list.yview)
        scroll.grid(row=0, column=1, rowspan=3, sticky="ns", pady=6)
        self.file_list.config(yscrollcommand=scroll.set)
        ttk.Button(files_box, text="Add PDFs...", command=self.add_files).grid(
            row=0, column=2, sticky="ew", padx=6, pady=(6, 2))
        ttk.Button(files_box, text="Remove selected", command=self.remove_selected).grid(
            row=1, column=2, sticky="ew", padx=6, pady=2)
        ttk.Button(files_box, text="Clear", command=self.clear_files).grid(
            row=2, column=2, sticky="ew", padx=6, pady=(2, 6))
        files_box.columnconfigure(0, weight=1)

        # --- 2. Month + process ---
        run_box = ttk.LabelFrame(self, text="2. Month")
        run_box.pack(fill="x", **pad)
        ttk.Label(run_box, text="Month:").pack(side="left", padx=(10, 4), pady=8)
        self.month_var = tk.StringVar(value=MONTH_NAMES[datetime.now().month])
        ttk.Combobox(run_box, textvariable=self.month_var, state="readonly", width=14,
                     values=list(MONTH_NAMES.values())).pack(side="left", pady=8)
        ttk.Label(run_box, text="Timesheets for any other month are rejected.",
                  foreground="gray40").pack(side="left", padx=12)
        self.process_btn = ttk.Button(run_box, text="Process", command=self.process)
        self.process_btn.pack(side="right", padx=10, pady=8)

        # --- 3. Results ---
        results_box = ttk.LabelFrame(self, text="3. Results")
        results_box.pack(fill="both", expand=True, **pad)
        self.summary_var = tk.StringVar(value="No files processed yet.")
        ttk.Label(results_box, textvariable=self.summary_var).pack(anchor="w", padx=8, pady=(6, 2))
        self.tree = ttk.Treeview(results_box, columns=("file", "status", "message"),
                                 show="headings", height=8)
        for col, title, width in (("file", "File", 280), ("status", "Status", 80),
                                  ("message", "Message", 460)):
            self.tree.heading(col, text=title)
            self.tree.column(col, width=width, anchor="w")
        self.tree.tag_configure("rejected", foreground="#b00020")
        self.tree.pack(fill="both", expand=True, padx=8, pady=(2, 8))

        # --- 4. Export ---
        bottom = ttk.Frame(self)
        bottom.pack(fill="x", **pad)
        self.export_btn = ttk.Button(bottom, text="Export to Excel...", command=self.export,
                                     state="disabled")
        self.export_btn.pack(side="right")

    # --- file list ---
    def add_files(self):
        paths = filedialog.askopenfilenames(title="Select timesheet PDFs",
                                            filetypes=[("PDF files", "*.pdf")])
        for path in paths:
            if path not in self.files:
                self.files.append(path)
                self.file_list.insert("end", Path(path).name)

    def remove_selected(self):
        for index in reversed(self.file_list.curselection()):
            self.file_list.delete(index)
            del self.files[index]

    def clear_files(self):
        self.file_list.delete(0, "end")
        self.files.clear()

    # --- processing ---
    def process(self):
        if not self.files:
            messagebox.showinfo("No files", "Add at least one timesheet PDF first.")
            return
        month = next(n for n, name in MONTH_NAMES.items() if name == self.month_var.get())
        files = list(self.files)
        self.process_btn.config(state="disabled")
        self.export_btn.config(state="disabled")
        self.summary_var.set("Processing...")
        self.tree.delete(*self.tree.get_children())

        def work():
            try:
                data = [(Path(p).name, Path(p).read_bytes()) for p in files]
                self.results_queue.put(("done", month, *build_report(data, month)))
            except Exception as e:
                self.results_queue.put(("error", str(e)))

        threading.Thread(target=work, daemon=True).start()
        self.after(100, self._poll)

    def _poll(self):
        try:
            message = self.results_queue.get_nowait()
        except queue.Empty:
            self.after(100, self._poll)
            return

        self.process_btn.config(state="normal")
        if message[0] == "error":
            self.summary_var.set("Processing failed.")
            messagebox.showerror("Error", message[1])
            return

        _, month, report, results = message
        self.report, self.report_month = report, month
        for r in results:
            self.tree.insert("", "end", values=(r["File"], r["Status"], r["Message"]),
                             tags=("rejected",) if r["Status"] != "OK" else ())
        ok = sum(r["Status"] == "OK" for r in results)
        text = f"{MONTH_NAMES[month]}: {ok} of {len(results)} file(s) mapped."
        if ok < len(results):
            text += " Rejected files are NOT in the report."
        self.summary_var.set(text)
        self.export_btn.config(state="normal" if ok else "disabled")

    # --- export ---
    def export(self):
        name = (f"Final_Timesheet_Report_{MONTH_NAMES[self.report_month]}_"
                f"{datetime.now():%Y-%m-%d_%H-%M-%S}.xlsx")
        path = filedialog.asksaveasfilename(
            title="Export report", initialfile=name, defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")])
        if not path:
            return
        try:
            Path(path).write_bytes(report_to_xlsx_bytes(self.report))
        except OSError as e:
            messagebox.showerror("Could not save", f"{e}\n\nIs the file open in Excel?")
            return
        messagebox.showinfo("Exported", f"Report saved to:\n{path}")


def selftest(month, pdf_paths, out_path):
    """Headless check used to verify a packaged .exe: map the PDFs, write a summary."""
    files = [(Path(p).name, Path(p).read_bytes()) for p in pdf_paths]
    report, results = build_report(files, int(month))
    lines = [f"{r['Status']} | {r['File']} | {r['Message']}" for r in results]
    lines.append("rows: " + ", ".join(f"{k}={len(v)}" for k, v in report.items()))
    lines.append(f"xlsx bytes: {len(report_to_xlsx_bytes(report))}")
    Path(out_path).write_text("\n".join(lines), encoding="utf-8")


if __name__ == "__main__":
    if len(sys.argv) > 3 and sys.argv[1] == "--selftest":
        selftest(sys.argv[2], sys.argv[4:], sys.argv[3])
    else:
        App().mainloop()
