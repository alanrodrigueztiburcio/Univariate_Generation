"""Minimal Tkinter GUI wrapper for the csv_reporter package.

This script provides a simple graphical front end for non‑technical
users to generate frequency tables and reports from a CSV file.
Users can select a CSV input file and an output folder, adjust a
couple of optional parameters, and run the underlying pipeline with
one click. Progress and error messages are displayed in the interface.

The GUI is designed for Windows environments where Microsoft Word is
installed, as the report generation relies on COM automation via
pywin32.
"""

import os
import threading
import queue
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd

try:
    # import functions from the csv_reporter package
    from csv_reporter.utils import select_columns
    from csv_reporter.freq_tables import compute_tables
    from csv_reporter.excel_report import build_excel_report
    from csv_reporter.word_report import build_word_report
except ImportError as e:
    raise RuntimeError(
        "The csv_reporter package must be installed and on the PYTHONPATH."
    ) from e


def main() -> None:
    """Create and run the Tkinter application."""
    root = tk.Tk()
    root.title("CSV Reporter GUI")

    # Message queue for status updates from worker thread
    msg_queue: queue.Queue[str] = queue.Queue()

    # --- Input file selection ---
    tk.Label(root, text="CSV file:").grid(row=0, column=0, sticky="w", pady=2, padx=2)
    csv_var = tk.StringVar()
    csv_entry = tk.Entry(root, textvariable=csv_var, width=50)
    csv_entry.grid(row=0, column=1, pady=2, padx=2, sticky="w")

    def browse_csv() -> None:
        path = filedialog.askopenfilename(
            title="Select CSV file", filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if path:
            csv_var.set(path)

    tk.Button(root, text="Browse", command=browse_csv).grid(row=0, column=2, pady=2, padx=2)

    # --- Output directory selection ---
    tk.Label(root, text="Output folder:").grid(row=1, column=0, sticky="w", pady=2, padx=2)
    out_var = tk.StringVar()
    out_entry = tk.Entry(root, textvariable=out_var, width=50)
    out_entry.grid(row=1, column=1, pady=2, padx=2, sticky="w")

    def browse_output() -> None:
        directory = filedialog.askdirectory(title="Select output folder")
        if directory:
            out_var.set(directory)

    tk.Button(root, text="Browse", command=browse_output).grid(row=1, column=2, pady=2, padx=2)

    # --- Optional parameters ---
    tk.Label(root, text="Start column:").grid(row=2, column=0, sticky="w", pady=2, padx=2)
    start_var = tk.StringVar(value="2")
    tk.Entry(root, textvariable=start_var, width=10).grid(row=2, column=1, sticky="w", pady=2, padx=2)

    tk.Label(root, text="Max unique values:").grid(row=3, column=0, sticky="w", pady=2, padx=2)
    max_unique_var = tk.StringVar(value="25")
    tk.Entry(root, textvariable=max_unique_var, width=10).grid(row=3, column=1, sticky="w", pady=2, padx=2)

    tk.Label(root, text="Missing label:").grid(row=4, column=0, sticky="w", pady=2, padx=2)
    na_var = tk.StringVar(value="Missing")
    tk.Entry(root, textvariable=na_var, width=20).grid(row=4, column=1, sticky="w", pady=2, padx=2)

    # --- Status area ---
    status_text = tk.Text(root, height=10, width=60, state="disabled")
    status_text.grid(row=6, column=0, columnspan=3, pady=5, padx=2)

    def append_status(msg: str) -> None:
        status_text.configure(state="normal")
        status_text.insert(tk.END, msg + "\n")
        status_text.see(tk.END)
        status_text.configure(state="disabled")

    def worker() -> None:
        """Background task to run the report generation."""
        csv_path = csv_var.get().strip()
        out_dir = out_var.get().strip()
        start_col_str = start_var.get().strip()
        max_unique_str = max_unique_var.get().strip()
        na_label = na_var.get().strip()
        # Validation
        if not csv_path:
            msg_queue.put("Error: Please select a CSV file.")
            return
        if not os.path.isfile(csv_path):
            msg_queue.put(f"Error: CSV file not found: {csv_path}")
            return
        if not out_dir:
            msg_queue.put("Error: Please select an output folder.")
            return
        if not os.path.isdir(out_dir):
            msg_queue.put(f"Error: Output folder not found: {out_dir}")
            return
        # Parse numeric parameters
        try:
            start_col = int(start_col_str)
        except ValueError:
            msg_queue.put("Error: Start column must be an integer.")
            return
        try:
            max_unique = int(max_unique_str)
        except ValueError:
            msg_queue.put("Error: Max unique values must be an integer.")
            return
        # Begin pipeline
        try:
            msg_queue.put("Reading CSV file...")
            df = pd.read_csv(csv_path)
            msg_queue.put(f"Loaded {len(df)} rows and {len(df.columns)} columns.")
            msg_queue.put("Selecting columns...")
            columns = select_columns(df, start_col=start_col, include_cols=None, exclude_cols=None, max_unique=max_unique)
            if not columns:
                msg_queue.put("No suitable columns found for analysis.")
                return
            msg_queue.put(f"Selected {len(columns)} columns: {', '.join(columns)}")
            msg_queue.put("Computing frequency tables...")
            tables = compute_tables(df, columns, na_label=na_label)
            excel_path = os.path.join(out_dir, "report.xlsx")
            doc_path = os.path.join(out_dir, "report.docx")
            msg_queue.put("Building Excel report...")
            build_excel_report(tables, excel_path)
            msg_queue.put(f"Excel report saved to {excel_path}")
            msg_queue.put("Building Word report (this may take a while)...")
            try:
                build_word_report(excel_path, doc_path)
                msg_queue.put(f"Word report saved to {doc_path}")
            except Exception as e:
                msg_queue.put(f"Error generating Word report: {e}")
                raise
            msg_queue.put("SUCCESS")
        except Exception as e:
            msg_queue.put(f"Error: {e}")

    def run_report() -> None:
        """Start the worker thread and poll for status updates."""
        status_text.configure(state="normal")
        status_text.delete(1.0, tk.END)
        status_text.configure(state="disabled")
        # Launch worker thread
        thread = threading.Thread(target=worker, daemon=True)
        thread.start()

        def poll() -> None:
            while not msg_queue.empty():
                message = msg_queue.get_nowait()
                if message == "SUCCESS":
                    messagebox.showinfo("Success", "Report generated successfully.")
                elif message.startswith("Error"):
                    messagebox.showerror("Error", message)
                append_status(message)
            if thread.is_alive():
                root.after(200, poll)

        root.after(200, poll)

    # --- Run button ---
    tk.Button(root, text="Run", command=run_report).grid(row=5, column=0, columnspan=3, pady=5)

    root.mainloop()


if __name__ == "__main__":
    main()