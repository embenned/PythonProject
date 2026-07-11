import logging
import os
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from app.comparison.models import ComparisonError, WorkbookLoadError
from app.comparison.workbook_comparator import WorkbookComparator

logger = logging.getLogger(__name__)


class CompareTab:
    _PREVIEW_LIMIT = 20

    def __init__(self, parent: ttk.Frame) -> None:
        self._parent = parent
        self._root = parent.winfo_toplevel()
        self._standard_path = tk.StringVar()
        self._project_path = tk.StringVar()
        self._status_text = tk.StringVar(value="Select the standard and project workbooks.")
        self._preview_title_text = tk.StringVar(value="No comparison run yet.")
        self._sheets_text = tk.StringVar(value="-")
        self._cells_text = tk.StringVar(value="-")
        self._deviations_text = tk.StringVar(value="-")
        self._time_text = tk.StringVar(value="-")
        self._output_text = tk.StringVar(value="-")
        self._preview_count_text = tk.StringVar(value="-")
        self._preview_note_text = tk.StringVar(value="Run a comparison to preview the first deviations before saving the report.")
        self._preview_rows: list[tuple[str, str, str, str, str]] = []
        self._build_ui()

    def _build_ui(self) -> None:
        self._parent.columnconfigure(0, weight=1)
        self._parent.rowconfigure(0, weight=1)

        container = ttk.Frame(self._parent, padding=20)
        container.grid(row=0, column=0, sticky="nsew")
        container.columnconfigure(1, weight=1)

        ttk.Label(container, text="Compare Workbook Against Standard", font=("Segoe UI", 11, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 14)
        )

        self._add_file_selector(container, row=1, label_text="Standard Workbook", variable=self._standard_path, command=self._browse_standard)
        self._add_file_selector(container, row=2, label_text="Project Workbook", variable=self._project_path, command=self._browse_project)

        self._generate_btn = ttk.Button(container, text="Generate Deviation Report", command=self._on_generate)
        self._generate_btn.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(8, 14))

        summary = ttk.LabelFrame(container, text="Summary", padding=12)
        summary.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(0, 12))
        summary.columnconfigure(1, weight=1)

        self._add_summary_row(summary, 0, "Total Sheets Compared", self._sheets_text)
        self._add_summary_row(summary, 1, "Total Cells Compared", self._cells_text)
        self._add_summary_row(summary, 2, "Total Deviations Found", self._deviations_text)
        self._add_summary_row(summary, 3, "Execution Time", self._time_text)
        self._add_summary_row(summary, 4, "Output File", self._output_text)

        preview = ttk.LabelFrame(container, text="Preview Before Saving", padding=12)
        preview.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=(0, 12))
        preview.columnconfigure(0, weight=1)
        preview.rowconfigure(2, weight=1)

        ttk.Label(preview, textvariable=self._preview_title_text, font=("Segoe UI", 10, "bold")).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(preview, textvariable=self._preview_count_text).grid(row=1, column=0, sticky="w", pady=(2, 6))

        self._preview_tree = ttk.Treeview(
            preview,
            columns=("change_type", "sheet_name", "cell_address", "standard_value", "project_value"),
            show="headings",
            height=10,
        )
        self._preview_tree.heading("change_type", text="Change Type")
        self._preview_tree.heading("sheet_name", text="Sheet Name")
        self._preview_tree.heading("cell_address", text="Cell Address")
        self._preview_tree.heading("standard_value", text="Standard Value")
        self._preview_tree.heading("project_value", text="Project Value")
        self._preview_tree.column("change_type", width=130, anchor="w")
        self._preview_tree.column("sheet_name", width=220, anchor="w")
        self._preview_tree.column("cell_address", width=90, anchor="center")
        self._preview_tree.column("standard_value", width=220, anchor="w")
        self._preview_tree.column("project_value", width=220, anchor="w")
        self._preview_tree.grid(row=2, column=0, sticky="nsew")

        preview_scrollbar = ttk.Scrollbar(preview, orient="vertical", command=self._preview_tree.yview)
        preview_scrollbar.grid(row=2, column=1, sticky="ns")
        self._preview_tree.configure(yscrollcommand=preview_scrollbar.set)

        ttk.Label(preview, textvariable=self._preview_note_text, wraplength=760, justify="left").grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(6, 0)
        )

        self._status_label = tk.Label(
            container,
            textvariable=self._status_text,
            wraplength=760,
            justify="left",
            anchor="w",
            fg="#333333",
            bg=self._root.cget("bg"),
        )
        self._status_label.grid(row=6, column=0, columnspan=3, sticky="ew")

    def _add_file_selector(self, parent, row: int, label_text: str, variable: tk.StringVar, command) -> None:
        ttk.Label(parent, text=f"{label_text}:").grid(row=row, column=0, sticky="w", pady=4)
        entry = ttk.Entry(parent, textvariable=variable)
        entry.grid(row=row, column=1, sticky="ew", padx=(8, 8), pady=4)
        ttk.Button(parent, text="Browse", command=command).grid(row=row, column=2, sticky="ew", pady=4)

    @staticmethod
    def _add_summary_row(parent, row: int, label_text: str, variable: tk.StringVar) -> None:
        ttk.Label(parent, text=f"{label_text}:").grid(row=row, column=0, sticky="w", pady=2)
        ttk.Label(parent, textvariable=variable).grid(row=row, column=1, sticky="w", pady=2)

    def _browse_standard(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Select Standard Workbook",
            filetypes=[("Excel Macro-Enabled Workbook", "*.xlsm"), ("Excel Workbook", "*.xlsx"), ("All files", "*.*")],
        )
        if file_path:
            self._standard_path.set(file_path)

    def _browse_project(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Select Project Workbook",
            filetypes=[("Excel Macro-Enabled Workbook", "*.xlsm"), ("Excel Workbook", "*.xlsx"), ("All files", "*.*")],
        )
        if file_path:
            self._project_path.set(file_path)

    def _set_status(self, message: str) -> None:
        self._status_text.set(message)
        self._root.update_idletasks()

    def _progress_callback(self, message: str) -> None:
        self._set_status(message)

    def _clear_summary(self) -> None:
        self._sheets_text.set("-")
        self._cells_text.set("-")
        self._deviations_text.set("-")
        self._time_text.set("-")
        self._output_text.set("-")

    def _clear_preview(self) -> None:
        self._preview_title_text.set("No comparison run yet.")
        self._preview_count_text.set("-")
        self._preview_note_text.set("Run a comparison to preview the first deviations before saving the report.")
        self._preview_rows = []
        for item in self._preview_tree.get_children():
            self._preview_tree.delete(item)

    def _populate_preview(self, deviations) -> None:
        preview_rows = deviations[: self._PREVIEW_LIMIT]
        self._preview_rows = [
            (
                item.change_type,
                item.sheet_name,
                item.cell_address,
                "" if item.standard_value is None else str(item.standard_value),
                "" if item.project_value is None else str(item.project_value),
            )
            for item in preview_rows
        ]

        self._preview_title_text.set("Preview of first deviations")
        self._preview_count_text.set(
            f"Showing {len(self._preview_rows)} of {len(deviations)} deviations before saving the report."
        )
        if deviations:
            self._preview_note_text.set(
                "Review the preview below. The report will be saved only after you confirm the file location."
            )
        else:
            self._preview_note_text.set("No deviations were found. A report can still be saved if needed.")

        for item in self._preview_tree.get_children():
            self._preview_tree.delete(item)

        for row in self._preview_rows:
            self._preview_tree.insert("", "end", values=row)

    def _on_generate(self) -> None:
        standard_path = self._standard_path.get().strip()
        project_path = self._project_path.get().strip()

        if not standard_path or not project_path:
            messagebox.showwarning("Missing Files", "Please select both workbook files first.")
            return

        if not os.path.exists(standard_path):
            messagebox.showerror("Missing File", f"Standard workbook not found:\n{standard_path}")
            return

        if not os.path.exists(project_path):
            messagebox.showerror("Missing File", f"Project workbook not found:\n{project_path}")
            return

        self._clear_summary()
        self._clear_preview()
        self._generate_btn.configure(state="disabled")
        self._set_status("Workbook loading...")
        start_time = time.perf_counter()

        comparator = None
        try:
            comparator = WorkbookComparator(
                standard_workbook_path=standard_path,
                project_workbook_path=project_path,
                progress_callback=self._progress_callback,
            )
            result = comparator.compare_sheets()
            compare_elapsed = round(time.perf_counter() - start_time, 2)

            self._sheets_text.set(str(result.summary.total_sheets_compared))
            self._cells_text.set(str(result.summary.total_cells_compared))
            self._deviations_text.set(str(result.summary.total_deviations_found))
            self._time_text.set(f"Comparison: {compare_elapsed:.2f} s")
            self._output_text.set("Not saved yet")
            self._populate_preview(result.deviations)

            should_save = messagebox.askyesno(
                "Save Deviation Report",
                f"Comparison finished with {result.summary.total_deviations_found} deviations.\n\n"
                "Do you want to save the deviation report now?",
            )
            if not should_save:
                self._set_status("Comparison completed. Report was not saved.")
                return

            output_path = filedialog.asksaveasfilename(
                title="Save deviation report",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile="Workbook_Deviations.xlsx",
            )
            if not output_path:
                self._set_status("Report save cancelled.")
                return

            self._set_status("Creating deviation report...")
            save_start = time.perf_counter()
            result.output_file = comparator.save_report(result, output_path)
            total_elapsed = round(time.perf_counter() - start_time, 2)
            save_elapsed = round(time.perf_counter() - save_start, 2)
            result.summary.execution_time_seconds = total_elapsed

            self._sheets_text.set(str(result.summary.total_sheets_compared))
            self._cells_text.set(str(result.summary.total_cells_compared))
            self._deviations_text.set(str(result.summary.total_deviations_found))
            self._time_text.set(f"Total: {total_elapsed:.2f} s (save: {save_elapsed:.2f} s)")
            self._output_text.set(result.output_file)
            self._set_status("Comparison completed successfully.")
        except WorkbookLoadError as exc:
            logger.exception("Workbook loading failed.")
            self._set_status(f"Error: {exc}")
            messagebox.showerror("Workbook Loading Failed", str(exc))
        except ComparisonError as exc:
            logger.exception("Comparison failed.")
            self._set_status(f"Error: {exc}")
            messagebox.showerror("Comparison Failed", str(exc))
        except Exception as exc:
            logger.exception("Unexpected comparison failure.")
            self._set_status(f"Unexpected error: {exc}")
            messagebox.showerror("Comparison Failed", str(exc))
        finally:
            if comparator is not None:
                comparator.close_workbooks()
            self._generate_btn.configure(state="normal")
