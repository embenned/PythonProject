import logging
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from app.models.train_type import DataValidationError
from app.ui.compare_tab import CompareTab

logger = logging.getLogger(__name__)

_STATUS_IDLE = ""
_COLOR_SUCCESS = "#1a7a1a"
_COLOR_ERROR = "#b00000"
_COLOR_IDLE = "#333333"


class MainWindow:

    def __init__(
        self,
        root: tk.Tk,
        train_type_service,
        excel_generator,
        data_update_service,
    ) -> None:
        self._root = root
        self._train_type_service = train_type_service
        self._excel_generator = excel_generator
        self._data_update_service = data_update_service

        self._selected_train_type = tk.StringVar()
        self._status_text = tk.StringVar(value=_STATUS_IDLE)
        self._status_color = _COLOR_IDLE

        self._build_ui()
        self._load_train_types()

    def _build_ui(self) -> None:
        self._root.title("Train Type Data Preparation Tool")
        self._root.resizable(True, True)
        self._root.minsize(980, 640)

        notebook = ttk.Notebook(self._root)
        notebook.pack(fill="both", expand=True)

        generate_frame = ttk.Frame(notebook, padding=20)
        compare_frame = ttk.Frame(notebook, padding=20)
        notebook.add(generate_frame, text="Generate Train Type")
        notebook.add(compare_frame, text="Compare Workbook")

        self._root.columnconfigure(0, weight=1)
        self._root.rowconfigure(0, weight=1)

        ttk.Label(generate_frame, text="Select Train Type:").grid(
            row=0, column=0, sticky="w", pady=(0, 4)
        )
        self._combo = ttk.Combobox(
            generate_frame,
            textvariable=self._selected_train_type,
            state="readonly",
            width=35,
        )
        self._combo.grid(row=1, column=0, sticky="ew", pady=(0, 16))

        self._generate_btn = ttk.Button(
            generate_frame,
            text="Generate Excel",
            command=self._on_generate,
        )
        self._generate_btn.grid(row=2, column=0, sticky="ew", pady=(0, 12))

        self._status_label = tk.Label(
            generate_frame,
            textvariable=self._status_text,
            wraplength=380,
            justify="left",
            fg=_COLOR_IDLE,
            bg=self._root.cget("bg"),
        )
        self._status_label.grid(row=3, column=0, sticky="ew")
        generate_frame.columnconfigure(0, weight=1)

        self._compare_tab = CompareTab(compare_frame)

    def _load_train_types(self) -> None:
        try:
            types = self._train_type_service.get_available_train_types()
            self._combo["values"] = types
            if types:
                self._combo.current(0)
            self._set_status("Data loaded successfully.", success=True)
        except Exception as exc:
            logger.exception("Failed to load train types.")
            self._set_status(f"Error loading train types: {exc}", success=False)
            self._generate_btn.configure(state="disabled")

    def _on_generate(self) -> None:
        train_type = self._selected_train_type.get().strip()
        if not train_type:
            messagebox.showwarning("No Selection", "Please select a train type first.")
            return

        output_path = filedialog.asksaveasfilename(
            title="Save generated Excel file",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="Train type.xlsx",
        )
        if not output_path:
            self._set_status("Generation cancelled.", success=None)
            return

        self._generate_btn.configure(state="disabled")
        self._set_status("Generating…", success=None)
        self._root.update_idletasks()

        try:
            values = self._train_type_service.get_train_type_values(train_type)
            self._excel_generator.generate(values, output_path)
            self._data_update_service.record_generation_event(train_type, output_path)
            filename = os.path.basename(output_path)
            self._set_status(f"Success! File saved: {filename}", success=True)
            logger.info("Generated Excel for '%s' at '%s'.", train_type, output_path)
        except DataValidationError as exc:
            logger.warning("Validation error during generation: %s", exc)
            self._set_status(f"Validation error: {exc}", success=False)
            messagebox.showerror("Invalid data", str(exc))
        except Exception as exc:
            logger.exception("Excel generation failed.")
            self._set_status(f"Error: {exc}", success=False)
            messagebox.showerror("Generation failed", str(exc))
        finally:
            self._generate_btn.configure(state="normal")

    def _set_status(self, message: str, success: bool | None) -> None:
        self._status_text.set(message)
        if success is True:
            color = _COLOR_SUCCESS
        elif success is False:
            color = _COLOR_ERROR
        else:
            color = _COLOR_IDLE
        self._status_label.configure(fg=color)

    def run(self) -> None:
        self._root.mainloop()
