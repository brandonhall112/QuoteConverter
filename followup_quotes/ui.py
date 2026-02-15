from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .app import generate_followup_workbook, make_run_config
from .config import FollowupError


class FollowupUI(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Follow-Up Quote Finder")
        self.geometry("700x380")
        self.resizable(True, True)

        self.quote_path = tk.StringVar()
        self.order_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.floor_value = tk.StringVar(value="1500")
        self.tolerance_value = tk.StringVar(value="1")

        self._build()

    def _build(self) -> None:
        frame = ttk.Frame(self, padding=16)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Quote Summary (.xlsx)").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.quote_path, width=70).grid(row=1, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(frame, text="Browse", command=self._browse_quotes).grid(row=1, column=1)

        ttk.Label(frame, text="Order Log (.xlsx)").grid(row=2, column=0, sticky="w", pady=(12, 0))
        ttk.Entry(frame, textvariable=self.order_path, width=70).grid(row=3, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(frame, text="Browse", command=self._browse_orders).grid(row=3, column=1)

        ttk.Label(frame, text="Output Workbook (.xlsx)").grid(row=4, column=0, sticky="w", pady=(12, 0))
        ttk.Entry(frame, textvariable=self.output_path, width=70).grid(row=5, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(frame, text="Save As", command=self._browse_output).grid(row=5, column=1)

        options = ttk.Frame(frame)
        options.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(16, 8))
        ttk.Label(options, text="Quote floor >").grid(row=0, column=0, sticky="w")
        ttk.Entry(options, textvariable=self.floor_value, width=12).grid(row=0, column=1, padx=(8, 20))
        ttk.Label(options, text="Tolerance ±").grid(row=0, column=2, sticky="w")
        ttk.Entry(options, textvariable=self.tolerance_value, width=12).grid(row=0, column=3, padx=(8, 0))

        ttk.Button(frame, text="Generate Follow-Up Workbook", command=self._run).grid(row=7, column=0, columnspan=2, pady=(16, 4), sticky="ew")

        ttk.Label(
            frame,
            text="Output sheets: Option A, Option B, Option C, _Meta (values only)",
            foreground="#444",
        ).grid(row=8, column=0, columnspan=2, sticky="w", pady=(8, 0))

        frame.columnconfigure(0, weight=1)

    def _browse_quotes(self) -> None:
        path = filedialog.askopenfilename(title="Select Quote Summary", filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.quote_path.set(path)

    def _browse_orders(self) -> None:
        path = filedialog.askopenfilename(title="Select Order Log", filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.order_path.set(path)

    def _browse_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Save Output Workbook",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="FollowUp_Output.xlsx",
        )
        if path:
            self.output_path.set(path)

    def _run(self) -> None:
        quotes = self.quote_path.get().strip()
        orders = self.order_path.get().strip()
        out = self.output_path.get().strip()

        if not quotes or not orders or not out:
            messagebox.showerror("Missing input", "Please select quotes file, orders file, and output path.")
            return

        try:
            cfg = make_run_config(
                quotes,
                orders,
                out,
                floor=float(self.floor_value.get()),
                tolerance=float(self.tolerance_value.get()),
            )
            result_path = generate_followup_workbook(cfg)
            messagebox.showinfo("Done", f"Workbook created:\n{result_path}")
        except FollowupError as exc:
            messagebox.showerror("Input or mapping error", str(exc))
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Unexpected error", f"{type(exc).__name__}: {exc}")


def main() -> int:
    app = FollowupUI()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
