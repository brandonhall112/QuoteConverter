from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

from followup_quotes.app import generate_followup_workbook, make_run_config
from followup_quotes.config import FollowupError


class FollowupUI(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Follow-Up Quote Finder")
        self.geometry("900x560")
        self.minsize(860, 520)

        self.quote_path = tk.StringVar()
        self.order_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.icon_path = tk.StringVar()
        self.floor_value = tk.StringVar(value="1500")
        self.tolerance_value = tk.StringVar(value="1")
        self.relative_tolerance_value = tk.StringVar(value="0.05")
        self.status_text = tk.StringVar(value="Ready")

        self._set_icon()
        self._configure_theme()
        self._build()

    def _set_icon(self) -> None:
        candidates = [
            self.icon_path.get().strip(),
            str(Path.cwd() / "app.ico"),
            str(Path.cwd() / "followup.ico"),
            str(Path(__file__).with_name("app.ico")),
        ]
        for icon in candidates:
            if icon and Path(icon).exists():
                try:
                    self.iconbitmap(icon)
                    self.icon_path.set(icon)
                    break
                except Exception:
                    continue

    def _configure_theme(self) -> None:
        self.configure(bg="#111827")
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure("App.TFrame", background="#111827")
        style.configure("Card.TFrame", background="#1f2937", relief="flat")
        style.configure("Header.TLabel", background="#111827", foreground="#f9fafb", font=("Segoe UI", 18, "bold"))
        style.configure("Sub.TLabel", background="#111827", foreground="#9ca3af", font=("Segoe UI", 10))
        style.configure("Label.TLabel", background="#1f2937", foreground="#e5e7eb", font=("Segoe UI", 10, "bold"))
        style.configure("Status.TLabel", background="#111827", foreground="#cbd5e1", font=("Segoe UI", 10))
        style.configure("TEntry", fieldbackground="#0b1220", foreground="#f8fafc", bordercolor="#334155", insertcolor="#f8fafc")
        style.configure("Primary.TButton", background="#2563eb", foreground="#ffffff", borderwidth=0, font=("Segoe UI", 10, "bold"))
        style.map("Primary.TButton", background=[("active", "#1d4ed8")])

    def _build_file_row(self, parent, row: int, label: str, var: tk.StringVar, cmd, button_text: str) -> None:
        ttk.Label(parent, text=label, style="Label.TLabel").grid(row=row, column=0, sticky="w", pady=(10 if row else 0, 4))
        ttk.Entry(parent, textvariable=var, width=88).grid(row=row + 1, column=0, sticky="ew", padx=(0, 12))
        ttk.Button(parent, text=button_text, command=cmd).grid(row=row + 1, column=1, sticky="ew")

    def _build(self) -> None:
        root = ttk.Frame(self, style="App.TFrame", padding=24)
        root.pack(fill="both", expand=True)

        ttk.Label(root, text="Follow-Up Quote Finder", style="Header.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(root, text="Compare quotes vs order log, then export a follow-up workbook.", style="Sub.TLabel").grid(row=1, column=0, sticky="w", pady=(2, 16))

        card = ttk.Frame(root, style="Card.TFrame", padding=16)
        card.grid(row=2, column=0, sticky="nsew")

        self._build_file_row(card, 0, "Quote Summary (.xlsx)", self.quote_path, self._browse_quotes, "Browse")
        self._build_file_row(card, 2, "Order Log (.xlsx)", self.order_path, self._browse_orders, "Browse")
        self._build_file_row(card, 4, "Follow-up Template (.xlsx, optional)", self.template_path, self._browse_template, "Browse")
        self._build_file_row(card, 6, "Application Icon (.ico, optional)", self.icon_path, self._browse_icon, "Browse")
        self._build_file_row(card, 8, "Output Workbook (.xlsx)", self.output_path, self._browse_output, "Save As")

        options = ttk.Frame(card, style="Card.TFrame")
        options.grid(row=10, column=0, columnspan=2, sticky="ew", pady=(16, 8))

        ttk.Label(options, text="Quote floor >", style="Label.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(options, textvariable=self.floor_value, width=14).grid(row=1, column=0, sticky="w", padx=(0, 12))
        ttk.Label(options, text="Absolute tolerance ±", style="Label.TLabel").grid(row=0, column=1, sticky="w")
        ttk.Entry(options, textvariable=self.tolerance_value, width=14).grid(row=1, column=1, sticky="w", padx=(0, 12))
        ttk.Label(options, text="Relative tolerance", style="Label.TLabel").grid(row=0, column=2, sticky="w")
        ttk.Entry(options, textvariable=self.relative_tolerance_value, width=14).grid(row=1, column=2, sticky="w")

        ttk.Button(root, text="Generate Follow-Up Workbook", style="Primary.TButton", command=self._run).grid(row=3, column=0, sticky="ew", pady=(16, 8))
        ttk.Label(root, textvariable=self.status_text, style="Status.TLabel").grid(row=4, column=0, sticky="w")

        root.columnconfigure(0, weight=1)
        root.rowconfigure(2, weight=1)
        card.columnconfigure(0, weight=1)

    def _browse_quotes(self) -> None:
        path = filedialog.askopenfilename(title="Select Quote Summary", filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.quote_path.set(path)

    def _browse_orders(self) -> None:
        path = filedialog.askopenfilename(title="Select Order Log", filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.order_path.set(path)

    def _browse_template(self) -> None:
        path = filedialog.askopenfilename(title="Select Follow-up Template", filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.template_path.set(path)

    def _browse_icon(self) -> None:
        path = filedialog.askopenfilename(title="Select Application Icon", filetypes=[("Icon files", "*.ico")])
        if path:
            self.icon_path.set(path)
            self._set_icon()

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
        template = self.template_path.get().strip() or None

        if not quotes or not orders or not out:
            messagebox.showerror("Missing input", "Please select quotes file, orders file, and output path.")
            return

        try:
            self.status_text.set("Generating workbook...")
            self.update_idletasks()

            cfg = make_run_config(
                quotes,
                orders,
                out,
                floor=float(self.floor_value.get()),
                tolerance=float(self.tolerance_value.get()),
                relative_tolerance=float(self.relative_tolerance_value.get()),
                template=template,
            )
            result_path = generate_followup_workbook(cfg)
            self.status_text.set(f"Done: {result_path}")
            messagebox.showinfo("Done", f"Workbook created:\n{result_path}")
        except FollowupError as exc:
            self.status_text.set("Failed: input/mapping issue")
            messagebox.showerror("Input or mapping error", str(exc))
        except Exception as exc:  # noqa: BLE001
            self.status_text.set("Failed: unexpected error")
            messagebox.showerror("Unexpected error", f"{type(exc).__name__}: {exc}")


def main() -> int:
    app = FollowupUI()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
