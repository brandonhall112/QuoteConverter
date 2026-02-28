from __future__ import annotations

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

from followup_quotes.app import generate_followup_workbook, make_run_config, resolve_template_path
from followup_quotes.config import FollowupError


class FollowupUI(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Follow-Up Quote Finder")
        self.geometry("960x620")
        self.minsize(920, 580)

        self.quote_path = tk.StringVar()
        self.order_path = tk.StringVar()
        self.template_path = tk.StringVar(value=self._default_template_label())
        self.output_path = tk.StringVar()
        self.icon_path = tk.StringVar()
        self.floor_value = tk.StringVar(value="1500")
        self.tolerance_value = tk.StringVar(value="1")
        self.relative_tolerance_value = tk.StringVar(value="0.05")
        self.status_text = tk.StringVar(value="Ready")

        self._set_app_icon()
        self._configure_theme()
        self._build()

    def _default_template_label(self) -> str:
        resolved = resolve_template_path(None)
        return str(resolved) if resolved else "(auto-detect: not found)"

    def _set_app_icon(self) -> None:
        root = Path(__file__).resolve().parent
        candidates = [
            Path.cwd() / "assets" / "app.ico",
            Path.cwd() / "assets" / "followup.ico",
            root / "assets" / "app.ico",
            root / "assets" / "followup.ico",
            root / "app.ico",
            root / "followup.ico",
            Path.cwd() / "app.ico",
            Path.cwd() / "followup.ico",
        ]
        for icon in candidates:
            if icon.exists():
                try:
                    self.iconbitmap(str(icon))
                    return
                except Exception:
                    continue

    def _configure_theme(self) -> None:
        c_bg = "#343551"
        c_panel = "#4c4b4c"
        c_card = "#5f5f5f"
        c_border = "#6790a0"
        c_accent = "#e54102"
        c_text = "#bebebe"
        c_text_dark = "#0c2949"
        c_ok = "#90a997"

        self.configure(bg=c_bg)
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure("Root.TFrame", background=c_bg)
        style.configure("Sidebar.TFrame", background=c_panel)
        style.configure("Card.TFrame", background=c_card)

        style.configure("Title.TLabel", background=c_bg, foreground=c_accent, font=("Segoe UI Semibold", 22))
        style.configure("Subtitle.TLabel", background=c_bg, foreground=c_text, font=("Segoe UI", 10))
        style.configure("Label.TLabel", background=c_card, foreground="#ffffff", font=("Segoe UI", 10, "bold"))
        style.configure("Hint.TLabel", background=c_card, foreground=c_text, font=("Segoe UI", 9))
        style.configure("Status.TLabel", background=c_bg, foreground=c_ok, font=("Segoe UI", 10, "bold"))

        style.configure(
            "TEntry",
            fieldbackground="#ffffff",
            foreground=c_text_dark,
            bordercolor=c_border,
            lightcolor=c_border,
            darkcolor=c_border,
            insertcolor=c_text_dark,
            padding=7,
        )
        style.configure(
            "Primary.TButton",
            background=c_accent,
            foreground="#ffffff",
            borderwidth=0,
            focusthickness=0,
            focuscolor=c_accent,
            font=("Segoe UI", 10, "bold"),
            padding=(12, 10),
        )
        style.map("Primary.TButton", background=[("active", "#e04426")])
        style.configure("Secondary.TButton", background="#42725e", foreground="#ffffff", font=("Segoe UI", 9, "bold"), padding=(10, 8))
        style.map("Secondary.TButton", background=[("active", "#90a997")])

    def _build_file_row(self, parent, row: int, label: str, var: tk.StringVar, cmd, btn_text: str = "Browse") -> None:
        ttk.Label(parent, text=label, style="Label.TLabel").grid(row=row, column=0, sticky="w", pady=(10 if row else 0, 4))
        ttk.Entry(parent, textvariable=var).grid(row=row + 1, column=0, sticky="ew", padx=(0, 10))
        ttk.Button(parent, text=btn_text, style="Secondary.TButton", command=cmd).grid(row=row + 1, column=1, sticky="ew")

    def _build(self) -> None:
        root = ttk.Frame(self, style="Root.TFrame", padding=20)
        root.pack(fill="both", expand=True)
        root.columnconfigure(1, weight=1)
        root.rowconfigure(0, weight=1)

        sidebar = ttk.Frame(root, style="Sidebar.TFrame", padding=16)
        sidebar.grid(row=0, column=0, sticky="nsw", padx=(0, 16))

        ttk.Label(sidebar, text="Follow-Up\nQuote Finder", foreground="#ffffff", background="#4c4b4c", font=("Segoe UI", 16, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Label(
            sidebar,
            text="Uses customer + grouped order totals\nwith absolute/relative tolerance\n\nOutputs:\n• Follow-Up\n• Per-rep tabs\n• _Meta",
            foreground="#bebebe",
            background="#4c4b4c",
            font=("Segoe UI", 10),
            justify="left",
        ).grid(row=1, column=0, sticky="w", pady=(12, 0))

        content = ttk.Frame(root, style="Root.TFrame")
        content.grid(row=0, column=1, sticky="nsew")
        content.columnconfigure(0, weight=1)

        ttk.Label(content, text="Generate Follow-Up Workbook", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(content, text="Template-driven output with automatic app icon and rep-based sheets", style="Subtitle.TLabel").grid(row=1, column=0, sticky="w", pady=(0, 12))

        card = ttk.Frame(content, style="Card.TFrame", padding=16)
        card.grid(row=2, column=0, sticky="nsew")
        card.columnconfigure(0, weight=1)

        self._build_file_row(card, 0, "Quote Summary (.xlsx)", self.quote_path, self._browse_quotes)
        self._build_file_row(card, 2, "Order Log (.xlsx)", self.order_path, self._browse_orders)
        self._build_file_row(card, 4, "Output Workbook (.xlsx)", self.output_path, self._browse_output, btn_text="Save As")

        ttk.Label(card, text="Template workbook (auto-detected; optional override)", style="Label.TLabel").grid(row=6, column=0, sticky="w", pady=(12, 4))
        ttk.Entry(card, textvariable=self.template_path).grid(row=7, column=0, sticky="ew", padx=(0, 10))
        ttk.Button(card, text="Override", style="Secondary.TButton", command=self._browse_template).grid(row=7, column=1, sticky="ew")
        ttk.Label(card, text="Put your .ico in followup_quotes/app.ico to brand the app icon automatically.", style="Hint.TLabel").grid(row=8, column=0, columnspan=2, sticky="w", pady=(6, 0))

        options = ttk.Frame(card, style="Card.TFrame")
        options.grid(row=9, column=0, columnspan=2, sticky="ew", pady=(14, 4))

        ttk.Label(options, text="Quote floor", style="Label.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(options, textvariable=self.floor_value, width=14).grid(row=1, column=0, sticky="w", padx=(0, 10))
        ttk.Label(options, text="Abs tolerance", style="Label.TLabel").grid(row=0, column=1, sticky="w")
        ttk.Entry(options, textvariable=self.tolerance_value, width=14).grid(row=1, column=1, sticky="w", padx=(0, 10))
        ttk.Label(options, text="Relative tolerance", style="Label.TLabel").grid(row=0, column=2, sticky="w")
        ttk.Entry(options, textvariable=self.relative_tolerance_value, width=14).grid(row=1, column=2, sticky="w")

        ttk.Button(content, text="Generate Workbook", style="Primary.TButton", command=self._run).grid(row=3, column=0, sticky="ew", pady=(14, 6))
        ttk.Label(content, textvariable=self.status_text, style="Status.TLabel").grid(row=4, column=0, sticky="w")

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
        template_raw = self.template_path.get().strip()
        template = None if template_raw in {"", "(auto-detect: not found)"} else template_raw

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
