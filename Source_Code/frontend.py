"""
Frontend GUI for the Automation Prototype.
Built with customtkinter for a modern look and built-in light/dark mode.
"""
import customtkinter as ctk
from tkinter import filedialog
import tkinter as tk
import threading
import sys
import io
from datetime import datetime

# ── Import backend functions ──────────────────────────────────────────────────
from AutomationPrototype import (
    automated_testing_activity,
    severity_counter,
    assessment_results,
    recommendations,
    appendix,
    finding_details,
    Logistics,
    DEFAULT_RECOMMENDATIONS_PATH,
    DEFAULT_EXECUTIVE_REPORT_PATH,
    DEFAULT_ACTIVITY_REPORT_PATH,
    DEFAULT_FINDINGS_REPORT_PATH,
    DEFAULT_TECHNICAL_REPORT_PATH,
)

# ── customtkinter appearance defaults ─────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

LOG_COLORS = {
    "INFO":   "#60a5fa",
    "OK":     "#4ade80",
    "WARN":   "#f87171",
    "PLAIN":  "#d1d5db",
    "HEADER": "#c084fc",
}


# ── Main Application ──────────────────────────────────────────────────────────
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Report Automation Tool")
        self.geometry("860x700")
        self.resizable(True, True)
        self.minsize(680, 540)
        self._build_ui()

    # ── UI Construction ───────────────────────────────────────────────────────
    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # ── Top bar ──────────────────────────────────────────────────────────
        top_bar = ctk.CTkFrame(self, corner_radius=0, height=54)
        top_bar.grid(row=0, column=0, sticky="ew")
        top_bar.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            top_bar,
            text="Report Automation Tool",
            font=ctk.CTkFont(size=16, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=20, pady=12)

        self.theme_btn = ctk.CTkButton(
            top_bar,
            text="☀  Light Mode",
            width=130,
            command=self._toggle_theme,
        )
        self.theme_btn.grid(row=0, column=1, padx=16, pady=10)

        # ── Main content frame ────────────────────────────────────────────────
        content = ctk.CTkFrame(self, fg_color="transparent")
        content.grid(row=1, column=0, sticky="nsew", padx=20, pady=(12, 16))
        content.grid_columnconfigure(0, weight=1)
        content.grid_rowconfigure(2, weight=1)

        # ── File selectors card ───────────────────────────────────────────────
        files_card = ctk.CTkFrame(content)
        files_card.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        files_card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            files_card,
            text="Report Files",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=14, pady=(10, 4))

        self.activity_var  = tk.StringVar(value=DEFAULT_ACTIVITY_REPORT_PATH)
        self.findings_var  = tk.StringVar(value=DEFAULT_FINDINGS_REPORT_PATH)
        self.technical_var = tk.StringVar(value=DEFAULT_TECHNICAL_REPORT_PATH)
        self.executive_var = tk.StringVar(value=DEFAULT_EXECUTIVE_REPORT_PATH)
        self.recommendation_var = tk.StringVar(value=DEFAULT_RECOMMENDATIONS_PATH)

        self._file_row(files_card, "Activity Report (PDF):",  self.activity_var,  [("PDF files", "*.pdf")], 1)
        self._file_row(files_card, "Findings Report (DOCX):", self.findings_var,  [("Word files", "*.docx")], 2)
        self._file_row(files_card, "Technical Report (PDF):", self.technical_var, [("PDF files", "*.pdf")], 3)
        self._file_row(files_card, "Executive Report (PDF):", self.executive_var, [("PDF files", "*.pdf")], 4)
        self._file_row(files_card, "Findings Details & Recommendation Summary (XLSX):", self.recommendation_var, [("XLSX files", "*.xlsx")], 5)



    # padding at bottom of card
        ctk.CTkLabel(files_card, text="").grid(row=4, column=0)

        # ── Action buttons ────────────────────────────────────────────────────
        actions = ctk.CTkFrame(content, fg_color="transparent")
        actions.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        actions.grid_columnconfigure((0, 1), weight=0)

        self.run_testing_btn = ctk.CTkButton(
            actions,
            text="▶  Run Automated Testing",
            font=ctk.CTkFont(size=12, weight="bold"),
            width=210,
            command=self._run_testing,
        )
        self.run_testing_btn.grid(row=0, column=0, padx=(0, 10))

        self.run_assessments_btn = ctk.CTkButton(
            actions,
            text="▶  Run Assessment Results",
            font=ctk.CTkFont(size=12, weight="bold"),
            width=210,
            command=self._run_assessments,
        )
        self.run_assessments_btn.grid(row=0, column=1, padx=(0, 10))

        self.run_recommendations_btn = ctk.CTkButton(
            actions,
            text="▶  Run Recommendation",
            font=ctk.CTkFont(size=12, weight="bold"),
            width=210,
            command=self._run_recommendations,
        )
        self.run_recommendations_btn.grid(row=0, column=2, padx=(0, 10))

        self.run_appendix_btn = ctk.CTkButton(
            actions,
            text="▶  Run Appendix",
            font=ctk.CTkFont(size=12, weight="bold"),
            width=210,
            command=self._run_appendix,
        )
        self.run_appendix_btn.grid(row=0, column=3, padx=(0, 10))

        self.run_findings_btn = ctk.CTkButton(
            actions,
            text="▶  Run Findings Details",
            font=ctk.CTkFont(size=12, weight="bold"),
            width=210,
            command=self._run_findings,
        )
        self.run_findings_btn.grid(row=0, column=4, padx=(0, 10))

        self.run_logistics_btn = ctk.CTkButton(
            actions,
            text="▶  Run Logistics",
            font=ctk.CTkFont(size=12, weight="bold"),
            width=210,
            command=self._run_logistics,
        )
        self.run_logistics_btn.grid(row=1, column=0, padx=(0, 10))

        self.count_vuln_btn = ctk.CTkButton(
            actions,
            text="🔍  Count Vulnerabilities",
            font=ctk.CTkFont(size=12, weight="bold"),
            width=200,
            fg_color="#16a34a",
            hover_color="#15803d",
            command=self._run_severity,
        )
        self.count_vuln_btn.grid(row=1, column=1, padx=(0, 10), pady=(10, 0), sticky="w")

        self.clear_btn = ctk.CTkButton(
            actions,
            text="✕  Clear Log",
            width=110,
            fg_color="transparent",
            border_width=1,
            text_color=("gray30", "gray70"),
            hover_color=("gray85", "gray25"),
            command=self._clear_log,
        )
        self.clear_btn.grid(row=1, column=2, padx=(0, 10), pady=(10, 0), sticky="w")

        # ── Log output ────────────────────────────────────────────────────────
        log_card = ctk.CTkFrame(content)
        log_card.grid(row=2, column=0, sticky="nsew")
        log_card.grid_columnconfigure(0, weight=1)
        log_card.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            log_card,
            text="Output Log",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=14, pady=(10, 4))

        self.log = tk.Text(
            log_card,
            font=("Cascadia Code", 9),
            wrap="word",
            state="disabled",
            relief="flat",
            borderwidth=0,
            bg="#0d1117",
            fg="#e6edf3",
            insertbackground="#e6edf3",
            selectbackground="#2563eb",
            padx=8, pady=6,
        )
        self.log.grid(row=1, column=0, sticky="nsew", padx=6, pady=(0, 6))

        scrollbar = ctk.CTkScrollbar(log_card, command=self.log.yview)
        scrollbar.grid(row=1, column=1, sticky="ns", pady=(0, 6), padx=(0, 4))
        self.log.configure(yscrollcommand=scrollbar.set)

        for tag, color in LOG_COLORS.items():
            self.log.tag_config(tag, foreground=color)
        self.log.tag_config("HEADER", foreground=LOG_COLORS["HEADER"],
                            font=("Cascadia Code", 9, "bold"))

        # ── Status bar ────────────────────────────────────────────────────────
        self.status_var = tk.StringVar(value="Ready.")
        ctk.CTkLabel(
            self, textvariable=self.status_var,
            font=ctk.CTkFont(size=10),
            anchor="w",
        ).grid(row=2, column=0, sticky="ew", padx=14, pady=(0, 4))

    def _file_row(self, parent, label_text: str, var: tk.StringVar, filetypes, row: int):
        ctk.CTkLabel(
            parent, text=label_text,
            font=ctk.CTkFont(size=11), anchor="w"
        ).grid(row=row, column=0, sticky="w", padx=(14, 8), pady=5)

        ctk.CTkEntry(parent, textvariable=var, font=ctk.CTkFont(size=11)) \
            .grid(row=row, column=1, sticky="ew", padx=(0, 8), pady=5)

        ctk.CTkButton(
            parent, text="Browse…", width=90,
            command=lambda v=var, ft=filetypes: self._browse(v, ft),
        ).grid(row=row, column=2, padx=(0, 14), pady=5)

    def _browse(self, var: tk.StringVar, filetypes):
        path = filedialog.askopenfilename(filetypes=filetypes)
        if path:
            var.set(path)

    # ── Logging helpers ───────────────────────────────────────────────────────
    def _log(self, text: str, tag: str = "PLAIN"):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log.config(state="normal")
        self.log.insert("end", f"[{ts}] ", "INFO")
        self.log.insert("end", text + "\n", tag)
        self.log.see("end")
        self.log.config(state="disabled")

    def _clear_log(self):
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")
        self.status_var.set("Log cleared.")

    # ── Redirect stdout into the log ──────────────────────────────────────────
    def _capture_stdout(self) -> io.StringIO:
        buf = io.StringIO()
        sys.stdout = buf
        return buf

    def _restore_stdout(self, buf: io.StringIO):
        sys.stdout = sys.__stdout__
        for line in buf.getvalue().splitlines():
            if line.strip():
                tag = "OK" if "✅" in line else ("WARN" if "❌" in line or "error" in line.lower() else "PLAIN")
                self._log(line, tag)

    # ── Busy state ────────────────────────────────────────────────────────────
    def _set_busy(self, busy: bool):
        state = "disabled" if busy else "normal"
        self.run_testing_btn.configure(state=state)
        self.run_assessments_btn.configure(state=state)
        self.run_recommendations_btn.configure(state=state)
        self.run_appendix_btn.configure(state=state)
        self.run_findings_btn.configure(state=state)
        self.count_vuln_btn.configure(state=state)
        self.clear_btn.configure(state=state)
        self.status_var.set("Running…" if busy else "Done.")

    # ── Actions ───────────────────────────────────────────────────────────────
    def _run_testing(self):
        activity = self.activity_var.get().strip()
        findings = self.findings_var.get().strip()
        if not activity or not findings:
            self._log("⚠  Please fill in both Activity Report and Findings Report paths.", "WARN")
            return
        self._log("── Automated Testing Activity ──", "HEADER")
        self._log(f"Activity : {activity}", "INFO")
        self._log(f"Findings : {findings}", "INFO")
        self._set_busy(True)
        prompt = ctk.CTkInputDialog(text = "Enter the page range number of the Activity Log under Activity Report: ", title = "Activity Log")
        page_range = prompt.get_input()
        if not page_range:
            self._log("⚠ Page range is required for Activity Log", "WARN")
            return
        def worker():
            buf = self._capture_stdout()
            try:
                automated_testing_activity(activity, findings, page_range)
                self._restore_stdout(buf)
                self.after(0, lambda: self._log("✅ Complete. Output saved to: " + findings, "OK"))
            except Exception as exc:
                self._restore_stdout(buf)
                self.after(0, lambda err=str(exc): self._log(f"❌ Error: {err}", "WARN"))
            finally:
                self.after(0, lambda: self._set_busy(False))

        threading.Thread(target=worker, daemon=True).start()

    def _run_assessments(self):
        executive = self.executive_var.get().strip()
        findings = self.findings_var.get().strip()
        if not executive or not findings:
            self._log("⚠  Please fill in both Executive Report and Findings Report paths.", "WARN")
            return
        self._log("── Assessment Results Summary ──", "HEADER")
        self._log(f"Executive : {executive}", "INFO")
        self._log(f"Findings : {findings}", "INFO")
        self._set_busy(True)
        prompt = ctk.CTkInputDialog(text = "Enter the page number of the Engagement Results Summary under Executive Report: ", title = "Assessment Results Summary")
        page_number = prompt.get_input()
        if not page_number:
            self._log("⚠ Page number required for Assessment Results Summary", "WARN")
            return
        def worker():
            buf = self._capture_stdout()
            try:
                assessment_results(executive, findings, page_number)
                self._restore_stdout(buf)
                self.after(0, lambda: self._log("✅ Complete. Output saved to: " + findings, "OK"))
            except Exception as exc:
                self._restore_stdout(buf)
                self.after(0, lambda err=str(exc): self._log(f"❌ Error: {err}", "WARN"))
            finally:
                self.after(0, lambda: self._set_busy(False))
        threading.Thread(target=worker, daemon=True).start()

    def _run_recommendations(self):
        recommendation = self.recommendation_var.get().strip()
        technical = self.technical_var.get().strip()
        findings = self.findings_var.get().strip()

        if not recommendation or not technical or not findings:
            self._log("⚠  Please fill in Recommendation Summary, Technical Report, and Findings Report paths.", "WARN")
            return
        self._log("── Recommendations ──", "HEADER")
        self._log(f"Recommendations : {recommendation}", "INFO")
        self._log(f"Technical: {technical}", "INFO")
        self._log(f"Findings : {findings}", "INFO")
        self._set_busy(True)
        prompt = ctk.CTkInputDialog(text = "Enter the environment for Recommendation. For example, either 'Internal' or 'External': ", title = "Recommendation")
        environment = prompt.get_input()
        if not environment:
            self._log("⚠ Environment is required for Recommendation", "WARN")
            return
        def worker():
            buf = self._capture_stdout()
            try:
                recommendations(recommendation, technical, findings, environment)
                self._restore_stdout(buf)
                self.after(0, lambda: self._log("✅ Complete. Output saved to: " + findings, "OK"))
            except Exception as exc:
                self._restore_stdout(buf)
                self.after(0, lambda err=str(exc): self._log(f"❌ Error: {err}", "WARN"))
            finally:
                self.after(0, lambda: self._set_busy(False))
        threading.Thread(target=worker, daemon=True).start()

    def _run_appendix(self):
        technical = self.technical_var.get().strip()
        findings = self.findings_var.get().strip()
        if not technical or not findings:
            self._log("⚠  Please fill in both Technical Report and Findings Report paths.", "WARN")
            return
        self._log("── Appendix ──", "HEADER")
        self._log(f"Technical: {technical}", "INFO")
        self._log(f"Findings : {findings}", "INFO")
        self._set_busy(True)
        def worker():
            buf = self._capture_stdout()
            try:
                appendix(technical, findings)
                self._restore_stdout(buf)
                self.after(0, lambda: self._log("✅ Complete. Output saved to: " + findings, "OK"))
            except Exception as exc:
                self._restore_stdout(buf)
                self.after(0, lambda err=str(exc): self._log(f"❌ Error: {err}", "WARN"))
            finally:
                self.after(0, lambda: self._set_busy(False))
        threading.Thread(target=worker, daemon=True).start()

    def _run_findings(self):
        find_info = self.recommendation_var.get().strip()
        technical = self.technical_var.get().strip()
        findings = self.findings_var.get().strip()

        if not find_info or not technical or not findings:
            self._log("⚠  Please fill in Findings Summary, Technical Report, and Findings Report paths.", "WARN")
            return
        self._log("── Findings Details ──", "HEADER")
        self._log(f"Findings Details: {find_info}", "INFO")
        self._log(f"Technical: {technical}", "INFO")
        self._log(f"Findings : {findings}", "INFO")
        self._set_busy(True)
        def worker():
            buf = self._capture_stdout()
            try:
                finding_details(technical, findings, find_info, "External")
                self._restore_stdout(buf)
                self.after(0, lambda: self._log("✅ Complete. Output saved to: " + findings, "OK"))
            except Exception as exc:
                self._restore_stdout(buf)
                self.after(0, lambda err=str(exc): self._log(f"❌ Error: {err}", "WARN"))
            finally:
                self.after(0, lambda: self._set_busy(False))
        threading.Thread(target=worker, daemon=True).start()

    def _run_logistics(self):
        technical = self.technical_var.get().strip()
        findings = self.findings_var.get().strip()
        if not technical or not findings:
            self._log("⚠  Please fill in both Technical Report and Findings Report paths.", "WARN")
            return
        self._log("── Logistics ──", "HEADER")
        self._log(f"Technical: {technical}", "INFO")
        self._log(f"Findings : {findings}", "INFO")
        self._set_busy(True)
        def worker():
            buf = self._capture_stdout()
            try:
                Logistics(technical, findings)
                self._restore_stdout(buf)
                self.after(0, lambda: self._log("✅ Complete. Output saved to: " + findings, "OK"))
            except Exception as exc:
                self._restore_stdout(buf)
                self.after(0, lambda err=str(exc): self._log(f"❌ Error: {err}", "WARN"))
            finally:
                self.after(0, lambda: self._set_busy(False))
        threading.Thread(target=worker, daemon=True).start()

    def _run_severity(self):
        technical = self.technical_var.get().strip()
        if not technical:
            self._log("⚠  Please fill in the Technical Report path.", "WARN")
            return
        self._log("── Vulnerability Count ──", "HEADER")
        self._log(f"Technical : {technical}", "INFO")
        self._set_busy(True)

        def worker():
            buf = self._capture_stdout()
            try:
                frames = severity_counter(technical)
                self._restore_stdout(buf)
                counts: dict[str, int] = {}
                for f in frames:
                    counts[f.rating] = counts.get(f.rating, 0) + 1
                self.after(0, lambda: self._log(f"Found {len(frames)} total vulnerabilities:", "OK"))
                for sev in ["Critical", "High", "Medium", "Low", "Informational"]:
                    if sev in counts:
                        self.after(0, lambda s=sev, c=counts[sev]: self._log(f"  {s:<16} {c}", "PLAIN"))
                for frame in frames:
                    self.after(0, lambda f=frame: self._log(f"  [{f.rating}]  {f.discoveredName}", "PLAIN"))
            except Exception as exc:
                self._restore_stdout(buf)
                self.after(0, lambda err=str(exc): self._log(f"❌ Error: {err}", "WARN"))
            finally:
                self.after(0, lambda: self._set_busy(False))

        threading.Thread(target=worker, daemon=True).start()

    # ── Theme toggle ──────────────────────────────────────────────────────────
    def _toggle_theme(self):
        current = ctk.get_appearance_mode().lower()
        new_mode = "dark" if current == "light" else "light"
        ctk.set_appearance_mode(new_mode)
        self.theme_btn.configure(text="☀  Light Mode" if new_mode == "dark" else "🌙  Dark Mode")

# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()
