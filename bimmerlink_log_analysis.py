import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import csv
import pandas as pd

import matplotlib
matplotlib.use("Agg")

import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import webbrowser

'''pip install pandas matplotlib'''

'''
win executable
pip install pyinstaller pandas matplotlib
pyinstaller --onefile --windowed --name BimmerLogAnalyzer bimmerlink_log_analysis.py
'''

def get_app_dir() -> str:

    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

class BimmerLogAnalyzerGUI(tk.Tk):

    def __init__(self):

        super().__init__()

        self.title("BMW/Mini BimmerLink, Log Analyzer v1.0")
        self.geometry("550x350")
        self.resizable(False, False)

        try:
            self.iconbitmap(default=None)
        except Exception:
            pass

        self.selected_csv_path = tk.StringVar(value="")

        self.export_pdf_path = tk.StringVar(value="")
        self.export_txt_path = tk.StringVar(value="")

        self._build_widgets()

    def _build_widgets(self):

        main_frame = tk.Frame(self, padx=20, pady=20)
        main_frame.pack(fill="both", expand=True)

        lbl_title = tk.Label(
            main_frame,
            text="BMW/Mini BimmerLink, Log Analyzer v1.0",
            font=("Segoe UI", 11, "bold")
        )
        lbl_title.grid(row=0, column=0, columnspan=3, sticky="w")

        lbl_import = tk.Label(
            main_frame,
            text="Import Exported Log .csv",
            font=("Segoe UI", 10)
        )
        lbl_import.grid(row=1, column=0, sticky="w", pady=(20, 0))

        btn_select = tk.Button(
            main_frame,
            text="Select",
            width=10,
            command=self.select_csv
        )
        btn_select.grid(row=1, column=1, sticky="w", padx=(10, 0), pady=(20, 0))

        self.btn_run = tk.Button(
            main_frame,
            text="RUN",
            width=10,
            state="disabled",
            command=self.run_analysis
        )
        self.btn_run.grid(row=1, column=2, sticky="w", padx=(10, 0), pady=(20, 0))

        lbl_selected_path = tk.Label(
            main_frame,
            textvariable=self.selected_csv_path,
            font=("Consolas", 9),
            fg="#333333",
            wraplength=500,
            justify="left"
        )
        lbl_selected_path.grid(row=2, column=0, columnspan=3, sticky="w", pady=(5, 15))

        lbl_export_title = tk.Label(
            main_frame,
            text="Report Export Location",
            font=("Segoe UI", 10)
        )
        lbl_export_title.grid(row=3, column=0, sticky="w")

        lbl_export_pdf = tk.Label(
            main_frame,
            textvariable=self.export_pdf_path,
            font=("Consolas", 9),
            fg="#333333",
            wraplength=500,
            justify="left"
        )
        lbl_export_pdf.grid(row=4, column=0, columnspan=3, sticky="w", pady=(5, 0))

        lbl_export_txt = tk.Label(
            main_frame,
            textvariable=self.export_txt_path,
            font=("Consolas", 9),
            fg="#333333",
            wraplength=500,
            justify="left"
        )
        lbl_export_txt.grid(row=5, column=0, columnspan=3, sticky="w", pady=(5, 0))

        link_label = tk.Label(
            main_frame,
            text="https://blog.armanasci.com",
            fg="blue",
            cursor="hand2",
            font=("Segoe UI", 9, "underline")
        )
        link_label.grid(row=7, column=0, columnspan=3, pady=(20, 0), sticky="w")
        link_label.bind("<Button-1>", self.open_link)

    def select_csv(self):

        file_path = filedialog.askopenfilename(
            title="Select exported log (.csv)",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )

        if not file_path:
            return

        _, ext = os.path.splitext(file_path)
        if ext.lower() != ".csv":
            messagebox.showerror(
                "Invalid file",
                "Please select CSV file."
            )
            return

        self.selected_csv_path.set(file_path)

        import_dir = os.path.dirname(file_path)

        base_name = os.path.splitext(os.path.basename(file_path))[0]

        pdf_path = os.path.join(import_dir, f"{base_name}_bimmerlink_report.pdf")
        txt_path = os.path.join(import_dir, f"{base_name}_bimmerlink_sensorlist.txt")

        self.export_pdf_path.set(pdf_path)
        self.export_txt_path.set(txt_path)

        self.btn_run.config(state="normal")

    def generate_pdf_report(self, csv_path: str, pdf_path: str):

        df = pd.read_csv(csv_path)

        time_col = "Time"
        engine_col = "Engine speed"

        if time_col not in df.columns or engine_col not in df.columns:
            raise ValueError("Time or Engine speed column count not found.")

        time = df[time_col]
        engine = df[engine_col]

        sensor_cols = [c for c in df.columns if c not in (time_col, engine_col)]

        # PDF writer
        with PdfPages(pdf_path) as pdf:
            for col in sensor_cols:
                series = pd.to_numeric(df[col], errors="coerce")

                min_val = series.min()
                max_val = series.max()
                avg_val = series.mean()

                if series.dropna().nunique() < 2:
                    continue

                step = max(len(df) // 1000, 1)
                idx = slice(None, None, step)
                t_plot = time.iloc[idx]
                s_plot = series.iloc[idx]
                e_plot = engine.iloc[idx]

                fig, ax1 = plt.subplots(figsize=(10, 4))

                ax1.plot(t_plot, s_plot, color="orange", label=f"{col}")
                ax1.set_xlabel("Time")
                ax1.set_ylabel(col)

                ax2 = ax1.twinx()
                ax2.plot(t_plot, e_plot, color="blue", label="Engine speed (rpm)")
                ax2.set_ylabel("Engine speed (rpm)")

                fig.text(
                    0.5,
                    0.02,
                    f"Min: {min_val:.2f}    Max: {max_val:.2f}    Average: {avg_val:.2f}",
                    ha="center",
                    fontsize=9,
                )

                lines = ax1.get_lines() + ax2.get_lines()
                labels = [line.get_label() for line in lines]
                fig.legend(
                    lines,
                    labels,
                    loc="lower center",
                    ncol=2,
                    bbox_to_anchor=(0.5, -0.12),
                )

                plt.title(f"{col} vs Engine speed")
                plt.tight_layout()

                pdf.savefig(fig)
                plt.close(fig)

    def open_link(self, _event=None):
        webbrowser.open("https://blog.armanasci.com")

    def run_analysis(self):

        csv_path = self.selected_csv_path.get().strip()
        txt_path = self.export_txt_path.get().strip()

        if not csv_path:
            messagebox.showwarning("Uyarı", "Lütfen önce bir CSV dosyası seçin.")
            return

        if not txt_path:
            messagebox.showwarning("Uyarı", "TXT export yolu bulunamadı.")
            return

        try:

            with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
                reader = csv.reader(f)
                header = next(reader, None)

            if header is None:
                messagebox.showerror("Error", "CSV is empty.")
                return

            with open(txt_path, "w", encoding="utf-8") as out_f:
                for col in header:
                    col_name = (col or "").strip()
                    out_f.write(col_name + "\n")

            pdf_path = self.export_pdf_path.get().strip()
            if pdf_path:
                self.generate_pdf_report(csv_path, pdf_path)

            messagebox.showinfo(
                "Success",
                f"Sensor/column list written to TXT file:\n{txt_path}\n\n"
                f"Graphical report generated as PDF:\n{pdf_path}"
            )

        except Exception as e:
            messagebox.showerror(
                "Error",
                f"An error occurred during processing:\n{e}"
            )


if __name__ == "__main__":
    app = BimmerLogAnalyzerGUI()
    app.mainloop()
