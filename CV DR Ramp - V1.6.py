# -*- coding: utf-8 -*-
"""
DR Tool CC – Matias v1.6
Stable CC ramp GUI with:
- Direct VISA connection to EL34143 and WT1800
- Linear Ramp / Custom List
- Up to 4 decimals, truncated (never rounded)
- One measurement per step
- Full dwell countdown
- Multiple ramps before saving
- Blank row separator between ramps
- Live log + progress bar
- Original Excel output format

Author: Matias Riquelme, CPI
"""

import os
import sys
import time
import threading
import queue
from datetime import datetime
from decimal import Decimal, ROUND_DOWN, InvalidOperation
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import pyvisa
import xlsxwriter

DWELL_SECONDS_DEFAULT = 20


def exec_dir() -> str:
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.abspath('.')


def truncate_4(value) -> Decimal:
    """
    Truncate to 4 decimals without rounding.
    """
    return Decimal(str(value)).quantize(Decimal("0.0001"), rounding=ROUND_DOWN)


def dec_to_str(d: Decimal) -> str:
    """
    Convert Decimal to plain string without scientific notation,
    keeping up to 4 decimals, trimming only trailing zeros.
    """
    s = format(d, "f")
    if "." in s:
        s = s.rstrip("0").rstrip(".")
    return s if s else "0"


class DRToolApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DR Tool CC – Matias v1.6")
        self.geometry("920x700")
        self.minsize(920, 700)

        # ---- Theme ----
        self.configure(bg="#f8fafc")
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        accent_blue = "#2563eb"
        accent_green = "#16a34a"
        accent_red = "#dc2626"
        text_color = "#1e293b"

        style.configure("TLabel", foreground=text_color, background="#f8fafc", font=("Segoe UI", 11))
        style.configure("Header.TLabelframe", background="#f1f5f9", borderwidth=1, relief="solid")
        style.configure("Header.TLabelframe.Label", foreground=accent_blue, background="#f1f5f9", font=("Segoe UI", 12, "bold"))
        style.configure("Card.TLabelframe", background="#f1f5f9", borderwidth=1, relief="solid")
        style.configure("Card.TLabelframe.Label", foreground=accent_blue, background="#f1f5f9", font=("Segoe UI", 12, "bold"))
        style.configure("TEntry", fieldbackground="#ffffff", foreground=text_color)
        style.configure("TButton", font=("Segoe UI", 10, "bold"))
        style.configure("Green.TButton", background=accent_green)
        style.map("Green.TButton", background=[("active", "#15803d")])
        style.configure("Accent.TButton", background=accent_blue)
        style.map("Accent.TButton", background=[("active", "#1d4ed8")])
        style.configure("Danger.TButton", background=accent_red)
        style.map("Danger.TButton", background=[("active", "#b91c1c")])

        # ---- Layout ----
        top = ttk.LabelFrame(self, text="Ramp Parameters (CC)", style="Header.TLabelframe")
        top.pack(fill="x", padx=14, pady=(14, 8))

        mid = ttk.LabelFrame(self, text="Control & Instruments", style="Card.TLabelframe")
        mid.pack(fill="x", padx=14, pady=8)

        bot = ttk.LabelFrame(self, text="Live Log", style="Card.TLabelframe")
        bot.pack(fill="both", expand=True, padx=14, pady=(8, 14))

        # ---- Parameters ----
        pgrid = ttk.Frame(top, padding=10)
        pgrid.pack(fill="x")

        self.mode = tk.StringVar(value="linear")

        self.var_i_start = tk.StringVar(value="0.1000")
        self.var_i_end = tk.StringVar(value="1.0000")
        self.var_i_step = tk.StringVar(value="0.1000")
        self.var_custom = tk.StringVar(value="")
        self.var_dwell = tk.StringVar(value=str(DWELL_SECONDS_DEFAULT))

        ttk.Radiobutton(
            pgrid, text="Linear Ramp", variable=self.mode, value="linear",
            command=self.update_mode_ui
        ).grid(row=0, column=0, sticky="w", padx=6, pady=(0, 8))

        ttk.Radiobutton(
            pgrid, text="Custom List", variable=self.mode, value="custom",
            command=self.update_mode_ui
        ).grid(row=0, column=1, sticky="w", padx=6, pady=(0, 8))

        def row(r, label, var, unit):
            ttk.Label(pgrid, text=label).grid(row=r, column=0, sticky="e", padx=6, pady=4)
            ent = ttk.Entry(pgrid, textvariable=var, width=14)
            ent.grid(row=r, column=1, sticky="w", padx=4)
            ttk.Label(pgrid, text=unit).grid(row=r, column=2, sticky="w", padx=4)
            return ent

        self.ent_start = row(1, "I start", self.var_i_start, "A")
        self.ent_end = row(2, "I end", self.var_i_end, "A")
        self.ent_step = row(3, "I step", self.var_i_step, "A")

        ttk.Label(pgrid, text="Custom List").grid(row=4, column=0, sticky="e", padx=6, pady=4)
        self.ent_custom = ttk.Entry(pgrid, textvariable=self.var_custom, width=50)
        self.ent_custom.grid(row=4, column=1, columnspan=2, sticky="we", padx=4)

        ttk.Label(pgrid, text="Delay").grid(row=5, column=0, sticky="e", padx=6, pady=4)
        self.ent_dwell = ttk.Entry(pgrid, textvariable=self.var_dwell, width=14)
        self.ent_dwell.grid(row=5, column=1, sticky="w", padx=4)
        ttk.Label(pgrid, text="s").grid(row=5, column=2, sticky="w", padx=4)

        ttk.Label(
            top,
            text="Supports up to 4 decimals (truncated, not rounded)",
            foreground=text_color,
            background="#f1f5f9",
            font=("Segoe UI", 10, "italic")
        ).pack(anchor="w", padx=20, pady=(2, 8))

        # ---- Controls ----
        cgrid = ttk.Frame(mid, padding=10)
        cgrid.pack(fill="x")
        self.btn_connect = ttk.Button(cgrid, text="Connect", style="Accent.TButton", command=self.on_connect)
        self.btn_start = ttk.Button(cgrid, text="Start", style="Green.TButton", command=self.on_start, state="disabled")
        self.btn_stop = ttk.Button(cgrid, text="Stop", style="Danger.TButton", command=self.on_stop, state="disabled")
        self.btn_connect.pack(side="left", padx=6)
        self.btn_start.pack(side="left", padx=6)
        self.btn_stop.pack(side="left", padx=6)

        self.var_status = tk.StringVar(value="Idle")
        self.lbl_status = ttk.Label(cgrid, textvariable=self.var_status, font=("Segoe UI", 12, "bold"))
        self.lbl_status.pack(side="left", padx=12)

        self.prog = ttk.Progressbar(cgrid, mode="determinate", length=240)
        self.prog.pack(side="right", padx=8)

        # ---- Live log ----
        logwrap = tk.Frame(bot, bg="#ffffff")
        logwrap.pack(fill="both", expand=True, padx=8, pady=8)
        self.txt = tk.Text(
            logwrap,
            height=18,
            bg="#ffffff",
            fg="#1e293b",
            insertbackground="#1e293b",
            relief="flat",
            font=("Consolas", 11),
        )
        self.txt.pack(fill="both", expand=True)
        self.txt.configure(state="disabled")

        # ---- Runtime state ----
        self.running = False
        self.worker = None
        self.msg_q = queue.Queue()
        self.after(100, self._drain_log)

        # ---- VISA / Instrument state ----
        self.rm_load = None
        self.rm_wt = None
        self.load = None
        self.wt = None
        self.el_connected = False
        self.wt_connected = False
        self.connected = False

        # ---- Results ----
        self.current_results = []
        self.all_results = []

        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.update_mode_ui()

    # --------------------------------------------------
    # Thread-safe log
    # --------------------------------------------------
    def _log(self, msg: str):
        self.msg_q.put(msg)

    def _drain_log(self):
        try:
            while True:
                msg = self.msg_q.get_nowait()
                self.txt.configure(state="normal")
                self.txt.insert("end", msg + "\n")
                self.txt.see("end")
                self.txt.configure(state="disabled")
        except queue.Empty:
            pass
        self.after(100, self._drain_log)

    def set_status(self, text: str):
        self.var_status.set(text)
        self.update_idletasks()

    # --------------------------------------------------
    # UI mode switch
    # --------------------------------------------------
    def update_mode_ui(self):
        if self.mode.get() == "linear":
            self.ent_start.config(state="normal")
            self.ent_end.config(state="normal")
            self.ent_step.config(state="normal")
            self.ent_custom.config(state="disabled")
        else:
            self.ent_start.config(state="disabled")
            self.ent_end.config(state="disabled")
            self.ent_step.config(state="disabled")
            self.ent_custom.config(state="normal")

    # --------------------------------------------------
    # VISA / Instrument helpers
    # --------------------------------------------------
    def connect_load(self):
        self.rm_load = pyvisa.ResourceManager()
        self.load = self.rm_load.open_resource('USB0::0x2A8D::0x3802::MY60260348::INSTR')
        self.load.timeout = 10000
        idn = self.load.query("*IDN?")
        self._log(f"[OK] EL34143 connected: {idn.strip()}")
        self.load.write("*CLS")
        self.load.write("*RST")
        self.el_connected = True

    def connect_wt(self):
        self.rm_wt = pyvisa.ResourceManager()
        self.wt = self.rm_wt.open_resource('USB0::0x0B21::0x0025::43325247323730303756::INSTR')
        self.wt.timeout = 10000
        idn = self.wt.query("*IDN?")
        self._log(f"[OK] WT1800 connected: {idn.strip()}")
        self.wt.write(":NUMERIC:FORMAT ASCII")
        self.wt.write(":INPUT:VOLTAGE:AUTO:ELEMENT1 ON")
        self.wt.write(":INPUT:CURRENT:AUTO:ELEMENT1 ON")
        self.wt_connected = True

    def load_prepare_cc(self, max_current: Decimal):
        """
        Prepare CC mode once per ramp.
        Choose range once, based on max current of the whole ramp.
        """
        self.load.write("FUNC CURR, (@1)")

        if max_current <= Decimal("0.6120"):
            self.load.write("CURR:RANG 0.612, (@1)")
            self._log("[INFO] EL34143 current range set to 0.612 A")
        else:
            self.load.write("CURR:RANG 6.12, (@1)")
            self._log("[INFO] EL34143 current range set to 6.12 A")

    def load_set_current(self, current: Decimal):
        """
        Set only the current point.
        Do not reconfigure mode/range on every step.
        """
        self.load.write(f"CURR {dec_to_str(current)}, (@1)")

    def load_on(self):
        if self.load is not None:
            self.load.write("INP ON, (@1)")

    def load_off(self):
        if self.load is not None:
            self.load.write("INP OFF, (@1)")

    def load_read_voltage(self) -> float:
        return float(self.load.query("MEAS:VOLT? (@1)"))

    def load_read_current(self) -> float:
        return float(self.load.query("MEAS:CURR? (@1)"))

    def load_read_power(self) -> float:
        return float(self.load.query("MEAS:POW? (@1)"))

    def wt_read_urms(self) -> float:
        self.wt.write(":NUMERIC:NORMAL:ITEM1 URMS,1")
        return float(self.wt.query(":NUMeric:NORMal:VALue? 1"))

    def wt_read_irms(self) -> float:
        self.wt.write(":NUMERIC:NORMAL:ITEM2 IRMS,1")
        return float(self.wt.query(":NUMeric:NORMal:VALue? 2"))

    def wt_read_power(self) -> float:
        self.wt.write(":NUMERIC:NORMAL:ITEM3 P,1")
        return float(self.wt.query(":NUMeric:NORMal:VALue? 3"))

    def wt_read_thd(self) -> float:
        self.wt.write(":NUMERIC:NORMAL:ITEM4 ITHD,1")
        return float(self.wt.query(":NUMeric:NORMal:VALue? 4"))

    def wt_read_pf(self) -> float:
        self.wt.write(":NUMERIC:NORMAL:ITEM5 LAMbda,1")
        return float(self.wt.query(":NUMeric:NORMal:VALue? 5"))

    def disconnect_instruments(self):
        try:
            self.load_off()
        except Exception:
            pass
        try:
            if self.wt is not None:
                self.wt.write(":COMMunicate:REMote OFF")
        except Exception:
            pass
        try:
            if self.wt is not None:
                self.wt.close()
        except Exception:
            pass
        try:
            if self.load is not None:
                self.load.close()
        except Exception:
            pass

        self.wt = None
        self.load = None
        self.connected = False
        self.el_connected = False
        self.wt_connected = False

    # --------------------------------------------------
    # Build sequence
    # --------------------------------------------------
    def build_linear_sequence(self, start_txt: str, end_txt: str, step_txt: str):
        start = truncate_4(start_txt)
        end = truncate_4(end_txt)
        step = truncate_4(step_txt)

        if step == Decimal("0.0000"):
            raise ValueError("Step cannot be 0.")

        values = []

        if start <= end:
            current = start
            while current <= end:
                values.append(truncate_4(current))
                current = truncate_4(current + step)

            if not values or values[-1] != end:
                values.append(end)
        else:
            current = start
            while current >= end:
                values.append(truncate_4(current))
                current = truncate_4(current - step)

            if not values or values[-1] != end:
                values.append(end)

        return values

    def parse_custom_list(self, text: str):
        values = []
        for item in text.split(","):
            item = item.strip()
            if item:
                values.append(truncate_4(item))

        if not values:
            raise ValueError("Custom List is empty.")

        return values

    def build_current_sequence(self):
        if self.mode.get() == "linear":
            return self.build_linear_sequence(
                self.var_i_start.get(),
                self.var_i_end.get(),
                self.var_i_step.get()
            )
        else:
            return self.parse_custom_list(self.var_custom.get())

    # --------------------------------------------------
    # Measurement helper
    # --------------------------------------------------
    def measure_one_point(self, step_index: int, current_set: Decimal):
        v_out = self.load_read_voltage()
        time.sleep(0.2)
        i_out = self.load_read_current()
        time.sleep(0.2)
        p_out = self.load_read_power()

        time.sleep(0.2)
        v_in = self.wt_read_urms()
        i_in = self.wt_read_irms()
        p_in = self.wt_read_power()
        pf = self.wt_read_pf()
        thd = self.wt_read_thd()

        eff = (p_out / p_in * 100) if p_in > 0 else 0.0

        row = {
            "Step": step_index,
            "I set (A)": float(current_set),
            "U in (V)": v_in,
            "I in (A)": i_in,
            "P in (W)": p_in,
            "V out (V)": v_out,
            "I out (A)": i_out,
            "P out (W)": p_out,
            "Efficiency(%)": round(eff, 4),
            "THD(%)": thd,
            "Power Factor": pf,
        }
        return row

    # --------------------------------------------------
    # Actions
    # --------------------------------------------------
    def on_connect(self):
        try:
            self.set_status("Connecting…")

            if self.connected:
                self._log("[INFO] Instruments already connected.")
                self.set_status("Connected")
                self.btn_start.config(state="normal")
                return

            self.connect_load()
            self.connect_wt()

            self.connected = True
            self.set_status("Connected")
            self._log("[OK] Instruments ready")
            self.btn_start.config(state="normal")

        except Exception as e:
            self.connected = False
            self.set_status("Idle")
            messagebox.showerror("Connection Error", str(e))
            self._log(f"[ERROR] Connection: {e}")

    def on_start(self):
        if not self.connected:
            messagebox.showwarning("Connection", "Connect instruments first.")
            return

        try:
            currents = self.build_current_sequence()
            if not currents:
                messagebox.showwarning("Parameters", "No valid points generated.")
                return

            dwell = float(self.var_dwell.get())
            if dwell < 0:
                messagebox.showwarning("Parameters", "Delay must be >= 0.")
                return

        except (ValueError, InvalidOperation) as e:
            messagebox.showwarning("Parameters", f"Invalid parameters.\n\n{e}")
            return

        self.running = True
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.btn_connect.config(state="disabled")
        self.prog.configure(value=0, maximum=len(currents))
        self.set_status("Running…")

        self.worker = threading.Thread(target=self._run_sequence, args=(currents, dwell), daemon=True)
        self.worker.start()

    def on_stop(self):
        self.running = False
        self.btn_stop.config(state="disabled")
        self.set_status("Stopping…")
        self._log("[INFO] Stop requested by user.")
        try:
            self.load_off()
        except Exception:
            pass

    def on_close(self):
        self.running = False
        try:
            self.disconnect_instruments()
        except Exception:
            pass
        self.destroy()

    # --------------------------------------------------
    # Core sequence
    # --------------------------------------------------
    def _run_sequence(self, currents, dwell):
        try:
            self.current_results = []

            max_current = max(currents)
            self._log(f"[INFO] Current sequence: {[dec_to_str(x) for x in currents]}")
            self._log(f"[INFO] Max current in ramp: {dec_to_str(max_current)} A")

            # Prepare CC only once per ramp
            self.load_prepare_cc(max_current)

            # First point
            first = currents[0]
            self.load_set_current(first)
            self.load_on()
            self._log(f"[STEP 1/{len(currents)}] Applying {dec_to_str(first)} A in CC")

            t0 = time.time()
            while self.running and (time.time() - t0) < dwell:
                remaining = max(0, int(dwell - (time.time() - t0) + 0.999))
                self.set_status(f"Capturing measurement… {remaining}s")
                time.sleep(0.2)

            if self.running:
                row = self.measure_one_point(1, first)
                self.current_results.append(row)
                self._log(
                    f"  -> Measurement: "
                    f"Set={row['I set (A)']:.4f} A | "
                    f"Vin={row['U in (V)']:.3f} V, Iin={row['I in (A)']:.3f} A, Pin={row['P in (W)']:.3f} W | "
                    f"Vout={row['V out (V)']:.3f} V, Iout={row['I out (A)']:.3f} A, Pout={row['P out (W)']:.3f} W | "
                    f"Eff={row['Efficiency(%)']:.2f}% THD={row['THD(%)']:.2f}% PF={row['Power Factor']:.3f}"
                )
                self.prog.configure(value=1)

            # Remaining points
            for idx, current in enumerate(currents[1:], start=2):
                if not self.running:
                    break

                self._log(f"[STEP {idx}/{len(currents)}] Applying {dec_to_str(current)} A in CC")
                self.load_set_current(current)

                t0 = time.time()
                while self.running and (time.time() - t0) < dwell:
                    remaining = max(0, int(dwell - (time.time() - t0) + 0.999))
                    self.set_status(f"Capturing measurement… {remaining}s")
                    time.sleep(0.2)

                if not self.running:
                    break

                row = self.measure_one_point(idx, current)
                self.current_results.append(row)
                self._log(
                    f"  -> Measurement: "
                    f"Set={row['I set (A)']:.4f} A | "
                    f"Vin={row['U in (V)']:.3f} V, Iin={row['I in (A)']:.3f} A, Pin={row['P in (W)']:.3f} W | "
                    f"Vout={row['V out (V)']:.3f} V, Iout={row['I out (A)']:.3f} A, Pout={row['P out (W)']:.3f} W | "
                    f"Eff={row['Efficiency(%)']:.2f}% THD={row['THD(%)']:.2f}% PF={row['Power Factor']:.3f}"
                )
                self.prog.configure(value=idx)

            self.load_off()

            if self.running and self.current_results:
                if self.all_results:
                    self.all_results.append(None)
                self.all_results.extend(self.current_results)

                self.set_status("Ramp finished")
                self._log("[OK] Ramp finished successfully.")

                save_now = messagebox.askyesno(
                    "Finish measurements",
                    "The current ramp has finished.\n\n"
                    "Did you finish your measurements and want to save ALL accumulated results?\n\n"
                    "If you answer 'No', you can run another ramp and keep accumulating data."
                )

                if save_now:
                    self.save_measurements_to_excel_original_format()
                    self.all_results.clear()
                    self.current_results.clear()
                    self._log("[OK] Results saved and memory cleared.")
                    messagebox.showinfo("Done", "Measurements saved successfully.")
                else:
                    self._log("[INFO] Results kept in memory for another ramp.")

                self.set_status("Connected")

            elif not self.running:
                self.set_status("Stopped")
                self._log("[INFO] Test stopped by user.")

        except Exception as e:
            self.set_status("Error")
            self._log(f"[ERROR] {e}")
            messagebox.showerror("Runtime Error", str(e))
            try:
                self.load_off()
            except Exception:
                pass

        finally:
            self.running = False
            self.btn_stop.config(state="disabled")
            self.btn_start.config(state="normal")
            self.btn_connect.config(state="normal")
            if self.connected:
                self.set_status("Connected")
            else:
                self.set_status("Idle")

    # --------------------------------------------------
    # Excel output (original format)
    # --------------------------------------------------
    def save_measurements_to_excel_original_format(self):
        if not self.all_results:
            messagebox.showwarning("No Data", "There is no accumulated data to save.")
            return

        time_display_excel = time.asctime(time.localtime(time.time()))
        safe_time = time_display_excel.replace(" ", "_").replace(":", "_")
        default_name = f"Export_file_{safe_time}_.xlsx"

        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Save measurement results as...",
            initialfile=default_name
        )

        if not filepath:
            messagebox.showwarning("Cancelled", "Results file was not saved.")
            return

        workbook = xlsxwriter.Workbook(filepath)
        worksheet = workbook.add_worksheet()

        worksheet.write('B3', 'Date')
        worksheet.write('C3', time_display_excel)
        worksheet.write('B4', 'Manufacturer')
        worksheet.write('B5', 'Model')
        worksheet.write('B6', 'Description')
        worksheet.write('B7', 'Current')
        worksheet.write('B8', 'Voltage range')
        worksheet.write('B9', 'Dimming')
        worksheet.write('B10', 'Setpoints')

        number = 14
        worksheet.write('B' + str(number), 'Test ')
        worksheet.write('D' + str(number), 'U in (V)')
        worksheet.write('E' + str(number), 'I in (A)')
        worksheet.write('F' + str(number), 'P in (W)')
        worksheet.write('G' + str(number), 'V out (V)')
        worksheet.write('H' + str(number), 'I out (A)')
        worksheet.write('I' + str(number), 'P out (W)')
        worksheet.write('J' + str(number), 'Efficiency(%)')
        worksheet.write('K' + str(number), 'THD(%)')
        worksheet.write('L' + str(number), ' Power Factor ')

        row_xls = 15

        for r in self.all_results:
            if r is None:
                row_xls += 1
                continue

            worksheet.write(row_xls, 3, r["U in (V)"])
            worksheet.write(row_xls, 4, r["I in (A)"])
            worksheet.write(row_xls, 5, r["P in (W)"])
            worksheet.write(row_xls, 6, r["V out (V)"])
            worksheet.write(row_xls, 7, r["I out (A)"])
            worksheet.write(row_xls, 8, r["P out (W)"])
            worksheet.write(row_xls, 9, r["Efficiency(%)"])
            worksheet.write(row_xls, 10, r["THD(%)"])
            worksheet.write(row_xls, 11, r["Power Factor"])
            row_xls += 1

        workbook.close()
        self._log(f"[OK] Results saved to: {filepath}")


if __name__ == "__main__":
    app = DRToolApp()
    app.mainloop()