"""
HMP4040 Ramp Frontend (CV/CC) with:
- CONNECT separate from START
- Wiring-aware control: SINGLE / SERIES / PARALLEL
- SERIES robust split: CH1 up to 32 V, remainder on CH2
- Manual HOLD between steps (PSU stays ON at setpoint): waits for NEXT STEP
- Auto snapshot AFTER stabilization time (before HOLD)
- Snapshots measured from PSU: V, A, P (no calculated power)
- Save results to Excel (.xlsx)

Install:
  python -m pip install pyvisa pyvisa-py openpyxl
"""

import threading
import time
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from openpyxl import Workbook


# ---------------------------
# Backend: R&S HMP4040 via PyVISA (auto-detect)
# ---------------------------
class HMP4040Controller:
    def __init__(self):
        self.rm = None
        self.inst = None
        self.idn = ""

    def connect(self) -> bool:
        try:
            import pyvisa
        except Exception as e:
            raise RuntimeError("Missing pyvisa. Install: python -m pip install pyvisa pyvisa-py") from e

        self.rm = pyvisa.ResourceManager()
        resources = list(self.rm.list_resources())
        if not resources:
            raise RuntimeError("No VISA resources found. Check USB/LAN and VISA installation.")

        last_err = None
        for r in resources:
            try:
                inst = self.rm.open_resource(r)
                inst.timeout = 5000
                inst.write_termination = "\n"
                inst.read_termination = "\n"
                idn = inst.query("*IDN?").strip()
                if ("HMP4040" in idn) or ("ROHDE" in idn) or ("SCHWARZ" in idn):
                    self.inst = inst
                    self.idn = idn
                    return True
                inst.close()
            except Exception as e:
                last_err = e
                continue

        raise RuntimeError(f"Could not detect HMP4040 via VISA. Last error: {last_err}")

    def disconnect(self):
        try:
            if self.inst:
                self.inst.close()
        finally:
            self.inst = None
            if self.rm:
                try:
                    self.rm.close()
                except Exception:
                    pass
                self.rm = None

    def _write(self, cmd: str):
        if not self.inst:
            raise RuntimeError("Instrument not connected.")
        self.inst.write(cmd)

    def _query(self, cmd: str) -> str:
        if not self.inst:
            raise RuntimeError("Instrument not connected.")
        return self.inst.query(cmd)

    def select_channel(self, ch: int):
        self._write(f"INST OUT{ch}")

    def set_output(self, ch: int, on: bool):
        self.select_channel(ch)
        self._write(f"OUTP {'ON' if on else 'OFF'}")

    def apply_cv(self, ch: int, voltage: float, current_set: float):
        self.select_channel(ch)
        self._write(f"VOLT {voltage}")
        self._write(f"CURR {current_set}")

    def apply_cc(self, ch: int, current_set: float, voltage_set: float):
        self.select_channel(ch)
        self._write(f"CURR {current_set}")
        self._write(f"VOLT {voltage_set}")

    def read_meas_voltage(self, ch: int) -> float:
        self.select_channel(ch)
        return float(self._query("MEAS:VOLT?").strip())

    def read_meas_current(self, ch: int) -> float:
        self.select_channel(ch)
        return float(self._query("MEAS:CURR?").strip())

    def read_meas_power(self, ch: int) -> float:
        """
        Try common power readback commands; adjust if your firmware uses another.
        """
        self.select_channel(ch)
        candidates = [
            "MEAS:POW?",
            "MEAS:POWer?",
            "MEAS:POW:DC?",
            "MEAS:POWer:DC?",
        ]
        last_err = None
        for cmd in candidates:
            try:
                return float(self._query(cmd).strip())
            except Exception as e:
                last_err = e
                continue
        raise RuntimeError(
            f"Could not read POWER using {candidates}. "
            f"Your HMP4040 may use different SCPI. Last error: {last_err}"
        )

    def off_all(self):
        for ch in (1, 2):
            try:
                self.set_output(ch, False)
            except Exception:
                pass


# ---------------------------
# Helpers
# ---------------------------
def frange(start, stop, step):
    vals = []
    if step == 0:
        return vals
    x = start
    if start <= stop:
        while x <= stop + 1e-12:
            vals.append(round(x, 6))
            x += step
    else:
        while x >= stop - 1e-12:
            vals.append(round(x, 6))
            x -= abs(step)
    return vals


def compute_totals(wiring_mode, v1, i1, p1, v2=0.0, i2=0.0, p2=0.0):
    if wiring_mode == "SINGLE":
        return v1, i1, p1
    if wiring_mode == "SERIES":
        v_total = v1 + v2
        i_total = (i1 + i2) / 2.0 if abs(i2) > 1e-12 else i1
        p_total = p1 + p2
        return v_total, i_total, p_total
    if wiring_mode == "PARALLEL":
        v_total = (v1 + v2) / 2.0 if abs(v2) > 1e-12 else v1
        i_total = i1 + i2
        p_total = p1 + p2
        return v_total, i_total, p_total
    raise ValueError(f"Unknown wiring_mode: {wiring_mode}")


# ---------------------------
# GUI App
# ---------------------------
class RampGUI(tk.Tk):
    V_MAX_PER_CH = 32.0
    I_MAX_PER_CH = 10.0

    def __init__(self):
        super().__init__()
        self.title("HMP4040 Ramp Frontend (CV/CC) + Manual Hold + Auto Snapshots")
        self.geometry("780x580")
        self.resizable(False, False)

        self.ctrl = HMP4040Controller()
        self.connected = False
        self.running = False
        self.stop_requested = False

        self.wiring_mode = "SINGLE"  # SINGLE / SERIES / PARALLEL

        # Manual hold control
        self.next_event = threading.Event()
        self.waiting_for_next = False

        # Storage
        self.records = []
        self.current_step_index = 0
        self.current_steps_total = 0
        self.current_set_x = None
        self.current_other = None

        self._build_ui()

    def _build_ui(self):
        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Mode").grid(row=0, column=0, sticky="w")
        self.mode = tk.StringVar(value="CV")
        ttk.Radiobutton(frm, text="Voltage Ramp (CV)", variable=self.mode, value="CV", command=self._refresh_labels).grid(
            row=0, column=1, sticky="w"
        )
        ttk.Radiobutton(frm, text="Current Ramp (CC)", variable=self.mode, value="CC", command=self._refresh_labels).grid(
            row=0, column=2, sticky="w"
        )

        ttk.Separator(frm).grid(row=1, column=0, columnspan=4, sticky="ew", pady=10)

        self.lbl_min = ttk.Label(frm, text="V min (V)")
        self.lbl_max = ttk.Label(frm, text="V max (V)")
        self.lbl_step = ttk.Label(frm, text="V step (V)")
        self.lbl_other = ttk.Label(frm, text="Current set (A)")

        self.lbl_min.grid(row=2, column=0, sticky="w")
        self.lbl_max.grid(row=3, column=0, sticky="w")
        self.lbl_step.grid(row=4, column=0, sticky="w")
        ttk.Label(frm, text="Warmup / Stabilization (s)").grid(row=5, column=0, sticky="w")
        self.lbl_other.grid(row=6, column=0, sticky="w")

        self.var_min = tk.StringVar(value="30")
        self.var_max = tk.StringVar(value="40")
        self.var_step = tk.StringVar(value="1")
        self.var_stab = tk.StringVar(value="5")
        self.var_other = tk.StringVar(value="1.0")

        ttk.Entry(frm, textvariable=self.var_min, width=12).grid(row=2, column=1, sticky="w")
        ttk.Entry(frm, textvariable=self.var_max, width=12).grid(row=3, column=1, sticky="w")
        ttk.Entry(frm, textvariable=self.var_step, width=12).grid(row=4, column=1, sticky="w")
        ttk.Entry(frm, textvariable=self.var_stab, width=12).grid(row=5, column=1, sticky="w")
        ttk.Entry(frm, textvariable=self.var_other, width=12).grid(row=6, column=1, sticky="w")

        ttk.Separator(frm).grid(row=7, column=0, columnspan=4, sticky="ew", pady=10)

        ttk.Label(frm, text="Per-channel limits (fixed for HMP4040):").grid(row=8, column=0, columnspan=4, sticky="w")
        ttk.Label(frm, text=f"V max/channel: {self.V_MAX_PER_CH:g} V").grid(row=9, column=0, columnspan=4, sticky="w")
        ttk.Label(frm, text=f"I max/channel: {self.I_MAX_PER_CH:g} A").grid(row=10, column=0, columnspan=4, sticky="w")

        ttk.Separator(frm).grid(row=11, column=0, columnspan=4, sticky="ew", pady=10)

        self.status = tk.StringVar(value="Ready. Not connected.")
        self.idn = tk.StringVar(value="")
        self.step_status = tk.StringVar(value="")

        ttk.Label(frm, textvariable=self.status).grid(row=12, column=0, columnspan=4, sticky="w")
        ttk.Label(frm, textvariable=self.idn).grid(row=13, column=0, columnspan=4, sticky="w")
        ttk.Label(frm, textvariable=self.step_status).grid(row=14, column=0, columnspan=4, sticky="w")

        btns = ttk.Frame(frm)
        btns.grid(row=15, column=0, columnspan=4, sticky="ew", pady=10)
        for c in range(7):
            btns.columnconfigure(c, weight=1)

        self.btn_connect = ttk.Button(btns, text="CONNECT", command=self.connect)
        self.btn_start = ttk.Button(btns, text="START", command=self.start, state="disabled")
        self.btn_next = ttk.Button(btns, text="NEXT STEP", command=self.next_step, state="disabled")
        self.btn_snapshot = ttk.Button(btns, text="SNAPSHOT NOW", command=self.snapshot_now, state="disabled")
        self.btn_stop = ttk.Button(btns, text="STOP (soft)", command=self.stop_soft, state="disabled")
        self.btn_estop = ttk.Button(btns, text="E-STOP (OFF now)", command=self.estop, state="disabled")
        self.btn_disconnect = ttk.Button(btns, text="DISCONNECT", command=self.disconnect, state="disabled")

        self.btn_connect.grid(row=0, column=0, padx=5, sticky="ew")
        self.btn_start.grid(row=0, column=1, padx=5, sticky="ew")
        self.btn_next.grid(row=0, column=2, padx=5, sticky="ew")
        self.btn_snapshot.grid(row=0, column=3, padx=5, sticky="ew")
        self.btn_stop.grid(row=0, column=4, padx=5, sticky="ew")
        self.btn_estop.grid(row=0, column=5, padx=5, sticky="ew")
        self.btn_disconnect.grid(row=0, column=6, padx=5, sticky="ew")

        ttk.Button(btns, text="SAVE EXCEL", command=self.save_excel).grid(row=1, column=1, padx=5, sticky="ew")
        ttk.Button(btns, text="CLEAR DATA", command=self.clear_data).grid(row=1, column=2, padx=5, sticky="ew")

        ttk.Separator(frm).grid(row=16, column=0, columnspan=4, sticky="ew", pady=10)
        ttk.Label(frm, text="Last snapshots (most recent at bottom):").grid(row=17, column=0, columnspan=4, sticky="w")

        self.preview = tk.Text(frm, width=95, height=12)
        self.preview.grid(row=18, column=0, columnspan=4, sticky="w")
        self.preview.configure(state="disabled")

        self._refresh_labels()

    def _refresh_labels(self):
        if self.mode.get() == "CV":
            self.lbl_min.config(text="V min (V)")
            self.lbl_max.config(text="V max (V)")
            self.lbl_step.config(text="V step (V)")
            self.lbl_other.config(text="Current set (A)")
        else:
            self.lbl_min.config(text="I min (A)")
            self.lbl_max.config(text="I max (A)")
            self.lbl_step.config(text="I step (A)")
            self.lbl_other.config(text="Voltage set (V)")

    def _parse_float(self, s, name):
        try:
            return float(s)
        except Exception:
            raise ValueError(f"Invalid {name}: '{s}'")

    def _wiring_advice(self, max_v_needed, max_i_needed):
        need_series = max_v_needed > self.V_MAX_PER_CH + 1e-9
        need_parallel = max_i_needed > self.I_MAX_PER_CH + 1e-9
        if need_series and need_parallel:
            return "NOT_PRACTICAL"
        if need_series:
            return "SERIES"
        if need_parallel:
            return "PARALLEL"
        return "SINGLE"

    # ---------------------------
    # CONNECT / DISCONNECT
    # ---------------------------
    def connect(self):
        if self.connected:
            messagebox.showinfo("Info", "Already connected.")
            return
        self.status.set("Connecting to HMP4040 (auto-detect)...")
        self.update_idletasks()
        try:
            self.ctrl.connect()
            self.connected = True
            self.idn.set(f"Detected: {self.ctrl.idn}")
            self.status.set("Connected. Ready.")
            self.btn_start.config(state="normal")
            self.btn_estop.config(state="normal")
            self.btn_disconnect.config(state="normal")
            self.btn_snapshot.config(state="normal")
            self.btn_snapshot.config(state="normal")
            self.btn_next.config(state="disabled")
            self.btn_stop.config(state="disabled")
            self.btn_snapshot.config(state="normal")
            self.btn_estop.config(state="normal")
        except Exception as e:
            self.connected = False
            self.idn.set("")
            self.status.set("Ready. Not connected.")
            messagebox.showerror("Connection error", str(e))

    def disconnect(self):
        if self.running:
            messagebox.showwarning("Warning", "Stop the ramp before disconnecting.")
            return
        try:
            self.ctrl.off_all()
        except Exception:
            pass
        try:
            self.ctrl.disconnect()
        except Exception:
            pass

        self.connected = False
        self.idn.set("")
        self.status.set("Disconnected. Ready.")
        self.step_status.set("")
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="disabled")
        self.btn_next.config(state="disabled")
        self.btn_estop.config(state="disabled")
        self.btn_disconnect.config(state="disabled")
        self.btn_snapshot.config(state="disabled")

    # ---------------------------
    # Apply setpoints according to wiring
    # ---------------------------
    def _apply_step_setpoints(self, x, other):
        mode = self.mode.get()

        if self.wiring_mode == "SINGLE":
            self.ctrl.set_output(1, True)
            self.ctrl.set_output(2, False)
            if mode == "CV":
                self.ctrl.apply_cv(1, voltage=x, current_set=other)
            else:
                self.ctrl.apply_cc(1, current_set=x, voltage_set=other)
            return

        # SERIES / PARALLEL: both channels ON
        self.ctrl.set_output(1, True)
        self.ctrl.set_output(2, True)

        if self.wiring_mode == "SERIES":
            # Robust split: CH1 up to 32V, remainder on CH2
            if mode == "CV":
                v1 = min(x, self.V_MAX_PER_CH)
                v2 = max(0.0, x - v1)
                self.ctrl.apply_cv(1, voltage=v1, current_set=other)
                self.ctrl.apply_cv(2, voltage=v2, current_set=other)
            else:
                v1 = min(other, self.V_MAX_PER_CH)
                v2 = max(0.0, other - v1)
                self.ctrl.apply_cc(1, current_set=x, voltage_set=v1)
                self.ctrl.apply_cc(2, current_set=x, voltage_set=v2)
            return

        if self.wiring_mode == "PARALLEL":
            if mode == "CV":
                self.ctrl.apply_cv(1, voltage=x, current_set=other / 2.0)
                self.ctrl.apply_cv(2, voltage=x, current_set=other / 2.0)
            else:
                self.ctrl.apply_cc(1, current_set=x / 2.0, voltage_set=other)
                self.ctrl.apply_cc(2, current_set=x / 2.0, voltage_set=other)
            return

        raise RuntimeError(f"Unknown wiring mode: {self.wiring_mode}")

    # ---------------------------
    # Snapshot reading (measured V/A/P from PSU)
    # ---------------------------
    def _read_snapshot(self):
        v1 = self.ctrl.read_meas_voltage(1)
        i1 = self.ctrl.read_meas_current(1)
        p1 = self.ctrl.read_meas_power(1)

        v2 = i2 = p2 = 0.0
        if self.wiring_mode in ("SERIES", "PARALLEL"):
            v2 = self.ctrl.read_meas_voltage(2)
            i2 = self.ctrl.read_meas_current(2)
            p2 = self.ctrl.read_meas_power(2)

        v_total, i_total, p_total = compute_totals(self.wiring_mode, v1, i1, p1, v2, i2, p2)

        rec = {
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "step": self.current_step_index,
            "steps_total": self.current_steps_total,
            "mode": self.mode.get(),
            "wiring": self.wiring_mode,
            "set_x": self.current_set_x,
            "set_other": self.current_other,
            "meas_v_total": v_total,
            "meas_i_total": i_total,
            "meas_p_total": p_total,
            "meas_v1": v1, "meas_i1": i1, "meas_p1": p1,
            "meas_v2": v2, "meas_i2": i2, "meas_p2": p2,
        }
        return rec

    def snapshot_now(self):
        if not self.connected:
            messagebox.showwarning("Not connected", "Connect first.")
            return
        try:
            rec = self._read_snapshot()
            self.records.append(rec)
            self._update_preview()
            self.status.set("Snapshot captured (manual).")
        except Exception as e:
            messagebox.showerror("Snapshot error", str(e))

    # ---------------------------
    # Manual hold: NEXT STEP releases waiting worker
    # ---------------------------
    def next_step(self):
        if not self.running or not self.waiting_for_next:
            return
        self.next_event.set()

    # ---------------------------
    # Ramp control
    # ---------------------------
    def start(self):
        if not self.connected:
            messagebox.showwarning("Not connected", "Click CONNECT first.")
            return
        if self.running:
            messagebox.showinfo("Info", "Ramp is already running.")
            return

        try:
            mn = self._parse_float(self.var_min.get(), "min")
            mx = self._parse_float(self.var_max.get(), "max")
            st = self._parse_float(self.var_step.get(), "step")
            warmup_s = self._parse_float(self.var_stab.get(), "warmup/stabilization time")
            other = self._parse_float(self.var_other.get(), "set value")

            if st == 0:
                raise ValueError("Step cannot be 0.")
            if warmup_s < 0:
                raise ValueError("Warmup time must be >= 0.")
            if other < 0:
                raise ValueError("Set value must be >= 0.")

            values = frange(mn, mx, st)
            if not values:
                raise ValueError("No steps generated. Check min/max/step.")
        except Exception as e:
            messagebox.showerror("Input error", str(e))
            return

        # Determine wiring needed based on maximum required totals
        if self.mode.get() == "CV":
            max_v_needed = max(mn, mx)
            max_i_needed = other
        else:
            max_i_needed = max(mn, mx)
            max_v_needed = other

        wiring = self._wiring_advice(max_v_needed, max_i_needed)
        if wiring == "NOT_PRACTICAL":
            messagebox.showerror(
                "Limits exceeded",
                "This setup exceeds BOTH per-channel V and I limits.\n"
                "Series and parallel simultaneously is usually not practical.\n"
                "Reduce values or use different supply/wiring."
            )
            return

        if wiring == "SERIES":
            msg = (
                f"Max voltage needed: {max_v_needed:.2f} V > {self.V_MAX_PER_CH:.2f} V/channel\n\n"
                "Wire CH1 + CH2 in SERIES (lab procedure), then press OK."
            )
            if not messagebox.askokcancel("Wiring required: SERIES", msg):
                return
        elif wiring == "PARALLEL":
            msg = (
                f"Max current needed: {max_i_needed:.2f} A > {self.I_MAX_PER_CH:.2f} A/channel\n\n"
                "Wire CH1 + CH2 in PARALLEL (lab procedure), then press OK."
            )
            if not messagebox.askokcancel("Wiring required: PARALLEL", msg):
                return

        self.wiring_mode = wiring

        # Start worker thread
        self.stop_requested = False
        self.running = True
        self.waiting_for_next = False
        self.next_event.clear()

        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.btn_next.config(state="normal")
        self.btn_disconnect.config(state="disabled")
        self.btn_snapshot.config(state="normal")

        self.status.set(f"Running... wiring={self.wiring_mode}")
        t = threading.Thread(target=self._run_ramp_worker, args=(values, warmup_s, other), daemon=True)
        t.start()

    def stop_soft(self):
        if self.running:
            self.stop_requested = True
            self.status.set("Soft stop requested (will stop at next safe point).")
            self.next_event.set()

    def estop(self):
        self.stop_requested = True
        try:
            self.ctrl.off_all()
        except Exception:
            pass
        self.status.set("E-STOP: Outputs OFF.")
        self.next_event.set()
        self._finish_run()

    def _run_ramp_worker(self, values, warmup_s, other):
        try:
            self.current_steps_total = len(values)

            for idx, x in enumerate(values, start=1):
                if self.stop_requested:
                    break

                self.current_step_index = idx
                self.current_set_x = x
                self.current_other = other

                # 1) Apply setpoints (PSU changes here)
                self._apply_step_setpoints(x, other)

                # 2) Warmup (PSU stays at setpoint)
                self.step_status.set(f"Step {idx}/{len(values)} applied. Warming up {warmup_s}s...")
                t0 = time.time()
                while time.time() - t0 < warmup_s:
                    if self.stop_requested:
                        break
                    time.sleep(0.1)
                if self.stop_requested:
                    break

                # 3) Auto snapshot AFTER stabilization
                try:
                    rec = self._read_snapshot()
                    self.records.append(rec)
                    self._update_preview()
                    self.status.set("Snapshot captured (auto, after stabilization).")
                except Exception as e:
                    messagebox.showerror("Snapshot error", str(e))

                # 4) HOLD (NO PSU COMMANDS HERE)
                self.waiting_for_next = True
                self.next_event.clear()
                self.step_status.set(
                    f"HOLD at Step {idx}/{len(values)} — PSU stays ON at this step. Press NEXT STEP to continue."
                )

                # Wait until NEXT STEP
                self.next_event.wait()
                self.waiting_for_next = False
                if self.stop_requested:
                    break

            # End: turn outputs off
            self.ctrl.off_all()
            if self.stop_requested:
                self.status.set("Stopped. Outputs OFF.")
            else:
                self.status.set("Completed. Outputs OFF.")
            self.step_status.set("")

        except Exception as e:
            try:
                self.ctrl.off_all()
            except Exception:
                pass
            messagebox.showerror("Runtime error", str(e))
            self.status.set("Error. Outputs OFF.")
            self.step_status.set("")
        finally:
            self._finish_run()

    def _finish_run(self):
        self.running = False
        self.waiting_for_next = False
        self.btn_start.config(state="normal" if self.connected else "disabled")
        self.btn_stop.config(state="disabled")
        self.btn_next.config(state="disabled")
        self.btn_disconnect.config(state="normal" if self.connected else "disabled")

    # ---------------------------
    # Data / Excel
    # ---------------------------
    def clear_data(self):
        self.records.clear()
        self._update_preview()

    def save_excel(self):
        if not self.records:
            messagebox.showinfo("No data", "No snapshots to save yet.")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
            title="Save snapshots to Excel",
        )
        if not path:
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "Snapshots"

        headers = [
            "timestamp",
            "step", "steps_total",
            "mode", "wiring",
            "set_x", "set_other",
            "meas_v_total", "meas_i_total", "meas_p_total",
            "meas_v1", "meas_i1", "meas_p1",
            "meas_v2", "meas_i2", "meas_p2",
        ]
        ws.append(headers)

        for r in self.records:
            ws.append([r.get(h, "") for h in headers])

        try:
            wb.save(path)
            messagebox.showinfo("Saved", f"Saved Excel:\n{path}")
        except Exception as e:
            messagebox.showerror("Save error", str(e))

    def _update_preview(self):
        lines = []
        tail = self.records[-8:]
        for r in tail:
            lines.append(
                f"{r['timestamp']} | Step {r['step']}/{r['steps_total']} | {r['mode']} {r['wiring']} | "
                f"SET x={r['set_x']} other={r['set_other']} | "
                f"MEAS V={r['meas_v_total']:.4g}  I={r['meas_i_total']:.4g}  P={r['meas_p_total']:.4g}"
            )

        self.preview.configure(state="normal")
        self.preview.delete("1.0", tk.END)
        self.preview.insert(tk.END, "\n".join(lines))
        self.preview.configure(state="disabled")


if __name__ == "__main__":
    app = RampGUI()
    app.mainloop()
