"""
HMP4040 Ramp Frontend (CV/CC) + Status/Timer + Auto Snapshot After Stabilization
+ Save Excel + GL_SpectroSoft - Lab automation (SPACE trigger + wait + auto-next)
Author: Matias Riquelme, CPI
"""
import threading
import time
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from openpyxl import Workbook
# --- GL SpectroSoft UI automation ---
from pywinauto import Desktop
from pywinauto.keyboard import send_keys

# ---------------------------
# GL SpectroSoft constants
# ---------------------------
GL_MAIN_TITLE = "GL_SpectroSoft - Lab"   # keyword match
GL_WAIT_AFTER_SPACE_S = 20.0            # <-- wait after SPACE before next step

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
               inst.timeout = 6000
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
   def _query_float(self, cmd: str) -> float:
       return float(self._query(cmd).strip())
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
       return self._query_float("MEAS:VOLT?")
   def read_meas_current(self, ch: int) -> float:
       self.select_channel(ch)
       return self._query_float("MEAS:CURR?")
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

def compute_totals(wiring_mode, v1, i1, v2=0.0, i2=0.0):
   if wiring_mode == "SINGLE":
       return v1, i1
   if wiring_mode == "SERIES":
       v_total = v1 + v2
       i_total = (i1 + i2) / 2.0 if abs(i2) > 1e-12 else i1
       return v_total, i_total
   if wiring_mode == "PARALLEL":
       v_total = (v1 + v2) / 2.0 if abs(v2) > 1e-12 else v1
       i_total = i1 + i2
       return v_total, i_total
   raise ValueError(f"Unknown wiring_mode: {wiring_mode}")

# ---------------------------
# GUI App
# ---------------------------
class RampGUI(tk.Tk):
   V_MAX_PER_CH = 32.0
   I_MAX_PER_CH = 10.0
   def __init__(self):
       super().__init__()
       self.title("HMP4040 Ramp Frontend (CV/CC) + Status + Timer + Snapshots + GL SPACE")
       self.geometry("900x660")
       self.resizable(False, False)
       self.ctrl = HMP4040Controller()
       self.connected = False
       self.running = False
       self.stop_requested = False
       self.wiring_mode = "SINGLE"  # SINGLE / SERIES / PARALLEL
       self.records = []
       self.current_step_index = 0
       self.current_steps_total = 0
       self.current_set_x = None
       self.current_other = None
       # UI status/timer
       self.phase = tk.StringVar(value="Ready")
       self.timer = tk.StringVar(value="")
       self.detail = tk.StringVar(value="")
       self.idn = tk.StringVar(value="")
       # Ramp type + custom points
       self.ramp_type = tk.StringVar(value="LINEAR")  # LINEAR / CUSTOM
       self.custom_points = []
       self._build_ui()
   # ---------- UI ----------
   def _build_ui(self):
       frm = ttk.Frame(self, padding=12)
       frm.pack(fill="both", expand=True)
       status_box = ttk.LabelFrame(frm, text="STATUS / TIMER", padding=10)
       status_box.grid(row=0, column=0, columnspan=4, sticky="ew")
       status_box.columnconfigure(0, weight=1)
       ttk.Label(status_box, textvariable=self.phase, font=("Segoe UI", 14, "bold")).grid(row=0, column=0, sticky="w")
       ttk.Label(status_box, textvariable=self.timer, font=("Segoe UI", 12)).grid(row=0, column=1, sticky="e")
       ttk.Label(status_box, textvariable=self.detail).grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))
       ttk.Label(status_box, textvariable=self.idn).grid(row=2, column=0, columnspan=2, sticky="w")
       ttk.Separator(frm).grid(row=1, column=0, columnspan=4, sticky="ew", pady=10)
       ttk.Label(frm, text="Mode").grid(row=2, column=0, sticky="w")
       self.mode = tk.StringVar(value="CC")
       ttk.Radiobutton(frm, text="Voltage Ramp (CV)", variable=self.mode, value="CV",
                       command=self._refresh_labels).grid(row=2, column=1, sticky="w")
       ttk.Radiobutton(frm, text="Current Ramp (CC)", variable=self.mode, value="CC",
                       command=self._refresh_labels).grid(row=2, column=2, sticky="w")
       ttk.Label(frm, text="Power in snapshot: CALC(V×I)", font=("Segoe UI", 9, "italic")).grid(row=2, column=3, sticky="w")
       # Ramp type box
       ramp_box = ttk.LabelFrame(frm, text="RAMP TYPE", padding=8)
       ramp_box.grid(row=3, column=0, columnspan=4, sticky="ew", pady=(6, 0))
       ramp_box.columnconfigure(3, weight=1)
       ttk.Radiobutton(ramp_box, text="Linear", variable=self.ramp_type, value="LINEAR",
                       command=self._refresh_ramp_ui).grid(row=0, column=0, sticky="w")
       ttk.Radiobutton(ramp_box, text="Custom list", variable=self.ramp_type, value="CUSTOM",
                       command=self._refresh_ramp_ui).grid(row=0, column=1, sticky="w")
       ttk.Label(ramp_box, text="Custom points (same units as X):", font=("Segoe UI", 9)).grid(
           row=1, column=0, columnspan=2, sticky="w", pady=(6, 2)
       )
       self.custom_entry = tk.StringVar(value="")
       self.ent_custom = ttk.Entry(ramp_box, textvariable=self.custom_entry, width=40)
       self.ent_custom.grid(row=2, column=0, columnspan=2, sticky="w")
       self.btn_add = ttk.Button(ramp_box, text="+ Add", command=self._custom_add)
       self.btn_remove = ttk.Button(ramp_box, text="- Remove", command=self._custom_remove)
       self.btn_clear = ttk.Button(ramp_box, text="Clear", command=self._custom_clear)
       self.btn_add.grid(row=2, column=2, padx=(8, 4), sticky="w")
       self.btn_remove.grid(row=2, column=3, padx=4, sticky="w")
       self.btn_clear.grid(row=2, column=4, padx=4, sticky="w")
       self.lst_custom = tk.Listbox(ramp_box, height=4, width=40, exportselection=False)
       self.lst_custom.grid(row=3, column=0, columnspan=2, sticky="w", pady=(6, 0))
       ttk.Label(
           ramp_box,
           text="Tip: paste comma-separated values (e.g., 0.001,0.003,0.006).",
           font=("Segoe UI", 8, "italic"),
       ).grid(row=4, column=0, columnspan=5, sticky="w", pady=(4, 0))
       # Parameter labels (keep existing behavior)
       self.lbl_min = ttk.Label(frm, text="I min (A)")
       self.lbl_max = ttk.Label(frm, text="I max (A)")
       self.lbl_step = ttk.Label(frm, text="I step (A)")
       self.lbl_other = ttk.Label(frm, text="Voltage set (V)")
       self.lbl_min.grid(row=5, column=0, sticky="w")
       self.lbl_max.grid(row=6, column=0, sticky="w")
       self.lbl_step.grid(row=7, column=0, sticky="w")
       ttk.Label(frm, text="Stabilization (s)").grid(row=8, column=0, sticky="w")
       self.lbl_other.grid(row=9, column=0, sticky="w")
       self.var_min = tk.StringVar(value="1.0")
       self.var_max = tk.StringVar(value="2.0")
       self.var_step = tk.StringVar(value="0.1")
       self.var_stab = tk.StringVar(value="5")
       self.var_other = tk.StringVar(value="2.0")
       # --- Grey-out capable entries (store refs) ---
       self.ent_min = ttk.Entry(frm, textvariable=self.var_min, width=12)
       self.ent_max = ttk.Entry(frm, textvariable=self.var_max, width=12)
       self.ent_step = ttk.Entry(frm, textvariable=self.var_step, width=12)
       self.ent_min.grid(row=5, column=1, sticky="w")
       self.ent_max.grid(row=6, column=1, sticky="w")
       self.ent_step.grid(row=7, column=1, sticky="w")
       ttk.Entry(frm, textvariable=self.var_stab, width=12).grid(row=8, column=1, sticky="w")
       ttk.Entry(frm, textvariable=self.var_other, width=12).grid(row=9, column=1, sticky="w")
       ttk.Label(frm, text=f"(Fixed limits: {self.V_MAX_PER_CH}V/ch, {self.I_MAX_PER_CH}A/ch)").grid(row=5, column=2, sticky="w")
       ttk.Separator(frm).grid(row=10, column=0, columnspan=4, sticky="ew", pady=10)
       btns = ttk.Frame(frm)
       btns.grid(row=11, column=0, columnspan=4, sticky="ew")
       for c in range(6):
           btns.columnconfigure(c, weight=1)
       self.btn_connect = ttk.Button(btns, text="CONNECT", command=self.connect)
       self.btn_start = ttk.Button(btns, text="START", command=self.start, state="disabled")
       self.btn_snapshot = ttk.Button(btns, text="SNAPSHOT NOW", command=self.snapshot_now, state="disabled")
       self.btn_stop = ttk.Button(btns, text="STOP (soft)", command=self.stop_soft, state="disabled")
       self.btn_estop = ttk.Button(btns, text="E-STOP", command=self.estop, state="disabled")
       self.btn_disconnect = ttk.Button(btns, text="DISCONNECT", command=self.disconnect, state="disabled")
       self.btn_connect.grid(row=0, column=0, padx=4, sticky="ew")
       self.btn_start.grid(row=0, column=1, padx=4, sticky="ew")
       self.btn_snapshot.grid(row=0, column=2, padx=4, sticky="ew")
       self.btn_stop.grid(row=0, column=3, padx=4, sticky="ew")
       self.btn_estop.grid(row=0, column=4, padx=4, sticky="ew")
       self.btn_disconnect.grid(row=0, column=5, padx=4, sticky="ew")
       ttk.Button(btns, text="SAVE EXCEL", command=self.save_excel).grid(row=1, column=2, padx=4, sticky="ew")
       ttk.Button(btns, text="CLEAR DATA", command=self.clear_data).grid(row=1, column=3, padx=4, sticky="ew")
       ttk.Separator(frm).grid(row=12, column=0, columnspan=4, sticky="ew", pady=10)
       # ----------------------------
       # PRO preview: scrollbars + wrap none (FIXED LAYOUT)
       # ----------------------------
       ttk.Label(frm, text="Last snapshots:").grid(row=13, column=0, columnspan=4, sticky="w")
       # IMPORTANT: allow expansion in grid (even though window is fixed size,
       # this still ensures scrollbars have proper width/behavior)
       frm.columnconfigure(0, weight=1)
       frm.rowconfigure(14, weight=1)
       preview_frame = ttk.Frame(frm)
       preview_frame.grid(row=14, column=0, columnspan=4, sticky="nsew")
       preview_frame.columnconfigure(0, weight=1)
       preview_frame.rowconfigure(0, weight=1)
       xscroll = ttk.Scrollbar(preview_frame, orient="horizontal")
       yscroll = ttk.Scrollbar(preview_frame, orient="vertical")
       self.preview = tk.Text(
           preview_frame,
           width=112,
           height=16,
           wrap="none",
           xscrollcommand=xscroll.set,
           yscrollcommand=yscroll.set,
       )
       xscroll.config(command=self.preview.xview)
       yscroll.config(command=self.preview.yview)
       self.preview.grid(row=0, column=0, sticky="nsew")
       yscroll.grid(row=0, column=1, sticky="ns")
       xscroll.grid(row=1, column=0, sticky="ew")
       self.preview.configure(state="disabled")
       self._refresh_labels()
       self._refresh_ramp_ui()
       self._ui_set("Ready", "", "Not connected.")
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
       self._refresh_ramp_ui()
   def _refresh_ramp_ui(self):
       is_custom = (self.ramp_type.get() == "CUSTOM")
       # Enable/disable custom widgets
       self.ent_custom.configure(state=("normal" if is_custom else "disabled"))
       self.lst_custom.configure(state=("normal" if is_custom else "disabled"))
       self.btn_add.configure(state=("normal" if is_custom else "disabled"))
       self.btn_remove.configure(state=("normal" if is_custom else "disabled"))
       self.btn_clear.configure(state=("normal" if is_custom else "disabled"))
       # Grey out linear fields when custom is selected (visual only)
       self.ent_min.configure(state=("disabled" if is_custom else "normal"))
       self.ent_max.configure(state=("disabled" if is_custom else "normal"))
       self.ent_step.configure(state=("disabled" if is_custom else "normal"))
   def _ui_set(self, phase: str, timer: str = "", detail: str = ""):
       def _do():
           self.phase.set(phase)
           self.timer.set(timer)
           self.detail.set(detail)
       self.after(0, _do)
   def _parse_float(self, s, name):
       try:
           return float(s)
       except Exception:
           raise ValueError(f"Invalid {name}: '{s}'")
   def _parse_custom_text(self, text: str):
       raw = (text or "").strip()
       if not raw:
           raise ValueError("Custom point is empty.")
       parts = [p.strip() for p in raw.replace(";", ",").replace("\n", ",").split(",")]
       vals = []
       for p in parts:
           if not p:
               continue
           subparts = p.split()
           for sp in subparts:
               if not sp:
                   continue
               vals.append(float(sp))
       if not vals:
           raise ValueError("No valid custom points found.")
       return vals
   def _custom_add(self):
       try:
           vals = self._parse_custom_text(self.custom_entry.get())
           changed = False
           for v in vals:
               if v < 0:
                   raise ValueError("Custom points must be >= 0.")
               rv = round(v, 6)
               if rv not in self.custom_points:
                   self.custom_points.append(rv)
                   changed = True
           if changed:
               self._custom_refresh_listbox()
           self.custom_entry.set("")
       except Exception as e:
           messagebox.showerror("Custom list", str(e))
   def _custom_remove(self):
       try:
           sel = self.lst_custom.curselection()
           if not sel:
               return
           idx = int(sel[0])
           if 0 <= idx < len(self.custom_points):
               self.custom_points.pop(idx)
               self._custom_refresh_listbox()
       except Exception:
           pass
   def _custom_clear(self):
       self.custom_points.clear()
       self._custom_refresh_listbox()
   def _custom_refresh_listbox(self):
       self.lst_custom.delete(0, tk.END)
       for v in self.custom_points:
           self.lst_custom.insert(tk.END, f"{v:g}")
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
   # ---------- Connect / Disconnect ----------
   def connect(self):
       if self.connected:
           messagebox.showinfo("Info", "Already connected.")
           return
       self._ui_set("Connecting…", "", "Auto-detecting HMP4040 via VISA…")
       try:
           self.ctrl.connect()
           self.connected = True
           self.idn.set(f"Detected: {self.ctrl.idn}")
           self._ui_set("Connected", "", "Ready. Set parameters, then START.")
           self.btn_start.config(state="normal")
           self.btn_estop.config(state="normal")
           self.btn_disconnect.config(state="normal")
           self.btn_snapshot.config(state="normal")
       except Exception as e:
           self.connected = False
           self.idn.set("")
           self._ui_set("Ready", "", "Not connected.")
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
       self._ui_set("Ready", "", "Not connected.")
       self.btn_start.config(state="disabled")
       self.btn_stop.config(state="disabled")
       self.btn_estop.config(state="disabled")
       self.btn_disconnect.config(state="disabled")
       self.btn_snapshot.config(state="disabled")
   # ---------- Apply setpoints ----------
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
       self.ctrl.set_output(1, True)
       self.ctrl.set_output(2, True)
       if self.wiring_mode == "SERIES":
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
   # ---------- GL SpectroSoft automation (SPACE only) ----------
   def _gl_find_main(self):
       wins = Desktop(backend="uia").windows()
       for w in wins:
           title = (w.window_text() or "")
           if GL_MAIN_TITLE.lower() in title.lower():
               return w
       raise RuntimeError(f"GL main window not found (keyword: '{GL_MAIN_TITLE}').")
   def gl_trigger_space_and_wait(self, wait_s: float):
       main = self._gl_find_main()
       try:
           main.restore()
       except Exception:
           pass
       main.set_focus()
       time.sleep(0.2)
       send_keys("{SPACE}")
       time.sleep(0.2)
       t_end = time.time() + float(wait_s)
       while time.time() < t_end:
           remaining = t_end - time.time()
           self._ui_set("GL Measuring…", f"{remaining:0.1f}s", "SPACE sent; waiting…")
           time.sleep(0.1)
   # ---------- Snapshot ----------
   def _read_snapshot(self):
       v1 = self.ctrl.read_meas_voltage(1)
       i1 = self.ctrl.read_meas_current(1)
       v2 = i2 = 0.0
       if self.wiring_mode in ("SERIES", "PARALLEL"):
           v2 = self.ctrl.read_meas_voltage(2)
           i2 = self.ctrl.read_meas_current(2)
       v_total, i_total = compute_totals(self.wiring_mode, v1, i1, v2, i2)
       p_total = v_total * i_total
       return {
           "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
           "step": self.current_step_index,
           "steps_total": self.current_steps_total,
           "mode": self.mode.get(),
           "wiring": self.wiring_mode,
           "p_mode": "CALC(VxI)",
           "set_x": self.current_set_x,
           "set_other": self.current_other,
           "meas_v_total": v_total,
           "meas_i_total": i_total,
           "meas_p_total": p_total,
           "meas_v1": v1,
           "meas_i1": i1,
           "meas_v2": v2,
           "meas_i2": i2,
       }
   def snapshot_now(self):
       if not self.connected:
           messagebox.showwarning("Not connected", "Connect first.")
           return
       try:
           self._ui_set("Capturing snapshot…", "", "Reading V / I from PSU…")
           rec = self._read_snapshot()
           self.records.append(rec)
           self._update_preview()
           self._ui_set("Connected", "", "Snapshot captured (manual).")
       except Exception as e:
           self._ui_set("Error", "", "Snapshot failed.")
           messagebox.showerror("Snapshot error", str(e))
   # ---------- Run control ----------
   def start(self):
       if not self.connected:
           messagebox.showwarning("Not connected", "Click CONNECT first.")
           return
       if self.running:
           messagebox.showinfo("Info", "Ramp is already running.")
           return
       try:
           stab = self._parse_float(self.var_stab.get(), "stabilization time")
           other = self._parse_float(self.var_other.get(), "set value")
           if stab < 0:
               raise ValueError("Stabilization time must be >= 0.")
           if other < 0:
               raise ValueError("Set value must be >= 0.")
           if self.ramp_type.get() == "LINEAR":
               mn = self._parse_float(self.var_min.get(), "min")
               mx = self._parse_float(self.var_max.get(), "max")
               st = self._parse_float(self.var_step.get(), "step")
               if st == 0:
                   raise ValueError("Step cannot be 0.")
               values = frange(mn, mx, st)
               if not values:
                   raise ValueError("No steps generated. Check min/max/step.")
           else:
               if not self.custom_points:
                   raise ValueError("Custom list is empty. Add points first.")
               values = list(self.custom_points)
       except Exception as e:
           messagebox.showerror("Input error", str(e))
           return
       # Wiring decision
       if self.mode.get() == "CV":
           max_v_needed = max(values) if values else 0.0
           max_i_needed = other
       else:
           max_i_needed = max(values) if values else 0.0
           max_v_needed = other
       wiring = self._wiring_advice(max_v_needed, max_i_needed)
       if wiring == "NOT_PRACTICAL":
           messagebox.showerror("Limits exceeded", "Exceeds BOTH per-channel V and I limits.")
           return
       if wiring == "SERIES":
           if not messagebox.askokcancel("Wiring required: SERIES", "Wire CH1+CH2 in SERIES, then OK."):
               return
       elif wiring == "PARALLEL":
           if not messagebox.askokcancel("Wiring required: PARALLEL", "Wire CH1+CH2 in PARALLEL, then OK."):
               return
       self.wiring_mode = wiring
       self.stop_requested = False
       self.running = True
       self.btn_start.config(state="disabled")
       self.btn_stop.config(state="normal")
       self.btn_disconnect.config(state="disabled")
       t = threading.Thread(target=self._run_worker, args=(values, stab, other), daemon=True)
       t.start()
   def stop_soft(self):
       if self.running:
           self.stop_requested = True
           self._ui_set("Stopping…", "", "Soft stop requested (will stop at safe point).")
   def estop(self):
       self.stop_requested = True
       try:
           self.ctrl.off_all()
       except Exception:
           pass
       self._ui_set("E-STOP", "", "Outputs OFF immediately.")
       self._finish_run()
   def _run_worker(self, values, stab, other):
       try:
           self.current_steps_total = len(values)
           for idx, x in enumerate(values, start=1):
               if self.stop_requested:
                   break
               self.current_step_index = idx
               self.current_set_x = x
               self.current_other = other
               self._ui_set("Applying step…", "", f"Step {idx}/{len(values)}  wiring={self.wiring_mode}")
               self._apply_step_setpoints(x, other)
               t_end = time.time() + stab
               while True:
                   if self.stop_requested:
                       break
                   remaining = t_end - time.time()
                   if remaining <= 0:
                       break
                   self._ui_set("Stabilizing…", f"{remaining:0.1f}s", f"Step {idx}/{len(values)}")
                   time.sleep(0.1)
               if self.stop_requested:
                   break
               self._ui_set("Capturing snapshot…", "", "Reading V / I from PSU…")
               rec = self._read_snapshot()
               self.records.append(rec)
               self._update_preview()
               try:
                   self.gl_trigger_space_and_wait(GL_WAIT_AFTER_SPACE_S)
                   self._ui_set("Auto-Advance", "", f"GL done; moving to next step ({idx+1}/{len(values)})")
               except Exception as e:
                   self._ui_set("GL Error", "", f"{e}")
                   messagebox.showwarning(
                       "GL automation failed",
                       f"{e}\n\nBring GL window to front (not minimized) and retry. Stopping for safety."
                   )
                   break
           self.ctrl.off_all()
           if self.stop_requested:
               self._ui_set("Stopped", "", "Outputs OFF.")
           else:
               self._ui_set("Completed", "", "Outputs OFF.")
       except Exception as e:
           try:
               self.ctrl.off_all()
           except Exception:
               pass
           self._ui_set("Error", "", "Runtime error; outputs OFF.")
           messagebox.showerror("Runtime error", str(e))
       finally:
           self._finish_run()
   def _finish_run(self):
       self.running = False
       self.btn_start.config(state="normal" if self.connected else "disabled")
       self.btn_stop.config(state="disabled")
       self.btn_disconnect.config(state="normal" if self.connected else "disabled")
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
           "p_mode",
           "set_x", "set_other",
           "meas_v_total", "meas_i_total", "meas_p_total",
           "meas_v1", "meas_i1",
           "meas_v2", "meas_i2",
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
       tail = self.records[-8:]
       lines = []
       for r in tail:
           lines.append(
               f"{r['timestamp']} | Step {r['step']}/{r['steps_total']} | {r['mode']} {r['wiring']} | "
               f"{r['p_mode']} | SET x={r['set_x']} other={r['set_other']} | "
               f"MEAS V={r['meas_v_total']:.4g}  I={r['meas_i_total']:.4g}  P={r['meas_p_total']:.4g}"
           )
       self.preview.configure(state="normal")
       self.preview.delete("1.0", tk.END)
       self.preview.insert(tk.END, "\n".join(lines))
       self.preview.configure(state="disabled")

if __name__ == "__main__":
   app = RampGUI()
   app.mainloop()