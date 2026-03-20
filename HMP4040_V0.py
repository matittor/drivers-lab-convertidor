import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time

# ---------------------------
# Backend: R&S HMP4040 via PyVISA
# ---------------------------
class HMP4040Controller:
    """
    Minimal, practical controller for HMP4040 using SCPI.
    Notes:
      - HMP series SCPI is pretty standard: INST OUTx, VOLT, CURR, OUTP.
      - If your unit uses slightly different commands, adjust in _write/_query or in apply_*.
    """

    def __init__(self, resource_hint: str = ""):
        self.resource_hint = resource_hint.strip()
        self.rm = None
        self.inst = None
        self.idn = ""

    def connect(self) -> bool:
        try:
            import pyvisa
        except Exception as e:
            raise RuntimeError(
                "PyVISA not installed. Run: python -m pip install pyvisa pyvisa-py"
            ) from e

        self.rm = pyvisa.ResourceManager()
        resources = list(self.rm.list_resources())

        if not resources:
            raise RuntimeError("No VISA resources found. Check USB/LAN connection and VISA installation.")

        # Prefer a hint if provided, otherwise scan for HMP4040 by *IDN?
        candidates = []
        if self.resource_hint:
            candidates = [r for r in resources if self.resource_hint in r]
            if not candidates:
                # allow exact resource as hint
                if self.resource_hint in resources:
                    candidates = [self.resource_hint]
        else:
            candidates = resources

        last_err = None
        for r in candidates:
            try:
                inst = self.rm.open_resource(r)
                inst.timeout = 5000
                inst.write_termination = "\n"
                inst.read_termination = "\n"
                idn = inst.query("*IDN?").strip()
                if "HMP4040" in idn or "ROHDE" in idn or "Rohde" in idn or "SCHWARZ" in idn:
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
                except:
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
        # HMP typically uses: INST OUT1/OUT2/OUT3/OUT4
        self._write(f"INST OUT{ch}")

    def set_output(self, ch: int, on: bool):
        self.select_channel(ch)
        self._write(f"OUTP {'ON' if on else 'OFF'}")

    def apply_cv(self, voltage: float, current_limit: float, ch: int = 1):
        # CV means we set VOLT setpoint and CURR limit
        self.select_channel(ch)
        self._write(f"VOLT {voltage}")
        self._write(f"CURR {current_limit}")

    def apply_cc(self, current: float, voltage_limit: float, ch: int = 1):
        # CC means we set CURR setpoint and VOLT limit (compliance)
        self.select_channel(ch)
        self._write(f"CURR {current}")
        self._write(f"VOLT {voltage_limit}")

    def off_all(self):
        # Immediate OFF on channels 1 & 2 (extend if you use 3/4)
        for ch in (1, 2):
            try:
                self.set_output(ch, False)
            except:
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


# ---------------------------
# GUI
# ---------------------------
class RampGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("HMP4040 Ramp Frontend (CV/CC) - Merged")
        self.geometry("640x520")
        self.resizable(False, False)

        # Backend controller (optionally set resource hint in GUI)
        self.ctrl = HMP4040Controller()
        self.running = False
        self.stop_requested = False

        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)

        # Mode
        ttk.Label(frm, text="Mode").grid(row=0, column=0, sticky="w")
        self.mode = tk.StringVar(value="CV")
        ttk.Radiobutton(frm, text="Voltage Ramp (CV)", variable=self.mode, value="CV", command=self._refresh_labels).grid(row=0, column=1, sticky="w")
        ttk.Radiobutton(frm, text="Current Ramp (CC)", variable=self.mode, value="CC", command=self._refresh_labels).grid(row=0, column=2, sticky="w")

        ttk.Separator(frm).grid(row=1, column=0, columnspan=3, sticky="ew", pady=10)

        # Inputs labels
        self.lbl_min = ttk.Label(frm, text="V min (V)")
        self.lbl_max = ttk.Label(frm, text="V max (V)")
        self.lbl_step = ttk.Label(frm, text="V step (V)")
        self.lbl_other = ttk.Label(frm, text="Current limit (A)")

        self.lbl_min.grid(row=2, column=0, sticky="w")
        self.lbl_max.grid(row=3, column=0, sticky="w")
        self.lbl_step.grid(row=4, column=0, sticky="w")
        ttk.Label(frm, text="Stabilization time (s)").grid(row=5, column=0, sticky="w")
        self.lbl_other.grid(row=6, column=0, sticky="w")

        self.var_min = tk.StringVar(value="30")
        self.var_max = tk.StringVar(value="40")
        self.var_step = tk.StringVar(value="1")
        self.var_stab = tk.StringVar(value="10")
        self.var_other = tk.StringVar(value="1.0")

        ttk.Entry(frm, textvariable=self.var_min, width=12).grid(row=2, column=1, sticky="w")
        ttk.Entry(frm, textvariable=self.var_max, width=12).grid(row=3, column=1, sticky="w")
        ttk.Entry(frm, textvariable=self.var_step, width=12).grid(row=4, column=1, sticky="w")
        ttk.Entry(frm, textvariable=self.var_stab, width=12).grid(row=5, column=1, sticky="w")
        ttk.Entry(frm, textvariable=self.var_other, width=12).grid(row=6, column=1, sticky="w")

        ttk.Separator(frm).grid(row=7, column=0, columnspan=3, sticky="ew", pady=10)

        # Per-channel limits
        ttk.Label(frm, text="Per-channel limits (edit if needed):").grid(row=8, column=0, columnspan=3, sticky="w")
        ttk.Label(frm, text="V max per channel").grid(row=9, column=0, sticky="w")
        ttk.Label(frm, text="I max per channel").grid(row=10, column=0, sticky="w")

        self.var_vch = tk.StringVar(value="32")
        self.var_ich = tk.StringVar(value="10")

        ttk.Entry(frm, textvariable=self.var_vch, width=12).grid(row=9, column=1, sticky="w")
        ttk.Entry(frm, textvariable=self.var_ich, width=12).grid(row=10, column=1, sticky="w")

        ttk.Separator(frm).grid(row=11, column=0, columnspan=3, sticky="ew", pady=10)

        # VISA resource hint
        ttk.Label(frm, text="VISA Resource Hint (optional):").grid(row=12, column=0, sticky="w")
        self.var_hint = tk.StringVar(value="")
        ttk.Entry(frm, textvariable=self.var_hint, width=40).grid(row=12, column=1, columnspan=2, sticky="w")
        ttk.Label(frm, text="(Ej: USB0::0x0AAD::...  o  TCPIP0::192.168.0.50::INSTR)").grid(row=13, column=1, columnspan=2, sticky="w")

        # Status + IDN
        self.status = tk.StringVar(value="Ready.")
        self.idn = tk.StringVar(value="")
        ttk.Label(frm, textvariable=self.status).grid(row=14, column=0, columnspan=3, sticky="w", pady=(10, 0))
        ttk.Label(frm, textvariable=self.idn).grid(row=15, column=0, columnspan=3, sticky="w")

        # Buttons
        btns = ttk.Frame(frm)
        btns.grid(row=16, column=0, columnspan=3, sticky="ew", pady=12)
        for c in range(4):
            btns.columnconfigure(c, weight=1)

        self.btn_start = ttk.Button(btns, text="START", command=self.start)
        self.btn_stop = ttk.Button(btns, text="STOP (soft)", command=self.stop_soft)
        self.btn_estop = ttk.Button(btns, text="E-STOP (OFF now)", command=self.estop)
        self.btn_disconnect = ttk.Button(btns, text="DISCONNECT", command=self.disconnect)

        self.btn_start.grid(row=0, column=0, padx=6, sticky="ew")
        self.btn_stop.grid(row=0, column=1, padx=6, sticky="ew")
        self.btn_estop.grid(row=0, column=2, padx=6, sticky="ew")
        self.btn_disconnect.grid(row=0, column=3, padx=6, sticky="ew")

        self._refresh_labels()

    def _refresh_labels(self):
        if self.mode.get() == "CV":
            self.lbl_min.config(text="V min (V)")
            self.lbl_max.config(text="V max (V)")
            self.lbl_step.config(text="V step (V)")
            self.lbl_other.config(text="Current limit (A)")
        else:
            self.lbl_min.config(text="I min (A)")
            self.lbl_max.config(text="I max (A)")
            self.lbl_step.config(text="I step (A)")
            self.lbl_other.config(text="Voltage limit (V)")

    def _parse_float(self, s, name):
        try:
            return float(s)
        except:
            raise ValueError(f"Invalid {name}: '{s}'")

    def _wiring_advice(self, max_v_needed, max_i_needed, v_ch, i_ch):
        need_series = max_v_needed > v_ch + 1e-9
        need_parallel = max_i_needed > i_ch + 1e-9
        if need_series and need_parallel:
            return "NOT_PRACTICAL"
        if need_series:
            return "SERIES"
        if need_parallel:
            return "PARALLEL"
        return "SINGLE"

    def _ensure_connected(self):
        # Apply hint to controller before connecting
        self.ctrl.resource_hint = self.var_hint.get().strip()

        self.status.set("Connecting to HMP4040...")
        self.update_idletasks()

        ok = self.ctrl.connect()
        if not ok:
            raise RuntimeError("Connection failed.")
        self.idn.set(f"Detected: {self.ctrl.idn}")

    def start(self):
        if self.running:
            messagebox.showinfo("Info", "Ramp is already running.")
            return

        try:
            mn = self._parse_float(self.var_min.get(), "min")
            mx = self._parse_float(self.var_max.get(), "max")
            st = self._parse_float(self.var_step.get(), "step")
            stab = self._parse_float(self.var_stab.get(), "stabilization time")
            other = self._parse_float(self.var_other.get(), "limit")
            v_ch = self._parse_float(self.var_vch.get(), "V max/channel")
            i_ch = self._parse_float(self.var_ich.get(), "I max/channel")

            if st == 0:
                raise ValueError("Step cannot be 0.")
            if stab < 0:
                raise ValueError("Stabilization time must be >= 0.")
            if other < 0:
                raise ValueError("Limit must be >= 0.")

            values = frange(mn, mx, st)
            if not values:
                raise ValueError("No steps generated. Check min/max/step.")
        except Exception as e:
            messagebox.showerror("Input error", str(e))
            return

        # Wiring advice from the front design logic
        if self.mode.get() == "CV":
            max_v_needed = max(mn, mx)
            max_i_needed = other
        else:
            max_i_needed = max(mn, mx)
            max_v_needed = other

        wiring = self._wiring_advice(max_v_needed, max_i_needed, v_ch, i_ch)
        if wiring == "NOT_PRACTICAL":
            messagebox.showerror(
                "Limits exceeded",
                "This setup exceeds BOTH per-channel V and I limits.\n"
                "Series and parallel simultaneously is usually not practical.\n"
                "Reduce limits or use different wiring/supply."
            )
            return

        if wiring == "SERIES":
            msg = (
                f"Max voltage needed: {max_v_needed:.2f} V > {v_ch:.2f} V (per channel)\n\n"
                "Please wire CH1 + CH2 in SERIES (per lab procedure), then press OK."
            )
            if not messagebox.askokcancel("Wiring required: SERIES", msg):
                self.status.set("Ready.")
                return
        elif wiring == "PARALLEL":
            msg = (
                f"Max current needed: {max_i_needed:.2f} A > {i_ch:.2f} A (per channel)\n\n"
                "Please wire CH1 + CH2 in PARALLEL (per lab procedure), then press OK."
            )
            if not messagebox.askokcancel("Wiring required: PARALLEL", msg):
                self.status.set("Ready.")
                return

        # Connect instrument
        try:
            self._ensure_connected()
        except Exception as e:
            messagebox.showerror("Connection error", str(e))
            self.status.set("Ready.")
            return

        # Start worker thread
        self.stop_requested = False
        self.running = True
        self.btn_start.config(state="disabled")
        self.status.set("Running...")
        t = threading.Thread(target=self._run_ramp, args=(values, stab, other), daemon=True)
        t.start()

    def stop_soft(self):
        if self.running:
            self.stop_requested = True
            self.status.set("Soft stop requested (will stop at next safe point).")

    def estop(self):
        self.stop_requested = True
        try:
            self.ctrl.off_all()
        except:
            pass
        self.status.set("E-STOP: Outputs OFF.")
        self._finish()

    def disconnect(self):
        # disconnect safely
        try:
            self.ctrl.off_all()
        except:
            pass
        try:
            self.ctrl.disconnect()
        except:
            pass
        self.idn.set("")
        if not self.running:
            self.status.set("Disconnected. Ready.")
        else:
            self.status.set("Disconnected (ramp running?).")

    def _run_ramp(self, values, stab, other):
        try:
            # Default: use CH1 only; if wired series/parallel externally, CH1 controls setpoints.
            self.ctrl.set_output(1, True)

            for idx, x in enumerate(values, start=1):
                if self.stop_requested:
                    break

                if self.mode.get() == "CV":
                    self.status.set(f"Step {idx}/{len(values)} -> CV: V={x}V, Ilim={other}A")
                    self.ctrl.apply_cv(voltage=x, current_limit=other, ch=1)
                else:
                    self.status.set(f"Step {idx}/{len(values)} -> CC: I={x}A, Vlim={other}V")
                    self.ctrl.apply_cc(current=x, voltage_limit=other, ch=1)

                # Stabilization
                t0 = time.time()
                while time.time() - t0 < stab:
                    if self.stop_requested:
                        break
                    time.sleep(0.1)

            # done: OFF
            self.ctrl.off_all()
            if self.stop_requested:
                self.status.set("Stopped. Outputs OFF.")
            else:
                self.status.set("Completed. Outputs OFF.")

        except Exception as e:
            try:
                self.ctrl.off_all()
            except:
                pass
            messagebox.showerror("Runtime error", str(e))
            self.status.set("Error. Outputs OFF.")
        finally:
            self._finish()

    def _finish(self):
        self.running = False
        self.btn_start.config(state="normal")


if __name__ == "__main__":
    app = RampGUI()
    app.mainloop()
