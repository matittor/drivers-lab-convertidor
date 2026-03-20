import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time

# ---------------------------
# Backend: R&S HMP4040 via PyVISA
# ---------------------------
class HMP4040Controller:
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
            raise RuntimeError("No VISA resources found. Check USB/LAN and VISA installation.")

        candidates = []
        if self.resource_hint:
            candidates = [r for r in resources if self.resource_hint in r]
            if self.resource_hint in resources:
                candidates = [self.resource_hint]
            if not candidates:
                candidates = [self.resource_hint]  # try open directly
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
                if "HMP4040" in idn or "ROHDE" in idn or "SCHWARZ" in idn:
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

    def select_channel(self, ch: int):
        self._write(f"INST OUT{ch}")

    def set_output(self, ch: int, on: bool):
        self.select_channel(ch)
        self._write(f"OUTP {'ON' if on else 'OFF'}")

    def apply_cv(self, voltage: float, current_set: float, ch: int = 1):
        # In the instrument, CURR acts as current limit.
        # UI calls it "Current set" as requested.
        self.select_channel(ch)
        self._write(f"VOLT {voltage}")
        self._write(f"CURR {current_set}")

    def apply_cc(self, current_set: float, voltage_set: float, ch: int = 1):
        # In CC, VOLT behaves like compliance limit.
        # UI calls it "Voltage set" (or limit) — your choice.
        self.select_channel(ch)
        self._write(f"CURR {current_set}")
        self._write(f"VOLT {voltage_set}")

    def off_all(self):
        for ch in (1, 2):
            try:
                self.set_output(ch, False)
            except:
                pass


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


class RampGUI(tk.Tk):
    # Fixed limits for THIS power supply
    V_MAX_PER_CH = 32.0
    I_MAX_PER_CH = 10.0

    def __init__(self):
        super().__init__()
        self.title("HMP4040 Ramp Frontend (CV/CC)")
        self.geometry("660x520")
        self.resizable(False, False)

        self.ctrl = HMP4040Controller()
        self.connected = False
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

        # Inputs
        self.lbl_min = ttk.Label(frm, text="V min (V)")
        self.lbl_max = ttk.Label(frm, text="V max (V)")
        self.lbl_step = ttk.Label(frm, text="V step (V)")
        self.lbl_other = ttk.Label(frm, text="Current set (A)")

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

        # Fixed limits display (NOT editable)
        ttk.Label(frm, text="Per-channel limits (fixed for HMP4040):").grid(row=8, column=0, columnspan=3, sticky="w")
        ttk.Label(frm, text=f"V max per channel: {self.V_MAX_PER_CH:g} V").grid(row=9, column=0, columnspan=3, sticky="w")
        ttk.Label(frm, text=f"I max per channel: {self.I_MAX_PER_CH:g} A").grid(row=10, column=0, columnspan=3, sticky="w")

        ttk.Separator(frm).grid(row=11, column=0, columnspan=3, sticky="ew", pady=10)

        # VISA hint (optional)
        ttk.Label(frm, text="VISA Resource Hint (optional):").grid(row=12, column=0, sticky="w")
        self.var_hint = tk.StringVar(value="")
        ttk.Entry(frm, textvariable=self.var_hint, width=44).grid(row=12, column=1, columnspan=2, sticky="w")
        ttk.Label(frm, text="(Ej: USB0::...  o  TCPIP0::192.168.x.x::INSTR)").grid(row=13, column=1, columnspan=2, sticky="w")

        # Status + IDN
        self.status = tk.StringVar(value="Ready. Not connected.")
        self.idn = tk.StringVar(value="")
        ttk.Label(frm, textvariable=self.status).grid(row=14, column=0, columnspan=3, sticky="w", pady=(10, 0))
        ttk.Label(frm, textvariable=self.idn).grid(row=15, column=0, columnspan=3, sticky="w")

        # Buttons
        btns = ttk.Frame(frm)
        btns.grid(row=16, column=0, columnspan=3, sticky="ew", pady=12)
        for c in range(5):
            btns.columnconfigure(c, weight=1)

        self.btn_connect = ttk.Button(btns, text="CONNECT", command=self.connect)
        self.btn_start = ttk.Button(btns, text="START", command=self.start, state="disabled")
        self.btn_stop = ttk.Button(btns, text="STOP (soft)", command=self.stop_soft, state="disabled")
        self.btn_estop = ttk.Button(btns, text="E-STOP (OFF now)", command=self.estop, state="disabled")
        self.btn_disconnect = ttk.Button(btns, text="DISCONNECT", command=self.disconnect, state="disabled")

        self.btn_connect.grid(row=0, column=0, padx=6, sticky="ew")
        self.btn_start.grid(row=0, column=1, padx=6, sticky="ew")
        self.btn_stop.grid(row=0, column=2, padx=6, sticky="ew")
        self.btn_estop.grid(row=0, column=3, padx=6, sticky="ew")
        self.btn_disconnect.grid(row=0, column=4, padx=6, sticky="ew")

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
            self.lbl_other.config(text="Voltage set (V)")  # if you prefer "Voltage limit", change this text

    def _parse_float(self, s, name):
        try:
            return float(s)
        except:
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

    # -------- CONNECT / DISCONNECT --------
    def connect(self):
        if self.connected:
            messagebox.showinfo("Info", "Already connected.")
            return

        self.ctrl.resource_hint = self.var_hint.get().strip()
        self.status.set("Connecting to HMP4040...")
        self.update_idletasks()

        try:
            ok = self.ctrl.connect()
            if not ok:
                raise RuntimeError("Connection failed.")
            self.connected = True
            self.idn.set(f"Detected: {self.ctrl.idn}")
            self.status.set("Connected. Ready.")
            self.btn_start.config(state="normal")
            self.btn_disconnect.config(state="normal")
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
        except:
            pass
        try:
            self.ctrl.disconnect()
        except:
            pass

        self.connected = False
        self.idn.set("")
        self.status.set("Disconnected. Ready.")
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="disabled")
        self.btn_disconnect.config(state="disabled")
        self.btn_estop.config(state="disabled")

    # -------- RAMP CONTROL --------
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
            stab = self._parse_float(self.var_stab.get(), "stabilization time")
            other = self._parse_float(self.var_other.get(), "set value")

            if st == 0:
                raise ValueError("Step cannot be 0.")
            if stab < 0:
                raise ValueError("Stabilization time must be >= 0.")
            if other < 0:
                raise ValueError("Set value must be >= 0.")

            values = frange(mn, mx, st)
            if not values:
                raise ValueError("No steps generated. Check min/max/step.")
        except Exception as e:
            messagebox.showerror("Input error", str(e))
            return

        # Wiring advice using fixed limits
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
                f"Max voltage needed: {max_v_needed:.2f} V > {self.V_MAX_PER_CH:.2f} V (per channel)\n\n"
                "Please wire CH1 + CH2 in SERIES (per lab procedure), then press OK."
            )
            if not messagebox.askokcancel("Wiring required: SERIES", msg):
                return
        elif wiring == "PARALLEL":
            msg = (
                f"Max current needed: {max_i_needed:.2f} A > {self.I_MAX_PER_CH:.2f} A (per channel)\n\n"
                "Please wire CH1 + CH2 in PARALLEL (per lab procedure), then press OK."
            )
            if not messagebox.askokcancel("Wiring required: PARALLEL", msg):
                return

        # Start worker thread
        self.stop_requested = False
        self.running = True
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.btn_disconnect.config(state="disabled")
        self.status.set("Running...")
        t = threading.Thread(target=self._run_ramp, args=(values, stab, other), daemon=True)
        t.start()

    def stop_soft(self):
        if self.running:
            self.stop_requested = True
            self.status.set("Soft stop requested (will stop at next point).")

    def estop(self):
        self.stop_requested = True
        try:
            self.ctrl.off_all()
        except:
            pass
        self.status.set("E-STOP: Outputs OFF.")
        self._finish()

    def _run_ramp(self, values, stab, other):
        try:
            self.ctrl.set_output(1, True)

            for idx, x in enumerate(values, start=1):
                if self.stop_requested:
                    break

                if self.mode.get() == "CV":
                    self.status.set(f"Step {idx}/{len(values)} -> CV: V={x}V, I(set)={other}A")
                    self.ctrl.apply_cv(voltage=x, current_set=other, ch=1)
                else:
                    self.status.set(f"Step {idx}/{len(values)} -> CC: I={x}A, V(set)={other}V")
                    self.ctrl.apply_cc(current_set=x, voltage_set=other, ch=1)

                t0 = time.time()
                while time.time() - t0 < stab:
                    if self.stop_requested:
                        break
                    time.sleep(0.1)

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
        if self.connected:
            self.btn_start.config(state="normal")
            self.btn_disconnect.config(state="normal")
            self.btn_stop.config(state="disabled")
        else:
            self.btn_start.config(state="disabled")
            self.btn_stop.config(state="disabled")
            self.btn_disconnect.config(state="disabled")


if __name__ == "__main__":
    app = RampGUI()
    app.mainloop()
