import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time

# ---------------------------
# TODO: Reemplaza esto por tu clase real de control HMP4040
# (con pyvisa o lo que uses). Aquí dejo un stub.
# ---------------------------
class HMP4040Controller:
    def connect(self):
        # Detect/Connect here
        return True

    def set_output(self, ch: int, on: bool):
        # OUTP ON/OFF
        pass

    def apply_cv(self, voltage: float, current_limit: float, ch: int = 1):
        # Set V + I limit for CV
        pass

    def apply_cc(self, current: float, voltage_limit: float, ch: int = 1):
        # Set I + V limit for CC
        pass

    def off_all(self):
        # Immediate OFF
        try:
            self.set_output(1, False)
            self.set_output(2, False)
        except:
            pass


def frange(start, stop, step):
    # Inclusive stop (si cae exacto)
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
    def __init__(self):
        super().__init__()
        self.title("HMP4040 Ramp Frontend (CV/CC)")
        self.geometry("560x420")
        self.resizable(False, False)

        self.ctrl = HMP4040Controller()
        self.running = False
        self.stop_requested = False

        # ---- Top: Mode + Limits
        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Mode").grid(row=0, column=0, sticky="w")
        self.mode = tk.StringVar(value="CV")
        ttk.Radiobutton(frm, text="Voltage Ramp (CV)", variable=self.mode, value="CV", command=self._refresh_labels).grid(row=0, column=1, sticky="w")
        ttk.Radiobutton(frm, text="Current Ramp (CC)", variable=self.mode, value="CC", command=self._refresh_labels).grid(row=0, column=2, sticky="w")

        sep = ttk.Separator(frm)
        sep.grid(row=1, column=0, columnspan=3, sticky="ew", pady=10)

        # Inputs
        self.lbl_min = ttk.Label(frm, text="V min (V)")
        self.lbl_max = ttk.Label(frm, text="V max (V)")
        self.lbl_step = ttk.Label(frm, text="V step (V)")

        self.lbl_min.grid(row=2, column=0, sticky="w")
        self.lbl_max.grid(row=3, column=0, sticky="w")
        self.lbl_step.grid(row=4, column=0, sticky="w")

        self.var_min = tk.StringVar(value="30")
        self.var_max = tk.StringVar(value="40")
        self.var_step = tk.StringVar(value="1")

        ttk.Entry(frm, textvariable=self.var_min, width=12).grid(row=2, column=1, sticky="w")
        ttk.Entry(frm, textvariable=self.var_max, width=12).grid(row=3, column=1, sticky="w")
        ttk.Entry(frm, textvariable=self.var_step, width=12).grid(row=4, column=1, sticky="w")

        ttk.Label(frm, text="Stabilization time (s)").grid(row=5, column=0, sticky="w")
        self.var_stab = tk.StringVar(value="10")
        ttk.Entry(frm, textvariable=self.var_stab, width=12).grid(row=5, column=1, sticky="w")

        # “Other limit” depending on mode
        self.lbl_other = ttk.Label(frm, text="Current limit (A)")  # in CV
        self.lbl_other.grid(row=6, column=0, sticky="w")
        self.var_other = tk.StringVar(value="1.0")
        ttk.Entry(frm, textvariable=self.var_other, width=12).grid(row=6, column=1, sticky="w")

        sep2 = ttk.Separator(frm)
        sep2.grid(row=7, column=0, columnspan=3, sticky="ew", pady=10)

        # Channel limits (editable)
        ttk.Label(frm, text="Per-channel limits (edit if needed):").grid(row=8, column=0, columnspan=3, sticky="w")

        ttk.Label(frm, text="V max per channel").grid(row=9, column=0, sticky="w")
        ttk.Label(frm, text="I max per channel").grid(row=10, column=0, sticky="w")

        self.var_vch = tk.StringVar(value="32")   # adjust if your unit differs
        self.var_ich = tk.StringVar(value="10")   # adjust if your unit differs

        ttk.Entry(frm, textvariable=self.var_vch, width=12).grid(row=9, column=1, sticky="w")
        ttk.Entry(frm, textvariable=self.var_ich, width=12).grid(row=10, column=1, sticky="w")

        # Status
        self.status = tk.StringVar(value="Ready.")
        ttk.Label(frm, textvariable=self.status).grid(row=11, column=0, columnspan=3, sticky="w", pady=(10, 0))

        # Buttons
        btns = ttk.Frame(frm)
        btns.grid(row=12, column=0, columnspan=3, sticky="ew", pady=12)
        btns.columnconfigure(0, weight=1)
        btns.columnconfigure(1, weight=1)
        btns.columnconfigure(2, weight=1)

        self.btn_start = ttk.Button(btns, text="START", command=self.start)
        self.btn_stop = ttk.Button(btns, text="STOP (soft)", command=self.stop_soft)
        self.btn_estop = ttk.Button(btns, text="E-STOP (OFF now)", command=self.estop)

        self.btn_start.grid(row=0, column=0, padx=6, sticky="ew")
        self.btn_stop.grid(row=0, column=1, padx=6, sticky="ew")
        self.btn_estop.grid(row=0, column=2, padx=6, sticky="ew")

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

        # If both exceeded, typical 2-ch series+parallel simultaneously is not practical
        if need_series and need_parallel:
            return "NOT_PRACTICAL"

        if need_series:
            return "SERIES"
        if need_parallel:
            return "PARALLEL"
        return "SINGLE"

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

            values = frange(mn, mx, st)
            if not values:
                raise ValueError("No steps generated. Check min/max/step.")
        except Exception as e:
            messagebox.showerror("Input error", str(e))
            return

        # Determine what maximum V/I is needed to advise wiring
        if self.mode.get() == "CV":
            max_v_needed = max(mn, mx)
            max_i_needed = other  # current limit
        else:
            max_i_needed = max(mn, mx)
            max_v_needed = other  # voltage limit

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
                return
        elif wiring == "PARALLEL":
            msg = (
                f"Max current needed: {max_i_needed:.2f} A > {i_ch:.2f} A (per channel)\n\n"
                "Please wire CH1 + CH2 in PARALLEL (per lab procedure), then press OK."
            )
            if not messagebox.askokcancel("Wiring required: PARALLEL", msg):
                return

        # Connect once
        self.status.set("Connecting to HMP4040...")
        try:
            ok = self.ctrl.connect()
            if not ok:
                raise RuntimeError("Connection failed.")
        except Exception as e:
            messagebox.showerror("Connection error", str(e))
            self.status.set("Ready.")
            return

        # Launch worker
        self.stop_requested = False
        self.running = True
        self.btn_start.config(state="disabled")
        t = threading.Thread(target=self._run_ramp, args=(values, stab, other), daemon=True)
        t.start()

    def stop_soft(self):
        if self.running:
            self.stop_requested = True
            self.status.set("Soft stop requested (will stop at next point).")

    def estop(self):
        # immediate OFF
        self.stop_requested = True
        try:
            self.ctrl.off_all()
        except:
            pass
        self.status.set("E-STOP: Outputs OFF.")
        self._finish()

    def _run_ramp(self, values, stab, other):
        try:
            # Turn output on (adjust for your actual channel strategy)
            self.ctrl.set_output(1, True)

            for idx, x in enumerate(values, start=1):
                if self.stop_requested:
                    break

                if self.mode.get() == "CV":
                    self.status.set(f"Step {idx}/{len(values)}  -> V={x}V, Ilim={other}A")
                    self.ctrl.apply_cv(voltage=x, current_limit=other, ch=1)
                else:
                    self.status.set(f"Step {idx}/{len(values)}  -> I={x}A, Vlim={other}V")
                    self.ctrl.apply_cc(current=x, voltage_limit=other, ch=1)

                # Stabilization delay
                t0 = time.time()
                while time.time() - t0 < stab:
                    if self.stop_requested:
                        break
                    time.sleep(0.1)

            # done: turn off
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
