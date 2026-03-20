"""
HMP4040 LED Characterization Runner (Console UI + Hotkeys)
- Supports CH1 only, CH1+CH2 SERIES (for >32V), CH1+CH2 PARALLEL (for >10A)
- PSU output stays ON between steps (no OFF between steps)
- Per-step stabilization + 2-way readback snapshot (V/I/P if available)
- Console “frontend” with non-blocking hotkeys (Windows):
    [E] Emergency stop (ramps to 0 and turns outputs OFF immediately)
    [R] Soft reset request (sets next step to 0V/0A then continues unless you also stop)
    [S] Save snapshots to CSV now (does NOT pause the ramp)
    [P] Print last snapshot (does NOT pause the ramp)

Notes:
- For SERIES/PARALLEL wiring, follow your lab’s approved procedure/manual.
- PARALLEL current sharing is best done using the instrument’s supported parallel/tracking mode if available.
"""

try:
    import time
    import serial
    import sys
    import csv
    import threading

    # winsound only on Windows
    try:
        import winsound
        HAS_WINSOUND = True
    except:
        HAS_WINSOUND = False

    # Non-blocking keyboard (Windows)
    try:
        import msvcrt
        HAS_MSVCRT = True
    except:
        HAS_MSVCRT = False

    ##################################################################
    # Limits (per channel)
    V_MAX_CH = 32.0
    I_MAX_CH = 10.0
    P_MAX_CH = 160.0

    MODE_SINGLE   = "SINGLE_CH1"
    MODE_SERIES   = "SERIES_CH1_CH2"
    MODE_PARALLEL = "PARALLEL_CH1_CH2"

    ##################################################################
    # Global control flags (hotkeys)
    stop_event = threading.Event()          # emergency stop
    soft_reset_request = threading.Event()  # request apply 0,0 at next safe point (doesn't necessarily stop)
    save_request = threading.Event()        # save CSV now
    print_request = threading.Event()       # print last snapshot

    ##################################################################
    # Helpers
    def beep(freq=3500, dur=500):
        if HAS_WINSOUND:
            winsound.Beep(freq, dur)

    def ask_float(prompt, min_val=None, allow_empty=False, default=None):
        while True:
            s = input(prompt).strip()
            if s == "" and allow_empty:
                return default
            try:
                v = float(s.replace(",", "."))
                if min_val is not None and v < min_val:
                    print(f"  -> Must be >= {min_val}")
                    continue
                return v
            except:
                print("  -> Please enter a valid number.")

    def ask_int(prompt, min_val=None, allow_empty=False, default=None):
        while True:
            s = input(prompt).strip()
            if s == "" and allow_empty:
                return default
            try:
                v = int(s)
                if min_val is not None and v < min_val:
                    print(f"  -> Must be >= {min_val}")
                    continue
                return v
            except:
                print("  -> Please enter a valid integer.")

    def ask_yes_no(prompt, yes=("y", "yes", "oui"), no=("n", "no", "non")):
        while True:
            s = input(prompt).strip().lower()
            if s in yes:
                return True
            if s in no:
                return False
            print("  -> Invalid. Use y/n or oui/non.")

    def scpi_write(dev, cmd):
        dev.write((cmd.strip() + "\n").encode())

    def scpi_query(dev, cmd, delay=0.15):
        try:
            dev.reset_input_buffer()
        except:
            pass
        scpi_write(dev, cmd)
        time.sleep(delay)
        return dev.readline().decode(errors="ignore").strip()

    def try_float(s):
        try:
            return float(s)
        except:
            return None

    def find_hmp4040(max_com=30):
        print(" Detecting HMP4040...")
        print(" Please Wait...")
        i = 0
        while True:
            try:
                comx = f"COM{i}"
                hmp = serial.Serial(comx, 9600, timeout=3)

                scpi_write(hmp, "*IDN?")
                ID = hmp.readline().decode(errors="ignore").strip()

                if ID.startswith("ROHDE&SCHWARZ"):
                    print("\n Detected:")
                    print(" ", ID)
                    return hmp

            except Exception:
                i += 1
                if i > max_com:
                    print("\n Not Detected!")
                    input(" To try again, press ENTER...")
                    i = 0

    ##################################################################
    # Channel control
    def hmp_select(hmp, ch):  # "OUT1" etc
        scpi_write(hmp, f"INST {ch}")

    def hmp_apply_ch(hmp, ch, v, i):
        hmp_select(hmp, ch)
        scpi_write(hmp, f"APPLY {v},{i}")

    def hmp_output_ch(hmp, ch, on):
        hmp_select(hmp, ch)
        scpi_write(hmp, "OUTP ON" if on else "OUTP OFF")

    def output_on(hmp, mode):
        if mode == MODE_SINGLE:
            hmp_output_ch(hmp, "OUT1", True)
        else:
            hmp_output_ch(hmp, "OUT1", True)
            hmp_output_ch(hmp, "OUT2", True)

    def output_off(hmp, mode):
        if mode == MODE_SINGLE:
            hmp_output_ch(hmp, "OUT1", False)
        else:
            hmp_output_ch(hmp, "OUT1", False)
            hmp_output_ch(hmp, "OUT2", False)

    ##################################################################
    # Mode selection and setpoint application (with per-channel power checks)
    def choose_mode(v_req, i_req, mode_policy="AUTO"):
        """
        mode_policy:
          - "AUTO": choose based on limits for each step (requires wiring compatible with chosen mode!)
          - "FIXED_SINGLE" / "FIXED_SERIES" / "FIXED_PARALLEL": force one mode for the whole run
        """
        if mode_policy == "FIXED_SINGLE":
            return MODE_SINGLE
        if mode_policy == "FIXED_SERIES":
            return MODE_SERIES
        if mode_policy == "FIXED_PARALLEL":
            return MODE_PARALLEL

        # AUTO
        if v_req <= V_MAX_CH and i_req <= I_MAX_CH:
            return MODE_SINGLE

        if v_req > V_MAX_CH and v_req <= 2 * V_MAX_CH:
            if i_req > I_MAX_CH:
                raise ValueError("Cannot exceed 10A in SERIES mode (current is the same through both channels).")
            return MODE_SERIES

        if i_req > I_MAX_CH and i_req <= 2 * I_MAX_CH:
            if v_req > V_MAX_CH:
                raise ValueError("Cannot exceed 32V in PARALLEL mode (voltage is the same on both channels).")
            return MODE_PARALLEL

        raise ValueError("Requested V/I exceeds what CH1+CH2 can provide.")

    def print_mode_note(mode):
        if mode == MODE_SERIES:
            print("\n*** SERIES MODE (CH1 + CH2) ***")
            print("Lab note: CH1 and CH2 must be wired in SERIES per approved lab procedure/manual.")
            print("Example: 36V total => CH1=32V, CH2=4V (script splits automatically).")
            print("Limits per channel: 0-32V, 0-10A, 160W.")
        elif mode == MODE_PARALLEL:
            print("\n*** PARALLEL MODE (CH1 + CH2) ***")
            print("Lab note: CH1 and CH2 must be wired in PARALLEL per approved lab procedure/manual.")
            print("Total current is shared; script sets half current on each channel.")
            print("Limits per channel: 0-32V, 0-10A, 160W.")
        else:
            print("\n*** SINGLE MODE (CH1) ***")
            print("Limits: 0-32V, 0-10A, 160W.")

    def apply_setpoints(hmp, mode, v_req, i_req):
        """
        Applies requested total setpoints using CH1/CH2.
        Returns dict describing what was applied per channel.
        """
        if mode == MODE_SINGLE:
            if v_req > V_MAX_CH or i_req > I_MAX_CH:
                raise ValueError("SINGLE exceeds per-channel limits.")
            if v_req * i_req > P_MAX_CH:
                raise ValueError(f"CH1 power {v_req*i_req:.1f}W exceeds 160W limit.")
            hmp_apply_ch(hmp, "OUT1", v_req, i_req)
            return {"mode": mode, "V1": v_req, "I1": i_req, "V2": 0.0, "I2": 0.0}

        if mode == MODE_SERIES:
            if v_req > 2 * V_MAX_CH:
                raise ValueError("SERIES exceeds 64V total with two channels.")
            if i_req > I_MAX_CH:
                raise ValueError("SERIES cannot exceed 10A.")
            v1 = min(V_MAX_CH, v_req)
            v2 = max(0.0, v_req - v1)

            if v1 * i_req > P_MAX_CH:
                raise ValueError(f"CH1 power {v1*i_req:.1f}W exceeds 160W limit.")
            if v2 * i_req > P_MAX_CH:
                raise ValueError(f"CH2 power {v2*i_req:.1f}W exceeds 160W limit.")

            hmp_apply_ch(hmp, "OUT1", v1, i_req)
            hmp_apply_ch(hmp, "OUT2", v2, i_req)
            return {"mode": mode, "V1": v1, "I1": i_req, "V2": v2, "I2": i_req}

        if mode == MODE_PARALLEL:
            if v_req > V_MAX_CH:
                raise ValueError("PARALLEL cannot exceed 32V.")
            if i_req > 2 * I_MAX_CH:
                raise ValueError("PARALLEL cannot exceed ~20A with two channels.")

            i_each = i_req / 2.0
            if v_req * i_each > P_MAX_CH:
                raise ValueError(f"Per-channel power {v_req*i_each:.1f}W exceeds 160W limit.")

            # Best practice: use instrument parallel/tracking mode if available (manual/procedure).
            hmp_apply_ch(hmp, "OUT1", v_req, i_each)
            hmp_apply_ch(hmp, "OUT2", v_req, i_each)
            return {"mode": mode, "V1": v_req, "I1": i_each, "V2": v_req, "I2": i_each}

        raise ValueError("Unknown mode.")

    ##################################################################
    # 2-way measurement snapshot (tries common SCPI variants)
    def measure_snapshot(hmp, mode):
        """
        Returns measured V/I/P. In SERIES/PARALLEL, also attempts per-channel readbacks.
        If commands not supported, values remain None.
        """
        def meas_for_channel(ch):
            hmp_select(hmp, ch)

            candidates_v = ["MEAS:VOLT?", "MEAS:VOLT? "+ch, "MEAS:VOLT? 1"]
            candidates_i = ["MEAS:CURR?", "MEAS:CURR? "+ch, "MEAS:CURR? 1"]
            candidates_p = ["MEAS:POW?",  "MEAS:POW? "+ch,  "MEAS:POW? 1"]

            out = {"V": None, "I": None, "P": None}

            for cmd in candidates_v:
                v = try_float(scpi_query(hmp, cmd))
                if v is not None:
                    out["V"] = v
                    break
            for cmd in candidates_i:
                i = try_float(scpi_query(hmp, cmd))
                if i is not None:
                    out["I"] = i
                    break
            for cmd in candidates_p:
                p = try_float(scpi_query(hmp, cmd))
                if p is not None:
                    out["P"] = p
                    break
            return out

        ch1 = meas_for_channel("OUT1")
        if mode == MODE_SINGLE:
            return {"V_meas": ch1["V"], "I_meas": ch1["I"], "P_meas": ch1["P"],
                    "V1_meas": ch1["V"], "I1_meas": ch1["I"], "P1_meas": ch1["P"],
                    "V2_meas": None,     "I2_meas": None,     "P2_meas": None}

        ch2 = meas_for_channel("OUT2")

        # Estimate total:
        if mode == MODE_SERIES:
            Vtot = None
            if ch1["V"] is not None and ch2["V"] is not None:
                Vtot = ch1["V"] + ch2["V"]
            Itot = ch1["I"]  # should be same
            Ptot = None
            if ch1["P"] is not None and ch2["P"] is not None:
                Ptot = ch1["P"] + ch2["P"]
            return {"V_meas": Vtot, "I_meas": Itot, "P_meas": Ptot,
                    "V1_meas": ch1["V"], "I1_meas": ch1["I"], "P1_meas": ch1["P"],
                    "V2_meas": ch2["V"], "I2_meas": ch2["I"], "P2_meas": ch2["P"]}

        # PARALLEL: voltage same, currents add (approx)
        Vpar = ch1["V"]
        Ipar = None
        if ch1["I"] is not None and ch2["I"] is not None:
            Ipar = ch1["I"] + ch2["I"]
        Ppar = None
        if ch1["P"] is not None and ch2["P"] is not None:
            Ppar = ch1["P"] + ch2["P"]

        return {"V_meas": Vpar, "I_meas": Ipar, "P_meas": Ppar,
                "V1_meas": ch1["V"], "I1_meas": ch1["I"], "P1_meas": ch1["P"],
                "V2_meas": ch2["V"], "I2_meas": ch2["I"], "P2_meas": ch2["P"]}

    ##################################################################
    # Non-blocking hotkey thread
    def hotkey_listener():
        if not HAS_MSVCRT:
            return
        while not stop_event.is_set():
            if msvcrt.kbhit():
                ch = msvcrt.getch()
                try:
                    key = ch.decode(errors="ignore").lower()
                except:
                    key = ""
                if key == "e":
                    stop_event.set()
                elif key == "r":
                    soft_reset_request.set()
                elif key == "s":
                    save_request.set()
                elif key == "p":
                    print_request.set()
            time.sleep(0.05)

    ##################################################################
    # Time wait that can respond to emergency stop
    def wait_seconds_interruptible(seconds, label=None):
        if label:
            print("\n" + label)
        start = time.time()
        while True:
            if stop_event.is_set():
                return False
            elapsed = time.time() - start
            left = seconds - elapsed
            if left <= 0:
                break
            if int(left) <= 10 or int(left) % 10 == 0:
                print(f"  {int(left)} s left\r", end="")
            time.sleep(1)
        print("  0 s left        ")
        return True

    def save_log_csv(run_log, filename="run_log_hmp4040.csv"):
        if not run_log:
            print("\n[Save] No snapshots to save yet.")
            return
        with open(filename, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=run_log[0].keys())
            writer.writeheader()
            writer.writerows(run_log)
        print(f"\n[Save] Saved: {filename}")

    ##################################################################
    # MAIN
    print("\nImporting files...")
    print(time.asctime(time.localtime(time.time())))

    # Detect instrument
    HMP4040 = find_hmp4040(max_com=30)

    # Init
    scpi_write(HMP4040, "*RST?")
    hmp_apply_ch(HMP4040, "OUT1", 0, 0)
    hmp_apply_ch(HMP4040, "OUT2", 0, 0)
    output_off(HMP4040, MODE_PARALLEL)  # turns off OUT1 & OUT2
    time.sleep(0.5)

    print("\n================= CONSOLE FRONTEND =================")
    print("Hotkeys (does NOT pause ramp):")
    print("  [E] Emergency stop (OFF immediately)")
    print("  [R] Soft reset request (apply 0/0 at next safe point)")
    print("  [S] Save snapshots to CSV now")
    print("  [P] Print last snapshot")
    print("====================================================\n")

    print(" Connect (follow lab procedure/manual):")
    print("  - Default: CH1 (OUT1)")
    print("  - If you need >32V: wire CH1+CH2 in SERIES")
    print("  - If you need >10A: wire CH1+CH2 in PARALLEL")
    input("\n Press ENTER when wiring is ready...")

    if not ask_yes_no(" Do you want to Continue? [y/n]: "):
        print("\n Procedure aborted!")
        scpi_write(HMP4040, "SYST:LOC")
        output_off(HMP4040, MODE_PARALLEL)
        sys.exit()

    ##################################################################
    # Sequence setup
    print("\n=== Sequence Setup ===")
    confirm_each = ask_yes_no(" Confirm each step before proceeding? (oui/non): ",
                             yes=("oui", "y", "yes"), no=("non", "n", "no"))

    # Mode policy: FIXED is safer because wiring must match the mode
    print("\nMode policy:")
    print("  1) FIXED (recommended) -> you choose ONE wiring mode for the whole run")
    print("  2) AUTO per step (advanced) -> only if your wiring supports changes (usually not practical)")
    pol = input("Choose [1/2]: ").strip()
    mode_policy = "FIXED_SINGLE"
    if pol == "2":
        mode_policy = "AUTO"
    else:
        print("\nChoose FIXED wiring mode:")
        print("  1) CH1 only (<=32V, <=10A)")
        print("  2) SERIES CH1+CH2 (<=64V, <=10A)")
        print("  3) PARALLEL CH1+CH2 (<=32V, <=20A approx)")
        mm = input("Choose [1/2/3]: ").strip()
        if mm == "2":
            mode_policy = "FIXED_SERIES"
        elif mm == "3":
            mode_policy = "FIXED_PARALLEL"
        else:
            mode_policy = "FIXED_SINGLE"

    # Define steps
    mode = input("\nChoose steps input: 1) Manual steps  2) Generate ramp  [1/2]: ").strip()
    if mode not in ("1", "2"):
        mode = "1"

    steps = []
    if mode == "1":
        print("\nEnter steps. Leave Voltage empty to finish.\n")
        step_idx = 1
        while True:
            v_in = input(f" Step {step_idx} - Voltage (V) [empty to end]: ").strip()
            if v_in == "":
                break
            try:
                v = float(v_in.replace(",", "."))
            except:
                print("  -> Invalid voltage.")
                continue
            i = ask_float(f" Step {step_idx} - Current (A): ", min_val=0)
            stab = ask_float(f" Step {step_idx} - Stabilization time (s): ", min_val=0)
            steps.append((v, i, stab))
            step_idx += 1
    else:
        print("\nGenerate ramp (linear).")
        v_start = ask_float(" Start Voltage (V): ", min_val=0)
        v_end   = ask_float(" End Voltage (V): ", min_val=0)
        i_start = ask_float(" Start Current (A): ", min_val=0)
        i_end   = ask_float(" End Current (A): ", min_val=0)
        n_steps = ask_int(" Number of steps (>=2): ", min_val=2)
        stab    = ask_float(" Stabilization time per step (s): ", min_val=0)

        for k in range(n_steps):
            alpha = k / (n_steps - 1)
            v = v_start + alpha * (v_end - v_start)
            i = i_start + alpha * (i_end - i_start)
            steps.append((round(v, 4), round(i, 6), stab))

    if not steps:
        print("\nNo steps defined. Exiting.")
        scpi_write(HMP4040, "SYST:LOC")
        output_off(HMP4040, MODE_PARALLEL)
        sys.exit()

    print("\n=== Steps Preview ===")
    for idx, (v, i, stab) in enumerate(steps, start=1):
        print(f" {idx:02d}) V={v}  I={i}  Stabilize={stab}s")

    if not ask_yes_no("\nStart sequence now? (y/n): "):
        print("\nAborted by user.")
        scpi_write(HMP4040, "SYST:LOC")
        output_off(HMP4040, MODE_PARALLEL)
        sys.exit()

    # If FIXED, print note once
    if mode_policy == "FIXED_SINGLE":
        fixed_mode = MODE_SINGLE
    elif mode_policy == "FIXED_SERIES":
        fixed_mode = MODE_SERIES
    elif mode_policy == "FIXED_PARALLEL":
        fixed_mode = MODE_PARALLEL
    else:
        fixed_mode = None

    if fixed_mode is not None:
        print_mode_note(fixed_mode)
        input("Press ENTER to confirm wiring matches the selected FIXED mode...")

    ##################################################################
    # Start hotkey thread
    if HAS_MSVCRT:
        t_hot = threading.Thread(target=hotkey_listener, daemon=True)
        t_hot.start()
    else:
        print("\n[Info] Non-blocking hotkeys not available on this platform (msvcrt missing).")

    ##################################################################
    # RUN (PSU stays ON; no OFF between steps)
    run_log = []

    # Determine initial mode for turning outputs ON
    initial_mode = fixed_mode if fixed_mode is not None else choose_mode(steps[0][0], steps[0][1], mode_policy="AUTO")
    print("\n Output ON")
    output_on(HMP4040, initial_mode)
    time.sleep(0.3)

    for idx, (v_set, i_set, stab) in enumerate(steps, start=1):
        if stop_event.is_set():
            break

        # Determine mode for this step
        step_mode = fixed_mode if fixed_mode is not None else choose_mode(v_set, i_set, mode_policy="AUTO")

        print(f"\n--- STEP {idx}/{len(steps)} ---  Mode={step_mode}")
        print(f" Target: V={v_set}  I={i_set}  Stabilize={stab}s")

        # Apply soft reset if requested (doesn't stop unless you also press E)
        if soft_reset_request.is_set():
            soft_reset_request.clear()
            print("\n[Reset] Applying 0V/0A (outputs remain ON).")
            try:
                apply_setpoints(HMP4040, step_mode, 0.0, 0.0)
            except Exception as ex:
                print("[Reset] Failed:", ex)

        if confirm_each:
            ok = ask_yes_no(" Ready to apply this step? (oui/non): ", yes=("oui","y","yes"), no=("non","n","no"))
            if not ok:
                print(" Stopping sequence (user said non).")
                break

        # Apply setpoints with validation
        applied = apply_setpoints(HMP4040, step_mode, v_set, i_set)
        print(f" Applied: {applied}")

        # Stabilize (interruptible by emergency stop)
        ok_wait = wait_seconds_interruptible(stab, label=f" Stabilizing for {stab}s...")
        if not ok_wait:
            break

        # Snapshot
        snap_meas = measure_snapshot(HMP4040, step_mode)
        snap = {
            "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
            "step": idx,
            "mode": step_mode,
            "V_req": v_set,
            "I_req": i_set,
            "stab_s": stab,
            "V1_set": applied["V1"],
            "I1_set": applied["I1"],
            "V2_set": applied["V2"],
            "I2_set": applied["I2"],
            **snap_meas
        }
        run_log.append(snap)

        print(" End-of-step snapshot:")
        print(f"  V_meas={snap['V_meas']}  I_meas={snap['I_meas']}  P_meas={snap['P_meas']}")

        # Handle non-blocking requests (do NOT pause ramp)
        if print_request.is_set():
            print_request.clear()
            last = run_log[-1] if run_log else None
            print("\n[Print] Last snapshot:")
            print(last if last else "No snapshots yet.")

        if save_request.is_set():
            save_request.clear()
            save_log_csv(run_log, filename="run_log_hmp4040.csv")

        if confirm_each and idx < len(steps):
            ok2 = ask_yes_no(" Ready for next step? (oui/non): ", yes=("oui","y","yes"), no=("non","n","no"))
            if not ok2:
                print(" Stopping sequence (user said non).")
                break

    ##################################################################
    # EMERGENCY STOP or normal shutdown
    if stop_event.is_set():
        print("\n[EMERGENCY STOP] Ramping to 0 and turning outputs OFF NOW.")
        try:
            # Attempt to set 0/0 on both channels first
            hmp_apply_ch(HMP4040, "OUT1", 0.0, 0.0)
            hmp_apply_ch(HMP4040, "OUT2", 0.0, 0.0)
        except:
            pass
        try:
            output_off(HMP4040, MODE_PARALLEL)
        except:
            pass
        beep()
    else:
        print("\n Shutting down...")
        try:
            # Set to 0 and turn off depending on fixed or last mode
            last_mode = fixed_mode if fixed_mode is not None else initial_mode
            apply_setpoints(HMP4040, last_mode, 0.0, 0.0)
            output_off(HMP4040, last_mode)
        except Exception as ex:
            print(" Shutdown warning:", ex)
            try:
                output_off(HMP4040, MODE_PARALLEL)
            except:
                pass
        beep()

    # Save log at end (optional)
    save_log_csv(run_log, filename="run_log_hmp4040.csv")

    try:
        scpi_write(HMP4040, "SYST:LOC")
    except:
        pass

    print("\nDone.")
    input("Press ENTER to close.")

except Exception as e:
    print("*******************************************")
    print("There was a problem. Close the window, and then try running the procedure again.")
    print("")
    print("Oops!", e.__class__, "occurred.")
    print(str(e))
    try:
        scpi_write(HMP4040, "SYST:LOC")
        scpi_write(HMP4040, "OUTP OFF")
    except:
        pass
    input("")
