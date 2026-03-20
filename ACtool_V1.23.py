
# ======================================================
# UNIFIED SOFTWARE – Voltage Ramp + WT1800/EL34143 Auto Measure
# Excel output cloned from original DR tool AC script
# Multi-ramp support with blank line separators
# Matias Riquelme, CPI – Axis Lighting
# ======================================================

import tkinter as tk
from tkinter import messagebox, filedialog
import time
import threading
import keyboard
import pygetwindow as gw
import pyvisa

import pandas as pd      # kept in case you want to reuse it later
import openpyxl          # same here
import xlsxwriter        # to replicate the original Excel format exactly

# ================= EL34143 DRIVER =====================
EL34143 = None
EL34143_USB = None
el_connected = False

def EL34143_init():
    """Initialize EL34143 electronic load over USB."""
    global EL34143, EL34143_USB, el_connected
    try:
        EL34143_USB = pyvisa.ResourceManager()
        # EL34143 electronic load via USB
        EL34143 = EL34143_USB.open_resource('USB0::0x2A8D::0x3802::MY60260348::INSTR')
        EL34143.timeout = 10000
        ID = EL34143.query("*IDN?")
        print("EL34143 CONNECTED:", ID)
        EL34143.write("*CLS")
        EL34143.write("*RST")
        el_connected = True
    except Exception as e:
        el_connected = False
        raise e

def EL34143_ON():
    if EL34143 is not None:
        # Same as in the original code
        EL34143.write("INP ON, (@1)")

def EL34143_OFF():
    if EL34143 is not None:
        EL34143.write("INP OFF, (@1)")

def EL34143_SET_CV(v):
    """Set constant voltage mode and target voltage."""
    EL34143.write("FUNC VOLT, (@1)")
    EL34143.write(f"VOLT {v}, (@1)")

# --- Output readings (same as in original EL34143.py) ---
def EL34143_READ_VOLTAGE():
    return float(EL34143.query("MEAS:VOLT? (@1)"))

def EL34143_READ_CURRENT():
    return float(EL34143.query("MEAS:CURR? (@1)"))

def EL34143_READ_POW():
    return float(EL34143.query("MEAS:POW? (@1)"))

# ================= WT1800 DRIVER ======================
WT1800 = None
wt_connected = False

def WT1800_init():
    """Initialize WT1800 power analyzer over USB."""
    global WT1800, wt_connected
    try:
        rm = pyvisa.ResourceManager()
        WT1800 = rm.open_resource('USB0::0x0B21::0x0025::43325247323730303756::INSTR')
        WT1800.timeout = 10000
        ID = WT1800.query("*IDN?")
        print("WT1800 CONNECTED:", ID)
        WT1800.write(":NUMERIC:FORMAT ASCII")
        WT1800.write(":INPUT:VOLTAGE:AUTO:ELEMENT1 ON")
        WT1800.write(":INPUT:CURRENT:AUTO:ELEMENT1 ON")
        wt_connected = True
    except Exception as e:
        wt_connected = False
        raise e

# --- WT1800 read functions (ALL MEASURED, NOTHING CALCULATED) ---

def WT1800_READ_URMS1():
    WT1800.write(":NUMERIC:NORMAL:ITEM1 URMS,1")
    return float(WT1800.query(":NUMeric:NORMal:VALue? 1"))

def WT1800_READ_IRMS1():
    WT1800.write(":NUMERIC:NORMAL:ITEM2 IRMS,1")
    return float(WT1800.query(":NUMeric:NORMal:VALue? 2"))

def WT1800_READ_POW():
    WT1800.write(":NUMERIC:NORMAL:ITEM3 P,1")
    return float(WT1800.query(":NUMeric:NORMal:VALue? 3"))

def WT1800_READ_THDi():
    WT1800.write(":NUMERIC:NORMAL:ITEM4 ITHD,1")
    return float(WT1800.query(":NUMeric:NORMal:VALue? 4"))

def WT1800_READ_PF():
    WT1800.write(":NUMERIC:NORMAL:ITEM5 LAMbda,1")
    return float(WT1800.query(":NUMeric:NORMal:VALue? 5"))

# ============ ENTER trigger for DR Tool AC (optional) ===========
stop_flag = False

def run_data_capture(num_steps, delay, initial_delay=0):
    """
    Sends ENTER to the DR tool AC window once per step.
    If you are no longer using DR tool AC, you can comment
    out the call to this function.
    """
    def loop():
        time.sleep(initial_delay)
        for i in range(num_steps):
            if stop_flag:
                break
            print(f"[ENTER] STEP {i+1}/{num_steps}")
            try:
                target = [w for w in gw.getWindowsWithTitle("DR tool AC") if w.title]
                if target:
                    target[0].activate()
                    time.sleep(0.5)
                    keyboard.send('enter')
                else:
                    print("Window DR tool AC not found")
            except Exception as e:
                print("Window error:", e)
            time.sleep(delay)
    threading.Thread(target=loop, daemon=True).start()

# ================= Voltage Ramp + WT1800/EL34143 Measurement ======================

# Results from the current ramp
results = []
# Accumulated results from ALL ramps (with None as a blank-line separator)
all_results = []

def measure_and_build_row(step_index, v_set):
    """
    Replicates the logic from the original script:
    - Reads Vout, Iout, Pout from EL34143
    - Reads Uin, Iin, Pin, THD, PF from WT1800
    - Computes Eff = (Pout/Pin)*100, rounded to 4 decimals
    - Returns a dict with all values (does NOT store it directly)
    """
    try:
        # OUT from EL34143
        uo = EL34143_READ_VOLTAGE()
        time.sleep(0.3)
        io = EL34143_READ_CURRENT()
        time.sleep(0.3)
        Po = EL34143_READ_POW()

        # IN from WT1800
        time.sleep(0.3)
        Vin = WT1800_READ_URMS1()
        Iin = WT1800_READ_IRMS1()
        Pin = WT1800_READ_POW()
        Thd = WT1800_READ_THDi()
        Pf = WT1800_READ_PF()

        # Efficiency (same as the original code)
        if Pin != 0:
            Eff = round(float((Po / Pin) * 100), 4)
        else:
            Eff = 0.0

        row = {
            "Step": step_index,
            "Uin": Vin,
            "Iin": Iin,
            "Pin": Pin,
            "Uout": uo,
            "Iout": io,
            "Pout": Po,
            "Eff": Eff,
            "THD": Thd,
            "PF": Pf
        }
        print("MEASURED:", row)
        return row

    except Exception as e:
        print("WT/EL measure error:", e)
        return None

def apply_voltage_ramp():
    """Main ramp logic + measurement for each step."""
    global stop_flag, results, all_results
    try:
        if not (el_connected and wt_connected):
            return messagebox.showerror("Connection", "Connect instruments first")

        v_min = float(entry_min.get())
        v_max = float(entry_max.get())
        step = float(entry_step.get())
        delay = float(entry_delay.get())

        if v_min >= v_max:
            return messagebox.showerror("Error", "Min must be < Max")

        # Reset flag and results for THIS ramp at start
        stop_flag = False
        results = []

        # Build voltage list, making sure we include v_max
        voltages = [round(v_min + i * step, 2)
                    for i in range(int((v_max - v_min) // step + 1))]
        if voltages[-1] != round(v_max, 2):
            voltages.append(round(v_max, 2))
        num_steps = len(voltages)
        print("Voltages sequence:", voltages)

        # 1) Set FIRST VOLTAGE
        first_v = voltages[0]
        EL34143_SET_CV(first_v)

        # 2) TURN LOAD ON HERE
        print("Sending INPUT ON")
        EL34143_ON()

        status_label.config(text=f"V Applied: {first_v} V (step 1)", fg="green")
        root.update()

        # 3) Initial delay for the first point to settle
        time.sleep(delay)

        # 4) Measure and store STEP 1
        row = measure_and_build_row(step_index=1, v_set=first_v)
        if row is not None:
            results.append(row)

        # 5) (Optional) Start ENTER trigger thread for DR tool AC
        run_data_capture(num_steps, delay, initial_delay=0)

        # 6) Apply remaining steps
        for idx, v in enumerate(voltages[1:], start=2):
            if stop_flag:
                break
            EL34143_SET_CV(v)
            status_label.config(text=f"V Applied: {v} V (step {idx})", fg="green")
            root.update()
            time.sleep(delay)
            row = measure_and_build_row(step_index=idx, v_set=v)
            if row is not None:
                results.append(row)

        # 7) Turn load OFF when ramp completes
        EL34143_OFF()

        if not stop_flag and results:
            # First accumulate THIS ramp into all_results
            # If there were previous ramps, add a blank-line separator
            if all_results:
                all_results.append(None)  # separator = blank row

            all_results.extend(results)

            status_label.config(text="RAMP FINISHED", fg="blue")

            # Ask if we should save ALL accumulated ramps now
            save_now = messagebox.askyesno(
                "Finish measurements",
                "The ramp has finished.\n\n"
                "Did you finish your measurements and want to save ALL results?\n\n"
                "If you answer 'No', you can run another ramp; "
                "the current data will remain accumulated."
            )

            if save_now:
                if all_results:
                    save_measurements_to_excel_original_format()
                    # After saving, clear everything for a new file
                    all_results.clear()
                    results.clear()
                    messagebox.showinfo("Done", "Voltage ramp completed and data saved.")
            else:
                messagebox.showinfo(
                    "Info",
                    "OK. You can adjust a new ramp and press 'Start Ramp'.\n\n"
                    "The data from this ramp has been added after a blank row."
                )

    except Exception as e:
        messagebox.showerror("Ramp Error", str(e))
        try:
            EL34143_OFF()
        except:
            pass

# ---------- Save results to Excel using ORIGINAL FORMAT ----------

def save_measurements_to_excel_original_format():
    """
    Uses all_results (global list) which contains:
      - dicts with measurement data
      - None as ramp separators (blank rows)
    and generates an Excel file with the SAME format as the original script.
    """
    global all_results

    if not all_results:
        messagebox.showwarning("No data", "There is no accumulated data to save.")
        return

    # Timestamp like "Thu Dec  8 22:54:52 2025"
    TimeDisplay_Excel = time.asctime(time.localtime(time.time()))

    # Default filename, same pattern as original: Export_file_<timestamp>_.xlsx
    safe_time = TimeDisplay_Excel.replace(" ", "_").replace(":", "_")
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

    # Create workbook and worksheet exactly like in the original code
    workbook = xlsxwriter.Workbook(filepath)
    worksheet = workbook.add_worksheet()

    # Header (same positions as the original script)
    worksheet.write('B3', 'Date')
    worksheet.write('C3', TimeDisplay_Excel)
    worksheet.write('B4', 'Manufacturer')
    worksheet.write('B5', 'Model')
    worksheet.write('B6', 'Description')
    worksheet.write('B7', 'Current')
    worksheet.write('B8', 'Voltage range')
    worksheet.write('B9', 'Dimming')
    worksheet.write('B10', 'Setpoints')

    # Measurement table headers (row 14)
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

    # Data starts at row 15 (same as "row = 15" in original)
    row_xls = 15

    for r in all_results:
        if r is None:
            # Blank row between ramps
            row_xls += 1
            continue

        worksheet.write(row_xls, 3, r["Uin"])   # D: U in (V)
        worksheet.write(row_xls, 4, r["Iin"])   # E: I in (A)
        worksheet.write(row_xls, 5, r["Pin"])   # F: P in (W)
        worksheet.write(row_xls, 6, r["Uout"])  # G: V out (V)
        worksheet.write(row_xls, 7, r["Iout"])  # H: I out (A)
        worksheet.write(row_xls, 8, r["Pout"])  # I: P out (W)
        worksheet.write(row_xls, 9, r["Eff"])   # J: Efficiency(%)
        worksheet.write(row_xls, 10, r["THD"])  # K: THD(%)
        worksheet.write(row_xls, 11, r["PF"])   # L: Power Factor
        row_xls += 1

    workbook.close()
    messagebox.showinfo("Success", f"Results saved to:\n{filepath}")
    print("Results saved to:", filepath)

# ============== STOP ==================
def emergency_stop():
    """Emergency stop button: stops ramp and turns the load OFF."""
    global stop_flag
    stop_flag = True
    try:
        EL34143_OFF()
    except:
        pass
    status_label.config(text="STOPPED", fg="red")

# ============== RESET =================
def reset_form():
    """Clear GUI input fields and reset status label."""
    global stop_flag
    stop_flag = False
    for e in (entry_min, entry_max, entry_step, entry_delay):
        e.delete(0, tk.END)
    status_label.config(text="Ready", fg="blue")

# ====== Connect Instruments First ======
def connect_instruments():
    """Connect both EL34143 and WT1800 before enabling the ramp."""
    status_label.config(text="Connecting EL34143...", fg="orange")
    root.update()
    EL34143_init()

    status_label.config(text="Connecting WT1800...", fg="orange")
    root.update()
    WT1800_init()

    status_label.config(text="CONNECTED ✓", fg="green")
    start_button.config(state=tk.NORMAL)
    for e in (entry_min, entry_max, entry_step, entry_delay):
        e.config(state=tk.NORMAL)

# ======================================================
#                      GUI
# ======================================================
root = tk.Tk()
root.title("Voltage Ramp + WT1800/EL34143 Auto Measure (DR AC Format)")

tk.Label(root, text="Min Voltage (V):").grid(row=0, column=0, padx=5, pady=5, sticky="e")
tk.Label(root, text="Max Voltage (V):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
tk.Label(root, text="Step (V):").grid(row=2, column=0, padx=5, pady=5, sticky="e")
tk.Label(root, text="Delay (s):").grid(row=3, column=0, padx=5, pady=5, sticky="e")

entry_min = tk.Entry(root); entry_min.grid(row=0, column=1, pady=5)
entry_max = tk.Entry(root); entry_max.grid(row=1, column=1, pady=5)
entry_step = tk.Entry(root); entry_step.grid(row=2, column=1, pady=5)
entry_delay = tk.Entry(root); entry_delay.grid(row=3, column=1, pady=5)

for e in (entry_min, entry_max, entry_step, entry_delay):
    e.config(state=tk.DISABLED)

tk.Button(root, text="Connect Instruments", bg="lightblue",
          command=connect_instruments).grid(row=4, column=0, columnspan=2, pady=10)

start_button = tk.Button(
    root,
    text="▶ Start Ramp",
    bg="lightgreen",
    command=lambda: threading.Thread(target=apply_voltage_ramp, daemon=True).start(),
    state=tk.DISABLED
)
start_button.grid(row=5, column=0, pady=10)

tk.Button(root, text="STOP!", bg="tomato", command=emergency_stop).grid(row=5, column=1, pady=10)
tk.Button(root, text="Reset", command=reset_form).grid(row=6, column=0, columnspan=2)

status_label = tk.Label(root, text="Connect instruments first", fg="blue")
status_label.grid(row=7, column=0, columnspan=2, pady=10)

root.mainloop()

