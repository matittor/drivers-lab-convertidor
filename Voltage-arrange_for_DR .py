# -*- coding: utf-8 -*-
"""
Created on Fri Oct 31 11:32:36 2025

@author: Matias R, CPI
"""
# -*- coding: utf-8 -*-
"""

"""
import pandas as pd
from tkinter import filedialog, messagebox, Tk
import os
 
def select_input_files():
    root = Tk()
    root.withdraw()
    files = filedialog.askopenfilenames(
        title="Select measurement files (120V and/or 277V)",
        filetypes=[("Excel files", "*.xlsx")]
    )
    return list(files)
 
def select_output_file():
    root = Tk()
    root.withdraw()
    output = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="Save output file as..."
    )
    return output
 
def read_data(filepath):
    xls = pd.ExcelFile(filepath)
    for sheet in xls.sheet_names:
        df_raw = xls.parse(sheet, header=None, dtype=str)
        for i, row in df_raw.iterrows():
            if row.astype(str).str.contains("U in \(V\)", regex=True).any():
                df = xls.parse(sheet, header=i, dtype=str)
                df = df.dropna(how='all')
                df = df[df['U in (V)'].notna()]
                df.columns = [str(c).strip().replace('\n', ' ') for c in df.columns]
                return df
    return pd.DataFrame()
 
def detect_voltage_type(filepath):
    name = os.path.basename(filepath).lower()
    if "120" in name:
        return "120"
    elif "277" in name:
        return "277"
    return None
 
def sort_by_voltage_and_current(df):
    vout_col = next((col for col in df.columns if 'v out' in col.lower() and '(v)' in col.lower()), None)
    iout_col = next((col for col in df.columns if 'i out' in col.lower() and '(a)' in col.lower()), None)
 
    if not vout_col or not iout_col:
        return df
 
    df['_Vout_group'] = df[vout_col].astype(float).round(1)
    df['_Iout_sort'] = df[iout_col].astype(float)
 
    df_sorted = df.sort_values(by=['_Vout_group', '_Iout_sort'])
    df_sorted = df_sorted.drop(columns=['_Vout_group', '_Iout_sort']).reset_index(drop=True)
 
    return df_sorted
 
# --- MAIN ---
if __name__ == "__main__":
 
    files = select_input_files()
    if not files:
        messagebox.showwarning("Cancelled", "No files were selected.")
        raise SystemExit()
 
    output = select_output_file()
    if not output:
        messagebox.showwarning("Cancelled", "No output file was selected.")
        raise SystemExit()
 
    data_120 = pd.DataFrame()
    data_277 = pd.DataFrame()
 
    for file in files:
        vtype = detect_voltage_type(file)
        df = read_data(file)
 
        if df.empty:
            print(f"[!] Empty or invalid file: {file}")
            continue
 
        df_sorted = sort_by_voltage_and_current(df)
 
        if vtype == "120":
            data_120 = df_sorted
        elif vtype == "277":
            data_277 = df_sorted
        else:
            print(f"[!] Unknown file type (not 120V or 277V): {file}")
 
    # Align DataFrames to maximum rows
    max_rows = max(len(data_120), len(data_277))
 
    data_120 = data_120.reset_index(drop=True)
    data_277 = data_277.reset_index(drop=True)
 
    df_120_pad = pd.concat([
        data_120,
        pd.DataFrame([[None] * len(data_120.columns)] * (max_rows - len(data_120)), columns=data_120.columns)
    ], ignore_index=True)
 
    df_277_pad = pd.concat([
        data_277,
        pd.DataFrame([[None] * len(data_277.columns)] * (max_rows - len(data_277)), columns=data_277.columns)
    ], ignore_index=True)
 
    df_result = pd.concat([
        df_120_pad,
        pd.DataFrame([[None]] * max_rows, columns=['Test']),
        df_277_pad
    ], axis=1)
 
    try:
        df_result.to_excel(output, index=False)
        messagebox.showinfo("Success", f"Output file saved successfully:\n{output}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save output file:\n{e}")
 
 
