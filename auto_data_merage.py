# -*- coding: utf-8 -*-
"""
Created on Thu May  8 16:45:32 2025

@author: user
"""

# -*- coding: utf-8 -*-
"""
Created on Thu Dec 26 20:36:53 2024
Last edit : 2025-05-08  統整所有功能，將指標寫入同一個 .xlsx

@author: user
"""
import os
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import filedialog

# ---------- 1. 資料合併 ----------
def merge_columns_with_spacing_and_abs(folder, target_files, cols):
    final = pd.DataFrame()

    for fname in target_files:
        fpath = os.path.join(folder, fname)
        if not os.path.exists(fpath):
            print(f"[warn] 找不到 {fname}")
            continue
        try:
            df = pd.read_excel(fpath, engine="xlrd")
            pick = df[cols].copy()
            pick.columns = [f"{c}_{os.path.splitext(fname)[0]}" for c in pick.columns]

            if final.empty:
                final = pick
            else:
                spacer = pd.DataFrame(index=final.index, columns=["", ""])
                final = pd.concat([final, spacer, pick], axis=1)
        except Exception as e:
            print(f"[err] 讀取 {fname} 失敗：{e}")

    # 補絕對值欄位
    for col in ("I_Gate_VDS_10V", "I_Gate_VDS_30V"):
        if col in final.columns:
            final[f"abs_{col}"] = final[col].abs()

    return final


# ---------- 2. 計算 gm / S.S / On-Off ----------
def calc_metrics(vds30_path, Vg_step=0.1):
    df = pd.read_excel(vds30_path)

    Imax, Imin = df["I_Drain"].max(), df["I_Drain"].min()
    onoff = Imax / Imin

    df["gm"] = (df["I_Drain"] - df["I_Drain"].shift(-1)) / Vg_step
    gm_max = df["gm"].max()

    df["S.S"] = Vg_step / (np.log10(df["I_Drain"]) - np.log10(df["I_Drain"].shift(-1)))
    ss_min = df["S.S"][:-1].abs().min()

    return pd.DataFrame(
        {"SS (V/dec)": [ss_min], "gm_max (A/V)": [gm_max], "On/Off": [onoff]}
    )


# ---------- 3. 主程式 ----------
if __name__ == "__main__":
    root = tk.Tk(); root.withdraw()
    folder = filedialog.askdirectory(title="選擇含 VDS 檔的資料夾")
    if not folder:
        print("未選擇資料夾，程式結束。"); quit()

    out_xlsx = os.path.join(folder, "combine_data_with_spacing_abs.xlsx")

    files = ["VDS_10V.xls", "VDS_30V.xls", "dual_VDS_10V.xls", "dual_VDS_30V.xls"]
    cols  = ["V_Gate", "I_Gate", "V_Drain", "I_Drain"]

    merged_df = merge_columns_with_spacing_and_abs(folder, files, cols)

    # 只要 merged_df 不是空的就寫出
    if not merged_df.empty:
        with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
            merged_df.to_excel(writer, sheet_name="MergedData", index=False)

            # 針對 VDS_30V.xls 計算指標，再寫到第二張 Sheet
            vds30_path = os.path.join(folder, "VDS_30V.xls")
            if os.path.exists(vds30_path):
                metrics_df = calc_metrics(vds30_path)
                metrics_df.to_excel(writer, sheet_name="Metrics", index=False)
                print("Metrics 已寫入 Sheet『Metrics』")
            else:
                print("[warn] 找不到 VDS_30V.xls，無法計算 Metrics")

        print(f"✅ 完成！結果存於：{out_xlsx}")
    else:
        print("⚠️ 沒有可合併的資料，未產生輸出檔。")
