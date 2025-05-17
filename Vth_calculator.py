# -*- coding: utf-8 -*-
"""
Created on Fri Mar 21 13:49:38 2025

@author: user
"""
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# 讀取 Excel 檔案中有 "V_Gate" 與 "I_Drain" 兩個欄位
excel_input = r"C:\Users\user\OneDrive\桌面\experiment\2025.04.21\bottom_a_algao_100nm_beta_50nm_top_beta_algao_20nm\Sn_0.1M_35sec\VDS_30V.xls"
df = pd.read_excel(excel_input)
VG_data = df['V_Gate']
ID_data = df['I_Drain']

# 繪製整體的 ID-VG 曲線，方便參考選點
plt.figure(figsize=(8, 6))
plt.plot(VG_data, ID_data, 'b-', label='ID vs VG')
plt.xlabel('VG')
plt.ylabel('ID')
plt.title('ID vs VG 整體曲線')
plt.grid(True)
plt.legend()
plt.show()

print("請根據上圖選取兩點，並輸入對應的 VG 與 ID 值。")

# 手動輸入第一組數值 (VG1, ID1)
VG1 = float(input("請輸入第一點的 VG 值："))
ID1 = float(input("請輸入第一點的 ID 值："))

# 手動輸入第二組數值 (VG2, ID2)
VG2 = float(input("請輸入第二點的 VG 值："))
ID2 = float(input("請輸入第二點的 ID 值："))

# 檢查兩個點的 VG 值是否相同（避免斜率無窮大）
if VG1 == VG2:
    print("兩個點的 VG 值不能相同，無法計算斜率。")
else:
    # 計算直線的斜率 m 與截距 b
    m = (ID2 - ID1) / (VG2 - VG1)
    b = ID1 - m * VG1

    print(f"計算得到直線斜率 m = {m}")
    print(f"計算得到直線截距 b = {b}")

    # 若直線非水平線 (m ≠ 0)，求直線與 x 軸交點
    if m != 0:
        x_intersect = -b / m
        intersection = (x_intersect, 0)
        print(f"計算得到與 x 軸的交點 (Vth) 為： {intersection}")
    else:
        if ID1 == 0:
            print("直線與 x 軸重合。")
            intersection = (VG1, 0)
        else:
            print("直線為水平線，不與 x 軸相交。")
            intersection = None

    # 為繪製直線，定義 x 值範圍（擴展兩點範圍 1 單位）
    x_line = np.linspace(min(VG1, VG2) - 1, max(VG1, VG2) + 1, 100)
    y_line = m * x_line + b

    # 繪製整體曲線、選定點、計算直線及交點
    plt.figure(figsize=(8, 6))
    plt.plot(VG_data, ID_data, 'b-', label='ID vs VG')
    plt.scatter([VG1, VG2], [ID1, ID2], color='red', label='選定的兩點')
    plt.plot(x_line, y_line, 'g--', label='通過選定點之直線')
    
    if intersection is not None:
        plt.scatter([intersection[0]], [intersection[1]], color='purple', label='與 x 軸交點 (Vth)')
    
    plt.xlabel('VG')
    plt.ylabel('ID')
    plt.title('直線與 x 軸交點 (Vth) 計算結果')
    plt.legend()
    plt.grid(True)
    plt.show()
    
    # 若有有效的交點，則取 Vth = 交點的 x 座標
    if intersection is not None:
        Vth = intersection[0]
    else:
        raise ValueError("沒有計算到有效的 Vth")

    

