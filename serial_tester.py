import serial
import numpy as np
import time
import msvcrt

# 設定虛擬串行端口COM6和鮑率（例如9600）
ser = serial.Serial('COM1', 9600)

# 生成正弦波數據
frequency1 = 1  # 第一個正弦波的頻率
frequency2 = 5  # 第二個正弦波的頻率
frequency3 = 10  # 第二個正弦波的頻率
sampling_rate = 100  # 每秒的取樣數
t = np.linspace(0, 1, sampling_rate)  # 時間軸

try:
    print("傳輸開始，按下任意鍵停止...")
    while True:
        # 檢查是否有按鍵被按下
        if msvcrt.kbhit():
            break

        # 計算當前時刻的正弦波值
        sine_wave1 = np.sin(2 * np.pi * frequency1 * t)
        sine_wave2 = np.sin(2 * np.pi * frequency2 * t)
        sine_wave3 = np.sin(2 * np.pi * frequency3 * t)

        # 傳輸數據
        for i in range(sampling_rate):
            data = f"{sine_wave1[i]},{sine_wave2[i]},{sine_wave3[i]}\n"
            ser.write(data.encode())
            time.sleep(0.1)  # 控制傳輸速度，每次發送間隔10ms

except KeyboardInterrupt:
    print("傳輸中斷")

finally:
    # 關閉串行端口
    ser.close()
    print("程式已關閉")