import math

from reportlab.platypus import Flowable, SimpleDocTemplate
from reportlab.lib.pagesizes import A4

# 這裡放入修正後的類別
class PrecisionGridChart(Flowable):
    def __init__(self, data_series, col_widths, height=200):
        Flowable.__init__(self)
        processed = []
        for v in data_series:
            try:
                val = float(v)
                if math.isnan(val): processed.append(None)
                else: processed.append(val)
            except (ValueError, TypeError):
                processed.append(None)
        
        self.data = processed
        self.col_widths = col_widths
        self.width = sum(col_widths[2:])
        self.height = height
        self.y_min, self.y_max = 0.0, 1.0  # 確保全域可用

    def get_xy(self, i, val):
        # 修正後的對齊邏輯：i+2 是為了跳過前兩欄標籤
        x = sum(self.col_widths[:i+2]) + self.col_widths[i+2] / 2
        denom = (self.y_max - self.y_min) if self.y_max != self.y_min else 1
        y = 10 + (val - self.y_min) / denom * (self.height - 20)
        return x, y

    def draw(self):
        # 這裡僅測試 get_xy 是否能被呼叫
        try:
            print("--> 正在測試 draw() 中的 get_xy 呼叫...")
            test_x, test_y = self.get_xy(0, 0.5)
            print(f"成功！測試座標為: ({test_x}, {test_y})")
        except Exception as e:
            print(f"失敗！錯誤訊息: {e}")

# --- 測試主程式 ---
if __name__ == "__main__":
    print("=== 開始測試 PrecisionGridChart ===")
    
    # 模擬資料：5 個問題的正確率
    test_data = [0.8, 0.5, None, 1.0, 0.2]
    # 模擬欄寬：前兩欄是標籤（例如 50, 50），後面是問題欄位
    test_widths = [50, 50, 100, 100, 100, 100, 100]
    
    try:
        # 1. 測試初始化
        chart = PrecisionGridChart(test_data, test_widths)
        print("1. 類別實例化：成功")
        
        # 2. 測試直接呼叫 get_xy (這是你之前噴錯的地方)
        print("2. 測試 get_xy 屬性...")
        if hasattr(chart, 'get_xy'):
            x, y = chart.get_xy(0, 0.8)
            print(f"   屬性檢查：成功 (座標: {x}, {y})")
        else:
            print("   屬性檢查：失敗 (依舊找不到 get_xy)")

        # 3. 模擬 ReportLab 繪製過程
        print("3. 執行模擬繪製...")
        chart.draw()
        
    except Exception as e:
        print(f"測試過程中發生崩潰: {e}")
    
    print("=== 測試結束 ===")