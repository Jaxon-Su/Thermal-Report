# Thermal Report Tool

## 使用方式

執行：

```text
C:\Qt_Project\Thermal Report\build\ThermalTool.exe
```

操作：

1. 選擇輸入 `.xlsx` 或 `.csv`。
2. 選擇輸出 `.xlsx` 路徑。
3. 按 `執行`。

預設輸出檔名：原檔名加 `_autoReport.xlsx`。

## 輸出

- `Average`：B:BI 最後 30 筆平均值。
- `Max`：B:BI 最後 30 筆最大值。
- `FAN`：符合 FAN lock 條件時輸出最後一筆 Data。

三張 sheet 都使用 `Report Template` 樣式：

- A1:A4：`Test Item:`、`Test Start Time:`、`Test End Time:`、`Condition:`
- A5:A64：1-60，對應來源 B:BI
- B 欄開始：每一欄是一個測試條件結果

## 主要規則

- 掃描每張資料 sheet；若檔案內有名為 `Report Template` 的工作表，則略過。
- 同一個 `Test Item Name:` 區塊內若有 `User Abort Test!`，整段不輸出。
- `Pin:`：取上方最後 30 筆 Average / Max；若同列 A 欄為 `Test FAIL!` 則跳過。
- `Release Short State`：取上方最後 30 筆 Average / Max。
- `OTP Occur`：只輸出到 `Max`，取上方最後 30 筆 Max。
- `OCP Not Working`：取上方最後 30 筆 Average / Max；Condition 會加上最近的 `% OCP Trip Loading Value:` 百分比。若 `Condition:` B 欄空白，改用 `Test Item Name:`。
- `Thermal Dynamic Test`：沒有 marker，改抓同一 Test Item 區塊最後 30 筆 Data。
- `Vout Abnormal`：Condition 包含 `FAN` 時，輸出到 `FAN` sheet，取上方最後一筆 Data。
