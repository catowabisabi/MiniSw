關於 SWP 到 BAS 轉換功能
=============================

## 功能說明

新增的 SWP → BAS 轉換器可以將 SolidWorks 宏文件 (.swp) 轉換為 VBA 基礎模塊文件 (.bas)。

## 使用方法

1. 點擊工具欄上的「🔄 SWP→BAS」按鈕
2. 選擇要轉換的 .swp 文件
3. 指定輸出的 .bas 文件位置
4. 系統會嘗試提取並轉換宏內容

## 轉換方式

由於 .swp 文件是微軟的 OLE 複合文檔（Structured Storage）二進制格式，完整的轉換需要：

### 方法一：使用 SolidWorks API
- 通過 SolidWorks 應用程序接口獲取宏信息
- 適用於已連接 SolidWorks 的情況

### 方法二：VBA 引擎（需要額外組件）
- 使用 Microsoft.Vbe.Interop 庫
- 可以直接操控 VBA 對象模型

### 當前實現：模板生成
- 創建標準 VBA 模板文件
- 包含基本的 SolidWorks API 連接代碼
- 用戶需要手動添加具體的宏邏輯

## 輸出文件格式

生成的 .bas 文件包含：
- 標準 VBA 模塊聲明
- SolidWorks API 連接代碼
- 主程序入口點 (main)
- 輔助函數框架
- 轉換信息註釋

## 使用建議

1. 轉換後的文件需要手動編輯以恢復完整功能
2. 檢查並調整 SolidWorks API 調用
3. 測試宏功能是否正常
4. 考慮使用版本控制管理轉換後的文件

## 技術限制

- 無法完全自動化提取二進制 .swp 內容
- 需要用戶手動完善轉換後的代碼
- 依賴於 SolidWorks API 的可用性