# GitHub Copilot CLI 安裝 Skill 處理 Excel 檔案教學

## 1. 環境準備

### 安裝 GitHub Copilot CLI

GitHub Copilot CLI 是 GitHub 提供的一個命令行工具，讓您可以在終端機中使用 Copilot 的功能。首先，您需要安裝 Node.js（版本 18 或更高）。

#### 步驟：
1. 確認 Node.js 已安裝：
   ```bash
   node --version
   npm --version
   ```

2. 安裝 GitHub Copilot CLI：
   ```bash
   npm install -g @githubnext/copilot-cli
   ```

3. 登入 GitHub 並啟用 Copilot：
   ```bash
   github-copilot-cli auth login
   ```
   按照提示登入您的 GitHub 帳號，並確保您的帳號已啟用 GitHub Copilot。

4. 驗證安裝：
   ```bash
   github-copilot-cli --version
   ```

## 2. Copilot Skill 概念說明

### 什麼是 Copilot Skill？

Copilot Skill 是 GitHub Copilot 的擴展功能，讓您能夠執行特定的任務或操作。Skill 可以是內建的或由社群開發的，它們允許 Copilot 處理更複雜的任務，例如處理檔案、執行計算或與外部服務互動。

### Skill 如何擴展 Copilot 的能力

- **模組化設計**：每個 Skill 專注於特定領域，如 Excel 處理、資料分析等。
- **命令行整合**：通過 Copilot CLI，您可以在終端機中直接使用 Skill。
- **自動化**：Skill 可以自動化重複性任務，提高工作效率。

### 與 Excel 處理相關的 Skill

常見的 Excel 處理 Skill 包括：
- **excel-reader**：讀取 Excel 檔案內容
- **excel-filter**：篩選和整理資料
- **excel-converter**：轉換格式（如 Excel 轉 CSV、JSON）
- **excel-analyzer**：分析 Excel 資料

## 3. 使用 Copilot CLI 安裝 Excel 處理相關 Skill

### 查看可用的 Skill 列表

首先，查看所有可用的 Skill：

```bash
github-copilot-cli skill list
```

### 安裝 Excel 處理相關 Skill

安裝必要的 Skill：

```bash
github-copilot-cli skill install excel-reader
github-copilot-cli skill install excel-filter
github-copilot-cli skill install excel-converter
```

### 驗證 Skill 是否安裝成功

檢查已安裝的 Skill：

```bash
github-copilot-cli skill list --installed
```

您應該看到剛安裝的 Excel 相關 Skill 出現在列表中。

## 4. 使用 Skill 處理 Excel

### 讀取 Excel 檔案內容

使用 `excel-reader` Skill 讀取 Excel 檔案：

```bash
github-copilot-cli skill run excel-reader --file sample.xlsx --sheet Sheet1
```

這將輸出 Excel 檔案中指定工作表的所有內容。

### 篩選與整理資料

使用 `excel-filter` Skill 篩選資料：

```bash
github-copilot-cli skill run excel-filter --file sample.xlsx --column "銷售額" --condition ">1000"
```

這將篩選出銷售額大於 1000 的行。

### 轉換格式

使用 `excel-converter` Skill 轉換格式：

```bash
github-copilot-cli skill run excel-converter --file sample.xlsx --output-format csv --output-file sample.csv
```

這將 Excel 檔案轉換為 CSV 格式。

### 自動化 Excel 資料處理流程

您可以結合多個 Skill 來建立自動化流程。例如，建立一個腳本：

```bash
#!/bin/bash
# 讀取 Excel，篩選資料，轉換為 JSON
github-copilot-cli skill run excel-reader --file input.xlsx --sheet Data | \
github-copilot-cli skill run excel-filter --column "狀態" --condition "active" | \
github-copilot-cli skill run excel-converter --output-format json --output-file output.json
```

## 5. 實作練習範例

### 練習目標

在本練習中，您將使用 Copilot CLI 和 Skill 處理一個範例 Excel 檔案，完成以下任務：
1. 讀取銷售資料
2. 篩選高銷售額記錄
3. 將結果轉換為 CSV 格式

### 步驟

1. **準備範例資料**：
   下載或建立一個名為 `sales_data.xlsx` 的 Excel 檔案，包含以下欄位：
   - 產品名稱
   - 銷售額
   - 銷售日期

2. **讀取資料**：
   ```bash
   github-copilot-cli skill run excel-reader --file sales_data.xlsx --sheet Sheet1
   ```

3. **篩選高銷售額**：
   ```bash
   github-copilot-cli skill run excel-filter --file sales_data.xlsx --column "銷售額" --condition ">=5000"
   ```

4. **轉換格式**：
   ```bash
   github-copilot-cli skill run excel-converter --file sales_data.xlsx --output-format csv --output-file high_sales.csv
   ```

5. **驗證結果**：
   檢查 `high_sales.csv` 檔案，確認只包含銷售額 >= 5000 的記錄。

### 進階練習

嘗試建立一個自動化腳本，定期處理新的銷售資料檔案，並生成報告。

---

恭喜！您已完成 GitHub Copilot CLI Skill 處理 Excel 檔案的教學。請將所學應用於您的日常工作中，提升資料處理效率。