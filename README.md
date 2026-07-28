# PPT to PDF Converter

這是一個用 Python + comtypes 自動將 PowerPoint 檔案批次轉換成 PDF 的工具。

## 功能特色
- 支援 `.ppt` 和 `.pptx` 格式
- 可一次轉換整個資料夾的檔案
- 使用者可自行輸入來源與輸出資料夾路徑
- 輸出 PDF 檔案名稱與原始 PPT 檔案相同

 安裝方式
1. 安裝 [Python 3.10+](https://www.python.org/downloads/)
2. 安裝必要套件：
   ```bash
   pip install comtypes
## 注意事項 
需要 Windows + 安裝 Microsoft PowerPoint

如果沒有 PowerPoint，可以改用 LibreOffice（需自行修改程式）

請確保輸入路徑正確，否則程式會找不到檔案
