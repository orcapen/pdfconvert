# 文件合併轉 PDF

Windows 桌面工具，可將多個 `.pptx` 與 `.docx` 拖曳至視窗、調整順序，再轉換並合併成單一 PDF。

## 功能

- 拖放或檔案選擇器加入 PPTX / DOCX
- 直接拖曳清單項目，或用「上移／下移」調整順序
- 自動依清單順序轉檔與合併
- 優先使用 LibreOffice；Windows 若未安裝 LibreOffice，則嘗試使用 Microsoft Office
- 原始文件不會被修改，過程中的暫存 PDF 會自動清除

## 安裝與執行

需先安裝 [uv](https://docs.astral.sh/uv/) 與 Python 3.11 以上版本。

```powershell
uv sync
uv run pdfconvert
```

轉檔引擎需符合下列其中一項：

1. 安裝 LibreOffice（建議，程式會自動尋找一般安裝位置）；或
2. Windows 已安裝 Microsoft Word 與 PowerPoint。

## 測試

```powershell
uv run pytest
```

## 建置 Windows EXE

先同步開發依賴，再執行打包腳本：

```powershell
uv sync --group dev
.\scripts\build.ps1
```

完成後的單一執行檔位於 `dist\pdfconvert.exe`。使用者不需另外安裝 Python，
但轉檔時仍需安裝 LibreOffice，或已安裝 Microsoft Word 與 PowerPoint。

## 使用方式

1. 將 PPTX 或 DOCX 拖入視窗。
2. 拖曳清單項目，或使用右側按鈕調整順序。
3. 選擇輸出 PDF 的位置。
4. 按下「轉檔並合併」。

若同名檔案位於不同資料夾也能正常處理；程式會為每個來源建立獨立暫存目錄。
