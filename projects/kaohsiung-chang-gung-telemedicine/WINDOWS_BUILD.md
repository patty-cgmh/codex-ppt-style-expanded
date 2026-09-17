# Windows 本機產生 PPTX

此專案不提交 `.pptx` 或 `.zip` 二進位檔。下載或 clone GitHub repository 後，可在 Windows 本機從文字來源重建第一版簡報。

## 系統需求

- Windows 10 或 Windows 11。
- Python 3.10 以上；建議由 [python.org](https://www.python.org/downloads/windows/) 安裝。
- 首次執行需要網路，以便由 PyPI 安裝 `python-pptx`；之後可重用專案內的 `.venv`。
- PowerPoint 並非建置必要條件，但建議用於開啟成品及進行最終視覺檢查。

安裝 Python 時請勾選 **Add python.exe to PATH**。若系統提供 Windows Python Launcher（`py`），BAT 會優先使用 `py -3`。

## 一鍵執行

1. 下載並解壓縮完整 GitHub repository，或用 Git clone。
2. 進入：

   ```text
   projects\kaohsiung-chang-gung-telemedicine
   ```

3. 雙擊：

   ```text
   Windows_build_PPT.bat
   ```

4. BAT 會自動：
   - 先執行 `cd /d "%~dp0"` 切換到 BAT 所在的專案資料夾，不依賴目前命令列目錄。
   - 建立或重用 `projects\kaohsiung-chang-gung-telemedicine\.venv`。
   - 讀取同一專案資料夾內的 `requirements.txt` 安裝套件。
   - 執行 `build_deck.py`。
   - 執行 `validate_deck.py`。
   - 產生下列檔案：

     ```text
     projects\kaohsiung-chang-gung-telemedicine\output\kaohsiung-chang-gung-telemedicine-v1.pptx
     ```

## 命令列執行

若不使用 BAT，可在 repository 根目錄的 Command Prompt 或 PowerShell 執行：

```bat
py -3 -m venv projects\kaohsiung-chang-gung-telemedicine\.venv
projects\kaohsiung-chang-gung-telemedicine\.venv\Scripts\python.exe -m pip install -r projects\kaohsiung-chang-gung-telemedicine\requirements.txt
projects\kaohsiung-chang-gung-telemedicine\.venv\Scripts\python.exe projects\kaohsiung-chang-gung-telemedicine\build_deck.py
projects\kaohsiung-chang-gung-telemedicine\.venv\Scripts\python.exe projects\kaohsiung-chang-gung-telemedicine\validate_deck.py
```

## 路徑與雲端相依性檢查

- `build_deck.py` 只從 Python 標準庫 `pathlib.Path(__file__)` 取得自身位置。
- 輸出目錄永遠是 `build_deck.py` 同層的 `output`，不存在時會自動建立。
- `validate_deck.py` 同樣以自身位置尋找 `output` 與 `illustration-prompts.md`。
- BAT 的 `.venv`、requirements、建置程式、驗證程式及輸出檔路徑全部以 BAT 所在專案資料夾為基準。
- 程式沒有 `/workspace`、`/home/oai`、Codex 暫存目錄、雲端 API 或環境變數相依性。
- 簡報內容、配色、文字及版面完全由 repository 內的 Python 原始碼建立。
- 友善線性醫療插圖仍保留為 IL-01 至 IL-06 原生 placeholder；建置過程不需要圖片生成服務。

## 字型說明

簡報指定 `Noto Sans TC`，並在設計規格中保留 `Microsoft JhengHei` 作 Office 相容字型。`python-pptx` 建置時不要求電腦已安裝指定字型；但 PowerPoint 開啟檔案時會依本機字型進行顯示或替代。若需與設計規格最一致，請先安裝 Noto Sans TC，或在 PowerPoint 中統一替換成 Microsoft JhengHei。

## 常見問題

### 顯示找不到 Python

重新安裝 Python 並勾選 **Add python.exe to PATH**，或確認下列任一命令可在 Command Prompt 執行：

```bat
py -3 --version
python --version
```

### 套件安裝失敗

確認電腦可以連線至 PyPI，並重新執行 BAT。若公司網路使用代理伺服器，需依組織規範設定 pip proxy。

### 需要重新產生

再次執行 BAT 即可。既有同名 PPTX 會由相同來源重新建立；Clinical Calm 的設計內容不會因本機建置流程而改變。

## Clinical Calm V2 視覺重製版

V1 保留為資訊架構基準。若要產生完成 Visual Redesign 的 V2，請在同一專案資料夾雙擊：

```text
Windows_build_PPT_v2.bat
```

V2 使用 `build_deck_v2.py` 與 `validate_deck_v2.py`，並輸出：

```text
output\kaohsiung-chang-gung-telemedicine-v2.pptx
```

V2 不覆蓋 V1；兩個檔名與建置入口相互獨立。
