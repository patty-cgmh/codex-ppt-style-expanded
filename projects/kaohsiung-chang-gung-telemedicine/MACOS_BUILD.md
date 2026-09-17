# macOS 本機產生 Clinical Calm V2 PPTX

本說明適用於從 GitHub **Download ZIP** 取得 repository 的 Mac 使用者。macOS 建置流程不會修改 Windows BAT，也不會改變簡報內容或 Clinical Calm 視覺設計。

## 產出檔案

執行成功後會產生：

```text
projects/kaohsiung-chang-gung-telemedicine/output/kaohsiung-chang-gung-telemedicine-v2.pptx
```

macOS 腳本使用 V2 專用的 `build_deck_v2.py` 與 `validate_deck_v2.py`，以確保產出檔名和 Windows 的 V2 建置流程一致。V1 檔案不會被覆蓋。

## 1. 從 GitHub 下載

1. 在 GitHub repository 頁面選擇 **Code → Download ZIP**。
2. 在 Finder 打開「下載項目」。
3. 雙擊下載的 ZIP 解壓縮。
4. 保留完整 repository 資料夾結構，不要只單獨複製 `.sh`。

## 2. 安裝或確認 Python 3

打開 Terminal，執行：

```bash
python3 --version
```

建議使用 Python 3.10 以上。如果顯示找不到命令，可選擇以下任一方式安裝：

- 從 [python.org macOS downloads](https://www.python.org/downloads/macos/) 安裝。
- 已安裝 Homebrew 時執行 `brew install python`。

## 3. 在 Terminal 進入專案資料夾

最簡單的方法是在 Terminal 輸入 `cd `（`cd` 後保留一個空格），再從 Finder 把下列資料夾拖入 Terminal 視窗：

```text
projects/kaohsiung-chang-gung-telemedicine
```

按 Return 執行。也可以手動輸入完整路徑，例如：

```bash
cd "$HOME/Downloads/codex-ppt-style-expanded/projects/kaohsiung-chang-gung-telemedicine"
```

## 4. 首次授予執行權限

GitHub Download ZIP 有時不會保留 shell script 的可執行權限。請執行：

```bash
chmod +x macOS_build_PPT.sh
```

## 5. 建置 V2

執行：

```bash
./macOS_build_PPT.sh
```

腳本會自動完成：

1. 切換至 `macOS_build_PPT.sh` 所在的專案資料夾。
2. 在專案資料夾建立或重用 `.venv`。
3. 使用 `.venv/bin/python` 安裝同層 `requirements.txt`。
4. 執行 `build_deck_v2.py`。
5. 執行 `validate_deck_v2.py`。
6. 確認 V2 PPTX 已存在，並在 Terminal 顯示完整路徑。

首次執行需要網路連線以便從 PyPI 安裝 `python-pptx`。之後會重用同一個 `.venv`。

## 6. 開啟輸出檔案

成功訊息會顯示 PPTX 的完整路徑。也可以在專案資料夾執行：

```bash
open output/kaohsiung-chang-gung-telemedicine-v2.pptx
```

建議使用 Microsoft PowerPoint for Mac 進行最終字型、換行及投影對比檢查。

## 常見問題

### `permission denied`

重新執行：

```bash
chmod +x macOS_build_PPT.sh
./macOS_build_PPT.sh
```

### `python3 was not found`

先安裝 Python 3，關閉並重新開啟 Terminal，再執行 `python3 --version` 確認。

### 建立 `.venv` 失敗

確認專案資料夾不是唯讀，且目前使用者對該資料夾有寫入權限。必要時刪除未完成的 `.venv` 後重新執行腳本：

```bash
rm -rf .venv
./macOS_build_PPT.sh
```

