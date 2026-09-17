# 第一版交付說明

## 本機生成檔案

- `.pptx` 與 `.zip` 不納入 repository；此目錄只保留文字說明。
- Windows 請執行上一層 `Windows_build_PPT.bat`。
- Windows 建置所需套件列於上一層專案資料夾的 `requirements.txt`。
- 成功後會在本目錄產生 `kaohsiung-chang-gung-telemedicine-v1.pptx`：15 頁、16:9、繁體中文第一版完整簡報。
- 執行 `Windows_build_PPT_v2.bat` 會另行產生 `kaohsiung-chang-gung-telemedicine-v2.pptx`；不覆蓋 V1。
- macOS 可執行上一層 `macOS_build_PPT.sh` 產生同一份 V2；完整步驟見 `MACOS_BUILD.md`。
- 重建程式為上一層 `build_deck.py`；插圖替換提示詞為上一層 `illustration-prompts.md`。

## 可編輯層級

第一版的標題、正文、`〔待補資料〕`、KPI 卡、服務網絡節點與連線、時間軸、流程節點、趨勢圖空白座標、熱力格及插圖 placeholder 均為 PowerPoint 原生物件。檔案內沒有以點陣圖片冒充圖表、流程或插圖。

## 插圖 placeholder

IL-01 至 IL-06 是正式插圖的保留位置，不代表取消插圖。取得圖片生成能力後，依 `illustration-prompts.md` 逐一生成，再以插圖替換虛線框；不得更動已核准的 Clinical Calm 配色、網格、原生資料物件或文字層級。

## 尚未完成的 QA

目前環境沒有 LibreOffice、PowerPoint 或其他 PPTX renderer，因而只能完成 OOXML／python-pptx 結構檢查，尚未做實際 PowerPoint 渲染後的逐頁視覺檢查。開啟檔案後仍應確認目標電腦的繁中字型替代、文字換行、陰影／圓角顯示及投影對比。
