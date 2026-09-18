# V3 五頁精修 Visual QA

結論：本次 portable render 與資料／檔案一致性檢查 PASS。等待 Microsoft PowerPoint QA；尚未宣稱已通過原生 PowerPoint 開啟與放映驗證。

## 修改範圍

僅修改 05、08、11、12、14。其餘 10 頁（01、02、03、04、06、07、09、10、13、15）的 slide XML 與重新輸出的 PNG 均與核准版逐位元一致。所有共用圖片、版型、主題、備註、關聯及內嵌 chart workbook 保持原始位元內容。

- 05：降低 2024／2025 年份及年度量的視覺權重，保留全部原文與數字；突出合作起步、門診表自動化、Dashboard、澎湖監獄及後續方向。沒有新增服務升降原因或成效判斷。
- 08：保留八個完整步驟與 HIS【執行中】註記；採兩列連續編號、細連線與開放留白，仍為 native editable shapes。
- 11：3,749 為 summary KPI，5 月 572 加大並以橘色突出；維持正式可編輯 native chart。沒有新增原因、成長或改善解釋。
- 12：左區屏東／高雄 16 家，右區澎湖／花蓮／臺東 7 家。院所及數值使用 18 pt。右區保留九科別；左區五科別逐格呈現，另外四科別全為缺資料，以明確分區註記表示全部 64 格缺值。沒有把缺值當作 0，沒有刪除院所或已提供的數值。
- 14：保留所有問題、行動、狀態文字，改成開放式 Problem → Action rows；已完成／執行中／規劃中使用小型文字標示。

## 資料核對

- 月資料：372、286、473、494、572、551、541、460；總和 3,749，官方口徑、不含腦中風 468。
- 4,217 對外報告口徑保持原頁完全不變。
- Heatmap：23 家；57 個正式資料格全部與來源一致，其中 17 個真實零值。
- 150 個缺值：86 個顯式「—」格，加上左區註記涵蓋的 64 個缺值；無補 0。
- Scale 保持 0／1–50／51–150／151–300／301–600；缺資料與 0 以不同符號及底色表示。
- 05、08、14 原始文字逐項保留核對通過。
- 19 個受保護來源及 prototype 檔案 hash 不變。

## 視覺及結構檢查

15 頁均由最終 PPTX 重新匯入並 render 為 1920 × 1080 PNG。五張修改頁逐張檢視，未見文字溢出、裁切或可見文字互相重疊；月圖八個資料標籤清楚，Heatmap 院所及 cell 數字可辨識。其餘十頁預覽位元一致。

Package integrity、頁數／頁幅／字型、native quantitative chart、chart cache／embedded workbook 結構檢查皆通過，hard findings 為 0。

Geometry checker 有 6 個 connector-over-text bounding-box 警示（08 一處、14 五處）；已逐張檢視，屬連線外接框與文字框範圍相交，實際可見筆畫未壓到文字，並非聲稱自動工具零警示。字型沿用核准版 Noto Sans CJK TC。

## 交付與待驗證

- kaohsiung-chang-gung-telemedicine-v3-refined.pptx
- v3-full-deck-contact-sheet.png
- previews/slide-01.png 至 slide-15.png
- data-and-lock-verification.json
- validation.json

請以 Microsoft PowerPoint 開啟、放映及檢查原生編輯／字型呈現。此次未 commit、push、建立 PR 或 merge；已核准原版保留。
