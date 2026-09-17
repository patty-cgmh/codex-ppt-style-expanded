# A 方案正文樣張製作單

## 樣張目的

先用第 6 頁「服務網絡」驗證 **臨床清晰・暖橘協作** 是否能同時呈現醫院正式感、跨院合作關係與友善醫療視覺。此樣張不使用示意院所、不填入虛構節點數或服務量。

## 代表頁內容

- **結論式標題**：高雄長庚串聯合作院所，逐步形成遠距醫療服務網絡
- **副標**：服務範圍、合作節點與連線關係將依院方核實資料呈現
- **地圖區**：高雄區域輪廓與地理細節 `〔待補資料〕`
- **主院區節點**：高雄長庚紀念醫院
- **合作院所節點**：`〔待補資料〕`
- **右側摘要卡**：
  - 合作院所數 `〔待補資料〕`
  - 涵蓋區域 `〔待補資料〕`
  - 服務科別 `〔待補資料〕`
- **頁尾洞察**：網絡成果須以實際院所、服務關係與統計期間共同解讀

## 版面規格

- 16:9，霧白背景。
- 標題區 `[0.06, 0.06, 0.88, 0.12]`。
- 地圖主區 `[0.06, 0.23, 0.62, 0.62]`，使用 PowerPoint 原生節點、連線與標籤。
- 摘要卡區 `[0.73, 0.25, 0.21, 0.53]`，三張白底圓角 KPI 卡。
- 橘色只用於合作節點與待補資料焦點；主院區使用深藍、服務連線使用青綠。
- 背景插圖只能位於右下邊緣，不侵入地圖、摘要卡或標題安全區。

## 無字背景提示詞

```text
Use case: productivity-visual
Asset type: 16:9 editable PowerPoint sample background
Primary request: Create a wordless background for a formal hospital telemedicine network outcomes report.
Scene/backdrop: clean mist-white clinical presentation canvas with a very pale blue-gray geographic/network atmosphere; no literal map boundaries.
Subject: a small, soft, friendly line-and-flat illustration at the extreme lower-right edge suggesting remote care collaboration between a clinician and a community care site through a secure digital connection.
Style/medium: modern clinical editorial illustration, restrained, trustworthy, rounded geometry, subtle depth, not childish.
Composition/framing: keep the top 6%-18% fully clear for the title; keep the left 6%-68% and central 23%-85% visually quiet for an editable map; keep the right 73%-94% mostly clear for three editable KPI cards; illustration must remain at the lower-right edge and may not enter these safe zones.
Color palette: navy #163A5F, white #FFFFFF, mist white #F6F9FC, provisional orange #F28C28, medical teal #2A8C8B, pale blue-gray #DCE8F1; orange under 10%.
Lighting/mood: bright, calm, welcoming, clinically credible.
Constraints: background and illustration only; all map nodes, lines, labels, metrics and conclusions will be added as native PowerPoint objects.
Avoid: text, letters, numbers, logos, watermarks, hospital emblems, map labels, fake charts, fake dashboards, UI screens, medical records, test values, dark monitoring-room aesthetics, neon glow, toy-like 3D characters.
```

## 生產狀態

- 風格：**approved** — 使用者已明確確認 Clinical Calm／臨床清晰，不再重新詮釋或更換 Style。
- 背景：**not required for v1** — 第一版依使用者指示，以 PowerPoint 原生 placeholder 保留友善線性醫療插圖位置；未取消插圖需求。
- PPTX：**built** — 已直接擴展為 15 頁第一版完整簡報，服務網絡頁保留原生節點、連線、標籤與 KPI 卡。
- 插圖：**pending replacement** — placeholder 對應提示詞收錄於 `illustration-prompts.md`，日後只替換插圖，不更動版型與原生資訊物件。
