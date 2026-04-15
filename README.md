# 每週廠商情報監控系統

成人用品產業自動化情報工具。每週自動搜尋 10 家廠商的最新動態，輸出 Excel 與 PDF 雙格式報告。

---

## 功能概覽

- 自動搜尋各廠商**新品發布、停產、價格異動、重要公告**
- 優先搜尋台灣通路（蝦皮、PChome、紅犀牛等），再查原廠官網
- 新品核實機制：標記 ✅ 確認新品 / ⚠️ 疑似非新品 / ❓ 待確認
- 輸出 Excel 報告（三工作表）＋ PDF 報告
- 主要 AI：Google Gemini 2.5-Flash；配額耗盡自動備援至 Claude 3.5 Haiku
- Claude 備援費用上限保護（預設 $0.20 USD／次）

---

## 監控廠商（10 家）

| # | 廠商 | 地區 |
|---|---|---|
| 1 | TENGA | 日本 |
| 2 | HARU | 日本 |
| 3 | UPKO | 中國 |
| 4 | Satisfyer | 德國 |
| 5 | SISTALK | 台灣 |
| 6 | We-Vibe | 加拿大 |
| 7 | 輕喃 | 台灣 |
| 8 | Puissante | 法國 |
| 9 | PxPxP | 日本 |
| 10 | LELO | 瑞典 |

---

## 環境需求

- Python 3.10 以上
- Google Gemini API Key（[取得連結](https://aistudio.google.com/app/apikey)）
- Anthropic Claude API Key（[取得連結](https://console.anthropic.com/settings/keys)，備援用）

---

## 安裝步驟

```bash
# 1. Clone 專案
git clone https://github.com/tmhahasomean/regular-surveys.git
cd regular-surveys

# 2. 安裝套件
pip install -r requirements.txt

# 3. 設定 API Key
cp .env.example .env
# 用任意編輯器開啟 .env，填入你的 API Key
```

`.env` 填寫範例：
```
GOOGLE_API_KEY=AIzaSy...你的金鑰...
ANTHROPIC_API_KEY=sk-ant-...你的金鑰...
```

---

## 使用方式

```bash
python vendor_report.py
```

執行完畢後，報告會輸出至：
```
reports/
└── 2026-04-15/
    ├── vendor_report_2026-04-15.xlsx
    ├── vendor_report_2026-04-15.pdf
    └── vendor_report_2026-04-15.html
```

---

## 報告內容

### Excel 三工作表

| 工作表 | 內容 |
|---|---|
| 每週總覽 | 各廠商活躍度（活躍／一般／無異動）、新品數、異動數 |
| 新品詳細 | 產品名稱、發布日期、核實狀態、台灣售價、來源通路 |
| 其他異動 | 停產、價格調整、公司公告、促銷活動 |

---

## 廠商維護

**新增或移除廠商：** 修改 `廠商清單.txt`（每行格式：`編號\t廠商名稱`）

**更新廠商官網或搜尋關鍵字：** 修改 `vendors.json`

若 `廠商清單.txt` 中的廠商在 `vendors.json` 找不到對應設定，系統會直接以廠商名稱進行關鍵字搜尋。

---

## 費用控制

Claude 備援費用上限可在 `vendors.json` 調整：

```json
"config": {
  "claude_cost_limit_usd": 0.20
}
```

達到上限後，已完成分析的廠商仍會正常輸出報告，未分析廠商會在報告中標記說明。

---

## 檔案結構

```
.
├── CLAUDE.md               # Claude AI 工作指令（自動載入）
├── vendor_report.py        # 主程式：搜尋、分析、生成報告
├── snapshot_fetcher.py     # 廠商官網快照擷取工具
├── vendors.json            # 廠商詳細設定（官網、搜尋關鍵字等）
├── 廠商清單.txt             # 本週監控廠商清單
├── .env.example            # API Key 設定範本
├── requirements.txt        # Python 套件清單
└── reports/                # 報告輸出目錄（不納入版本控制）
```
