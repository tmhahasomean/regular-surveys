"""
快照抓取腳本
讀取 vendors.json 中每個廠商的 snapshot_pages，
下載頁面純文字並儲存至 snapshots/{vendor_id}/{日期}.txt。
保留最近 4 週快照，自動清理更舊的檔案。
"""
import json
import os
import re
import time
from datetime import date, timedelta

import requests
from bs4 import BeautifulSoup

SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
VENDOR_DIR   = SCRIPT_DIR
VENDORS_JSON = os.path.join(SCRIPT_DIR, "vendors.json")
SNAPSHOT_DIR = os.path.join(SCRIPT_DIR, "snapshots")
TODAY        = date.today().isoformat()
KEEP_WEEKS   = 4

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
}

# HTML 標籤屬性：這些區塊通常是導覽/頁首/頁尾雜訊，略過
NOISE_TAGS = {"nav", "header", "footer", "script", "style", "noscript", "aside"}


def _extract_embedded_products(soup: "BeautifulSoup") -> str:
    """
    從 <script> 標籤中提取嵌入的產品 JSON 資料。
    支援以下常見電商平台格式：
      1. 91APP（UPKO 台灣站等）：app.value('products', JSON.parse('...'))
      2. JSON-LD：<script type="application/ld+json"> 結構化資料
      3. Shopify：window.theme.product = {...} 或 var meta = {...}
    回傳可讀的產品清單文字，無法提取則回傳空字串。
    """
    import json as _json

    lines = []

    for script in soup.find_all("script"):
        content = script.string or ""
        if not content.strip():
            continue

        # ── 格式 1：91APP 平台（upko.com.tw、其他台灣電商）──
        # 格式：app.value('products', JSON.parse('JSON字串'))
        m = re.search(r"app\.value\(['\"]products['\"],\s*JSON\.parse\(['\"](.+?)['\"]\)\)", content, re.S)
        if m:
            try:
                raw = m.group(1).encode("utf-8").decode("unicode_escape")
                products = _json.loads(raw)
                for p in products[:50]:  # 最多取 50 筆
                    name  = p.get("title") or p.get("name") or ""
                    price = p.get("price") or p.get("min_price") or ""
                    if name:
                        lines.append(f"【商品】{name}" + (f"  NT${price}" if price else ""))
                if lines:
                    return "\n".join(lines)
            except Exception:
                pass

        # ── 格式 2：JSON-LD 結構化資料 ──
        if script.get("type") == "application/ld+json":
            try:
                data = _json.loads(content)
                # 可能是單一物件或陣列
                items = data if isinstance(data, list) else [data]
                for item in items:
                    dtype = item.get("@type", "")
                    if dtype in ("Product", "ItemList"):
                        name = item.get("name") or ""
                        price = ""
                        offers = item.get("offers", {})
                        if isinstance(offers, dict):
                            price = offers.get("price", "")
                        if name:
                            lines.append(f"【商品】{name}" + (f"  {price}" if price else ""))
            except Exception:
                pass

        # ── 格式 3：Shopify / 通用 window 變數 ──
        # 格式：window.ShopifyAnalytics.meta.product = {...}
        m = re.search(r'"title"\s*:\s*"([^"]+)".*?"price"\s*:\s*"?(\d+)"?', content, re.S)
        if m and "product" in content.lower():
            try:
                # 簡單提取所有 "title": "..." 模式
                titles = re.findall(r'"title"\s*:\s*"([^"]{3,80})"', content)
                for t in titles[:20]:
                    if not any(kw in t.lower() for kw in ["script", "function", "window", "var "]):
                        lines.append(f"【商品】{t}")
                if lines:
                    return "\n".join(lines)
            except Exception:
                pass

    return "\n".join(lines)


def fetch_page_text(url: str, timeout: int = 15) -> str:
    """
    下載頁面並回傳純文字。
    若頁面有效文字少於 200 字元，先嘗試從 <script> 提取嵌入產品 JSON；
    兩者皆失敗（可能是完全 JS 渲染）則回傳空字串。
    """
    try:
        resp = requests.get(url, headers=HEADERS, timeout=timeout)
        resp.raise_for_status()
        resp.encoding = resp.apparent_encoding or "utf-8"

        soup = BeautifulSoup(resp.text, "html.parser")

        # ── 優先嘗試從 script 提取嵌入的產品 JSON（在移除 script 前執行）──
        embedded = _extract_embedded_products(soup)

        # 移除雜訊區塊
        for tag in soup.find_all(NOISE_TAGS):
            tag.decompose()

        # 嘗試找主要內容區塊
        main = (
            soup.find("main")
            or soup.find(id=re.compile(r"(content|product|main)", re.I))
            or soup.find(class_=re.compile(r"(content|product|collection|article)", re.I))
            or soup.body
        )

        if not main:
            return embedded if len(embedded) >= 200 else ""

        # 取得純文字，合併多餘空白
        lines = []
        for text in main.stripped_strings:
            text = text.strip()
            if text and len(text) > 1:
                # 跳過 AngularJS 樣板語法（{{ ... }}）
                if re.match(r"^\{\{.*\}\}$", text):
                    continue
                lines.append(text)

        html_text = "\n".join(lines)

        # 若 HTML 純文字足夠，合併嵌入 JSON 一起回傳（去除 AngularJS 噪音後）
        if len(html_text) >= 200:
            if embedded:
                return html_text + "\n\n--- 嵌入商品資料 ---\n" + embedded
            return html_text

        # HTML 純文字不足 → 退而使用嵌入 JSON
        if len(embedded) >= 200:
            return embedded

        return ""

    except Exception as e:
        print(f"    [抓取失敗] {url}: {e}")
        return ""


def fetch_vendor_snapshot(vendor_id: str, snapshot_pages: list[str]) -> str:
    """
    抓取廠商所有 snapshot_pages，合併成一個文字快照。
    """
    parts = []
    for url in snapshot_pages:
        print(f"    抓取：{url}", flush=True)
        text = fetch_page_text(url)
        if text:
            parts.append(f"=== {url} ===\n{text}")
        else:
            base = url.split("/")[0] + "//" + url.split("/")[2]
            # 備援1：嘗試首頁
            if url.rstrip("/") != base:
                print(f"    [備援] 嘗試首頁：{base}/", flush=True)
                text = fetch_page_text(base + "/")
                if text:
                    parts.append(f"=== {base}/ (首頁) ===\n{text}")
                    time.sleep(2)
                    continue
            # 備援2：嘗試 sitemap.xml
            sitemap_url = base + "/sitemap.xml"
            print(f"    [備援] 嘗試 sitemap：{sitemap_url}", flush=True)
            text = fetch_page_text(sitemap_url)
            if text:
                parts.append(f"=== {sitemap_url} (sitemap) ===\n{text}")
        time.sleep(2)  # 避免對同一網站過快請求

    return "\n\n".join(parts)


def save_snapshot(vendor_id: str, text: str) -> str:
    """儲存快照，回傳儲存路徑。"""
    vendor_dir = os.path.join(SNAPSHOT_DIR, vendor_id)
    os.makedirs(vendor_dir, exist_ok=True)
    path = os.path.join(vendor_dir, f"{TODAY}.txt")
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)
    return path


def cleanup_old_snapshots(vendor_id: str):
    """刪除超過 KEEP_WEEKS 週的舊快照。"""
    vendor_dir = os.path.join(SNAPSHOT_DIR, vendor_id)
    if not os.path.isdir(vendor_dir):
        return
    cutoff = (date.today() - timedelta(weeks=KEEP_WEEKS)).isoformat()
    for fname in os.listdir(vendor_dir):
        if fname.endswith(".txt") and fname[:10] < cutoff:
            os.remove(os.path.join(vendor_dir, fname))
            print(f"    [清理] 刪除舊快照：{fname}")


def run_all() -> dict[str, bool]:
    """
    執行所有廠商的快照抓取。
    回傳 {vendor_id: True/False}，True 代表成功取得非空快照。
    """
    with open(VENDORS_JSON, encoding="utf-8") as f:
        config = json.load(f)

    results = {}
    vendors = config.get("vendors", [])

    for vendor in vendors:
        vid   = vendor["id"]
        name  = vendor["name"]
        pages = vendor.get("snapshot_pages", [])

        if not pages:
            print(f"  [{name}] 無 snapshot_pages，跳過")
            results[vid] = False
            continue

        print(f"  → [{name}] 開始抓取（{len(pages)} 頁）", flush=True)
        text = fetch_vendor_snapshot(vid, pages)

        if text:
            path = save_snapshot(vid, text)
            cleanup_old_snapshots(vid)
            print(f"    ✅ 儲存：{path}（{len(text)} 字元）")
            results[vid] = True
        else:
            print(f"    ❌ 所有頁面均無法取得有效內容，將走 web_search 備援")
            results[vid] = False

    return results


if __name__ == "__main__":
    print(f"[快照抓取] 開始，日期：{TODAY}")
    os.makedirs(SNAPSHOT_DIR, exist_ok=True)
    results = run_all()
    success = sum(1 for v in results.values() if v)
    print(f"\n[快照抓取] 完成：{success}/{len(results)} 家成功")
