"""
从 Ozon API 批量获取商品图片，更新 buyer_mapping 表
"""
import requests
import json
import time
import urllib.request
import sys

SUPABASE_URL = "https://jlzsonjjfgojmwgghxbl.supabase.co"
SERVICE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpsenNvbmpqZmdvam13Z2doeGJsIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc3NjkxNjY4OSwiZXhwIjoyMDkyNDkyNjg5fQ.qTAqHLBeWyVUfV8uxdP2-55EFI7kyh4aJ2RGJHrhQTo"
OZON_CLIENT_ID = "3306389"
OZON_API_KEY = "bd757233-6d64-4e7c-9d36-3db020b88533"

def sp_request(method, path, body=None):
    url = f"{SUPABASE_URL}/rest/v1/{path}"
    headers = {
        "apikey": SERVICE_KEY,
        "Authorization": f"Bearer {SERVICE_KEY}",
        "Content-Type": "application/json"
    }
    data = json.dumps(body).encode() if body else None
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            raw = resp.read()
            return json.loads(raw) if raw else None
    except Exception as e:
        print(f"  Supabase error: {e}")
        return None

def main():
    print("获取需要拉取图片的 SKU 列表...")
    sys.stdout.flush()

    # 直接获取有 SKU 但没图片的记录
    all_data = []
    from_offset = 0
    page = 1000
    while True:
        print(f"  加载分页 offset={from_offset}...")
        sys.stdout.flush()
        data = sp_request('GET', f"buyer_mapping?select=id,sku&sku=neq.&image_url=is.null&order=id.asc&limit={page}&offset={from_offset}")
        if not data or len(data) == 0:
            break
        all_data.extend(data)
        if len(data) < page:
            break
        from_offset += page

    print(f"找到 {len(all_data)} 条待处理记录")
    sys.stdout.flush()

    if len(all_data) == 0:
        print("所有图片已拉取完毕！")
        return

    # 按 SKU 去重
    sku_to_ids = {}
    for r in all_data:
        sku = r['sku']
        if sku not in sku_to_ids:
            sku_to_ids[sku] = []
        sku_to_ids[sku].append(r['id'])

    unique_skus = list(sku_to_ids.keys())
    print(f"唯一 SKU 数: {len(unique_skus)}")

    # Ozon API 配置
    ozon_headers = {
        "Client-Id": OZON_CLIENT_ID,
        "Api-Key": OZON_API_KEY,
        "Content-Type": "application/json"
    }

    batch_size = 10
    total = len(unique_skus)
    processed = 0
    found = 0
    start_time = time.time()

    for i in range(0, total, batch_size):
        batch = unique_skus[i:i + batch_size]
        numeric_skus = [int(s) for s in batch if s.isdigit()]
        batch_start = time.time()

        # 调用 Ozon API
        items = []
        for attempt in range(3):
            try:
                resp = requests.post(
                    "https://api-seller.ozon.ru/v3/product/info/list",
                    headers=ozon_headers,
                    json={"sku": numeric_skus},
                    timeout=60
                )
                if resp.status_code == 200:
                    items = resp.json().get("items", [])
                    break
                elif resp.status_code == 429:
                    wait = 2 ** attempt
                    print(f"  频率限制，等待 {wait}s...")
                    time.sleep(wait)
                else:
                    print(f"  Ozon API status {resp.status_code}: {resp.text[:200]}")
                    break
            except requests.exceptions.Timeout:
                print(f"  超时 (attempt {attempt+1}/3)")
                time.sleep(2)
            except Exception as e:
                print(f"  Ozon API error: {e}")
                time.sleep(1)

        batch_found = 0
        for item in items:
            sources = item.get("sources", [])
            for src in sources:
                sku_key = str(src.get("sku", ""))
                if sku_key in sku_to_ids:
                    primary = item.get("primary_image", "")
                    images = item.get("images", [])

                    if isinstance(primary, list) and len(primary) > 0:
                        primary = primary[0]
                    if not primary and isinstance(images, list) and len(images) > 0:
                        primary = images[0]

                    if primary:
                        for rid in sku_to_ids[sku_key]:
                            sp_request('PATCH', f"buyer_mapping?id=eq.{rid}", {"image_url": primary})
                        found += len(sku_to_ids[sku_key])
                        batch_found += len(sku_to_ids[sku_key])
                    break

        processed += len(batch)
        elapsed = time.time() - start_time
        batch_time = time.time() - batch_start
        eta = (elapsed / processed * (total - processed)) if processed > 0 else 0
        pct = processed * 100 // total
        print(f"  {processed}/{total} ({pct}%) | 本批:{batch_found}张 | 累计:{found}张 | {batch_time:.1f}s | ETA:{eta:.0f}s")
        sys.stdout.flush()

        if i + batch_size < total:
            time.sleep(0.15)

    print(f"\n完成！共为 {found} 条记录设置图片")

if __name__ == "__main__":
    main()
