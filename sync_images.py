"""
GitHub Actions 同步脚本：从 Ozon API 获取商品图片并更新 Supabase
支持命令行参数指定 SKU，或自动检测缺少图片的 SKU
"""
import requests
import json
import os
import sys
import time

SUPABASE_URL = os.environ["SUPABASE_URL"]
SUPABASE_SERVICE_KEY = os.environ["SUPABASE_SERVICE_KEY"]
OZON_CLIENT_ID = os.environ["OZON_CLIENT_ID"]
OZON_API_KEY = os.environ["OZON_API_KEY"]

HEADERS = {
    "apikey": SUPABASE_SERVICE_KEY,
    "Authorization": f"Bearer {SUPABASE_SERVICE_KEY}",
    "Content-Type": "application/json"
}

def sp_request(method, path, body=None):
    url = f"{SUPABASE_URL}/rest/v1/{path}"
    resp = requests.request(method, url, headers=HEADERS, json=body, timeout=30)
    if resp.status_code in (200, 201):
        return resp.json() if resp.text else None
    else:
        print(f"  Supabase error {resp.status_code}: {resp.text[:200]}")
        return None

def get_missing_skus(specified_skus=None):
    """获取缺少图片的 SKU 及其对应的记录 ID"""
    if specified_skus:
        # 手动指定的 SKU
        skus = [s.strip() for s in specified_skus.split(",") if s.strip()]
        # 查询这些 SKU 对应的记录
        all_data = []
        for sku in skus:
            data = sp_request('GET', f"buyer_mapping?select=id,sku&sku=eq.{sku}&image_url=is.null")
            if data:
                all_data.extend(data)
        return all_data

    # 自动检测：获取所有有 SKU 但没图片的记录
    all_data = []
    from_offset = 0
    page = 1000
    while True:
        data = sp_request('GET', f"buyer_mapping?select=id,sku&sku=neq.&image_url=is.null&order=id.asc&limit={page}&offset={from_offset}")
        if not data:
            break
        all_data.extend(data)
        if len(data) < page:
            break
        from_offset += page
    return all_data

def fetch_images(skus):
    """调用 Ozon API 获取商品图片"""
    numeric_skus = [int(s) for s in skus if s.isdigit()]
    if not numeric_skus:
        return []

    headers = {
        "Client-Id": OZON_CLIENT_ID,
        "Api-Key": OZON_API_KEY,
        "Content-Type": "application/json"
    }

    for attempt in range(3):
        try:
            resp = requests.post(
                "https://api-seller.ozon.ru/v3/product/info/list",
                headers=headers,
                json={"sku": numeric_skus},
                timeout=60
            )
            if resp.status_code == 200:
                return resp.json().get("items", [])
            elif resp.status_code == 429:
                wait = 2 ** attempt
                print(f"  频率限制，等待 {wait}s...")
                time.sleep(wait)
            else:
                print(f"  Ozon API error {resp.status_code}: {resp.text[:200]}")
                return []
        except requests.exceptions.Timeout:
            print(f"  超时 (attempt {attempt+1}/3)")
            time.sleep(2)
        except Exception as e:
            print(f"  Ozon API error: {e}")
            time.sleep(1)
    return []

def main():
    specified = sys.argv[1] if len(sys.argv) > 1 else ""
    print(f"指定 SKU: {specified if specified else '(自动检测)'}")

    records = get_missing_skus(specified or None)
    print(f"待处理记录: {len(records)}")

    if not records:
        print("所有图片已同步！")
        return

    # 按 SKU 去重
    sku_to_ids = {}
    for r in records:
        sku = r['sku']
        if sku not in sku_to_ids:
            sku_to_ids[sku] = []
        sku_to_ids[sku].append(r['id'])

    unique_skus = list(sku_to_ids.keys())
    print(f"唯一 SKU: {len(unique_skus)}")

    batch_size = 10
    total = len(unique_skus)
    processed = 0
    found = 0

    for i in range(0, total, batch_size):
        batch = unique_skus[i:i + batch_size]
        items = fetch_images(batch)

        for item in items:
            sources = item.get("sources", [])
            for src in sources:
                sku_key = str(src.get("sku", ""))
                if sku_key in sku_to_ids:
                    primary = item.get("primary_image", "")
                    images = item.get("images", [])

                    if isinstance(primary, list) and primary:
                        primary = primary[0]
                    if not primary and isinstance(images, list) and images:
                        primary = images[0]

                    if primary:
                        for rid in sku_to_ids[sku_key]:
                            sp_request('PATCH', f"buyer_mapping?id=eq.{rid}", {"image_url": primary})
                        found += len(sku_to_ids[sku_key])
                    break

        processed += len(batch)
        pct = processed * 100 // total
        print(f"  {processed}/{total} ({pct}%) - 累计 {found} 张")

        if i + batch_size < total:
            time.sleep(0.15)

    print(f"完成！共更新 {found} 条记录的图片")

if __name__ == "__main__":
    main()
