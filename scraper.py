import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
import time
import random
from datetime import datetime
import os

# 1. Daftar Kata Kunci Pencarian (10 Group K-Pop)
KEYWORDS = [
    "bts lightstick", "blackpink lightstick", "exo lightstick", 
    "seventeen lightstick", "ateez lightstick", "h2h lightstick", 
    "straykidz lightstick", "aespa lightstick", "riize lightstick", 
    "baby monster lightstick"
]

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept-Language": "en-US,en;q=0.9"
}

output_filename = "DATA_ebay_lightstick_RAW_ONLY.xlsx"

# 2. Cek apakah file Excel sudah ada sebelumnya (untuk penggabungan/append data harian)
if os.path.exists(output_filename):
    try:
        existing_df = pd.read_excel(output_filename)
        existing_rows = existing_df.to_dict('records')
    except Exception:
        existing_rows = []
else:
    existing_rows = []

new_rows = []
timestamp_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

print(f"=== Memulai Scraping eBay [{timestamp_str}] ===")

for keyword in KEYWORDS:
    print(f"\n[+] Scraping kata kunci: {keyword}")
    count = 0
    page = 1
    
    while count < 200 and page <= 5:
        url = f"https://www.ebay.com/sch/i.html?_nkw={keyword.replace(' ', '+')}&_sacat=0&LH_ItemCondition=3000|1500|2000|1000&_pgn={page}"
        
        try:
            res = requests.get(url, headers=headers, timeout=10)
            if res.status_code != 200:
                print(f"  [-] Status code {res.status_code}, berhenti pada halaman {page}")
                break
                
            soup = BeautifulSoup(res.text, 'html.parser')
            items = soup.select('.s-item')
            
            for item in items:
                if count >= 200:
                    break
                    
                # Judul Produk
                title_elem = item.select_one('.s-item__title')
                title_text = title_elem.text.strip() if title_elem else ""
                if not title_text or "Shop on eBay" in title_text:
                    continue
                
                # Harga Mentah
                price_elem = item.select_one('.s-item__price')
                raw_price = price_elem.text.strip() if price_elem else ""
                
                # Kondisi Mentah
                cond_elem = item.select_one('.SECONDARY_INFO')
                raw_cond = cond_elem.text.strip() if cond_elem else "Used"
                
                # Ongkos Kirim Mentah
                shipping_elem = item.select_one('.s-item__shipping')
                raw_shipping = shipping_elem.text.strip() if shipping_elem else "Free Shipping"
                
                # Informasi Penjual (ID, Feedback Mentah, Reputasi Mentah)
                seller_elem = item.select_one('.s-item__seller-info')
                seller_id = ""
                raw_fb_count = ""
                raw_reputation = ""
                
                if seller_elem:
                    seller_info_str = seller_elem.text.strip()
                    seller_parts = seller_info_str.split(' ')
                    seller_id = seller_parts[0] if seller_parts else "unknown"
                    
                    rep_match = re.search(r'([\d\.]+)%', seller_info_str)
                    if rep_match:
                        raw_reputation = f"{rep_match.group(1)}%"
                    
                    fb_match = re.search(r'\(([\d,]+)\)', seller_info_str)
                    if fb_match:
                        raw_fb_count = fb_match.group(1)
                
                # Negara Penjual Mentah
                location_elem = item.select_one('.s-item__location')
                raw_country = location_elem.text.replace("from ", "").strip() if location_elem else "South Korea"
                
                # Format 18 Kolom Sesuai Excel SPSS (Dummy dikosongkan)
                new_rows.append({
                    "수집_타임스탬프": timestamp_str,
                    "검색_키워드": keyword,
                    "BoyGroup_Dummy": "",
                    "BigAgency_Dummy": "",
                    "상품_제목": title_text,
                    "중고_판매_가격": raw_price,
                    "Price_USD": "",
                    "상품_상태": raw_cond,
                    "Cond_LikeNew_Dummy": "",
                    "배송료": raw_shipping,
                    "Shipping_Cost_USD": "",
                    "판매자_아이디": seller_id,
                    "판매자_피드백_수": raw_fb_count,
                    "Feedback_Count": "",
                    "판매자_긍정_피드백(%)": raw_reputation,
                    "Reputation_Y": "",
                    "판매자_국가": raw_country,
                    "Country_Dummy": ""
                })
                count += 1
            
            print(f"  -> Terkumpul {count}/200 baris dari halaman {page}")
            page += 1
            time.sleep(random.uniform(1.2, 2.5))
            
        except Exception as e:
            print(f"  [-] Error: {e}")
            break

# 3. Gabungkan data baru dengan data lama (jika ada) dan simpan ke file Excel
all_rows = existing_rows + new_rows
df_final = pd.DataFrame(all_rows)
df_final.to_excel(output_filename, index=False)

print(f"\n[=== Selesai! Total {len(df_final)} baris data tersimpan di '{output_filename}' ===]")