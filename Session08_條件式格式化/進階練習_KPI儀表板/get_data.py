import requests
import os
import re
import time
from urllib.parse import unquote

def download_multi_year_gov_data(start_year, end_year, save_dir="miaoli_population_data"):
    """
    批次下載指定年份區間的政府公開資料。
    """
    # 建立儲存資料夾
    if not os.path.exists(save_dir):
        os.makedirs(save_dir)

    # 設定 Header 模擬瀏覽器
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36"
    }

    # 開始批次下載迴圈
    for year in range(start_year, end_year + 1):
        url = f"https://mlhr.miaoli.gov.tw/xlsgetfile_{year}.php"
        print(f"🔄 準備下載 {year} 年度資料...")
        print(f"   連線至：{url}")

        try:
            response = requests.get(url, headers=headers, stream=True, timeout=15)
            response.raise_for_status() # 若遭遇 404/500 等錯誤，會直接觸發 Exception

            # 嘗試解析原始檔名
            content_disposition = response.headers.get('Content-Disposition', '')
            filename = f"{year}年苗栗縣人口統計.xls" # 預設檔名

            if content_disposition:
                filenames = re.findall('filename="?([^";]+)"?', content_disposition)
                if filenames:
                    filename = unquote(filenames[0])

            filepath = os.path.join(save_dir, filename)

            # 寫入檔案
            with open(filepath, 'wb') as file:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        file.write(chunk)

            print(f"   ✅ 成功儲存：{filename}")

        except requests.exceptions.HTTPError as errh:
            print(f"   ⚠️ 找不到 {year} 年度資料或伺服器錯誤 (可能尚未提供): {errh}")
        except Exception as e:
            print(f"   ❌ 下載 {year} 年度時發生未知錯誤: {e}")

        # 禮貌性延遲：避免對伺服器造成負擔或被視為惡意攻擊
        if year != end_year:
            print("   ⏳ 休息 3 秒後繼續...\n")
            time.sleep(3)

    print("\n🎉 所有批次下載任務已完成！")

# 執行 104 年至 115 年的批次下載
download_multi_year_gov_data(104, 115)