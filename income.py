import pandas as pd
import requests

def get_financial_report(stock_id, year, season, report_type='is'):
    """
    report_type: 
    'is' = 綜合損益表
    'bs' = 資產負債表
    'cf' = 現金流量表
    'equity' = 權益變動表
    """
    # 對應的網址路徑 (以綜合損益表為範例)
    url_map = {
        'is': 'https://mops.twse.com.tw/mops/web/ajax_t164sb04',
        'bs': 'https://mops.twse.com.tw/mops/web/ajax_t164sb03',
        'cf': 'https://mops.twse.com.tw/mops/web/ajax_t164sb05',
        'equity': 'https://mops.twse.com.tw/mops/web/ajax_t164sb06'
    }
    
    payload = {
        'encodeURIComponent': '1',
        'step': '1',
        'firstin': '1',
        'off': '1',
        'keyword4': '',
        'code1': '',
        'TYPEK': 'sii',
        'checkbtn': '',
        'queryName': 'co_id',
        'inpuType': 'co_id',
        'TYPEK2': '',
        'co_id': stock_id,
        'year': str(year - 1911), # 轉換為民國年
        'season': str(season)
    }

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }

    res = requests.post(url_map[report_type], data=payload, headers=headers)
    
    # 使用 pandas 解析表格
    try:
        dfs = pd.read_html(res.text)
        # 通常主要資料在第10個表格左右，具體需視報表類型而定
        return dfs
    except:
        return "無法獲取資料，請檢查年份或代號。"

# 執行範例：獲取聯發科 2024 年第 3 季損益表
# --- 52 行開始替換 ---
data = get_financial_report('2454', 2023, 4, 'is')

# 1. 檢查 data 是不是清單（代表抓取成功）
if isinstance(data, list):
    print(f"✅ 成功抓取！總共發現 {len(data)} 個表格")
    
    # 2. 判斷有沒有足夠的表格數量
    if len(data) > 10:
        target_df = data[10]
        
        # 3. 確保 target_df 真的不是字串，而是 DataFrame
        if hasattr(target_df, 'head'):
            import pandas as pd
            pd.set_option('display.max_columns', None)
            pd.set_option('display.expand_frame_repr', False)
            
            print("\n--- 聯發科財務報表內容 ---")
            print(target_df.head(20))
            
            # 存成 Excel 方便查看
            target_df.to_excel("mediatek_report.xlsx")
            print("\n💾 資料已成功存入 mediatek_report.xlsx")
        else:
            print("❌ 抓到的內容格式不正確，可能是網頁結構改變了。")
    else:
        print(f"❌ 抓到的表格數量不足（僅 {len(data)} 個），請檢查年份或代號。")
else:
    # 如果 data 是字串，印出錯誤訊息
    print(f"❌ 抓取失敗：{data}")

