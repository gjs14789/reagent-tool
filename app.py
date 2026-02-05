import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
from datetime import datetime

# ==========================================
# 0. 設定與 Log 功能
# ==========================================
LOG_FILE = "process_log.txt"

def write_log(filename, status, message=""):
    """寫入操作紀錄 (時間, 檔名, 狀態, 訊息)"""
    time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{time_str}] 檔案: {filename} | 狀態: {status} | 訊息: {message}\n"
    
    # 寫入檔案 (append模式)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(log_entry)
    except Exception as e:
        print(f"Log 寫入失敗: {e}")

# ==========================================
# 1. 欄位映射設定 (根據您的 Excel 實際標頭)
# ==========================================
INPUT_MAPPING = {
    "id": "產品品號",       
    "name": "品名",         
    "order": "製令單號",    
    "numerator": "已生產量", 
    "denominator": "預計產量" 
}

# 最終輸出順序
FINAL_COLUMNS_ORDER = [
    "index", "製令單別", "單別名稱", "製令單號", "季度", "急料", "開單日期", "列印", "星期", 
    "性質", "狀態碼", "類型", "物料型態", "系列項目", "項目分類", "產品品號", "品名", 
    "規格", "單位", "BOM版次", "預計產量", "已領套數", "產率", "已生產量", "報廢數量", 
    "備註", "BOM日期", "預計開工", "星期2", "預計完工", "星期3", "實際開工", "星期4", 
    "實際完工", "星期5", "確認日", "確認者", "名稱", "生產廠別", "廠別名稱", "入庫庫別", 
    "庫別名稱", "生產線別", "線別名稱", "加工廠商", "廠商名稱", "稅別碼", "稅別名稱", 
    "生管/採購人員", "人員姓名", "幣別", "課稅別", "營業稅率", "價格條件", "付款條件代號", 
    "付款條件名稱", "預計批號", "送貨地址", "匯率", "加工單位", "計劃批號", "母製令單別", 
    "母製令單號", "訂單單別", "訂單單號", "訂單序號", "客戶代號", "客戶簡稱", "客戶單號", 
    "客戶品號", "確認碼", "簽核狀態", "傳送次數", "EBO拋轉狀態", "版次", "專案代號", 
    "專案名稱", "SMES整合", "SMES拋轉紀錄碼", "ISO單號"
]

# ==========================================
# 2. 核心邏輯函式
# ==========================================

def get_stock_status(val):
    s = str(val).strip()
    return s if len(s) > 0 else ""

def classify_product(row):
    """回傳 (MainCategory, SubCategory) 的元組"""
    p_name = str(row.get(INPUT_MAPPING["name"], "")).lower().strip()
    stock_status = str(row.get("物料型態", "")).lower()
    
    main_cat = "核酸萃取"
    sub_cat = ""

    if stock_status != "a":
        return "非試劑類", ""

    if "extraction" in p_name or "cartridge" in p_name:
        main_cat = "核酸萃取"
    elif any(x in p_name for x in ["pockit", "iq", "dntp", "enzyme", "trehalose", "sedingin", "camap"]):
        main_cat = "配方試劑"
    elif "taco" in p_name:
        main_cat = "核酸萃取"
    elif "ivd" in p_name:
        main_cat = "IVD"
    
    if main_cat == "核酸萃取":
        if "cartridge" in p_name:
            sub_cat = "POCKIT Central (相關)"
        elif "extraction" in p_name:
            sub_cat = "核酸萃取"
        else:
            sub_cat = "核酸萃取" 
    elif main_cat == "配方試劑":
        if any(x in p_name for x in ["enzyme", "dntp", "iq plus", "pockit"]):
            sub_cat = "IQ Plus、POCKIT"
        elif "pockit central" in p_name or "sedingin" in p_name:
            sub_cat = "POCKIT Central"
        elif any(x in p_name for x in ["camap", "iq200", "iq 2000"]):
            sub_cat = "IQ 2000"
        elif "iq real" in p_name:
            sub_cat = "IQ real"

    return main_cat, sub_cat

def get_quarter(order_val):
    try:
        s = str(order_val).strip()
        if len(s) < 6: return ""
        month = int(s[4:6])
        if 1 <= month <= 3: return "Q1"
        if 4 <= month <= 6: return "Q2"
        if 7 <= month <= 9: return "Q3"
        if 10 <= month <= 12: return "Q4"
        return ""
    except:
        return ""

def process_data(df):
    # 1. Index
    df.reset_index(drop=True, inplace=True)
    df.index += 1
    df['index'] = df.index

    # 2. Check Columns
    required = list(INPUT_MAPPING.values())
    missing = [col for col in required if col not in df.columns]
    if missing:
        return None, f"缺少欄位: {missing}"

    # 3. Logic
    df['物料型態'] = df[INPUT_MAPPING["id"]].apply(get_stock_status)

    # *** 修正點：將元組拆解為兩個獨立欄位，避免 Excel 寫入錯誤 ***
    classification_results = df.apply(classify_product, axis=1).tolist()
    df['系列項目'] = [res for res in classification_results]
    df['項目分類'] = [res[1] for res in classification_results]

    df['季度'] = df[INPUT_MAPPING["order"]].apply(get_quarter)
    df['年份'] = df[INPUT_MAPPING["order"]].astype(str).str[:4]

    # Calc Yield
    def calc_yield(row):
        try:
            num = float(row.get(INPUT_MAPPING["numerator"], 0))
            den = float(row.get(INPUT_MAPPING["denominator"], 0))
            return num / den if den != 0 else 0
        except:
            return 0
    df['產率'] = df.apply(calc_yield, axis=1)

    stats = df['年份'].value_counts().sort_index().to_dict()

    # Reorder Columns
    final_df = pd.DataFrame()
    for col in FINAL_COLUMNS_ORDER:
        if col in df.columns:
            final_df[col] = df[col]
        else:
            final_df[col] = ""

    return final_df, stats

# ==========================================
# 3. Streamlit UI
# ==========================================

st.set_page_config(page_title="製造命令處理工具", page_icon="🏭")
st.title("🏭 製造命令單頭資料前處理")

uploaded_file = st.file_uploader("請上傳 Excel 檔案", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        wb = openpyxl.load_workbook(uploaded_file)
        sheet_names = wb.sheetnames
        
        selected_sheet = st.selectbox("請選擇工作表：", sheet_names)
        
        if st.button("開始處理"):
            with st.spinner('正在處理...'):
                try:
                    # 讀取資料 (header=2 表示第3列是標題)
                    df_raw = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=2)
                    
                    # 執行處理
                    result_df, stats = process_data(df_raw.copy())
                    
                    if result_df is not None:
                        # 命名新 Sheet
                        base_name = f"{selected_sheet}的處理結果"
                        count = 1
                        new_sheet_name = f"{base_name}({count})"
                        while new_sheet_name in wb.sheetnames:
                            count += 1
                            new_sheet_name = f"{base_name}({count})"
                        
                        # 寫入資料
                        ws_new = wb.create_sheet(new_sheet_name)
                        for r in dataframe_to_rows(result_df, index=False, header=True):
                            ws_new.append(r)
                        
                        # 設定表格
                        max_col_letter = openpyxl.utils.get_column_letter(len(result_df.columns))
                        max_row = len(result_df) + 1
                        tab = Table(displayName=f"Table_{datetime.now().strftime('%Y%m%d%H%M%S')}", 
                                    ref=f"A1:{max_col_letter}{max_row}")
                        tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
                        ws_new.add_table(tab)
                        
                        # 設定產率格式
                        if "產率" in result_df.columns:
                            yield_idx = result_df.columns.get_loc("產率") + 1
                            col_letter = openpyxl.utils.get_column_letter(yield_idx)
                            for cell in ws_new[col_letter]:
                                if cell.row > 1: cell.number_format = '0.00%'

                        # 準備下載
                        virtual_workbook = io.BytesIO()
                        wb.save(virtual_workbook)
                        virtual_workbook.seek(0)
                        
                        # 記錄 Log: 成功
                        write_log(uploaded_file.name, "Success", f"處理 {len(result_df)} 筆資料")
                        
                        st.success("✅ 處理完成！")
                        st.write("📊 統計結果：", stats)
                        st.download_button(
                            "📥 下載結果檔案",
                            data=virtual_workbook,
                            file_name=f"Processed_{uploaded_file.name}",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        # 記錄 Log: 失敗 (欄位錯誤)
                        error_msg = stats # process_data 返回 None 時，stats 是錯誤訊息
                        write_log(uploaded_file.name, "Failed", error_msg)
                        st.error(error_msg)

                except Exception as e:
                    # 記錄 Log: 失敗 (程式例外)
                    write_log(uploaded_file.name, "Error", str(e))
                    st.error(f"執行錯誤：{str(e)}")

    except Exception as e:
        st.error(f"檔案讀取錯誤：{str(e)}")

# 顯示 Log 查看器 (可選)
if st.checkbox("查看執行紀錄 (Log)"):
    try:
        with open(LOG_FILE, "r", encoding="utf-8") as f:
            st.text(f.read())
    except FileNotFoundError:
        st.info("尚無紀錄")
