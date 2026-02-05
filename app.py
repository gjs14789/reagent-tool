import streamlit as st
import pandas as pd
import io
import xlsxwriter

# ==========================================
# 1. 欄位映射設定 (CONFIG)
# ==========================================
# 請依照您 Excel 實際的 "第3列" 標題名稱修改右邊的字串
COLUMN_MAP = {
    "id": "品號料號",       # 對應 VBA 用來取第1碼判斷庫存狀態的欄位
    "name": "品名",         # 對應 VBA 用來判斷 extraction/pockit 的欄位
    "order": "單據編號",    # 對應 VBA 用來判斷年份(前4)與月份(5,6)的欄位
    "plan_qty": "預產量",   # 對應 VBA 計算產率的分子
    "exp_qty": "預計入庫"   # 對應 VBA 計算產率的分母
}

# ==========================================
# 2. 核心邏輯函式 (邏輯源自 '整理試劑資料_To_1.bas')
# ==========================================

def get_stock_status(val):
    """
    對應 VBA: If Len(...) > 0 Then ... Left(..., 1)
    取得品號的第一個字元作為庫存狀態。
    """
    s = str(val).strip()
    return s if len(s) > 0 else ""

def classify_product(row):
    """
    對應 VBA: 產品類別與次分類判斷邏輯
    VBA 原始碼使用 If...Like... 進行關鍵字匹配
    """
    # 取得品名並轉小寫，方便比對
    p_name = str(row.get(COLUMN_MAP["name"], "")).lower().strip()
    # 取得剛算出來的庫存狀態
    stock_status = str(row.get("庫存狀態", "")).lower()
    
    main_cat = "核酸萃取" # VBA Else 預設值 (推測亂碼為核酸萃取)
    sub_cat = ""

    # 邏輯 A: 若庫存狀態不是 "a"，則標記為非試劑 (推測亂碼含意)
    if stock_status != "a":
        return "非試劑類", ""

    # 邏輯 B: 主分類判斷
    if "extraction" in p_name or "cartridge" in p_name:
        main_cat = "核酸萃取"
    elif any(x in p_name for x in ["pockit", "iq", "dntp", "enzyme", "trehalose", "sedingin", "camap"]):
        main_cat = "配方試劑"
    elif "taco" in p_name:
        main_cat = "核酸萃取"
    elif "ivd" in p_name:
        main_cat = "IVD"
    
    # 邏輯 C: 次分類判斷
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
    """
    對應 VBA: Mid(..., 5, 2) 判斷月份並轉為 Q1-Q4
    """
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
    """執行主要的資料轉換流程"""
    
    # 1. 建立 Index 欄位 (對應 VBA: tbl.ListColumns(1).Name = "index")
    df.reset_index(drop=True, inplace=True)
    df.index += 1
    df.insert(0, 'index', df.index)

    # 檢查必要欄位是否存在
    required = list(COLUMN_MAP.values())
    missing = [col for col in required if col not in df.columns]
    if missing:
        return None, f"錯誤：找不到欄位 {missing}。請檢查 Excel 標題列（第3列）是否正確，或修改程式碼中的 COLUMN_MAP 設定。"

    # 2. 處理庫存狀態 (VBA: ƫA)
    df['庫存狀態'] = df[COLUMN_MAP["id"]].apply(get_stock_status)

    # 3. 處理分類 (VBA: tC & ؤ)
    # 使用 apply 同時計算主分類與次分類
    classification_result = df.apply(classify_product, axis=1)
    df['產品類別'] = [res for res in classification_result]
    df['次分類'] = [res[13] for res in classification_result]

    # 4. 處理季度 (VBA: u)
    df['季度'] = df[COLUMN_MAP["order"]].apply(get_quarter)

    # 5. 計算產率 (VBA: v Formula)
    # Python 直接計算數值，若分母為 0 則填 0
    def calc_yield(row):
        try:
            p = float(row[COLUMN_MAP["plan_qty"]])
            e = float(row[COLUMN_MAP["exp_qty"]])
            return p / e if e != 0 else 0
        except:
            return 0
    
    df['產率'] = df.apply(calc_yield, axis=1)

    # 6. 統計年份 (VBA: Dictionary 統計)
    # 從單據編號前4碼取得年份
    df['年份'] = df[COLUMN_MAP["order"]].astype(str).str[:4]
    stats = df['年份'].value_counts().sort_index().to_dict()

    return df, stats

# ==========================================
# 3. Streamlit 介面邏輯 (UI)
# ==========================================

st.set_page_config(page_title="試劑資料處理工具", page_icon="🧪")

st.title("🧪 製造命令資料處理工具")
st.markdown("""
本工具將自動執行以下動作：
1. 讀取 Excel **第 3 列** 作為標題。
2. 依據 **品名** 關鍵字自動分類 (核酸萃取/配方試劑等)。
3. 計算 **產率** 與 **季度**。
4. 產生包含統計資訊的 Excel 報表。
""")

# 對應 frmDataProcessor 的檔案選擇
uploaded_file = st.file_uploader("請上傳 Excel 檔案 (.xlsx, .xlsm)", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        # 讀取 Excel 檔案結構
        xls = pd.ExcelFile(uploaded_file)
        
        # 對應 frmSheetPicker 的工作表選擇
        sheet_name = st.selectbox("請選擇要處理的工作表：", xls.sheet_names)
        
        # 執行按鈕
        if st.button("開始處理 (Run Processing)"):
            with st.spinner('正在分析資料...'):
                # 讀取資料，header=2 表示 Excel 的第 3 列是標題
                df_raw = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=2)
                
                # 執行處理
                result_df, stats = process_data(df_raw.copy())
                
                if result_df is not None:
                    # 顯示成功訊息與統計 (對應 VBA MsgBox)
                    st.success("✅ 資料處理完成！")
                    
                    st.subheader("📊 年度統計報告")
                    stats_df = pd.DataFrame(list(stats.items()), columns=['年份', '筆數'])
                    st.table(stats_df)
                    
                    st.subheader("📝 結果預覽")
                    st.dataframe(result_df.head())
                    
                    # 產生 Excel 下載 (保留 VBA 的 ListObject 表格風格)
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        sheet_out_name = "處理結果"
                        result_df.to_excel(writer, index=False, sheet_name=sheet_out_name)
                        
                        # 取得 xlsxwriter 物件進行格式化
                        workbook = writer.book
                        worksheet = writer.sheets[sheet_out_name]
                        (max_row, max_col) = result_df.shape
                        
                        # 加入 Excel 表格 (ListObject)
                        column_settings = [{'header': col} for col in result_df.columns]
                        worksheet.add_table(0, 0, max_row, max_col - 1, {
                            'columns': column_settings,
                            'style': 'TableStyleMedium9', # 類似 VBA 的藍白樣式
                            'name': 'ResultTable'
                        })
                        
                        # 設定產率欄位為百分比格式 (0.00%)
                        percent_fmt = workbook.add_format({'num_format': '0.00%'})
                        if '產率' in result_df.columns:
                            idx = result_df.columns.get_loc('產率')
                            worksheet.set_column(idx, idx, None, percent_fmt)

                    buffer.seek(0)
                    
                    st.download_button(
                        label="📥 下載處理後的 Excel",
                        data=buffer,
                        file_name=f"Processed_{uploaded_file.name}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("處理失敗，請檢查欄位對照設定。")

    except Exception as e:
        st.error(f"讀取檔案時發生錯誤：{e}")
