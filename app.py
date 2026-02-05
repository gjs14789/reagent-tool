import streamlit as st
import pandas as pd
import io
import xlsxwriter

# ==========================================
# 1. 欄位映射設定 (已根據您提供的標頭修正)
# ==========================================
COLUMN_MAP = {
    "id": "產品品號",       # 用於截取第1碼判斷庫存狀態 (VBA: 判斷是否為 'a')
    "name": "品名",         # 用於關鍵字分類 (VBA: extraction/pockit...)
    "order": "製令單號",    # 用於判斷年份(前4碼)與月份(5,6碼)
    "numerator": "已生產量", # 分子 (產率計算用) [1]
    "denominator": "預計產量" # 分母 (產率計算用) [1]
}

# ==========================================
# 2. 核心邏輯函式 (源自 '整理試劑資料_To_1.bas')
# ==========================================

def get_stock_status(val):
    """
    對應 VBA: 取品號第1碼 (Left(..., 1))
    """
    s = str(val).strip()
    return s if len(s) > 0 else ""

def classify_product(row):
    """
    對應 VBA: 產品類別與次分類判斷邏輯 [2]-[3]
    """
    # 取得欄位值並轉小寫，方便比對
    p_name = str(row.get(COLUMN_MAP["name"], "")).lower().strip()
    stock_status = str(row.get("庫存狀態", "")).lower()
    
    main_cat = "核酸萃取" # VBA Else 預設值 [4]
    sub_cat = ""

    # --- 主分類判斷 ---
    # 邏輯: 若庫存狀態不是 "a"，則標記為非試劑 (VBA: <> "a" Then "非試劑") [2]
    if stock_status != "a":
        return "非試劑類", ""

    # VBA: Like *extraction* Or *cartridge* [2]
    if "extraction" in p_name or "cartridge" in p_name:
        main_cat = "核酸萃取"
    # VBA: Like *pockit*, *iq*, *dntp*... [2]
    elif any(x in p_name for x in ["pockit", "iq", "dntp", "enzyme", "trehalose", "sedingin", "camap"]):
        main_cat = "配方試劑"
    # VBA: Like *taco* [4]
    elif "taco" in p_name:
        main_cat = "核酸萃取"
    # VBA: Like *ivd* [4]
    elif "ivd" in p_name:
        main_cat = "IVD"
    
    # --- 次分類判斷 [3] ---
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
    對應 VBA: Mid(..., 5, 2) 判斷月份並轉為 Q1-Q4 [5]
    """
    try:
        s = str(order_val).strip()
        # 假設單號格式前4碼是年，5-6碼是月 (例如 202310...)
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
    """執行資料轉換流程"""
    
    # 1. 建立 Index 欄位 [6]
    df.reset_index(drop=True, inplace=True)
    df.index += 1
    df.insert(0, 'index', df.index)

    # 檢查必要欄位是否存在
    required = list(COLUMN_MAP.values())
    missing = [col for col in required if col not in df.columns]
    if missing:
        return None, f"❌ 錯誤：在 Excel 第 3 列找不到這些標頭：{missing}。請確認您上傳的檔案格式。"

    # 2. 處理庫存狀態 (VBA: Left(ProductNo, 1)) [7]
    df['庫存狀態'] = df[COLUMN_MAP["id"]].apply(get_stock_status)

    # 3. 處理分類 (VBA: 透過品名關鍵字分類) [2-4]
    classification_result = df.apply(classify_product, axis=1)
    df['產品類別'] = [res for res in classification_result]
    df['次分類'] = [res[1] for res in classification_result]

    # 4. 處理季度 (VBA: 從單號取月份) [5]
    df['季度'] = df[COLUMN_MAP["order"]].apply(get_quarter)

    # 5. 計算產率 (VBA: IFERROR(分子/分母, "?")) [5]
    def calc_yield(row):
        try:
            num = float(row.get(COLUMN_MAP["numerator"], 0))   # 已生產量
            den = float(row.get(COLUMN_MAP["denominator"], 0)) # 預計產量
            return num / den if den != 0 else 0
        except:
            return 0
    
    df['產率'] = df.apply(calc_yield, axis=1)

    # 6. 統計年份 (VBA: Dictionary 統計) [8]
    # 假設單號前4碼為年份
    df['年份'] = df[COLUMN_MAP["order"]].astype(str).str[:4]
    stats = df['年份'].value_counts().sort_index().to_dict()

    return df, stats

# ==========================================
# 3. Streamlit 介面邏輯
# ==========================================

st.set_page_config(page_title="製造命令分析工具", page_icon="⚙️")

st.title("⚙️ 製造命令單頭資料前處理")
st.markdown("""
本工具將自動讀取 Excel **第 3 列** 標頭，並執行以下 VBA 邏輯：
1. **庫存狀態**：取 `產品品號` 第一碼。
2. **分類**：依據 `品名` 關鍵字 (如 extraction, pockit)。
3. **季度**：依據 `製令單號` 判定。
4. **產率**：`已生產量` / `預計產量`。
""")

uploaded_file = st.file_uploader("請上傳 Excel 檔案", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        # 讀取 Excel 結構
        xls = pd.ExcelFile(uploaded_file)
        
        # 讓使用者選擇工作表 (對應 frmSheetPicker) [9]
        sheet_name = st.selectbox("請選擇要處理的工作表：", xls.sheet_names)
        
        if st.button("開始執行 (Execute)"):
            with st.spinner('正在分析資料...'):
                # 關鍵修正：header=2 代表讀取 Excel 的第 3 列 (0, 1, 2)
                df_raw = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=2)
                
                # 執行處理
                result_df, stats = process_data(df_raw.copy())
                
                if result_df is not None:
                    st.success(f"✅ 處理完成！共 {len(result_df)} 筆資料")
                    
                    # 顯示統計 (對應 VBA MsgBox) [10]
                    st.subheader("📊 年度統計")
                    stats_df = pd.DataFrame(list(stats.items()), columns=['年份', '筆數'])
                    st.table(stats_df)
                    
                    # 預覽資料
                    st.subheader("📝 結果預覽")
                    st.dataframe(result_df.head())
                    
                    # 產生 Excel 下載
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        sheet_out = "處理結果"
                        result_df.to_excel(writer, index=False, sheet_name=sheet_out)
                        
                        # 格式化輸出 (還原 VBA ListObject 風格)
                        workbook = writer.book
                        worksheet = writer.sheets[sheet_out]
                        (max_row, max_col) = result_df.shape
                        
                        # 加入 Excel 表格樣式
                        column_settings = [{'header': col} for col in result_df.columns]
                        worksheet.add_table(0, 0, max_row, max_col - 1, {
                            'columns': column_settings,
                            'style': 'TableStyleMedium9',
                            'name': 'ResultTable'
                        })
                        
                        # 設定產率為百分比格式 [10]
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
                    st.stop() # 停止執行並顯示上方的錯誤訊息

    except Exception as e:
        st.error(f"發生錯誤：{str(e)}")
