# reagent-tool
🏭 製造命令單頭資料前處理工具 (Streamlit Web App)
 
將傳統 Excel VBA (.xlsm) 巨集工具現代化，轉移為基於 Python 與 Streamlit 的網頁應用程式。本工具自動化處理製造命令資料，執行分類、產率計算與格式化報表輸出。
✨ 主要功能 (Features)
• 🚀 現代化介面：使用 Streamlit 網頁介面取代舊有的 VBA UserForm，支援拖曳上傳，無需啟用 Excel 巨集。
• 📊 自動化資料清洗：
    ◦ 物料型態判斷：依據 產品品號 首碼自動判斷 (如 a 為試劑相關)。
    ◦ 關鍵字分類：自動依據 品名 進行多層次分類 (如：核酸萃取、配方試劑、POCKIT Central 等)。
    ◦ 時間週期計算：從 製令單號 解析年份與季度 (Q1-Q4)。
• 🧮 產率計算：自動計算 已生產量 / 預計產量 並處理除數為零的異常。
• 📑 智慧輸出：
    ◦ 保留原始資料：不破壞原始檔案，處理結果將新增為獨立的工作表。
    ◦ 防呆命名：自動偵測工作表名稱，若重複則自動編號 (e.g., 處理結果(1)).
    ◦ 格式化報表：自動套用 Excel 表格樣式 (ListObject) 與百分比格式。
    ◦ 欄位重整：依據指定順序重新排列並補齊缺失欄位。
🛠️ 技術轉移對照 (Migration Log)
本專案將原有的 Excel VBA 模組 (RunFromTool.bas, frmDataProcessor.frm) 邏輯完全移植至 Python：
功能模組
原 VBA 實作
Python / Streamlit 實作
啟動入口
ShowDataProcessorForm
streamlit run app.py
檔案選擇
frmDataProcessor (FileDialog)
st.file_uploader
工作表選取
frmSheetPicker (ComboBox)
st.selectbox
核心邏輯
整理試劑資料_To (Loop processing)
pandas (Vectorized operations)
關鍵字比對
Like *keyword*
str.contains('keyword')
報表輸出
ws.Copy, ListObjects.Add
openpyxl + xlsxwriter
⚙️ 安裝與執行 (Installation & Usage)
1. 複製專案
git clone https://github.com/您的帳號/reagent-data-processor.git
cd reagent-data-processor
2. 安裝依賴套件
確保您已安裝 Python，然後執行：
pip install -r requirements.txt
3. 啟動應用程式
streamlit run app.py
啟動後，瀏覽器將自動開啟本地伺服器網址 (通常為 http://localhost:8501)。
📂 專案結構
.
├── app.py                # 主程式碼 (包含 UI 與 資料處理邏輯)
├── requirements.txt      # Python 依賴庫清單
└── README.md             # 專案說明文件
📋 輸入資料需求
程式預設讀取 Excel 的 第 3 列 (Row 3) 作為標題列，且必須包含以下關鍵欄位（程式會自動偵測）：
• 產品品號
• 品名
• 製令單號
• 已生產量
• 預計產量
🚀 部署 (Deployment)
本專案已設定為可直接部署於 Streamlit Cloud：
1. 將程式碼 Push 至 GitHub。
2. 登入 Streamlit Cloud 並連結 Repository。
3. 設定 Main file path 為 app.py 即可上線。
