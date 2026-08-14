import streamlit as st
import pandas as pd
import io 

st.set_page_config(page_title="盤點資料自動彙整", layout="wide")

st.title("📦 庫存盤點自動彙整系統")
st.write("請將各單位填寫完畢的 Excel 盤點檔上傳，系統會自動幫您合併並計算總表。")
st.info("💡 追蹤秘訣：系統會自動將「檔案名稱-分頁名稱」當作來源，建議同仁將檔名取為【調劑台名稱_姓名】。")

st.warning("⚠️ 注意：上傳的 Excel 檔案裡面，第一列的表頭必須精準包含這三個名稱：【料位號】、【藥品名】、【數量】")

uploaded_files = st.file_uploader("📂 批次上傳所有 Excel 盤點檔", type=["xlsx", "xls"], accept_multiple_files=True)

if uploaded_files:
    all_data = []
    error_files = []

    for file in uploaded_files:
        try:
            # 使用記憶體緩衝區，防止底層 C 語言引擎閃退
            file_buffer = io.BytesIO(file.getvalue())
            xls_dict = pd.read_excel(file_buffer, sheet_name=None, engine='openpyxl')
            
            file_has_valid_sheet = False 
            
            for sheet_name, df in xls_dict.items():
                if all(col in df.columns for col in ["料位號", "藥品名", "數量"]):
                    # 加上 .copy() 消除 Pandas 的黃色警告
                    df = df[["料位號", "藥品名", "數量"]].copy()
                    
                    df["藥品名"] = df["藥品名"].astype(str).str.replace(r'[\r\n]+', '', regex=True).str.replace('_x000D_', '').str.strip()
                    df["料位號"] = df["料位號"].astype(str).str.replace(r'[\r\n]+', '', regex=True).str.replace('_x000D_', '').str.strip()
                    
                    df["來源檔名 (分頁)"] = f"{file.name} - {sheet_name}"
                    
                    all_data.append(df)
                    file_has_valid_sheet = True
            
            if not file_has_valid_sheet:
                error_files.append(file.name)
                
        except Exception as e:
            print(f"檔案 {file.name} 讀取失敗: {e}")
            error_files.append(file.name)

    if error_files:
        st.error(f"❌ 下列檔案讀取失敗或裡面完全沒有必備欄位 (料位號/藥品名/數量)：\n {', '.join(error_files)}")

    if all_data:
        master_df = pd.concat(all_data, ignore_index=True)
        master_df['數量'] = pd.to_numeric(master_df['數量'], errors='coerce').fillna(0)
        
        source_df = master_df.sort_values(by=["料位號"])
        
        summary_df = master_df.groupby(["料位號", "藥品名"], as_index=False)["數量"].sum()
        summary_df = summary_df.sort_values(by="料位號") 

        st.success(f"✅ 成功合併了 {len(uploaded_files)} 個檔案中的所有分頁！結果如下：")

        tab1, tab2 = st.tabs(["📊 1. 盤點總表 (已加總)", "🔍 2. 追蹤來源明細"])

        with tab1:
            st.subheader("盤點總表")
            # 替換為最新的 width="stretch" 語法
            st.dataframe(summary_df, width="stretch", hide_index=True)
            
            csv_summary = summary_df.to_csv(index=False).encode('utf-8-sig')
            st.download_button("📥 下載盤點總表 (CSV)", data=csv_summary, file_name="1_盤點總表.csv", mime="text/csv", type="primary")

        with tab2:
            st.subheader("追蹤來源總表 (所有原始明細)")
            # 替換為最新的 width="stretch" 語法
            st.dataframe(source_df, width="stretch", hide_index=True)
            
            csv_source = source_df.to_csv(index=False).encode('utf-8-sig')
            st.download_button("📥 下載追蹤來源總表 (CSV)", data=csv_source, file_name="2_追蹤來源總表.csv", mime="text/csv")