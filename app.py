import streamlit as st
import pandas as pd

st.set_page_config(page_title="多表跨Sheet拼接工具", layout="wide")
st.title("🏭 跨表组装：Sheet3(A,C) + Sheet4(B)")

uploaded_file = st.file_uploader("上传 Excel 文件 (.xlsx, .xlsm)", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        # 1. 解析 Excel 结构
        xl_file = pd.ExcelFile(uploaded_file, engine='openpyxl')
        sheet_names = xl_file.sheet_names
        
        # 检查表数量够不够
        if len(sheet_names) < 4:
            st.error(f"❌ 文件只有 {len(sheet_names)} 个表，无法读取第 4 张表！请检查文件。")
            st.stop()
            
        # 锁定第3张和第4张表 (索引分别为2和3)
        sheet3_name = sheet_names[2]
        sheet4_name = sheet_names[3]
        
        st.success(f"✅ 已锁定源数据表：\n1. 第3张表：【{sheet3_name}】\n2. 第4张表：【{sheet4_name}】")

        # --- 第一步：处理 Sheet 3 (A列和C列) ---
        st.info("正在处理 Sheet3 数据...")
        df_s3 = pd.read_excel(uploaded_file, sheet_name=sheet3_name, header=None, usecols="A,C")
        # 清洗：炸开合并单元格
        df_s3_clean = df_s3.ffill()
        # 重置索引，确保拼接时对齐
        df_s3_clean.reset_index(drop=True, inplace=True)

        # --- 第二步：处理 Sheet 4 (B列) ---
        st.info("正在处理 Sheet4 数据...")
        df_s4 = pd.read_excel(uploaded_file, sheet_name=sheet4_name, header=None, usecols="B")
        # 清洗
        df_s4_clean = df_s4.ffill()
        # 重置索引
        df_s4_clean.reset_index(drop=True, inplace=True)

        # --- 第三步：横向拼接 (Concatenate) ---
        # axis=1 表示横着拼（左右拼接）
        # 现在的顺序是：[Sheet3的第一列(原A), Sheet3的第二列(原C), Sheet4的第一列(原B)]
        final_df = pd.concat([df_s3_clean.iloc[:, 0], df_s3_clean.iloc[:, 1], df_s4_clean.iloc[:, 0]], axis=1)
        
        # 给个临时列名方便预览
        final_df.columns = ["Sheet3_A列", "Sheet3_C列", "Sheet4_B列"]

        # --- 4. 预览与下载 ---
        st.markdown("### 📊 拼接结果预览")
        st.dataframe(final_df.head(15))
        
        # 导出不带表头的 CSV
        csv_data = final_df.to_csv(index=False, header=False, encoding='utf-8-sig').encode('utf-8-sig')
        
        st.download_button(
            label="📥 下载最终合并结果 CSV",
            data=csv_data,
            file_name="跨Sheet合并结果.csv",
            mime="text/csv"
        )
        
    except Exception as e:
        st.error(f"发生未知错误: {e}")
else:
    st.info("👆 请上传文件，系统将自动抓取 Sheet3 和 Sheet4 的特定列进行合并")
