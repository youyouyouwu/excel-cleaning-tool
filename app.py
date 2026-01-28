import streamlit as st
import pandas as pd

st.set_page_config(page_title="纯净数据提取工具", layout="wide")
st.title("🏭 基础数据提取 (5列精准版)")

uploaded_file = st.file_uploader("上传 Excel 文件 (.xlsx, .xlsm)", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        # 1. 解析 Excel 结构
        xl_file = pd.ExcelFile(uploaded_file, engine='openpyxl')
        sheet_names = xl_file.sheet_names
        
        # 检查表数量
        if len(sheet_names) < 4:
            st.error("❌ 文件Sheet数量不足4个，无法定位目标表。")
            st.stop()
            
        # 锁定表名 (按固定位置：第3张和第4张)
        sheet3_name = sheet_names[2]
        sheet4_name = sheet_names[3]
        
        st.success(f"已锁定源数据表：1.【{sheet3_name}】  2.【{sheet4_name}】")

        # ========================================================
        # 步骤 A: 处理 Sheet3 (混合提取)
        # ========================================================
        st.info("正在提取 Sheet3 (A/C炸开，E/F保留原样)...")
        
        # 一次性读取 A,C,E,F 四列，强制文本格式
        df_s3 = pd.read_excel(
            uploaded_file, 
            sheet_name=sheet3_name, 
            header=None, 
            usecols="A,C,E,F", 
            dtype=str
        )
        # 赋予临时列名，防止混淆
        df_s3.columns = ["Raw_A", "Raw_C", "Raw_E", "Raw_F"]
        
        # --- 关键清洗逻辑 ---
        # 1. 炸开组：A列 和 C列 (向下填充)
        df_s3["Raw_A"] = df_s3["Raw_A"].ffill()
        df_s3["Raw_C"] = df_s3["Raw_C"].ffill()
        
        # 2. 原样组：E列 和 F列 (不做任何操作，保持原汁原味)
        
        # 重置索引，确保对齐
        df_s3.reset_index(drop=True, inplace=True)

        # ========================================================
        # 步骤 B: 处理 Sheet4 (单列炸开)
        # ========================================================
        st.info("正在提取 Sheet4 (B列炸开)...")
        
        df_s4 = pd.read_excel(
            uploaded_file, 
            sheet_name=sheet4_name, 
            header=None, 
            usecols="B", 
            dtype=str
        )
        # B列全部填充
        df_s4 = df_s4.ffill()
        df_s4.reset_index(drop=True, inplace=True)

        # ========================================================
        # 步骤 C: 最终组装 (拼接)
        # ========================================================
        # 顺序：Sheet3-A -> Sheet3-C -> Sheet4-B -> Sheet3-E -> Sheet3-F
        final_df = pd.concat([
            df_s3["Raw_A"], 
            df_s3["Raw_C"], 
            df_s4.iloc[:, 0], 
            df_s3["Raw_E"], 
            df_s3["Raw_F"]
        ], axis=1)
        
        # 清理因为强制文本模式产生的 "nan" 字符串
        final_df = final_df.replace("nan", "")

        # ========================================================
        # 步骤 D: 预览与导出
        # ========================================================
        st.markdown("---")
        st.subheader("📋 最终数据预览")
        
        # 设置显示给用户的列名 (你可以根据业务改这里)
        final_df.columns = ["产品/A列", "C列信息", "B
