import streamlit as st
import pandas as pd

st.set_page_config(page_title="数据清洗最终版", layout="wide")
st.title("🏭 混合提取模式：炸开填充 + 原样保留")

uploaded_file = st.file_uploader("上传 Excel 文件 (.xlsx, .xlsm)", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        # 1. 解析结构
        xl_file = pd.ExcelFile(uploaded_file, engine='openpyxl')
        sheet_names = xl_file.sheet_names
        
        if len(sheet_names) < 4:
            st.error("❌ 文件Sheet数量不足4个")
            st.stop()
            
        # 锁定表名 (按位置)
        sheet3_name = sheet_names[2]
        sheet4_name = sheet_names[3]
        
        st.success(f"已锁定源表：【{sheet3_name}】 和 【{sheet4_name}】")

        # ========================================================
        # 核心逻辑 A: 处理 Sheet3 (混合模式)
        # ========================================================
        # 一次性读取 A, C, E, F 四列
        # usecols="A,C,E,F" -> 读进来后顺序是 A, C, E, F (索引 0,1,2,3)
        st.info("正在处理 Sheet3 (A/C列炸开，E/F列保持原样)...")
        
        df_s3 = pd.read_excel(
            uploaded_file, 
            sheet_name=sheet3_name, 
            header=None, 
            usecols="A,C,E,F", 
            dtype=str
        )
        
        # 给列起个内部代号，方便操作
        df_s3.columns = ["Col_A", "Col_C", "Col_E", "Col_F"]
        
        # --- 局部炸开逻辑 ---
        # 只对 A列 和 C列 进行向下填充 (ffill)
        df_s3["Col_A"] = df_s3["Col_A"].ffill()
        df_s3["Col_C"] = df_s3["Col_C"].ffill()
        # E列 和 F列 咱们不动它，保持原样
        
        df_s3.reset_index(drop=True, inplace=True)

        # ========================================================
        # 核心逻辑 B: 处理 Sheet4 (炸开模式)
        # ========================================================
        st.info("正在处理 Sheet4 (B列炸开)...")
        
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
        # 核心逻辑 C: 最终组装
        # ========================================================
        # 现在的顺序要求：
        # 1. Sheet3-A (炸开)
        # 2. Sheet3-C (炸开)
        # 3. Sheet4-B (炸开)
        # 4. Sheet3-E (原样)
        # 5. Sheet3-F (原样)
        
        final_df = pd.concat([
            df_s3["Col_A"], 
            df_s3["Col_C"], 
            df_s4.iloc[:, 0], # Sheet4只有一列
            df_s3["Col_E"], 
            df_s3["Col_F"]
        ], axis=1)
        
        # 清理 'nan' 文本 -> 变回空值
        final_df = final_df.replace("nan", "")

        # 预览
        st.subheader("数据预览 (前15行)")
        st.dataframe(final_df.head(15))
        
        # 导出
        csv_data = final_df.to_csv(index=False, header=False, encoding='utf-8-sig').encode('utf-8-sig')
        
        st.download_button(
            label="📥 下载 CSV (包含 A,C,B,E,F 五列)",
            data=csv_data,
            file_name="提取结果_5列完整版.csv",
            mime="text/csv"
        )
        
    except Exception as e:
        st.error(f"发生错误: {e}")
else:
    st.info("👆 请上传文件")
