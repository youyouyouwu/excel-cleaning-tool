import streamlit as st
import pandas as pd

st.set_page_config(page_title="数据清洗最终版", layout="wide")
st.title("🏭 6列精准提取：S3(A,C,E,F) + S4(B,I)")
st.markdown("### ✅ 配置更新：Sheet4-I列 已设为【原样保留】")

uploaded_file = st.file_uploader("上传 Excel 文件 (.xlsx, .xlsm)", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        # 1. 解析 Excel 结构
        xl_file = pd.ExcelFile(uploaded_file, engine='openpyxl')
        sheet_names = xl_file.sheet_names
        
        if len(sheet_names) < 4:
            st.error("❌ 文件Sheet数量不足4个")
            st.stop()
            
        sheet3_name = sheet_names[2]
        sheet4_name = sheet_names[3]
        
        st.success(f"已锁定：1.【{sheet3_name}】  2.【{sheet4_name}】")

        # ========================================================
        # 步骤 A: 处理 Sheet3 (4列)
        # ========================================================
        st.info("正在提取 Sheet3 (A/C炸开，E/F原样)...")
        
        df_s3 = pd.read_excel(
            uploaded_file, 
            sheet_name=sheet3_name, 
            header=None, 
            usecols="A,C,E,F", 
            dtype=str
        )
        df_s3.columns = ["S3_A", "S3_C", "S3_E", "S3_F"]
        
        # --- Sheet3 清洗逻辑 ---
        # A列、C列 -> 炸开
        df_s3["S3_A"] = df_s3["S3_A"].ffill()
        df_s3["S3_C"] = df_s3["S3_C"].ffill()
        # E列、F列 -> 保持原样 (不动)
        
        df_s3.reset_index(drop=True, inplace=True)

        # ========================================================
        # 步骤 B: 处理 Sheet4 (2列：B和I)
        # ========================================================
        st.info("正在提取 Sheet4 (B列炸开，I列原样)...")
        
        df_s4 = pd.read_excel(
            uploaded_file, 
            sheet_name=sheet4_name, 
            header=None, 
            usecols="B,I", 
            dtype=str
        )
        
        # 容错处理：防止I列没数据导致列数不够
        if df_s4.shape[1] == 1:
            df_s4["S4_I"] = ""
            df_s4.columns = ["S4_B", "S4_I"]
        else:
            df_s4.columns = ["S4_B", "S4_I"]
        
        # --- Sheet4 关键修改 ---
        # 1. B列 -> 继续炸开 (它是分类信息)
        df_s4["S4_B"] = df_s4["S4_B"].ffill()
        
        # 2. 🛑 I列 -> 原样保留！(注释掉了之前的 ffill)
        # df_s4["S4_I"] = df_s4["S4_I"].ffill()  <-- 已禁用
        
        df_s4.reset_index(drop=True, inplace=True)

        # ========================================================
        # 步骤 C: 最终 6 列组装
        # ========================================================
        # 顺序：S3-A, S3-C, S4-B, S4-I, S3-E, S3-F
        
        final_df = pd.concat([
            df_s3["S3_A"],       # 第1列
            df_s3["S3_C"],       # 第2列
            df_s4["S4_B"],       # 第3列
            df_s4["S4_I"],       # 第4列 (原样)
            df_s3["S3_E"],       # 第5列 (原样)
            df_s3["S3_F"]        # 第6列 (原样)
        ], axis=1)
        
        # 清理文本模式带来的 nan 字符
        final_df = final_df.replace("nan", "")

        # ========================================================
        # 步骤 D: 预览与导出
        # ========================================================
        st.subheader("📋 6列数据预览 (I列空值已保留)")
        
        final_df.columns = [
            "A列(炸开)", 
            "B列(炸开)", 
            "C列(炸开)", 
            "D列(S4-I原样)", 
            "E列(S3-E原样)", 
            "F列(S3-F原样)"
        ]
        
        st.dataframe(final_df.head(15))
        
        csv_data = final_df.to_csv(index=False, header=False, encoding='utf-8-sig').encode('utf-8-sig')
        
        st.download_button(
            label="📥 下载 CSV (最终版)",
            data=csv_data,
            file_name="6列数据_I列原样.csv",
            mime="text/csv"
        )
        
    except Exception as e:
        st.error(f"发生错误: {e}")
else:
    st.info("👆 请上传文件")
