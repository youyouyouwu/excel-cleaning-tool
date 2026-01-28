import streamlit as st
import pandas as pd

st.set_page_config(page_title="多表跨Sheet拼接工具", layout="wide")
st.title("🏭 跨表组装：Sheet3(A,C) + Sheet4(B)")
st.markdown("### ✅ 已启用：强制文本模式 (防科学计数法)")

uploaded_file = st.file_uploader("上传 Excel 文件 (.xlsx, .xlsm)", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        # 1. 解析 Excel 结构
        xl_file = pd.ExcelFile(uploaded_file, engine='openpyxl')
        sheet_names = xl_file.sheet_names
        
        if len(sheet_names) < 4:
            st.error(f"❌ 文件只有 {len(sheet_names)} 个表，无法读取第 4 张表！")
            st.stop()
            
        sheet3_name = sheet_names[2] # 第3张
        sheet4_name = sheet_names[3] # 第4张
        
        st.success(f"已锁定：1.【{sheet3_name}】  2.【{sheet4_name}】")

        # ========================================================
        # 核心修改点：加入 dtype=str
        # 这告诉 Pandas：别自作聪明，把所有内容都当成“文本”读进来
        # ========================================================

        # --- 第一步：处理 Sheet 3 (A列和C列) ---
        st.info("正在读取 Sheet3 (强制文本模式)...")
        df_s3 = pd.read_excel(
            uploaded_file, 
            sheet_name=sheet3_name, 
            header=None, 
            usecols="A,C", 
            dtype=str  # <--- 关键！禁止转为数字
        )
        
        # 炸开合并单元格 (ffill 对文本也有效)
        df_s3_clean = df_s3.ffill()
        df_s3_clean.reset_index(drop=True, inplace=True)

        # --- 第二步：处理 Sheet 4 (B列) ---
        st.info("正在读取 Sheet4 (强制文本模式)...")
        df_s4 = pd.read_excel(
            uploaded_file, 
            sheet_name=sheet4_name, 
            header=None, 
            usecols="B", 
            dtype=str  # <--- 关键！禁止转为数字
        )
        
        df_s4_clean = df_s4.ffill()
        df_s4_clean.reset_index(drop=True, inplace=True)

        # --- 第三步：拼接 ---
        # 顺序：Sheet3-A -> Sheet3-C -> Sheet4-B
        final_df = pd.concat([df_s3_clean.iloc[:, 0], df_s3_clean.iloc[:, 1], df_s4_clean.iloc[:, 0]], axis=1)
        final_df.columns = ["列1_来自S3_A", "列2_来自S3_C", "列3_来自S4_B"]

        # --- 第四步：清理残留的 "nan" 字符串 ---
        # 因为强制用了文本模式，原本的空值可能会变成字符串 "nan"，这里把它们变回真正的空
        # 这样导出CSV时就是空的，而不是显示 "nan"
        final_df = final_df.replace("nan", "")
        
        # --- 预览与下载 ---
        st.subheader("数据预览 (所见即所得)")
        st.dataframe(final_df.head(15))
        
        csv_data = final_df.to_csv(index=False, header=False, encoding='utf-8-sig').encode('utf-8-sig')
        
        st.download_button(
            label="📥 下载结果 CSV (无科学计数法)",
            data=csv_data,
            file_name="跨Sheet合并结果_文本版.csv",
            mime="text/csv"
        )
        
    except Exception as e:
        st.error(f"发生错误: {e}")
else:
    st.info("👆 请上传文件")
