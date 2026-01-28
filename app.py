import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="数据清洗最终版", layout="wide")
st.title("🏭 13列精准提取：S3(多列) + S4(B,I,L)")
st.markdown("### ✅ 配置：新增 S4 的 L 列 (输出到 M 列)")

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
        # 步骤 A: 处理 Sheet3 (读取 10 列)
        # ========================================================
        st.info("正在提取 Sheet3 数据...")
        
        cols_to_read = "A,C,E,F,N,O,AE,AB,AC,AF"
        
        df_s3 = pd.read_excel(
            uploaded_file, 
            sheet_name=sheet3_name, 
            header=None, 
            usecols=cols_to_read, 
            dtype=str
        )
        
        # 补全列防报错
        while df_s3.shape[1] < 10:
            df_s3[f"Auto_{df_s3.shape[1]}"] = ""
            
        df_s3 = df_s3.iloc[:, :10]
        
        df_s3.columns = [
            "S3_A", "S3_C", "S3_E", "S3_F", "S3_N", "S3_O", 
            "S3_AB", "S3_AC", "S3_AE", "S3_AF"
        ]
        
        # --- Sheet3 清洗 ---
        df_s3["S3_A"] = df_s3["S3_A"].ffill() # 炸开
        df_s3["S3_C"] = df_s3["S3_C"].ffill() # 炸开
        
        df_s3.reset_index(drop=True, inplace=True)

        # ========================================================
        # 步骤 B: 处理 Sheet4 (B, I, L)
        # ========================================================
        st.info("正在提取 Sheet4 数据 (新增 L 列)...")
        
        # 🟢 修改点：增加读取 L 列
        df_s4 = pd.read_excel(
            uploaded_file, 
            sheet_name=sheet4_name, 
            header=None, 
            usecols="B,I,L",  # <--- 加上 L
            dtype=str
        )
        
        # 补全列防报错 (防止L列没数据导致列数不够)
        while df_s4.shape[1] < 3:
            df_s4[f"S4_Auto_{df_s4.shape[1]}"] = ""
        
        # 确保只取前3列
        df_s4 = df_s4.iloc[:, :3]
        
        # 命名
        df_s4.columns = ["S4_B", "S4_I", "S4_L"]
        
        # --- Sheet4 清洗 ---
        df_s4["S4_B"] = df_s4["S4_B"].ffill() # B列炸开
        # I列 -> 原样保留
        # L列 -> 原样保留 (新加的)
        
        df_s4.reset_index(drop=True, inplace=True)

        # ========================================================
        # 步骤 C: 最终 13 列组装
        # ========================================================
        # 目标顺序：
        # A-L (前12列保持不变)
        # M: S4_L (新)
        
        final_df = pd.concat([
            df_s3["S3_A"],    # A
            df_s3["S3_C"],    # B
            df_s4["S4_B"],    # C
            df_s4["S4_I"],    # D (ID列)
            df_s3["S3_E"],    # E
            df_s3["S3_F"],    # F
            df_s3["S3_N"],    # G
            df_s3["S3_O"],    # H
            df_s3["S3_AE"],   # I
            df_s3["S3_AB"],   # J
            df_s3["S3_AC"],   # K
            df_s3["S3_AF"],   # L
            df_s4["S4_L"]     # M (新成员：Sheet4的L列)
        ], axis=1)
        
        final_df = final_df.replace("nan", "")

        # 🧹 ID 列深度清洗 (保留这个好功能，方便您核对)
        def clean_id(val):
            s = str(val).strip()
            if s.replace('.', '', 1).isdigit() and '.' in s:
                try:
                    return str(int(float(s)))
                except:
                    return s
            return s
        final_df.iloc[:, 3] = final_df.iloc[:, 3].apply(clean_id) # 第4列是ID列

        # ========================================================
        # 步骤 D: 预览与导出
        # ========================================================
        st.subheader("📋 13列数据预览")
        
        # 预览表头
        final_df.columns = [
            "A:产品", "B:C列", "C:S4-B", "D:ID", "E:S3-E", "F:S3-F",
            "G:S3-N", "H:S3-O", "I:S3-AE", "J:S3-AB", "K:S3-AC", "L:S3-AF",
            "M:S4-L列(新)" 
        ]
        
        st.dataframe(final_df.head(15))
        
        csv_data = final_df.to_csv(index=False, header=False, encoding='utf-8-sig').encode('utf-8-sig')
        
        st.download_button(
            label="📥 下载 CSV (13列完整版)",
            data=csv_data,
            file_name="13列数据提取结果.csv",
            mime="text/csv"
        )
        
    except Exception as e:
        st.error(f"发生错误: {e}")
else:
    st.info("👆 请上传文件")
