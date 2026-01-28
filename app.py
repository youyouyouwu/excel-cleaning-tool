import streamlit as st
import pandas as pd

st.set_page_config(page_title="数据清洗最终版", layout="wide")
st.title("🏭 12列精准提取：S3(多列) + S4(B,I)")
st.markdown("### ✅ 配置：新增 S3 的 N/O/AE/AB/AC/AF 列 (原样保留)")

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
        # 步骤 A: 处理 Sheet3 (读取所有需要的列)
        # ========================================================
        st.info("正在提取 Sheet3 数据...")
        
        # 技巧：usecols 可以乱序写，但 Pandas 读进来会按 Excel 原始顺序(左->右)排列
        # 我们这里把所有要用的列都写上
        cols_to_read = "A,C,E,F,N,O,AE,AB,AC,AF"
        
        df_s3 = pd.read_excel(
            uploaded_file, 
            sheet_name=sheet3_name, 
            header=None, 
            usecols=cols_to_read, 
            dtype=str
        )
        
        # ⚠️ 关键：Pandas 读取后的列顺序是 Excel 的物理顺序：
        # A, C, E, F, N, O, AB, AC, AE, AF  (注意 AB 在 AE 前面)
        # 我们必须按这个顺序给它们起内部代号，后面才能拼对
        df_s3.columns = [
            "S3_A", "S3_C", "S3_E", "S3_F", "S3_N", "S3_O", 
            "S3_AB", "S3_AC", "S3_AE", "S3_AF"
        ]
        
        # --- Sheet3 清洗逻辑 ---
        # 1. 炸开组 (分类信息)
        df_s3["S3_A"] = df_s3["S3_A"].ffill()
        df_s3["S3_C"] = df_s3["S3_C"].ffill()
        
        # 2. 原样组 (E, F, N, O, AE, AB, AC, AF)
        # 这些列保持不动，防止数据篡改
        
        df_s3.reset_index(drop=True, inplace=True)

        # ========================================================
        # 步骤 B: 处理 Sheet4 (B, I)
        # ========================================================
        st.info("正在提取 Sheet4 数据...")
        
        df_s4 = pd.read_excel(
            uploaded_file, 
            sheet_name=sheet4_name, 
            header=None, 
            usecols="B,I", 
            dtype=str
        )
        
        if df_s4.shape[1] == 1:
            df_s4["S4_I"] = ""
        
        df_s4.columns = ["S4_B", "S4_I"]
        
        # Sheet4 清洗
        df_s4["S4_B"] = df_s4["S4_B"].ffill() # B列炸开
        # I列原样 (不动)
        
        df_s4.reset_index(drop=True, inplace=True)

        # ========================================================
        # 步骤 C: 最终 12 列组装 (严格按您要求的顺序)
        # ========================================================
        # 目标顺序：
        # A, B, C, D, E, F (前6列旧逻辑)
        # G: S3_N
        # H: S3_O
        # I: S3_AE  <-- 注意顺序
        # J: S3_AB
        # K: S3_AC
        # L: S3_AF
        
        final_df = pd.concat([
            df_s3["S3_A"],    # Result A
            df_s3["S3_C"],    # Result B
            df_s4["S4_B"],    # Result C
            df_s4["S4_I"],    # Result D
            df_s3["S3_E"],    # Result E
            df_s3["S3_F"],    # Result F
            df_s3["S3_N"],    # Result G (新)
            df_s3["S3_O"],    # Result H (新)
            df_s3["S3_AE"],   # Result I (新) -> 您的要求
            df_s3["S3_AB"],   # Result J (新)
            df_s3["S3_AC"],   # Result K (新)
            df_s3["S3_AF"]    # Result L (新)
        ], axis=1)
        
        final_df = final_df.replace("nan", "")

        # ========================================================
        # 步骤 D: 预览与导出
        # ========================================================
        st.subheader("📋 12列数据预览")
        
        # 设置表头用于预览 (不影响CSV)
        final_df.columns = [
            "A:产品(炸)", "B:C列(炸)", "C:S4-B(炸)", "D:S4-I", "E:S3-E", "F:S3-F",
            "G:S3-N", "H:S3-O", "I:S3-AE", "J:S3-AB", "K:S3-AC", "L:S3-AF"
        ]
        
        st.dataframe(final_df.head(15))
        
        csv_data = final_df.to_csv(index=False, header=False, encoding='utf-8-sig').encode('utf-8-sig')
        
        st.download_button(
            label="📥 下载 CSV (12列完整版)",
            data=csv_data,
            file_name="12列数据提取结果.csv",
            mime="text/csv"
        )
        
    except Exception as e:
        st.error(f"发生错误: {e}")
else:
    st.info("👆 请上传文件")
