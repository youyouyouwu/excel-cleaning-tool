import streamlit as st
import pandas as pd

st.set_page_config(page_title="数据清洗最终版", layout="wide")
st.title("🏭 最终版：Sheet3(A,C) + Sheet4(B) -> CSV")

uploaded_file = st.file_uploader("上传 Excel 文件 (.xlsx, .xlsm)", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        # 1. 解析结构
        xl_file = pd.ExcelFile(uploaded_file, engine='openpyxl')
        sheet_names = xl_file.sheet_names
        
        if len(sheet_names) < 4:
            st.error("❌ 文件Sheet数量不足4个")
            st.stop()
            
        sheet3_name = sheet_names[2]
        sheet4_name = sheet_names[3]
        
        st.success(f"已锁定：【{sheet3_name}】 和 【{sheet4_name}】")

        # 2. 读取数据 (关键：dtype=str 保证数据底层不丢失精度)
        df_s3 = pd.read_excel(uploaded_file, sheet_name=sheet3_name, header=None, usecols="A,C", dtype=str)
        df_s4 = pd.read_excel(uploaded_file, sheet_name=sheet4_name, header=None, usecols="B", dtype=str)
        
        # 3. 填充清洗
        df_s3 = df_s3.ffill()
        df_s4 = df_s4.ffill()
        
        # 重置索引防止错位
        df_s3.reset_index(drop=True, inplace=True)
        df_s4.reset_index(drop=True, inplace=True)

        # 4. 拼接 (A+C+B)
        final_df = pd.concat([df_s3.iloc[:, 0], df_s3.iloc[:, 1], df_s4.iloc[:, 0]], axis=1)
        
        # 5. 去除 'nan' 文本
        final_df = final_df.replace("nan", "")

        # 6. 预览
        st.subheader("数据预览 (如果您看到长数字完整显示，说明数据是安全的)")
        st.dataframe(final_df.head(15))
        
        # 7. 导出标准 CSV
        csv_data = final_df.to_csv(index=False, header=False, encoding='utf-8-sig').encode('utf-8-sig')
        
        st.download_button(
            label="📥 下载最终 CSV (数据绝对安全)",
            data=csv_data,
            file_name="清洗结果_Final.csv",
            mime="text/csv"
        )
        
    except Exception as e:
        st.error(f"发生错误: {e}")
else:
    st.info("👆 请上传文件")
