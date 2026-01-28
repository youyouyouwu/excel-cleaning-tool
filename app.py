import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="完美格式导出工具", layout="wide")
st.title("🏭 跨表组装：智能适中列宽版")

uploaded_file = st.file_uploader("上传 Excel 文件 (.xlsx, .xlsm)", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        # 1. 解析 Excel 结构
        xl_file = pd.ExcelFile(uploaded_file, engine='openpyxl')
        sheet_names = xl_file.sheet_names
        
        if len(sheet_names) < 4:
            st.error(f"❌ 文件只有 {len(sheet_names)} 个表，无法读取第 4 张表！")
            st.stop()
            
        sheet3_name = sheet_names[2]
        sheet4_name = sheet_names[3]
        
        st.success(f"已锁定源表：【{sheet3_name}】 和 【{sheet4_name}】")

        # --- 2. 读取数据 (强制文本模式 dtype=str) ---
        df_s3 = pd.read_excel(uploaded_file, sheet_name=sheet3_name, header=None, usecols="A,C", dtype=str)
        df_s4 = pd.read_excel(uploaded_file, sheet_name=sheet4_name, header=None, usecols="B", dtype=str)
        
        # 清洗填充
        df_s3_clean = df_s3.ffill()
        df_s3_clean.reset_index(drop=True, inplace=True)
        
        df_s4_clean = df_s4.ffill()
        df_s4_clean.reset_index(drop=True, inplace=True)

        # 拼接
        final_df = pd.concat([df_s3_clean.iloc[:, 0], df_s3_clean.iloc[:, 1], df_s4_clean.iloc[:, 0]], axis=1)
        final_df.columns = ["Sheet3_A", "Sheet3_C", "Sheet4_B"]
        
        # 去除 'nan'
        final_df = final_df.replace("nan", "")

        # --- 3. 预览 ---
        st.subheader("数据预览")
        st.dataframe(final_df.head(15))

        # --- 4. 导出 Excel (优化列宽逻辑) ---
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            final_df.to_excel(writer, index=False, header=False, sheet_name='清洗结果')
            
            workbook = writer.book
            worksheet = writer.sheets['清洗结果']
            
            # 文本格式 (防科学计数法)
            text_format = workbook.add_format({'num_format': '@'})
            
            # --- 💡 智能适中列宽逻辑 ---
            for idx, col in enumerate(final_df.columns):
                # 1. 估算该列最大长度
                series = final_df[col].astype(str)
                
                # 计算“视觉长度”技巧：如果不全是英文，稍微把宽度乘个系数，因为中文比英文宽
                # 这里简单处理：取字符长度的最大值
                max_len = series.map(len).max()
                
                if pd.isna(max_len):
                    max_len = 10
                
                # 2. 设定算法：长度 + 2个字符的边距
                calc_width = max_len + 2
                
                # 3. 🛑 关键限制：设置下限10，上限40
                # 如果算出来是 100，我也只给 40，防止太夸张
                # 如果算出来是 2，我强制给 10，防止看不见
                final_width = max(10, min(calc_width, 40))
                
                worksheet.set_column(idx, idx, final_width, text_format)
                
        output.seek(0)
        
        st.download_button(
            label="📥 下载适中列宽 Excel (.xlsx)",
            data=output,
            file_name="清洗结果_适中列宽.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"发生错误: {e}")
else:
    st.info("👆 请上传文件")
