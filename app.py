import streamlit as st
import pandas as pd
import io # 需要用到 IO流

st.set_page_config(page_title="完美格式导出工具", layout="wide")
st.title("🏭 跨表组装：自动调整列宽 + 防科学计数法")

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
        # 即使这里强制文本，如果存成 CSV，Excel 打开还是会可能会变回去
        # 但我们这次存成 xlsx，就能完美保持住
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
        
        # 去除 'nan' 字符
        final_df = final_df.replace("nan", "")

        # --- 3. 预览 ---
        st.subheader("数据预览")
        st.dataframe(final_df.head(15))

        # --- 4. 核心升级：导出为带格式的 Excel (.xlsx) ---
        # 创建一个内存里的 Excel 文件
        output = io.BytesIO()
        
        # 使用 xlsxwriter 引擎，因为它支持设置格式
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # 写入数据，不带索引，不带表头(header=False)
            final_df.to_excel(writer, index=False, header=False, sheet_name='清洗结果')
            
            # 获取 workbook 和 worksheet 对象
            workbook = writer.book
            worksheet = writer.sheets['清洗结果']
            
            # 定义一个“纯文本”格式，防止Excel自作聪明变科学计数法
            text_format = workbook.add_format({'num_format': '@'})
            
            # --- 智能调整列宽逻辑 ---
            for idx, col in enumerate(final_df.columns):
                # 计算这一列最长的一行有多少个字符
                # map(str, ...) 确保所有内容转字符串，防止报错
                series = final_df[col].astype(str)
                # 找出这一列最长的内容长度
                max_len = series.map(len).max()
                
                # 如果全是空的，给个默认宽度 10
                if pd.isna(max_len):
                    max_len = 10
                
                # 设置稍微宽一点点，保证能看全 (比如 +2)
                # 限制一下最大宽度，防止有一行写作文导致列宽 500
                final_width = min(max_len + 2, 60) 
                
                # 应用列宽 和 文本格式
                # set_column(开始列, 结束列, 宽度, 格式)
                worksheet.set_column(idx, idx, final_width, text_format)
                
        # 准备下载
        output.seek(0)
        
        st.download_button(
            label="📥 下载完美格式 Excel (.xlsx)",
            data=output,
            file_name="清洗结果_自动列宽.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"发生错误: {e}")
else:
    st.info("👆 请上传文件")
