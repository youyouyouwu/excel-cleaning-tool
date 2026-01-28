import streamlit as st
import pandas as pd

st.set_page_config(page_title="双列提取工具", layout="wide")
st.title("🎯 锁定提取：第3张表 - A列 & C列")

uploaded_file = st.file_uploader("上传 Excel 文件 (.xlsx, .xlsm)", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        # 1. 锁定第3张表 (逻辑不变)
        xl_file = pd.ExcelFile(uploaded_file, engine='openpyxl')
        sheet_names = xl_file.sheet_names
        
        target_index = 2
        if len(sheet_names) > target_index:
            target_sheet_name = sheet_names[target_index]
            st.success(f"已锁定第 3 张表：【{target_sheet_name}】")
        else:
            target_sheet_name = sheet_names[-1]
            st.warning(f"警告：文件少于3个表，已选择：【{target_sheet_name}】")

        # 2. 读取数据 (关键修改：同时抓 A 和 C)
        st.info("正在提取 A列 和 C列，并执行【炸开合并单元格】...")
        
        # usecols="A,C"：告诉 Python 只把这两列抓进内存
        # 抓进来后：DataFrame 的第1列就是原A列，第2列就是原C列
        df = pd.read_excel(uploaded_file, sheet_name=target_sheet_name, header=None, usecols="A,C")
        
        # 给它个临时名字方便你预览分辨
        df.columns = ["源表_A列", "源表_C列"]
        
        # 3. 核心清洗：双列同时向下填充 (ffill)
        # 这行代码会分别把 A列的空值用A列上方填满，C列的用C列上方填满，互不干扰
        df_cleaned = df.ffill()
        
        # 4. 预览对比
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("清洗前 (可能含空值)")
            st.dataframe(df.head(15))
        with col2:
            st.subheader("清洗后 (结果预览)")
            # 这里展示的就是最终要导出的样子：左边是原A的内容，右边是原C的内容
            st.dataframe(df_cleaned.head(15))
            
        # 5. 导出 CSV
        # 结果文件说明：第一列 = 原A列清洗版，第二列 = 原C列清洗版
        # header=False：不带表头，纯数据导出
        csv_data = df_cleaned.to_csv(index=False, header=False, encoding='utf-8-sig').encode('utf-8-sig')
        
        st.markdown("---")
        st.download_button(
            label="📥 下载双列结果 CSV",
            data=csv_data,
            file_name=f"{target_sheet_name}_AC提取.csv",
            mime="text/csv"
        )
        
    except Exception as e:
        st.error(f"处理出错: {e}")
else:
    st.info("👆 请上传文件，将提取【第3张表】的 A列与C列")
