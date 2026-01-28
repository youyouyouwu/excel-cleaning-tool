import streamlit as st
import pandas as pd

st.set_page_config(page_title="固定位置提取工具", layout="wide")
st.title("🎯 锁定提取：Excel 第3张表 - A列")

uploaded_file = st.file_uploader("上传 Excel 文件 (.xlsx, .xlsm)", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        # 1. 获取所有 Sheet 名称列表
        xl_file = pd.ExcelFile(uploaded_file, engine='openpyxl')
        sheet_names = xl_file.sheet_names
        
        # -------------------------------------------------------
        # 核心修改：强制锁定第 3 张表 (索引为 2，因为计算机从0开始数)
        # -------------------------------------------------------
        target_index = 2  # 0是第1张，1是第2张，2是第3张
        
        # 安全检查：万一文件里只有1张表，防止报错
        if len(sheet_names) > target_index:
            target_sheet_name = sheet_names[target_index]
            st.success(f"已锁定第 3 张表，检测到表名为：【{target_sheet_name}】")
        else:
            # 如果表不够3张，默认取最后一张
            target_sheet_name = sheet_names[-1]
            st.warning(f"警告：文件少于3个表，已自动选择最后一张：【{target_sheet_name}】")

        # 2. 读取数据 (只读 A 列)
        st.info(f"正在读取 A 列并执行【炸开合并单元格】...")
        
        df = pd.read_excel(uploaded_file, sheet_name=target_sheet_name, header=None, usecols="A")
        df.columns = ["原始数据"]
        
        # 3. 炸开合并单元格 (向下填充)
        df["清洗后数据"] = df["原始数据"].ffill()
        
        # 4. 展示前20行供检查
        st.subheader("数据预览 (前20行)")
        st.dataframe(df.head(20))
            
        # 5. 导出 CSV
        result_df = df[["清洗后数据"]]
        # 这里的 header=False 表示导出的 CSV 不带表头，纯数据
        # 如果你需要表头，把 header=False 改成 header=["业务数据"]
        csv_data = result_df.to_csv(index=False, header=False, encoding='utf-8-sig').encode('utf-8-sig')
        
        st.download_button(
            label=f"📥 下载 CSV ({target_sheet_name}_A列.csv)",
            data=csv_data,
            file_name=f"{target_sheet_name}_A列.csv",
            mime="text/csv"
        )
        
    except Exception as e:
        st.error(f"处理出错: {e}")

else:
    st.info("👆 请上传文件，我将自动提取【第 3 张表】的 A 列")
