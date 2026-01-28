import streamlit as st
import pandas as pd

st.set_page_config(page_title="数据清洗与可视化", layout="wide")
st.title("🏭 全能版：清洗 + 可视化分析")

uploaded_file = st.file_uploader("上传 Excel 文件 (.xlsx, .xlsm)", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        # --- 1. 读取与清洗 (保持之前的逻辑不变) ---
        xl_file = pd.ExcelFile(uploaded_file, engine='openpyxl')
        sheet_names = xl_file.sheet_names
        
        if len(sheet_names) < 4:
            st.error("❌ Sheet数量不足")
            st.stop()
            
        sheet3_name = sheet_names[2]
        sheet4_name = sheet_names[3]
        
        # 读取 Sheet3 (A,C,E,F)
        df_s3 = pd.read_excel(uploaded_file, sheet_name=sheet3_name, header=None, usecols="A,C,E,F", dtype=str)
        df_s3.columns = ["Col_A", "Col_C", "Col_E", "Col_F"]
        
        # 炸开 A, C
        df_s3["Col_A"] = df_s3["Col_A"].ffill()
        df_s3["Col_C"] = df_s3["Col_C"].ffill()
        df_s3.reset_index(drop=True, inplace=True)
        
        # 读取 Sheet4 (B)
        df_s4 = pd.read_excel(uploaded_file, sheet_name=sheet4_name, header=None, usecols="B", dtype=str)
        df_s4 = df_s4.ffill()
        df_s4.reset_index(drop=True, inplace=True)

        # 拼接 (给列起个直观的名字，方便后面选)
        final_df = pd.concat([
            df_s3["Col_A"], 
            df_s3["Col_C"], 
            df_s4.iloc[:, 0], 
            df_s3["Col_E"], 
            df_s3["Col_F"]
        ], axis=1)
        
        # 设置展示给用户看的列名
        column_names = [
            "第1列(原S3-A)", 
            "第2列(原S3-C)", 
            "第3列(原S4-B)", 
            "第4列(原S3-E)", 
            "第5列(原S3-F)"
        ]
        final_df.columns = column_names
        final_df = final_df.replace("nan", "")

        # --- 2. 展示数据与下载 ---
        col1, col2 = st.columns([1, 1])
        with col1:
            st.subheader("📋 清洗结果预览")
            st.dataframe(final_df.head(10))
            
        with col2:
            st.subheader("📥 下载数据")
            csv_data = final_df.to_csv(index=False, header=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button(
                label="下载最终 CSV",
                data=csv_data,
                file_name="清洗结果_可视化版.csv",
                mime="text/csv"
            )

        st.markdown("---")

        # ========================================================
        # 核心新增：📊 数据可视化区域
        # ========================================================
        st.header("📊 数据可视化分析")
        
        # 1. 创建一个用于绘图的副本 (以免破坏原始数据的文本格式)
        plot_df = final_df.copy()
        
        # 2. 让用户选择 X 轴 (分类) 和 Y 轴 (数值)
        c1, c2, c3 = st.columns(3)
        with c1:
            x_axis = st.selectbox("选择 X 轴 (分类/名称)", column_names, index=0)
        with c2:
            # 默认选第4列和第5列作为数值，因为它们通常是金额
            default_y = [column_names[3], column_names[4]]
            y_axis_list = st.multiselect("选择 Y 轴 (数值/金额)", column_names, default=default_y)
        with c3:
            chart_type = st.radio("图表类型", ["柱状图 (Bar)", "折线图 (Line)", "面积图 (Area)"], horizontal=True)

        # 3. 数据转换：将选中的 Y 轴列强制转为数字
        if y_axis_list:
            try:
                for col in y_axis_list:
                    # errors='coerce' 意思是：如果遇到无法转成数字的文字，就强制变成 0
                    plot_df[col] = pd.to_numeric(plot_df[col], errors='coerce').fillna(0)
                
                # 4. 聚合数据 (可选)
                # 比如：同一个“产品A”出现了多次，我们需要把它加总
                st.caption(f"正在按【{x_axis}】合并计算总和...")
                chart_data = plot_df.groupby(x_axis)[y_axis_list].sum()
                
                # 5. 绘图
                if chart_type == "柱状图 (Bar)":
                    st.bar_chart(chart_data)
                elif chart_type == "折线图 (Line)":
                    st.line_chart(chart_data)
                else:
                    st.area_chart(chart_data)
                    
            except Exception as e:
                st.warning("⚠️ 无法生成图表，请检查您选择的 Y 轴是否包含数字。")
        else:
            st.info("请在上方选择至少一列作为 Y 轴数值")
            
    except Exception as e:
        st.error(f"发生错误: {e}")
else:
    st.info("👆 请上传文件")
