import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="数据清洗实验室", layout="wide")
st.title("🛠️ 业务报表清洗 - 调试模式")

# 1. 支持上传 xlsm (带宏文件)
uploaded_file = st.file_uploader("请上传你的基础表格 (.xlsx, .xlsm)", type=["xlsx", "xlsm"])

if uploaded_file:
    st.markdown("---")
    st.subheader("1. 寻找表头 (Header)")
    
    # 让用户动态调整表头位置，找到数据的“第一行”
    header_idx = st.slider("请拖动滑块，直到下方的表格【第一行】显示为正确的中文列名", 0, 10, 0)
    
    try:
        # 强制用 openpyxl 引擎读取，兼容 xlsm
        df_raw = pd.read_excel(uploaded_file, header=header_idx, engine='openpyxl')
        
        # 显示前 20 行数据让用户看
        st.dataframe(df_raw.head(20))
        
        st.markdown("---")
        st.subheader("2. 请根据上表回答我的问题")
        st.info("数据已加载。现在请在聊天框告诉我下一步的要求。")
        
        # 显示列名列表，方便复制
        with st.expander("点击查看所有识别到的列名"):
            st.write(list(df_raw.columns))
            
    except Exception as e:
        st.error(f"读取出错: {e}")
