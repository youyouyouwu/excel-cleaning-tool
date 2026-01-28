import streamlit as st
import pandas as pd
import io

# --- 页面配置 ---
st.set_page_config(page_title="业务报表清洗工具", layout="wide")

st.title("📊 业务报表转标准 CSV 清洗工具")
st.markdown("### 专治：合并单元格、多级表头、垃圾数据行")

# --- 侧边栏：上传与配置 ---
with st.sidebar:
    st.header("1. 上传文件")
    uploaded_file = st.file_uploader("上传 Excel 文件 (.xlsx)", type=["xlsx"])
    
    st.header("2. 清洗规则配置")
    # 既然是业务报表，表头通常不是第一行
    header_row = st.number_input("列名在第几行？(索引从0开始，0代表第一行)", min_value=0, value=0, step=1)
    
    fill_merged = st.checkbox("填充合并单元格 (推荐)", value=True, help="将合并单元格的内容填充到拆分后的所有格子里")
    drop_footer = st.text_input("删除包含此关键词的行 (如：合计, 总计)", value="合计")

# --- 核心清洗函数 ---
def clean_data(file, header_idx, do_fill, drop_kw):
    # 1. 读取数据
    try:
        # 即使 Excel 有合并单元格，Pandas 读入后通常是 "值, NaN, NaN" 的形式
        df = pd.read_excel(file, header=header_idx)
    except Exception as e:
        return None, f"读取失败: {e}"

    # 2. 处理合并单元格 (核心逻辑)
    # 业务报表中，如果 A1:A3 是合并的，Pandas 读出来 A1是有值的，A2-A3是 NaN
    # 使用 ffill() 向下填充即可完美还原业务逻辑
    if do_fill:
        # 只对 Object (文本) 类型的列进行填充，防止误伤数字列的空值（视具体业务而定，通常全填充更安全）
        df = df.ffill()

    # 3. 清理全空的行和列
    df.dropna(how='all', axis=0, inplace=True) # 删空行
    df.dropna(how='all', axis=1, inplace=True) # 删空列

    # 4. 删除包含特定关键词的行 (比如底部的小计、合计)
    if drop_kw:
        # 检查第一列是否包含关键词（通常合计都在第一列写着）
        mask = df.iloc[:, 0].astype(str).str.contains(drop_kw, na=False)
        df = df[~mask]

    return df, None

# --- 主界面逻辑 ---
if uploaded_file:
    # 1. 预览原始数据（方便用户确定表头在第几行）
    st.subheader("原始数据预览 (未清洗)")
    # 读取前10行不带header，方便用户数行数
    raw_preview = pd.read_excel(uploaded_file, header=None, nrows=10)
    st.dataframe(raw_preview)
    st.info(f"👆 请查看上表，确认真实的列名（表头）在第几行，并在左侧侧边栏设置。")

    # 2. 执行清洗
    if st.button("开始清洗数据"):
        with st.spinner('正在重组数据结构...'):
            cleaned_df, error = clean_data(uploaded_file, header_row, fill_merged, drop_footer)
            
            if error:
                st.error(error)
            else:
                st.success("清洗完成！")
                
                # 3. 展示结果
                col1, col2 = st.columns(2)
                with col1:
                    st.write(f"数据行数: {cleaned_df.shape[0]}")
                with col2:
                    st.write(f"数据列数: {cleaned_df.shape[1]}")
                
                st.dataframe(cleaned_df.head(50))
                
                # 4. 导出 CSV
                # encoding='utf-8-sig' 解决中文乱码问题
                csv = cleaned_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
                
                st.download_button(
                    label="📥 下载标准 CSV 文件",
                    data=csv,
                    file_name=f"cleaned_{uploaded_file.name.split('.')[0]}.csv",
                    mime="text/csv"
                )
else:
    st.info("👈 请在左侧上传你的业务报表 Excel")