import streamlit as st
import pandas as pd

st.set_page_config(page_title="综合管理表格数据清洗", layout="wide")
st.title("🏭 13列精准提取：S3 + S4 + S5(L) + N(S3-B)")
st.markdown("### ✅ 配置更新：M 列提取 Sheet5 的 L 列；新增 N 列：Sheet3(1.COUPANG) 的 B 列")

uploaded_file = st.file_uploader("上传 Excel 文件 (.xlsx, .xlsm)", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        # ========================================================
        # 进度条 + 状态区（可视化）
        # ========================================================
        progress = st.progress(0)
        status = st.empty()

        def set_step(pct: int, msg: str):
            progress.progress(pct)
            status.info(msg)

        # ========================================================
        # 1. 解析 Excel 结构
        # ========================================================
        set_step(5, "正在解析 Excel 结构...")
        xl_file = pd.ExcelFile(uploaded_file, engine="openpyxl")
        sheet_names = xl_file.sheet_names

        # ⚠️ 关键检查：需要读取第5张表，所以总数不能少于5
        if len(sheet_names) < 5:
            st.error(f"❌ 文件只有 {len(sheet_names)} 个Sheet，无法读取 Sheet5 (第5张表)！")
            st.stop()

        sheet3_name = sheet_names[2]
        sheet4_name = sheet_names[3]
        sheet5_name = sheet_names[4]

        st.success(f"已锁定源表：\n1. {sheet3_name}\n2. {sheet4_name}\n3. {sheet5_name}")

        # ========================================================
        # 步骤 A: 处理 Sheet3（读取 11 列：新增 B）
        # ========================================================
        set_step(20, "正在提取 Sheet3 数据（A,B,C,E,F,N,O,AE,AB,AC,AF）...")
        cols_to_read = "A,B,C,E,F,N,O,AE,AB,AC,AF"

        df_s3 = pd.read_excel(
            uploaded_file,
            sheet_name=sheet3_name,
            header=None,
            usecols=cols_to_read,
            dtype=str,
            engine="openpyxl"
        )

        # 补全列防报错（现在应为 11 列）
        while df_s3.shape[1] < 11:
            df_s3[f"Auto_{df_s3.shape[1]}"] = ""

        df_s3 = df_s3.iloc[:, :11]
        df_s3.columns = [
            "S3_A", "S3_B", "S3_C", "S3_E", "S3_F", "S3_N", "S3_O",
            "S3_AE", "S3_AB", "S3_AC", "S3_AF"
        ]

        # 清洗：A/C 炸开；B 也炸开（保证“所有结果”完整）
        df_s3["S3_A"] = df_s3["S3_A"].ffill()
        df_s3["S3_B"] = df_s3["S3_B"].ffill()
        df_s3["S3_C"] = df_s3["S3_C"].ffill()
        df_s3.reset_index(drop=True, inplace=True)

        set_step(45, "✅ Sheet3 提取完成，正在提取 Sheet4 数据（B,I）...")

        # ========================================================
        # 步骤 B: 处理 Sheet4（B, I）
        # ========================================================
        df_s4 = pd.read_excel(
            uploaded_file,
            sheet_name=sheet4_name,
            header=None,
            usecols="B,I",
            dtype=str,
            engine="openpyxl"
        )

        # 容错：确保两列存在
        if df_s4.shape[1] < 2:
            df_s4["S4_I"] = ""

        df_s4.columns = ["S4_B", "S4_I"]

        # 清洗：B 炸开，I 原样
        df_s4["S4_B"] = df_s4["S4_B"].ffill()
        df_s4.reset_index(drop=True, inplace=True)

        set_step(65, "✅ Sheet4 提取完成，正在提取 Sheet5 数据（L列）...")

        # ========================================================
        # 步骤 C: 处理 Sheet5（只读取 L 列）
        # ========================================================
        df_s5 = pd.read_excel(
            uploaded_file,
            sheet_name=sheet5_name,
            header=None,
            usecols="L",
            dtype=str,
            engine="openpyxl"
        )

        # 容错：万一 Sheet5 空的连 L 列都没有
        if df_s5.shape[1] == 0:
            df_s5["S5_L"] = ""
        else:
            df_s5.columns = ["S5_L"]

        df_s5.reset_index(drop=True, inplace=True)

        set_step(80, "✅ Sheet5 提取完成，正在组装最终结果...")

        # ========================================================
        # 步骤 D: 最终 14 列组装（新增 N=Sheet3 B列）
        # ========================================================
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
            df_s5["S5_L"],    # M (Sheet5 的 L)
            df_s3["S3_B"],    # N ✅ (Sheet3 的 B)
        ], axis=1)

        # 清理 nan 字符串
        final_df = final_df.replace("nan", "")

        # 🧹 ID 列深度清洗（第4列，索引3）
        def clean_id(val):
            s = str(val).strip()
            if s.replace(".", "", 1).isdigit() and "." in s:
                try:
                    return str(int(float(s)))
                except:
                    return s
            return s

        final_df.iloc[:, 3] = final_df.iloc[:, 3].apply(clean_id)

        # ========================================================
        # 步骤 E: 预览与导出
        # ========================================================
        final_df.columns = [
            "A:产品", "B:C列", "C:S4-B", "D:ID", "E:S3-E", "F:S3-F",
            "G:S3-N", "H:S3-O", "I:S3-AE", "J:S3-AB", "K:S3-AC", "L:S3-AF",
            "M:Sheet5-L列", "N:Sheet3-B列"
        ]

        progress.progress(100)
        status.success("🎉 全部完成！可预览全量数据并下载 CSV")

        st.subheader("📋 数据预览（全量）")
        st.dataframe(
            final_df,
            use_container_width=True,
            height=1000  # ✅ 想更长就改大，比如 1200/1500
        )

        csv_data = final_df.to_csv(index=False, header=False, encoding="utf-8-sig").encode("utf-8-sig")

        st.download_button(
            label="📥 下载 CSV（新增N列）",
            data=csv_data,
            file_name="13列数据提取结果.csv",
            mime="text/csv"
        )

    except Exception as e:
        st.error(f"发生错误: {e}")
else:
    st.info("👆 请上传文件")
