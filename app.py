import streamlit as st
import pandas as pd

st.set_page_config(page_title="综合管理表格数据清洗", layout="wide")
st.title("🏭 精准提取：S3 + S4 + S5(L) + N(S3-B) + O(原样DISPIMG公式)")
st.markdown("### ✅ 新增：当 Sheet3 的 B 列为 Y 时，将 Sheet3 的 D 列（合并单元格图片函数）原样输出到 O 列")

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
        # 步骤 A: 处理 Sheet3（新增读取 D 列：图片函数列，且 D 需要炸开）
        # ========================================================
        set_step(20, "正在提取 Sheet3 数据（A,B,C,D,E,F,N,O,AE,AB,AC,AF）...")

        # ✅ 新增 D
        cols_to_read = "A,B,C,D,E,F,N,O,AE,AB,AC,AF"

        df_s3 = pd.read_excel(
            uploaded_file,
            sheet_name=sheet3_name,
            header=None,
            usecols=cols_to_read,
            dtype=str,
            engine="openpyxl"
        )

        # 补全列防报错（现在应为 12 列）
        while df_s3.shape[1] < 12:
            df_s3[f"Auto_{df_s3.shape[1]}"] = ""

        df_s3 = df_s3.iloc[:, :12]
        df_s3.columns = [
            "S3_A", "S3_B", "S3_C", "S3_D",
            "S3_E", "S3_F", "S3_N", "S3_O",
            "S3_AE", "S3_AB", "S3_AC", "S3_AF"
        ]

        # ✅ 维持原逻辑：A/C 向下填充（炸开）
        df_s3["S3_A"] = df_s3["S3_A"].ffill()
        df_s3["S3_C"] = df_s3["S3_C"].ffill()

        # ✅ D 列是合并单元格图片函数：必须炸开，否则逐行对位会丢失
        df_s3["S3_D"] = df_s3["S3_D"].ffill()

        # ❌ B 列不要 ffill（否则 Y 会向下扩散）
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

        # 清洗：B 炸开，I 原样（保持你原逻辑）
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

        if df_s5.shape[1] == 0:
            df_s5["S5_L"] = ""
        else:
            df_s5.columns = ["S5_L"]

        df_s5.reset_index(drop=True, inplace=True)

        set_step(80, "✅ Sheet5 提取完成，正在生成 O 列（原样DISPIMG公式）并组装最终结果...")

        # ========================================================
        # 步骤 D1: 生成 O 列（当 B=Y 时，原样输出 D 列公式）
        # ========================================================
        def as_clean_text(x) -> str:
            if x is None:
                return ""
            s = str(x)
            if s.lower() in ("nan", "none"):
                return ""
            return s

        def make_o_formula(b_val, d_val) -> str:
            b = as_clean_text(b_val).strip().upper()
            if b != "Y":
                return ""
            # ✅ 原样输出（WPS 打开可显示图片）
            return as_clean_text(d_val).strip()

        df_s3["S3_O_FORMULA"] = [
            make_o_formula(b, d) for b, d in zip(df_s3["S3_B"], df_s3["S3_D"])
        ]

        # ========================================================
        # 步骤 D2: 最终列组装（新增 O 列）
        # ========================================================
        final_df = pd.concat([
            df_s3["S3_A"],         # A
            df_s3["S3_C"],         # B
            df_s4["S4_B"],         # C
            df_s4["S4_I"],         # D (ID列)
            df_s3["S3_E"],         # E
            df_s3["S3_F"],         # F
            df_s3["S3_N"],         # G
            df_s3["S3_O"],         # H
            df_s3["S3_AE"],        # I
            df_s3["S3_AB"],        # J
            df_s3["S3_AC"],        # K
            df_s3["S3_AF"],        # L
            df_s5["S5_L"],         # M
            df_s3["S3_B"],         # N（原样，不扩散）
            df_s3["S3_O_FORMULA"]  # O ✅（原样公式）
        ], axis=1)

        # 清理 nan/None 字符串（注意：不能动公式本身）
        final_df = final_df.replace(["nan", "None"], "")

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
            "M:Sheet5-L列", "N:Sheet3-B列", "O:图片函数(当B=Y原样输出)"
        ]

        progress.progress(100)
        status.success("🎉 全部完成！可预览全量数据并下载 CSV")

        st.subheader("📋 数据预览（全量）")
        st.dataframe(
            final_df,
            use_container_width=True,
            height=1100
        )

        # ✅ 注意：CSV 本身不会“计算公式”，但保留公式文本，WPS 打开后能按你的需求处理
        csv_data = final_df.to_csv(index=False, header=False, encoding="utf-8-sig").encode("utf-8-sig")

        st.download_button(
            label="📥 下载 CSV（新增O列）",
            data=csv_data,
            file_name="13列数据提取结果.csv",
            mime="text/csv"
        )

    except Exception as e:
        st.error(f"发生错误: {e}")
else:
    st.info("👆 请上传文件")
