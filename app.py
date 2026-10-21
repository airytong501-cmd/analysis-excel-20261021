# ============================================
# 資料整理程式 - Streamlit 版（含統計 Sheet）
# 功能：
# 1. 上傳 Excel
# 2. 自動判斷完整 / 不完整資料
# 3. 分成三個 Sheet 輸出（完整、不完整、統計）
# 4. 提供下載按鈕
# ============================================

import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# --------------------------
# 頁面設定
# --------------------------
st.set_page_config(
    page_title="Excel 資料完整性整理工具",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Excel 資料完整性整理工具")
st.markdown("""
上傳 Excel 檔案後，程式會自動判斷每筆資料是否完整，並產生包含三個工作表的結果檔：
- **完整資料**
- **不完整資料**
- **統計資料**（基本統計、缺失情況、分布比較、數值統計）
""")

# --------------------------
# 判斷函式（與原 Colab 版相同）
# --------------------------
def is_null_like(value):
    """判斷是否為空值、NULL、N/A、空白等"""
    if pd.isna(value):
        return True
    s = str(value).strip().lower()
    null_keywords = {
        '', 'null', 'n/a', 'na', 'nan', 'none', 'none_provided',
        'nil', '#na', '#n/a', 'unknown', '-', '--', 'n.a.', 'n.a'
    }
    return s in null_keywords

def is_valid_date(value):
    """檢查日期格式是否合理"""
    if is_null_like(value):
        return False
    s = str(value).strip()

    # 常見無效日期
    if re.search(r'1990-00-00|0000-00-00|02-31|02/31|2023-02-31', s):
        return False

    formats = [
        '%Y-%m-%d', '%Y/%m/%d', '%d/%m/%Y', '%d-%m-%Y',
        '%Y.%m.%d', '%d.%m.%Y', '%Y-%m-%d %H:%M:%S'
    ]
    for fmt in formats:
        try:
            datetime.strptime(s[:10], fmt)
            return True
        except:
            continue
    return False

def is_valid_phone(value):
    """檢查電話格式是否大致合理（含數字且長度足夠）"""
    if is_null_like(value):
        return False
    s = str(value).strip()
    digits = re.sub(r'[^\d]', '', s)
    return len(digits) >= 8

def is_valid_email(value):
    """簡單檢查 Email 格式"""
    if is_null_like(value):
        return False
    s = str(value).strip()
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(pattern, s))

def is_valid_age(value):
    """檢查年齡是否合理（0~120）"""
    if is_null_like(value):
        return False
    try:
        age = float(str(value).strip())
        return 0 <= age <= 120
    except:
        return False

def is_valid_salary(value):
    """檢查薪資是否為合理正數"""
    if is_null_like(value):
        return False
    try:
        s = str(value).replace('$', '').replace(',', '').strip()
        salary = float(s)
        return salary > 0
    except:
        return False

def is_complete_row(row):
    """
    判斷一筆資料是否完整。
    只要有以下任一問題，就視為不完整：
    - 重要欄位為空 / NULL / N/A
    - 日期格式錯誤
    - 電話格式不合理
    - Email 格式錯誤
    - 年齡異常
    - 薪資異常
    - 明顯大小寫混亂或特殊字元
    """
    # --- 重要欄位不能為空 ---
    important_cols = [
        'Record_ID', 'First_Name', 'Last_Name', 'Email_Address',
        'Phone_Number', 'Department', 'Employment_Status', 'Salary_Monthly'
    ]
    for col in important_cols:
        if col in row.index and is_null_like(row[col]):
            return False

    # --- 日期檢查 ---
    date_cols = ['DOB', 'Hire_Date', 'Contract_End_Date', 'Registration_Date', 'Last_Purchase_Date']
    for col in date_cols:
        if col in row.index and not is_null_like(row[col]):
            if not is_valid_date(row[col]):
                return False

    # --- 電話檢查 ---
    if 'Phone_Number' in row.index:
        if not is_valid_phone(row['Phone_Number']):
            return False

    # --- Email 檢查 ---
    if 'Email_Address' in row.index:
        if not is_valid_email(row['Email_Address']):
            return False

    # --- 年齡檢查 ---
    if 'Age' in row.index:
        if not is_valid_age(row['Age']):
            return False

    # --- 薪資檢查 ---
    if 'Salary_Monthly' in row.index:
        if not is_valid_salary(row['Salary_Monthly']):
            return False

    # --- 特殊字元或明顯髒資料標記 ---
    if 'Internal_Notes' in row.index:
        notes = str(row['Internal_Notes']).lower()
        if any(x in notes for x in ['test row', 'special chars', 'duplicate suspected']):
            return False

    # --- 如果原本就有 Data_Quality_Label 且為 DIRTY，也視為不完整 ---
    if 'Data_Quality_Label' in row.index:
        if str(row['Data_Quality_Label']).strip().upper() == 'DIRTY':
            return False

    return True

def generate_stats(df, df_complete, df_incomplete):
    """產生統計資料"""
    total_rows = len(df)
    complete_count = len(df_complete)
    incomplete_count = len(df_incomplete)

    # 基本統計
    stats_basic = pd.DataFrame({
        '項目': [
            '原始總筆數',
            '完整資料筆數',
            '不完整資料筆數',
            '完整資料比例 (%)',
            '不完整資料比例 (%)'
        ],
        '數值': [
            total_rows,
            complete_count,
            incomplete_count,
            round(complete_count / total_rows * 100, 2) if total_rows > 0 else 0,
            round(incomplete_count / total_rows * 100, 2) if total_rows > 0 else 0
        ]
    })

    # 各欄位缺失統計
    missing_stats = []
    for col in df.columns:
        null_count = df[col].apply(is_null_like).sum()
        missing_stats.append({
            '欄位名稱': col,
            '缺失筆數': null_count,
            '缺失比例 (%)': round(null_count / total_rows * 100, 2) if total_rows > 0 else 0
        })
    df_missing = pd.DataFrame(missing_stats).sort_values('缺失筆數', ascending=False)

    # 完整 vs 不完整 在關鍵欄位的比較
    compare_cols = ['Gender', 'Department', 'Employment_Status', 'VIP_Tier', 'Currency']
    compare_list = []

    for col in compare_cols:
        if col in df.columns:
            complete_dist = df_complete[col].value_counts(dropna=False).head(10)
            for val, cnt in complete_dist.items():
                compare_list.append({
                    '欄位': col,
                    '類別': str(val),
                    '完整資料筆數': cnt,
                    '不完整資料筆數': 0
                })

            incomplete_dist = df_incomplete[col].value_counts(dropna=False).head(10)
            for val, cnt in incomplete_dist.items():
                found = False
                for item in compare_list:
                    if item['欄位'] == col and item['類別'] == str(val):
                        item['不完整資料筆數'] = cnt
                        found = True
                        break
                if not found:
                    compare_list.append({
                        '欄位': col,
                        '類別': str(val),
                        '完整資料筆數': 0,
                        '不完整資料筆數': cnt
                    })

    df_compare = pd.DataFrame(compare_list)

    # 年齡與薪資統計（僅完整資料）
    numeric_stats = []
    if 'Age' in df_complete.columns:
        try:
            ages = pd.to_numeric(df_complete['Age'], errors='coerce').dropna()
            if len(ages) > 0:
                numeric_stats.append({
                    '指標': 'Age（完整資料）',
                    '平均值': round(ages.mean(), 2),
                    '中位數': round(ages.median(), 2),
                    '最小值': ages.min(),
                    '最大值': ages.max()
                })
        except:
            pass

    if 'Salary_Monthly' in df_complete.columns:
        try:
            salaries = df_complete['Salary_Monthly'].astype(str).str.replace(r'[\$,]', '', regex=True)
            salaries = pd.to_numeric(salaries, errors='coerce').dropna()
            if len(salaries) > 0:
                numeric_stats.append({
                    '指標': 'Salary_Monthly（完整資料）',
                    '平均值': round(salaries.mean(), 2),
                    '中位數': round(salaries.median(), 2),
                    '最小值': salaries.min(),
                    '最大值': salaries.max()
                })
        except:
            pass

    df_numeric = pd.DataFrame(numeric_stats)

    return stats_basic, df_missing, df_compare, df_numeric

def create_excel_output(df_complete, df_incomplete, stats_basic, df_missing, df_compare, df_numeric):
    """建立含三個 Sheet 的 Excel 檔案（回傳 bytes）"""
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Sheet 1: 完整資料
        df_complete.to_excel(writer, sheet_name='完整資料', index=False)

        # Sheet 2: 不完整資料
        df_incomplete.to_excel(writer, sheet_name='不完整資料', index=False)

        # Sheet 3: 統計資料
        stats_basic.to_excel(writer, sheet_name='統計資料', index=False, startrow=0)

        worksheet = writer.sheets['統計資料']

        # 缺失統計
        start_row = len(stats_basic) + 3
        worksheet.cell(row=start_row, column=1, value='【各欄位缺失情況統計】')
        df_missing.to_excel(writer, sheet_name='統計資料', index=False, startrow=start_row)

        # 分布比較
        start_row2 = start_row + len(df_missing) + 3
        worksheet.cell(row=start_row2, column=1, value='【關鍵欄位完整 vs 不完整分布比較】')
        df_compare.to_excel(writer, sheet_name='統計資料', index=False, startrow=start_row2)

        # 數值統計
        if len(df_numeric) > 0:
            start_row3 = start_row2 + len(df_compare) + 3
            worksheet.cell(row=start_row3, column=1, value='【完整資料數值欄位統計】')
            df_numeric.to_excel(writer, sheet_name='統計資料', index=False, startrow=start_row3)

    output.seek(0)
    return output

# --------------------------
# Streamlit UI
# --------------------------
uploaded_file = st.file_uploader(
    "請上傳 Excel 檔案（.xlsx 或 .xls）",
    type=['xlsx', 'xls'],
    help="支援 .xlsx 與 .xls 格式"
)

if uploaded_file is not None:
    try:
        # 讀取檔案
        df = pd.read_excel(uploaded_file)
        st.success(f"成功讀取檔案：**{uploaded_file.name}**")
        st.info(f"原始資料共 **{len(df)}** 筆，**{len(df.columns)}** 個欄位")

        # 顯示前幾筆預覽
        with st.expander("📋 原始資料預覽（前 10 筆）", expanded=False):
            st.dataframe(df.head(10), use_container_width=True)

        # 處理按鈕
        if st.button("🚀 開始分析與整理", type="primary", use_container_width=True):
            with st.spinner("正在分析資料完整性，請稍候..."):
                # 分類
                complete_mask = df.apply(is_complete_row, axis=1)
                df_complete = df[complete_mask].copy()
                df_incomplete = df[~complete_mask].copy()

                # 統計
                stats_basic, df_missing, df_compare, df_numeric = generate_stats(
                    df, df_complete, df_incomplete
                )

                # 產生 Excel
                excel_bytes = create_excel_output(
                    df_complete, df_incomplete,
                    stats_basic, df_missing, df_compare, df_numeric
                )

            # 顯示結果摘要
            st.success("分析完成！")

            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("完整資料", f"{len(df_complete)} 筆")
            with col2:
                st.metric("不完整資料", f"{len(df_incomplete)} 筆")
            with col3:
                complete_pct = round(len(df_complete) / len(df) * 100, 2) if len(df) > 0 else 0
                st.metric("完整比例", f"{complete_pct}%")

            # 顯示統計表格
            st.subheader("📈 基本統計")
            st.dataframe(stats_basic, use_container_width=True, hide_index=True)

            with st.expander("各欄位缺失情況"):
                st.dataframe(df_missing, use_container_width=True, hide_index=True)

            with st.expander("關鍵欄位完整 vs 不完整分布"):
                st.dataframe(df_compare, use_container_width=True, hide_index=True)

            if len(df_numeric) > 0:
                with st.expander("完整資料數值欄位統計"):
                    st.dataframe(df_numeric, use_container_width=True, hide_index=True)

            # 下載按鈕
            st.download_button(
                label="⬇️ 下載整理後的 Excel 檔案",
                data=excel_bytes,
                file_name="整理後資料_完整與不完整_含統計.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )

            st.markdown("""
            **檔案說明：**
            - **完整資料**：通過所有檢查的乾淨記錄
            - **不完整資料**：有問題的記錄
            - **統計資料**：基本統計、缺失統計、分布比較、數值統計
            """)

    except Exception as e:
        st.error(f"讀取或處理檔案時發生錯誤：{str(e)}")
        st.exception(e)

else:
    st.info("👆 請先上傳 Excel 檔案")

# 側邊欄說明
with st.sidebar:
    st.header("使用說明")
    st.markdown("""
    1. 上傳你的 Excel 檔案
    2. 點擊「開始分析與整理」
    3. 檢視結果摘要與統計
    4. 下載包含三個工作表的結果檔

    **判斷為「不完整」的條件包括：**
    - 重要欄位為空（Record_ID、姓名、Email、電話、部門、狀態、薪資等）
    - 日期格式錯誤
    - 電話少於 8 位數字
    - Email 格式不正確
    - 年齡不在 0~120
    - 薪資不是正數
    - Internal_Notes 含有測試/重複標記
    - Data_Quality_Label 為 DIRTY
    """)
    st.markdown("---")
    st.caption("Streamlit 版 · 適合放上 GitHub + Streamlit Cloud")