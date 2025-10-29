import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import io

st.title("メニュー名上書きツール")

uploaded_menu_file = st.file_uploader("メニュー名変更依頼ファイルをアップロード（パスワード解除して）", type="xlsx")
uploaded_code_file = st.file_uploader("媒体コード発番依頼ファイルをアップロード（パスワード解除して）", type="xlsx")

if uploaded_menu_file and uploaded_code_file:
    menu_wb = load_workbook(uploaded_menu_file)
    code_wb = load_workbook(uploaded_code_file)

    st.subheader("シート名の確認")
    menu_sheet_name = st.selectbox("メニュー名変更依頼ファイルのシートを選択", menu_wb.sheetnames)
    code_sheet_name = st.selectbox("媒体コード発番依頼ファイルのシートを選択", code_wb.sheetnames)

    menu_ws = menu_wb[menu_sheet_name]
    code_ws = code_wb[code_sheet_name]

    # データ抽出
    menu_data = []
    for row in menu_ws.iter_rows(min_row=2, values_only=True):
        media_code = row[7]  # H列
        menu_name = row[8]   # I列
        point = row[10]      # K列
        media_name = row[11] # L列
        menu_data.append((media_code, menu_name, point, media_name))

    # プレビュー用データ
    preview_rows = []
    for row in code_ws.iter_rows(min_row=2):
        target_code = row[1].value  # B列
        for m_code, m_name, m_point, m_media_name in menu_data:
            if target_code == m_code:
                preview_rows.append({
                    "媒体コード": target_code,
                    "旧メニュー名": row[2].value,
                    "新メニュー名": m_name,
                    "旧ポイント": row[4].value,
                    "新ポイント": m_point,
                    "旧媒体名": row[5].value,
                    "新媒体名": m_media_name
                })
                break

    if preview_rows:
        st.subheader("更新プレビュー")
        st.dataframe(pd.DataFrame(preview_rows))

    if st.button("上書き実行"):
        rose_fill = PatternFill(start_color="FF66CC", end_color="FF66CC", fill_type="solid")
        updated_count = 0
        for row in code_ws.iter_rows(min_row=2):
            target_code = row[1].value  # B列
            for m_code, m_name, m_point, m_media_name in menu_data:
                if target_code == m_code:
                    row[2].value = m_name       # C列
                    row[2].fill = rose_fill
                    row[4].value = m_point      # E列
                    row[4].fill = rose_fill
                    row[5].value = m_media_name # F列
                    row[5].fill = rose_fill
                    updated_count += 1
                    break

        st.success(f"{updated_count} 件のセルを更新✨更新セルはピンクです🦩エクスポートして確認してください！")

        output = io.BytesIO()
        code_wb.save(output)
        st.download_button(
            label="更新済みファイルをダウンロード",
            data=output.getvalue(),
            file_name="更新済み_媒体コード発番依頼.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )