import streamlit as st
import pandas as pd
import io

# ページの設定とタイトル
st.set_page_config(page_title="訪問介護・実績管理システム", layout="wide")
st.title("📊 訪問介護事業所 12ヶ月データ管理・Excel出力アプリ")

st.markdown("""
各月のデータを入力して「データを保存」を押すと、表にリアルタイムで反映されます。
1年分の推移をまとめて確認・修正し、最後にまとめてExcelファイルとしてダウンロードできます。
""")

# 1. データの初期化（ブラウザを開いている間、データを保持する仕組み）
months = [f"{i}月" for i in range(1, 13)]
metrics = ["総売上", "介護給付売上", "支援給付売上", "顧客数", "顧客単価", "時間単価", "サ責数", "総提供時間"]

if "kaigo_data" not in st.session_state:
    # 初期データ（最初はすべて0で作成。サンプルとして一部に数値を入れることも可能）
    init_data = {m: [0] * len(metrics) for m in months}
    df_init = pd.DataFrame(init_data, index=metrics)
    st.session_state.kaigo_data = df_init

# 2画面に分割 (左: データ入力と保存 / 右: 12ヶ月の推移プレビューとExcel出力)
col1, col2 = st.columns([1, 2])

with col1:
    st.header("📝 月次データの入力・更新")
    
    # 入力対象の月を選択
    selected_month = st.selectbox("データを入力・修正する月を選択", months)
    
    # 選択された月の現在の値を取得（上書き編集しやすいようにする）
    current_df = st.session_state.kaigo_data
    
    st.markdown(f"### 📍 {selected_month} の数値を入力")
    kaigo_sales = st.number_input("介護給付売上 (円)", min_value=0, value=int(current_df.loc["介護給付売上", selected_month]), step=10000)
    shien_sales = st.number_input("支援給付売上 (円)", min_value=0, value=int(current_df.loc["支援給付売上", selected_month]), step=10000)
    customer_count = st.number_input("顧客数 (人)", min_value=0, value=int(current_df.loc["顧客数", selected_month]), step=1)
    total_hours = st.number_input("総提供時間 (時間)", min_value=0, value=int(current_df.loc["総提供時間", selected_month]), step=10)
    saseki_count = st.number_input("サ責数 (人)", min_value=0, value=int(current_df.loc["サ責数", selected_month]), step=1)

    # 計算項目のロジック
    total_sales = kaigo_sales + shien_sales
    customer_unit_price = int(total_sales / customer_count) if customer_count > 0 else 0
    hourly_unit_price = int(total_sales / total_hours) if total_hours > 0 else 0

    # データを保存するボタン
    if st.button(f"💾 {selected_month} のデータを表に反映・保存する", type="primary"):
        # セッション状態のデータフレームを更新
        st.session_state.kaigo_data.loc["総売上", selected_month] = total_sales
        st.session_state.kaigo_data.loc["介護給付売上", selected_month] = kaigo_sales
        st.session_state.kaigo_data.loc["支援給付売上", selected_month] = shien_sales
        st.session_state.kaigo_data.loc["顧客数", selected_month] = customer_count
        st.session_state.kaigo_data.loc["顧客単価", selected_month] = customer_unit_price
        st.session_state.kaigo_data.loc["時間単価", selected_month] = hourly_unit_price
        st.session_state.kaigo_data.loc["サ責数", selected_month] = saseki_count
        st.session_state.kaigo_data.loc["総提供時間", selected_month] = total_hours
        st.success(f"🎉 {selected_month} のデータを保存しました！右側の表を確認してください。")

with col2:
    st.header("📊 12ヶ月の推移一覧（プレビュー）")
    
    # 閲覧用に見やすくフォーマット（金額にカンマをつけるなど）
    display_df = st.session_state.kaigo_data.copy()
    
    # 画面で見やすいように数字を整形（文字列に変えるため、Excel出力用とは別に用意）
    for m in months:
        display_df[m] = display_df[m].map(lambda x: f"{int(x):,}")
        
    st.dataframe(display_df, use_container_width=True, height=350)
    
    st.header("📥 全データの一括Excel出力")
    st.markdown("現在表示されている1月〜12月のデータがすべて入ったExcelファイルを出力します。")
    
    # Excelファイルの生成処理
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        # インデックス（管理項目）を含めて書き出し
        export_df = st.session_state.kaigo_data.copy()
        export_df.index.name = "管理項目"
        export_df.to_excel(writer, sheet_name="年間売上管理シート")
    
    # ダウンロードボタン
    st.download_button(
        label="🟢 12ヶ月分のExcelファイルをダウンロード",
        data=buffer.getvalue(),
        file_name="houmon_kaigo_annual_sales.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
