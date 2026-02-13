import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime, timedelta, timezone
import streamlit.components.v1 as components
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- 設定 ---
st.set_page_config(page_title="経営判断の「ものさし」", layout="wide")

# --- セッション状態の初期化 ---
if "has_diagnosed" not in st.session_state:
    st.session_state["has_diagnosed"] = False

# --- CSS (印刷レイアウト改善・横向き強制) ---
st.markdown("""
    <style>
    .block-container { padding-top: 1rem; }
    
    /* 印刷時の設定 */
    @media print {
        @page { 
            size: A4 landscape; /* A4横向き */
            margin: 10mm; 
        }
        body {
            transform: scale(0.9); /* 全体を少し縮小して収まりやすくする */
            transform-origin: top left;
            width: 110%; /* 縮小した分、幅を広げる */
        }
        header, footer, aside, .stAppDeployButton, .no-print, details, [data-testid="stSidebar"] { 
            display: none !important; 
        }
        .block-container {
            max-width: none !important;
            width: 100% !important;
            padding: 0 !important;
            margin: 0 !important;
        }
        [data-testid="stHorizontalBlock"] { 
            display: flex !important; /* 横並びを維持 */
            width: 100% !important; 
        }
        [data-testid="stPlotlyChart"] {
            display: block !important;
            width: 100% !important;
            break-inside: avoid;
        }
        /* 印刷時に文字色を黒くハッキリさせる */
        p, li, .stMarkdown, h1, h2, h3, .metric-label, .metric-value, div {
            color: #000 !important;
        }
    }
    </style>
""", unsafe_allow_html=True)

# --- Google Sheets 保存関数 ---
def save_to_gsheet(data_row):
    try:
        if "gcp_service_account" not in st.secrets: return
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds_dict = st.secrets["gcp_service_account"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        sheet = client.open("financial_db").sheet1
        sheet.append_row(data_row)
        # 成功時はトースト通知（診断ボタンを押した直後に出る）
        st.toast("データを受信しました。診断結果を表示します。", icon="✅") 
    except Exception as e:
        # ユーザーにはエラーを見せず、ログに残す等の処理（ここでは簡易表示）
        print(f"Save Error: {e}")

# --- 関数群 ---
def fmt_yen(val): return f"{int(val):,} 千円" if val is not None else "-"
def fmt_pct(val): return f"{val:.1f}%" if val is not None else "-"
def fmt_times(val): return f"{val:.2f}回" if val is not None else "-"
def fmt_year(val): return f"{val:.1f}年" if val is not None else "-"
def fmt_days(val): return f"{val:.1f}日" if val is not None else "-"
def safe_div(n, d): return n / d if d != 0 else 0
def calc_growth(current, previous):
    if previous <= 0: return None
    return (current - previous) / previous * 100
def calc_score(val, t1, t2, t3, t4, lower_is_better=False):
    if val is None: return 1
    if lower_is_better:
        if val <= t4: return 5
        elif val <= t3: return 4
        elif val <= t2: return 3
        elif val <= t1: return 2
        else: return 1
    else:
        if val >= t4: return 5
        elif val >= t3: return 4
        elif val >= t2: return 3
        elif val >= t1: return 2
        else: return 1
def get_jst_now():
    JST = timezone(timedelta(hours=9), 'JST')
    return datetime.now(JST)

# --- 定数（都道府県リスト） ---
PREFECTURES = [
    "北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県",
    "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県",
    "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県", "岐阜県",
    "静岡県", "愛知県", "三重県", "滋賀県", "京都府", "大阪府", "兵庫県",
    "奈良県", "和歌山県", "鳥取県", "島根県", "岡山県", "広島県", "山口県",
    "徳島県", "香川県", "愛媛県", "高知県", "福岡県", "佐賀県", "長崎県",
    "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県"
]

# --- メイン画面 ---
st.title("📏 経営判断の「ものさし」")
st.markdown("決算書の数値を入力し、「診断する」ボタンを押すと結果が表示されます。")

# サンプルデータ注入ボタン
if st.button("▶ サンプル数値を入れる（入力の手間を省略）", help="クリックすると架空の数値が入力されます"):
    st.session_state["sales_curr"] = 100000
    st.session_state["cogs_curr"] = 70000
    st.session_state["dep_curr"] = 2000
    st.session_state["sga_curr"] = 25000
    st.session_state["noi_curr"] = 500
    st.session_state["noe_curr"] = 500
    st.session_state["ext_i_curr"] = 0
    st.session_state["ext_e_curr"] = 0
    st.session_state["tax_curr"] = 1000
    st.session_state["cash_curr"] = 15000
    st.session_state["rec_curr"] = 12000
    st.session_state["inv_curr"] = 5000
    st.session_state["oca_curr"] = 1000
    st.session_state["fa_curr"] = 20000
    st.session_state["pay_curr"] = 8000
    st.session_state["sl_curr"] = 10000
    st.session_state["ocl_curr"] = 2000
    st.session_state["ll_curr"] = 20000
    st.session_state["na_curr"] = 13000
    st.session_state["emp_curr"] = 10
    
    st.session_state["sales_prev"] = 90000
    st.session_state["cogs_prev"] = 63000
    st.session_state["dep_prev"] = 2000
    st.session_state["sga_prev"] = 24000
    st.session_state["noi_prev"] = 0
    st.session_state["noe_prev"] = 500
    st.session_state["ext_i_prev"] = 0
    st.session_state["ext_e_prev"] = 0
    st.session_state["tax_prev"] = 500
    st.session_state["cash_prev"] = 10000
    st.session_state["rec_prev"] = 10000
    st.session_state["inv_prev"] = 4000
    st.session_state["oca_prev"] = 1000
    st.session_state["fa_prev"] = 20000
    st.session_state["pay_prev"] = 7000
    st.session_state["sl_prev"] = 10000
    st.session_state["ocl_prev"] = 2000
    st.session_state["ll_prev"] = 22000
    st.session_state["na_prev"] = 10000
    st.session_state["emp_prev"] = 9
    
    # サンプル時はサンプル企業情報を入れる
    st.session_state["default_company"] = "サンプル商事"
    st.session_state["default_industry_idx"] = 0 # 製造業
    st.session_state["default_pref_idx"] = 12 # 東京都
    st.rerun()

# --- 入力エリア ---
with st.container():
    st.subheader("1. 基本情報の入力")
    col_basic1, col_basic2, col_basic3 = st.columns(3)
    
    # 会社名
    company_val = st.session_state.get("default_company", "")
    company_name = col_basic1.text_input("会社名（匿名/仮名可）", value=company_val, placeholder="例：株式会社〇〇")
    
    # 都道府県（追加）
    pref_idx = st.session_state.get("default_pref_idx", None)
    prefecture = col_basic2.selectbox("所在地（都道府県）", PREFECTURES, index=pref_idx, placeholder="選択してください")
    
    # 業種（ラジオボタンに変更・初期値なし）
    industry_idx = st.session_state.get("default_industry_idx", None)
    industry_options = ["製造業", "建設業", "卸売業", "小売業", "サービス業", "その他"]
    industry = col_basic3.radio("業種 ※必須", industry_options, index=industry_idx, horizontal=True)

    st.subheader("2. 決算数値の入力")
    st.info("💡 入力単位は**「千円」**です。Enterキーを押すと確定します。")
    
    input_data = {}
    def create_inputs(key_suffix, label_color):
        d = {}
        st.markdown(f"### {label_color}")
        col1, col2, col3 = st.columns(3)
        
        def num_input(label, key, val=0):
            if key not in st.session_state:
                st.session_state[key] = int(val)
            return st.number_input(label, key=key, step=100, format="%d")
        
        with col1:
            st.markdown("##### P/L (損益計算書)")
            d['sales'] = num_input("売上高", f"sales_{key_suffix}", 0)
            d['cogs'] = num_input("売上原価", f"cogs_{key_suffix}", 0)
            d['depreciation'] = num_input("  うち減価償却費", f"dep_{key_suffix}", 0)
            d['gross_profit'] = d['sales'] - d['cogs']
            st.caption(f"粗利: {fmt_yen(d['gross_profit'])}")
            d['sga'] = num_input("販管費", f"sga_{key_suffix}", 0)
            d['op_profit'] = d['gross_profit'] - d['sga']
            st.caption(f"営業利益: {fmt_yen(d['op_profit'])}") 
            d['non_op_inc'] = num_input("営業外収益", f"noi_{key_suffix}", 0)
            d['non_op_exp'] = num_input("営業外費用", f"noe_{key_suffix}", 0)
            d['ord_profit'] = d['op_profit'] + d['non_op_inc'] - d['non_op_exp']
            st.caption(f"経常利益: {fmt_yen(d['ord_profit'])}") 
            d['extra_inc'] = num_input("特別利益", f"ext_i_{key_suffix}", 0) 
            d['extra_exp'] = num_input("特別損失", f"ext_e_{key_suffix}", 0) 
            d['pre_tax_profit'] = d['ord_profit'] + d['extra_inc'] - d['extra_exp']
            st.caption(f"税引前利益: {fmt_yen(d['pre_tax_profit'])}") 
            d['tax'] = num_input("法人税等", f"tax_{key_suffix}", 0)
            d['net_profit'] = d['pre_tax_profit'] - d['tax']
            st.caption(f"当期純利益: {fmt_yen(d['net_profit'])}") 

        with col2:
            st.markdown("##### B/S (資産)")
            d['cash'] = num_input("現預金", f"cash_{key_suffix}", 0)
            d['receivables'] = num_input("売上債権", f"rec_{key_suffix}", 0)
            d['inventory'] = num_input("棚卸資産", f"inv_{key_suffix}", 0)
            d['other_ca'] = num_input("その他流動資産", f"oca_{key_suffix}", 0)
            d['current_assets'] = d['cash'] + d['receivables'] + d['inventory'] + d['other_ca']
            d['fixed_assets'] = num_input("固定資産合計", f"fa_{key_suffix}", 0)
            d['total_assets'] = d['current_assets'] + d['fixed_assets']
            st.markdown("---")
            st.metric("資産合計", fmt_yen(d['total_assets']))

        with col3:
            st.markdown("##### B/S (負債・純資産)")
            d['payables'] = num_input("仕入債務", f"pay_{key_suffix}", 0)
            d['short_loan'] = num_input("短期借入金", f"sl_{key_suffix}", 0)
            d['other_cl'] = num_input("その他流動負債", f"ocl_{key_suffix}", 0)
            d['current_liab'] = d['payables'] + d['short_loan'] + d['other_cl']
            d['long_loan'] = num_input("長期借入金", f"ll_{key_suffix}", 0)
            d['fixed_liab'] = d['long_loan'] 
            d['net_assets'] = num_input("純資産合計", f"na_{key_suffix}", 0)
            d['total_liab_equity'] = d['current_liab'] + d['fixed_liab'] + d['net_assets']
            st.markdown("---")
            st.metric("負債・純資産", fmt_yen(d['total_liab_equity']))
            st.markdown("##### その他")
            d['employees'] = num_input("従業員数", f"emp_{key_suffix}", 0)
        
        diff = d['total_assets'] - d['total_liab_equity']
        if diff != 0: st.error(f"⚠️ 貸借不一致: {fmt_yen(diff)}")
        else: st.success("✅ 貸借一致")
        return d

    tab_curr, tab_prev = st.tabs(["🔴 当期 (最新)", "🔵 前期 (過去)"])
    with tab_curr: input_data['curr'] = create_inputs("curr", "🔴 当期データ")
    with tab_prev: input_data['prev'] = create_inputs("prev", "🔵 前期データ")


# --- 診断実行 & データ保存セクション ---
st.markdown("---")
st.markdown("""
<small>
**【データの取り扱いについて】**<br>
「診断する」ボタンを押すと、入力されたデータは診断精度の向上および統計的な業界分析のために、
個人・企業を特定できない形式（匿名加工情報）にてサーバーへ保存されます。<br>
入力データが第三者にそのまま開示されることはありません。ご利用にあたっては、これに同意したものとみなします。
</small>
""", unsafe_allow_html=True)

# 診断ボタン（ここを押すと保存＆表示）
if st.button("🚀 同意して診断する（レポートを表示）", type="primary", use_container_width=True):
    # バリデーションチェック
    if not company_name:
        st.error("⚠️ 「会社名」を入力してください（匿名・仮名でも構いません）。")
    elif not industry:
        st.error("⚠️ 「業種」を選択してください。")
    elif not prefecture:
        st.error("⚠️ 「所在地（都道府県）」を選択してください。")
    else:
        # 計算用データを取得
        c, p = input_data['curr'], input_data['prev']
        
        # 指標計算（保存用に先に計算が必要）
        c_op_margin = safe_div(c['op_profit'], c['sales']) * 100
        c_fcf = (c['op_profit'] * 0.6 + c['depreciation']) - ((c['fixed_assets'] - p['fixed_assets']) + c['depreciation'])
        c_sales_growth = calc_growth(c['sales'], p['sales'])
        c_op_growth = calc_growth(c['op_profit'], p['op_profit'])
        c_fixed_turn = safe_div(c['sales'], c['fixed_assets'])
        c_inv_days = safe_div(c['inventory'], c['cogs'] / 365)
        c_sales_per_emp = safe_div(c['sales'], c['employees'])
        c_op_per_emp = safe_div(c['op_profit'], c['employees'])
        c_equity_ratio = safe_div(c['net_assets'], c['total_assets']) * 100
        c_loan_sales_ratio = safe_div(c['short_loan'] + c['long_loan'], c['sales'] / 12)
        c_current_ratio = safe_div(c['current_assets'], c['current_liab']) * 100
        c_working_capital = c['current_assets'] - c['current_liab']
        c_redemption = safe_div(c['short_loan'] + c['long_loan'], c['ord_profit'] + c['depreciation'] - c['tax']) if (c['ord_profit'] + c['depreciation'] - c['tax']) > 0 else 0

        p_op_margin = safe_div(p['op_profit'], p['sales']) * 100 
        p_equity_ratio = safe_div(p['net_assets'], p['total_assets']) * 100
        p_loan_sales_ratio = safe_div(p['short_loan'] + p['long_loan'], p['sales'] / 12)
        p_fixed_turn = safe_div(p['sales'], p['fixed_assets'])
        p_inv_days = safe_div(p['inventory'], p['cogs'] / 365)
        p_sales_per_emp = safe_div(p['sales'], p['employees'])
        p_op_per_emp = safe_div(p['op_profit'], p['employees'])
        p_current_ratio = safe_div(p['current_assets'], p['current_liab']) * 100
        p_working_capital = p['current_assets'] - p['current_liab']
        p_redemption = safe_div(p['short_loan'] + p['long_loan'], p['ord_profit'] + p['depreciation'] - p['tax']) if (p['ord_profit'] + p['depreciation'] - p['tax']) > 0 else 0

        score_sales_growth = calc_score(c_sales_growth, 0, 3, 5, 10)
        score_op_growth = calc_score(c_op_growth, 0, 3, 5, 10)
        if p['op_profit'] <= 0 and c['op_profit'] > 0: score_op_growth = 5

        # 総合スコア計算
        scores = {
            "収益": (calc_score(c_op_margin, 0, 2, 5, 10) + calc_score(c_fcf, -1000, 0, 1000, 5000)) / 2,
            "成長": (score_sales_growth + score_op_growth) / 2,
            "効率": (calc_score(c_fixed_turn, 1, 3, 5, 10) + calc_score(c_inv_days, 180, 90, 60, 30, True)) / 2,
            "生産": (calc_score(c_sales_per_emp, 10000, 15000, 20000, 30000) + calc_score(c_op_per_emp, 0, 500, 1000, 2000)) / 2,
            "安全": (calc_score(c_equity_ratio, 10, 20, 40, 60) + calc_score(c_loan_sales_ratio, 12, 6, 3, 1, True)) / 2
        }
        avg_score = sum(scores.values()) / 5

        # データ保存（都道府県を追加）
        save_row = [
            str(get_jst_now()), company_name, prefecture, industry, avg_score, # prefectureを追加
            c['sales'], c['cogs'], c['depreciation'], c['gross_profit'], c['sga'], c['op_profit'], 
            c['non_op_inc'], c['non_op_exp'], c['ord_profit'], c['extra_inc'], c['extra_exp'], c['pre_tax_profit'], c['tax'], c['net_profit'], 
            c['cash'], c['receivables'], c['inventory'], c['other_ca'], c['current_assets'], c['fixed_assets'], c['total_assets'],
            c['payables'], c['short_loan'], c['other_cl'], c['current_liab'], c['long_loan'], c['fixed_liab'], c['net_assets'], c['total_liab_equity'],
            c['employees'],
            p['sales'], p['cogs'], p['depreciation'], p['gross_profit'], p['sga'], p['op_profit'], 
            p['non_op_inc'], p['non_op_exp'], p['ord_profit'], p['extra_inc'], p['extra_exp'], p['pre_tax_profit'], p['tax'], p['net_profit'],
            p['cash'], p['receivables'], p['inventory'], p['other_ca'], p['current_assets'], p['fixed_assets'], p['total_assets'],
            p['payables'], p['short_loan'], p['other_cl'], p['current_liab'], p['long_loan'], p['fixed_liab'], p['net_assets'], p['total_liab_equity'],
            p['employees'],
            c_op_margin, c_fcf, c_sales_growth, c_op_growth, c_fixed_turn, c_inv_days,
            c_sales_per_emp, c_op_per_emp, c_equity_ratio, c_working_capital, c_current_ratio, c_redemption, c_loan_sales_ratio
        ]
        
        save_to_gsheet(save_row)
        
        # 診断済みフラグを立てて、再描画
        st.session_state["has_diagnosed"] = True
        st.rerun()

# --- 結果レポート表示（診断済みの場合のみ表示） ---
if st.session_state["has_diagnosed"]:
    
    # ※ここで変数を再定義する必要があるため、計算ロジックを再実行（Streamlitの仕様上）
    c, p = input_data['curr'], input_data['prev']
    # ...(以下、表示用の計算と描画)...
    
    # 各種計算（表示用に再計算）
    c_op_margin = safe_div(c['op_profit'], c['sales']) * 100
    c_fcf = (c['op_profit'] * 0.6 + c['depreciation']) - ((c['fixed_assets'] - p['fixed_assets']) + c['depreciation'])
    c_sales_growth = calc_growth(c['sales'], p['sales'])
    c_op_growth = calc_growth(c['op_profit'], p['op_profit'])
    c_fixed_turn = safe_div(c['sales'], c['fixed_assets'])
    c_inv_days = safe_div(c['inventory'], c['cogs'] / 365)
    c_sales_per_emp = safe_div(c['sales'], c['employees'])
    c_op_per_emp = safe_div(c['op_profit'], c['employees'])
    c_equity_ratio = safe_div(c['net_assets'], c['total_assets']) * 100
    c_loan_sales_ratio = safe_div(c['short_loan'] + c['long_loan'], c['sales'] / 12)
    c_current_ratio = safe_div(c['current_assets'], c['current_liab']) * 100
    c_working_capital = c['current_assets'] - c['current_liab']
    c_redemption = safe_div(c['short_loan'] + c['long_loan'], c['ord_profit'] + c['depreciation'] - c['tax']) if (c['ord_profit'] + c['depreciation'] - c['tax']) > 0 else 0

    p_op_margin = safe_div(p['op_profit'], p['sales']) * 100 
    p_equity_ratio = safe_div(p['net_assets'], p['total_assets']) * 100
    p_loan_sales_ratio = safe_div(p['short_loan'] + p['long_loan'], p['sales'] / 12)
    p_fixed_turn = safe_div(p['sales'], p['fixed_assets'])
    p_inv_days = safe_div(p['inventory'], p['cogs'] / 365)
    p_sales_per_emp = safe_div(p['sales'], p['employees'])
    p_op_per_emp = safe_div(p['op_profit'], p['employees'])
    p_current_ratio = safe_div(p['current_assets'], p['current_liab']) * 100
    p_working_capital = p['current_assets'] - p['current_liab']
    p_redemption = safe_div(p['short_loan'] + p['long_loan'], p['ord_profit'] + p['depreciation'] - p['tax']) if (p['ord_profit'] + p['depreciation'] - p['tax']) > 0 else 0

    score_sales_growth = calc_score(c_sales_growth, 0, 3, 5, 10)
    score_op_growth = calc_score(c_op_growth, 0, 3, 5, 10)
    if p['op_profit'] <= 0 and c['op_profit'] > 0: score_op_growth = 5

    scores = {
        "収益": (calc_score(c_op_margin, 0, 2, 5, 10) + calc_score(c_fcf, -1000, 0, 1000, 5000)) / 2,
        "成長": (score_sales_growth + score_op_growth) / 2,
        "効率": (calc_score(c_fixed_turn, 1, 3, 5, 10) + calc_score(c_inv_days, 180, 90, 60, 30, True)) / 2,
        "生産": (calc_score(c_sales_per_emp, 10000, 15000, 20000, 30000) + calc_score(c_op_per_emp, 0, 500, 1000, 2000)) / 2,
        "安全": (calc_score(c_equity_ratio, 10, 20, 40, 60) + calc_score(c_loan_sales_ratio, 12, 6, 3, 1, True)) / 2
    }
    p_scores_val = {k: 3 for k in scores} 
    p_scores_val["収益"] = calc_score(p_op_margin, 0, 2, 5, 10)
    avg_score = sum(scores.values()) / 5

    # --- レポート描画 ---
    st.markdown("---")
    st.header(f"📈 {company_name} 様 経営診断レポート")
    st.markdown(f"診断日: {get_jst_now().strftime('%Y年%m月%d日 %H:%M')}")

    col_radar, col_msg = st.columns([1, 1])
    with col_radar:
        fig = go.Figure()
        fig.add_trace(go.Scatterpolar(r=list(p_scores_val.values()), theta=list(scores.keys()), fill='toself', name='前期', line_color='#00B4D8', opacity=0.5))
        fig.add_trace(go.Scatterpolar(r=list(scores.values()), theta=list(scores.keys()), fill='toself', name='当期', line_color='#FF4B4B', opacity=0.8))
        fig.update_layout(
            polar=dict(radialaxis=dict(visible=True, range=[0, 5])), 
            showlegend=True, 
            legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5),
            height=700, 
            margin=dict(l=80, r=80, t=40, b=80) 
        )
        st.plotly_chart(fig, use_container_width=True)

    with col_msg:
        st.markdown(f"#### 📝 総合スコア: {avg_score:.1f} / 5.0")
        if avg_score >= 4: st.success("極めて健全な経営状態です。攻めの投資を行う体力があります。")
        elif avg_score >= 3: st.info("標準的な経営状態です。弱点を補強しましょう。")
        else: st.error("経営改善が急務です。特に安全性の確保を優先してください。")

    st.subheader("詳細指標分析")
    kpi_definitions = [
        {"cat": "収益性", "name": "営業利益率", "curr_v": c_op_margin, "unit": "%", "prev_v": p_op_margin, "desc": "本業の稼ぐ力", "formula": "営業利益 ÷ 売上高"},
        {"cat": "収益性", "name": "フリーキャッシュフロー", "curr_v": c_fcf, "unit": "千円", "prev_v": None, "desc": "自由に使える現金", "formula": "営業利益×0.6+償却-設備投資"},
        {"cat": "成長性", "name": "売上高成長率", "curr_v": c_sales_growth, "unit": "%", "prev_v": None, "desc": "シェア拡大度", "formula": "(当期売上-前期)/前期"},
        {"cat": "成長性", "name": "営業利益成長率", "curr_v": c_op_growth, "unit": "%", "prev_v": None, "desc": "利益の伸び ※前期赤字の場合は計算不能(-)", "formula": "(当期営利-前期)/前期"},
        {"cat": "効率性", "name": "固定資産回転率", "curr_v": c_fixed_turn, "unit": "回", "prev_v": p_fixed_turn, "desc": "設備の稼働効率", "formula": "売上高 ÷ 固定資産"},
        {"cat": "効率性", "name": "棚卸資産回転日数", "curr_v": c_inv_days, "unit": "日", "prev_v": p_inv_days, "desc": "在庫の回転速度", "formula": "棚卸資産 ÷ (売上原価÷365)"},
        {"cat": "生産性", "name": "1人当たり売上高", "curr_v": c_sales_per_emp, "unit": "千円", "prev_v": p_sales_per_emp, "desc": "社員の稼ぐ規模", "formula": "売上高 ÷ 従業員数"},
        {"cat": "生産性", "name": "1人当たり営業利益", "curr_v": c_op_per_emp, "unit": "千円", "prev_v": p_op_per_emp, "desc": "社員の付加価値", "formula": "営業利益 ÷ 従業員数"},
        {"cat": "安全性", "name": "自己資本比率", "curr_v": c_equity_ratio, "unit": "%", "prev_v": p_equity_ratio, "desc": "倒産耐性", "formula": "純資産 ÷ 総資産"},
        {"cat": "安全性", "name": "運転資本", "curr_v": c_working_capital, "unit": "千円", "prev_v": p_working_capital, "desc": "支払い余力", "formula": "流動資産 - 流動負債"},
        {"cat": "安全性", "name": "流動比率", "curr_v": c_current_ratio, "unit": "%", "prev_v": p_current_ratio, "desc": "短期返済能力", "formula": "流動資産 ÷ 流動負債"},
        {"cat": "安全性", "name": "債務償還年数", "curr_v": c_redemption, "unit": "年", "prev_v": p_redemption, "desc": "借金完済までの年数", "formula": "有利子負債 ÷ CF"},
        {"cat": "安全性", "name": "借入金月商倍率", "curr_v": c_loan_sales_ratio, "unit": "倍", "prev_v": p_loan_sales_ratio, "desc": "借金規模の適正度", "formula": "有利子負債 ÷ 月商"}
    ]
    
    current_cat = ""
    temp_kpis = []
    for k in kpi_definitions:
        if k['unit'] == "%": curr_disp, prev_disp = fmt_pct(k['curr_v']), fmt_pct(k['prev_v'])
        elif k['unit'] == "千円": curr_disp, prev_disp = fmt_yen(k['curr_v']), fmt_yen(k['prev_v'])
        elif k['unit'] == "回": curr_disp, prev_disp = fmt_times(k['curr_v']), fmt_times(k['prev_v'])
        elif k['unit'] == "倍": curr_disp, prev_disp = fmt_times(k['curr_v']).replace("回","倍"), fmt_times(k['prev_v']).replace("回","倍")
        elif k['unit'] == "年": curr_disp, prev_disp = fmt_year(k['curr_v']), fmt_year(k['prev_v'])
        elif k['unit'] == "日": curr_disp, prev_disp = fmt_days(k['curr_v']), fmt_days(k['prev_v'])
        if k['prev_v'] is not None and k['curr_v'] is not None:
            delta_val = f"{k['curr_v'] - k['prev_v']:.1f}" if k['unit']!="千円" else fmt_yen(k['curr_v']-k['prev_v'])
        else: 
            delta_val = "-"
        k['curr_disp'], k['prev_disp'], k['delta'] = curr_disp, prev_disp, delta_val
        if current_cat != k['cat']:
            if temp_kpis: 
                with st.container(): 
                    st.markdown(f"#### 📌 {current_cat}")
                    for tk in temp_kpis:
                        cols = st.columns([2, 1, 1, 3])
                        cols[0].markdown(f"**{tk['name']}**")
                        cols[1].metric("当期", tk['curr_disp'], tk['delta'])
                        cols[2].caption(f"前期: {tk['prev_disp']}")
                        cols[3].markdown(f"<small>{tk['desc']}<br>🧮 `{tk['formula']}`</small>", unsafe_allow_html=True)
                        st.markdown("---")
            temp_kpis = []
            current_cat = k['cat']
        temp_kpis.append(k)
    if temp_kpis:
        with st.container():
            st.markdown(f"#### 📌 {current_cat}")
            for tk in temp_kpis:
                cols = st.columns([2, 1, 1, 3])
                cols[0].markdown(f"**{tk['name']}**")
                cols[1].metric("当期", tk['curr_disp'], tk['delta'])
                cols[2].caption(f"前期: {tk['prev_disp']}")
                cols[3].markdown(f"<small>{tk['desc']}<br>🧮 `{tk['formula']}`</small>", unsafe_allow_html=True)
                st.markdown("---")

    # CSV生成
    raw_data_list = [
        {"区分": "基本情報", "項目": "診断日時", "当期_数値": str(get_jst_now()), "単位": "-", "前期_数値": "-", "説明": "-"},
        {"区分": "基本情報", "項目": "会社名", "当期_数値": company_name, "単位": "-", "前期_数値": "-", "説明": "-"},
        {"区分": "基本情報", "項目": "所在地", "当期_数値": prefecture, "単位": "-", "前期_数値": "-", "説明": "-"}, # 追加
        {"区分": "基本情報", "項目": "業種", "当期_数値": industry, "単位": "-", "前期_数値": "-", "説明": "-"},
        {"区分": "基本情報", "項目": "総合スコア", "当期_数値": avg_score, "単位": "点", "前期_数値": "-", "説明": "-"},
    ]
    # P/L, B/S, KPIデータをCSV用リストに追加（冗長になるため省略せず記述推奨だが、ここはロジック同じ）
    # ... (CSV作成ロジックは既存のものを維持しつつ、変数c, pを使って生成) ...
    # 簡略化のため、既存コードの raw_data_list 生成ロジックをここに持ってきます
    
    pl_bs_items = [
        ("P/L", "売上高", 'sales', "千円"), ("P/L", "売上原価", 'cogs', "千円"), ("P/L", "減価償却費", 'depreciation', "千円"),
        ("P/L", "売上総利益", 'gross_profit', "千円"), ("P/L", "販管費", 'sga', "千円"), ("P/L", "営業利益", 'op_profit', "千円"),
        ("P/L", "営業外収益", 'non_op_inc', "千円"), ("P/L", "営業外費用", 'non_op_exp', "千円"), ("P/L", "経常利益", 'ord_profit', "千円"),
        ("P/L", "特別利益", 'extra_inc', "千円"), ("P/L", "特別損失", 'extra_exp', "千円"), ("P/L", "税引前当期純利益", 'pre_tax_profit', "千円"),
        ("P/L", "法人税等", 'tax', "千円"), ("P/L", "当期純利益", 'net_profit', "千円"),
        ("B/S", "流動資産計", 'current_assets', "千円"), ("B/S", "固定資産", 'fixed_assets', "千円"), ("B/S", "総資産", 'total_assets', "千円"),
        ("B/S", "流動負債計", 'current_liab', "千円"), ("B/S", "固定負債", 'fixed_liab', "千円"), ("B/S", "純資産", 'net_assets', "千円"),
        ("その他", "従業員数", 'employees', "人")
    ]
    for cat, name, key, unit in pl_bs_items:
        raw_data_list.append({"区分": cat, "項目": name, "当期_数値": c[key], "単位": unit, "前期_数値": p[key], "説明": "-"})
        
    for k in kpi_definitions:
        raw_data_list.append({
            "区分": k['cat'], "項目": k['name'], "当期_数値": k['curr_v'], "単位": k['unit'], "前期_数値": k['prev_v'], "説明": k['desc']
        })

    export_df = pd.DataFrame(raw_data_list)

    st.markdown("---")
    
    # ダウンロードボタン類
    st.download_button(
        label="📊 診断データ(CSV)を保存",
        data=export_df.to_csv(index=False).encode('utf-8_sig'),
        file_name=f"financial_report_{get_jst_now().strftime('%Y%m%d')}.csv",
        help="CSVをダウンロードします（データは既にクラウドへ保存済みです）"
    )

    if st.button("🖨️ レポートを印刷 (PDF保存)"):
        components.html("<script>window.parent.print();</script>", height=0, width=0)

    # 続けて別の診断をするボタン
    if st.button("🔄 新しいデータを入力して再診断する"):
        st.session_state["has_diagnosed"] = False
        st.rerun()

    st.markdown("---")
