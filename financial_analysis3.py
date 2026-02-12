import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime
import streamlit.components.v1 as components
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- 設定 ---
st.set_page_config(page_title="経営判断の「ものさし」", layout="wide")

# --- CSS (印刷レイアウト・最大化版) ---
st.markdown("""
    <style>
    .block-container { padding-top: 2rem; }
    @media print {
        header, footer, aside, .stAppDeployButton, .no-print, details, [data-testid="stSidebar"] { 
            display: none !important; 
        }
        @page { 
            margin: 5mm; 
            size: A4 portrait; 
        }
        .block-container {
            max-width: none !important;
            width: 100% !important;
            padding: 0 !important;
            margin: 0 !important;
        }
        [data-testid="stHorizontalBlock"] { 
            display: block !important; 
            width: 100% !important; 
        }
        [data-testid="stPlotlyChart"] {
            display: block !important;
            width: 100% !important;
            height: 700px !important; 
            page-break-inside: avoid;
            overflow: visible !important;
            margin-bottom: 0px !important;
        }
        .js-plotly-plot, .plot-container {
            height: 700px !important; 
            width: 100% !important;
        }
        p, li, .stMarkdown, h1, h2, h3, .metric-label, .metric-value {
            color: #000 !important;
        }
    }
    </style>
""", unsafe_allow_html=True)

# --- Google Sheets 保存関数 ---
def save_to_gsheet(data_dict):
    try:
        if "gcp_service_account" not in st.secrets: return
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds_dict = st.secrets["gcp_service_account"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        sheet = client.open("financial_db").sheet1
        row = [
            str(datetime.now()),
            data_dict.get("company_name", "-"),
            data_dict.get("industry", "-"),
            data_dict.get("curr_sales", 0),
            data_dict.get("curr_op_profit", 0),
            data_dict.get("total_score", 0),
            data_dict.get("loan_sales_ratio", 0)
        ]
        sheet.append_row(row)
    except Exception as e:
        print(f"Data Save Error: {e}")

# --- 関数群 (1行記述で圧縮) ---
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

# --- メイン画面 ---
st.title("📏 経営判断の「ものさし」 by かんぎアドバイザーズ")
st.markdown("数値を入れると**リアルタイム**で診断結果が変化します。")

# 入力エリア
with st.expander("📝 データの入力・修正（クリックで開閉）", expanded=True):
    st.info("💡 入力単位は**「千円」**です。Enterキーを押すと即座に反映されます。")
    col_basic1, col_basic2 = st.columns(2)
    company_name = col_basic1.text_input("会社名", "サンプル商事")
    industry = col_basic2.selectbox("業種", ["製造業", "建設業", "卸売業", "小売業", "サービス業", "その他"])
    input_data = {}
    
    def create_inputs(key_suffix, label_color):
        d = {}
        st.markdown(f"### {label_color}")
        col1, col2, col3 = st.columns(3)
        def num_input(label, key, val=0):
            return st.number_input(label, value=int(val), step=100, format="%d", key=key)
        with col1:
            st.markdown("##### P/L (損益計算書)")
            d['sales'] = num_input("売上高", f"sales_{key_suffix}", 100000)
            d['cogs'] = num_input("売上原価", f"cogs_{key_suffix}", 70000)
            d['depreciation'] = num_input("  うち減価償却費", f"dep_{key_suffix}", 2000)
            d['gross_profit'] = d['sales'] - d['cogs']
            st.caption(f"粗利: {fmt_yen(d['gross_profit'])}")
            d['sga'] = num_input("販管費", f"sga_{key_suffix}", 25000)
            d['op_profit'] = d['gross_profit'] - d['sga']
            st.caption(f"営利: {fmt_yen(d['op_profit'])}") 
            d['non_op_inc'] = num_input("営業外収益", f"noi_{key_suffix}", 0)
            d['non_op_exp'] = num_input("営業外費用", f"noe_{key_suffix}", 500)
            d['ord_profit'] = d['op_profit'] + d['non_op_inc'] - d['non_op_exp']
            d['tax'] = num_input("法人税等", f"tax_{key_suffix}", 500)
        with col2:
            st.markdown("##### B/S (資産)")
            d['cash'] = num_input("現預金", f"cash_{key_suffix}", 15000)
            d['receivables'] = num_input("売上債権", f"rec_{key_suffix}", 12000)
            d['inventory'] = num_input("棚卸資産", f"inv_{key_suffix}", 5000)
            d['other_ca'] = num_input("その他流動資産", f"oca_{key_suffix}", 1000)
            d['current_assets'] = d['cash'] + d['receivables'] + d['inventory'] + d['other_ca']
            d['fixed_assets'] = num_input("固定資産合計", f"fa_{key_suffix}", 20000)
            d['total_assets'] = d['current_assets'] + d['fixed_assets']
            st.markdown("---")
            st.metric("資産合計", fmt_yen(d['total_assets']))
        with col3:
            st.markdown("##### B/S (負債・純資産)")
            d['payables'] = num_input("仕入債務", f"pay_{key_suffix}", 8000)
            d['short_loan'] = num_input("短期借入金", f"sl_{key_suffix}", 10000)
            d['other_cl'] = num_input("その他流動負債", f"ocl_{key_suffix}", 2000)
            d['current_liab'] = d['payables'] + d['short_loan'] + d['other_cl']
            d['long_loan'] = num_input("長期借入金", f"ll_{key_suffix}", 20000)
            d['fixed_liab'] = d['long_loan'] 
            d['net_assets'] = num_input("純資産合計", f"na_{key_suffix}", 13000)
            d['total_liab_equity'] = d['current_liab'] + d['fixed_liab'] + d['net_assets']
            st.markdown("---")
            st.metric("負債・純資産", fmt_yen(d['total_liab_equity']))
            st.markdown("##### その他")
            d['employees'] = st.number_input(f"従業員数", value=10, step=1, format="%d", key=f"emp_{key_suffix}")
        diff = d['total_assets'] - d['total_liab_equity']
        if diff != 0: st.error(f"⚠️ 貸借不一致: {fmt_yen(diff)}")
        else: st.success("✅ 貸借一致")
        return d

    tab_curr, tab_prev = st.tabs(["🔴 当期 (最新)", "🔵 前期 (過去)"])
    with tab_curr: input_data['curr'] = create_inputs("curr", "🔴 当期データ")
    with tab_prev: input_data['prev'] = create_inputs("prev", "🔵 前期データ")

# --- 計算ロジック ---
c, p = input_data['curr'], input_data['prev']
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

# --- レポート表示 ---
st.markdown("---")
st.header(f"📈 {company_name} 様 経営診断レポート")
st.markdown(f"診断日: {datetime.now().strftime('%Y年%m月%d日')}")

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
    avg_score = sum(scores.values()) / 5
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

# CSVデータ作成
raw_data_list = [
    {"区分": "財務データ(P/L)", "項目": "売上高", "当期_数値": c['sales'], "単位": "千円", "前期_数値": p['sales'], "説明": "-"},
    {"区分": "財務データ(P/L)", "項目": "売上原価", "当期_数値": c['cogs'], "単位": "千円", "前期_数値": p['cogs'], "説明": "-"},
    {"区分": "財務データ(P/L)", "項目": "販管費", "当期_数値": c['sga'], "単位": "千円", "前期_数値": p['sga'], "説明": "-"},
    {"区分": "財務データ(P/L)", "項目": "営業利益", "当期_数値": c['op_profit'], "単位": "千円", "前期_数値": p['op_profit'], "説明": "-"},
    {"区分": "財務データ(P/L)", "項目": "経常利益", "当期_数値": c['ord_profit'], "単位": "千円", "前期_数値": p['ord_profit'], "説明": "-"},
    {"区分": "財務データ(B/S)", "項目": "流動資産", "当期_数値": c['current_assets'], "単位": "千円", "前期_数値": p['current_assets'], "説明": "-"},
    {"区分": "財務データ(B/S)", "項目": "固定資産", "当期_数値": c['fixed_assets'], "単位": "千円", "前期_数値": p['fixed_assets'], "説明": "-"},
    {"区分": "財務データ(B/S)", "項目": "流動負債", "当期_数値": c['current_liab'], "単位": "千円", "前期_数値": p['current_liab'], "説明": "-"},
    {"区分": "財務データ(B/S)", "項目": "固定負債", "当期_数値": c['fixed_liab'], "単位": "千円", "前期_数値": p['fixed_liab'], "説明": "-"},
    {"区分": "財務データ(B/S)", "項目": "純資産", "当期_数値": c['net_assets'], "単位": "千円", "前期_数値": p['net_assets'], "説明": "-"},
    {"区分": "財務データ(B/S)", "項目": "有利子負債合計", "当期_数値": c['short_loan'] + c['long_loan'], "単位": "千円", "前期_数値": p['short_loan'] + p['long_loan'], "説明": "-"},
]
indicator_list = []
for k in kpi_definitions:
    indicator_list.append({
        "区分": k['cat'], "項目": k['name'], "当期_数値": k['curr_v'], "単位": k['unit'], "前期_数値": k['prev_v'], "説明": k['desc']
    })
export_df = pd.DataFrame(raw_data_list + indicator_list)
save_data = {
    "company_name": company_name, "industry": industry,
    "curr_sales": c['sales'], "curr_op_profit": c['op_profit'],
    "loan_sales_ratio": c_loan_sales_ratio, "total_score": avg_score
}
st.markdown("---")
st.download_button(
    label="📊 診断データ(CSV)を保存",
    data=export_df.to_csv(index=False).encode('utf-8_sig'),
    file_name=f"financial_report_{datetime.now().strftime('%Y%m%d')}.csv",
    on_click=save_to_gsheet,
    args=(save_data,),
    help="CSVをダウンロードし、結果を保存します"
)
if st.button("🖨️ レポートを印刷 (PDF保存)"):
    try:
        save_to_gsheet(save_data)
    except:
        pass
    components.html("<script>window.parent.print();</script>", height=0, width=0)
st.markdown("---")
st.caption("""
**【データの取り扱いについて】**
本システムに入力されたデータは、診断精度の向上および統計的な業界分析のために、個人・企業を特定できない形式（匿名加工情報）にて
サーバーへ保存・活用される場合があります。入力データが第三者にそのまま開示されることはありません。
ご利用にあたっては、上記に同意したものとみなします。
""")