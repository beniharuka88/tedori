import streamlit as st
import pandas as pd
import plotly.express as px
import io
from openpyxl.drawing.image import Image
import openpyxl
from streamlit_option_menu import option_menu

# --- データ定義 (2025年度版) ---
# (このセクションのコードは前回から変更ありません)
health_insurance_rates = pd.DataFrame({
    '都道府県': ['北海道', '東京都', '神奈川県', '埼玉県', '千葉県', '愛知県', '大阪府', '福岡県'],
    '料率(40歳未満)': [10.02, 9.98, 9.93, 9.75, 9.73, 10.04, 10.34, 10.15],
    '料率(40歳以上)': [11.82, 11.78, 11.73, 11.55, 11.53, 11.84, 12.14, 11.95]
})
other_rates = {
    '厚生年金': 18.30,
    '雇用保険': 0.60
}
srm_table = pd.DataFrame({
    '報酬月額(以上)': [0, 93000, 101000, 107000, 114000, 122000, 130000, 138000, 146000, 155000, 165000, 175000, 185000, 195000, 210000, 230000, 250000, 270000, 290000, 310000, 330000, 350000, 370000, 395000, 425000, 455000, 485000, 515000, 545000, 575000, 605000, 635000],
    '標準報酬月額': [88000, 98000, 104000, 110000, 118000, 126000, 134000, 142000, 150000, 160000, 170000, 180000, 190000, 200000, 220000, 240000, 260000, 280000, 300000, 320000, 340000, 360000, 380000, 410000, 440000, 470000, 500000, 530000, 560000, 590000, 620000, 650000]
})
income_tax_table = pd.DataFrame({
    '課税対象額(以上)': [0, 88000, 91000, 94000, 248000, 251000, 329000],
    '税額': [2680, 2770, 2870, 2980, 8770, 8920, 13720],
    '税率(%)': [0, 2.042, 4.084, 4.084, 6.126, 8.168, 10.210]
})

# --- 計算関数 (変更なし) ---


def calculate_net_income(total_income, prefecture, age):
    if total_income == 0:
        return {'net_income': 0, 'social_insurance_total': 0, 'taxable_income_for_tax_rate': 0}
    srm_row = srm_table[srm_table['報酬月額(以上)'] <= total_income].iloc[-1]
    srm = srm_row['標準報酬月額']
    rate_col = '料率(40歳以上)' if age >= 40 else '料率(40歳未満)'
    health_rate = health_insurance_rates[health_insurance_rates['都道府県']
                                         == prefecture][rate_col].iloc[0] / 100
    health_fee = round(srm * health_rate / 2)
    pension_rate = other_rates['厚生年金'] / 100
    pension_fee = round(srm * pension_rate / 2)
    employment_rate = other_rates['雇用保険'] / 100
    employment_fee = round(total_income * employment_rate)
    social_insurance_total = health_fee + pension_fee + employment_fee
    taxable_income = total_income - social_insurance_total
    income_tax_row = income_tax_table[income_tax_table['課税対象額(以上)']
                                      <= taxable_income].iloc[-1]
    income_tax = income_tax_row['税額']
    net_income = total_income - social_insurance_total - income_tax
    return {'net_income': net_income, 'social_insurance_total': social_insurance_total, 'taxable_income_for_tax_rate': taxable_income}


def calculate_bonus_net(bonus_gross, monthly_taxable_for_rate, prefecture, age):
    if bonus_gross == 0:
        return 0
    standard_bonus = int(bonus_gross / 1000) * 1000
    rate_col = '料率(40歳以上)' if age >= 40 else '料率(40歳未満)'
    health_rate = health_insurance_rates[health_insurance_rates['都道府県']
                                         == prefecture][rate_col].iloc[0] / 100
    pension_rate = other_rates['厚生年金'] / 100
    employment_rate = other_rates['雇用保険'] / 100
    bonus_health_fee = round(standard_bonus * health_rate / 2)
    bonus_pension_fee = round(standard_bonus * pension_rate / 2)
    bonus_employment_fee = round(bonus_gross * employment_rate)
    bonus_social_insurance_total = bonus_health_fee + \
        bonus_pension_fee + bonus_employment_fee
    tax_rate_row = income_tax_table[income_tax_table['課税対象額(以上)']
                                    <= monthly_taxable_for_rate].iloc[-1]
    tax_rate = tax_rate_row['税率(%)'] / 100
    bonus_income_tax = round(
        (bonus_gross - bonus_social_insurance_total) * tax_rate)
    net_bonus = bonus_gross - bonus_social_insurance_total - bonus_income_tax
    return max(0, net_bonus)


def calculate_resident_tax(annual_income, annual_social_insurance):
    if annual_income == 0:
        return 0
    if annual_income <= 1625000:
        salary_deduction = 550000
    elif annual_income <= 1800000:
        salary_deduction = annual_income * 0.4 - 100000
    elif annual_income <= 3600000:
        salary_deduction = annual_income * 0.3 + 80000
    elif annual_income <= 6600000:
        salary_deduction = annual_income * 0.2 + 440000
    elif annual_income <= 8500000:
        salary_deduction = annual_income * 0.1 + 1100000
    else:
        salary_deduction = 1950000
    basic_deduction = 430000
    taxable_income = annual_income - salary_deduction - \
        annual_social_insurance - basic_deduction
    if taxable_income < 0:
        taxable_income = 0
    taxable_income = int(taxable_income / 1000) * 1000
    annual_resident_tax = taxable_income * 0.1 + 5000
    monthly_resident_tax = int(annual_resident_tax / 12 / 100) * 100
    return monthly_resident_tax


def create_excel_report(summary_df, details_df, charts):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        summary_df.to_excel(writer, sheet_name='サマリー', index=False)
        details_df.to_excel(writer, sheet_name='シミュレーション詳細', index=False)
        writer.book.create_sheet('グラフ')
        worksheet = writer.sheets['グラフ']
        for chart_name, fig in charts.items():
            if fig is None:
                continue
            try:
                img_bytes = fig.to_image(format="png", scale=2)
                img = Image(io.BytesIO(img_bytes))
                if chart_name == 'pie_a':
                    worksheet.add_image(img, 'A1')
                elif chart_name == 'pie_b':
                    worksheet.add_image(img, 'J1')
                elif chart_name == 'line_trend':
                    worksheet.add_image(img, 'A25')
            except Exception as e:
                st.error(f"グラフの画像変換中にエラー: {e}。kaleidoライブラリを確認してください。")
    processed_data = output.getvalue()
    return processed_data


# --- Streamlit アプリのUI部分 ---
st.set_page_config(layout="wide", page_title="最強の家計診断ツール",
                   initial_sidebar_state="expanded")

st.title('家計シュミレータ―')
st.markdown("---")

# --- ★★★ Session Stateの初期化 ★★★ ---
if 'results' not in st.session_state:
    st.session_state.results = None
if 'unlocked' not in st.session_state:
    st.session_state.unlocked = False

# --- サイドバー (入力専用) ---
with st.sidebar:
    st.header('STEP 1: お二人の情報を入力')
    with st.expander("Aさんの情報", expanded=True):
        income_a = st.number_input(
            '総支給額（月給）', min_value=0, value=287000, step=1000, key='income_a')
        bonus_a = st.number_input(
            'ボーナス年額（支給額）', min_value=0, value=500000, step=10000, key='bonus_a')
        pref_a = st.selectbox(
            '都道府県', options=health_insurance_rates['都道府県'].tolist(), index=2, key='pref_a')
        age_a = st.number_input(
            '年齢', min_value=18, max_value=70, value=23, key='age_a')
        auto_tax_a = st.checkbox(
            '2年目以降として計算（住民税・ボーナス）', value=True, key='auto_tax_a')
        resident_tax_a = st.number_input(
            '住民税（月額）', min_value=0, value=0, step=500, key='tax_a', disabled=auto_tax_a)

    with st.expander("Bさんの情報", expanded=True):
        income_b = st.number_input(
            '総支給額（月給）', min_value=0, value=226000, step=1000, key='income_b')
        bonus_b = st.number_input(
            'ボーナス年額（支給額）', min_value=0, value=300000, step=10000, key='bonus_b')
        pref_b = st.selectbox(
            '都道府県', options=health_insurance_rates['都道府県'].tolist(), index=2, key='pref_b')
        age_b = st.number_input(
            '年齢', min_value=18, max_value=70, value=23, key='age_b')
        auto_tax_b = st.checkbox(
            '2年目以降として計算（住民税・ボーナス）', value=True, key='auto_tax_b')
        resident_tax_b = st.number_input(
            '住民税（月額）', min_value=0, value=0, step=500, key='tax_b', disabled=auto_tax_b)

    st.header('STEP 2: 二人の生活費（月額）を入力')
    rent = st.number_input('🏠 家賃', min_value=0, value=110000, step=1000)
    food = st.number_input('🍳 食費', min_value=0, value=80000, step=1000)
    utilities = st.number_input(
        '💡 水道・光熱費', min_value=0, value=20000, step=1000)
    communication = st.number_input(
        '📱 通信費', min_value=0, value=12000, step=1000)
    entertainment = st.number_input(
        '🎉 交際費・娯楽費', min_value=0, value=40000, step=1000)
    other_expenses = st.number_input(
        '💸 その他雑費', min_value=0, value=25000, step=1000)

    st.header('STEP 3: 費用の分担方法')
    split_ratio_a = st.slider('Aさんの負担割合（％）', 0, 100, 50)

    st.header('STEP 4: 将来の貯蓄シミュレーション')
    sim_years = st.slider('シミュレーション期間（年）', min_value=1, max_value=30, value=10)
    annual_increase_rate = st.number_input(
        '年間の貯金増加率（%）', min_value=0.0, max_value=10.0, value=2.0, step=0.1, help="昇給などによる、貯金原資の年間増加率を設定します。")

    st.markdown("---")
    password = st.text_input("拡張モードを有効化", type="password")
    if password == "invest":
        st.session_state.unlocked = True

    if st.session_state.unlocked:
        st.sidebar.success("拡張モードが解放されました！")
        with st.sidebar.expander("【裏機能】コア・サテライト戦略", expanded=True):
            investor = st.radio("戦略を適用する人を選択", ('Aさん', 'Bさん'))
            core_investment_ratio = st.slider(
                "年間貯金のうち、コア投資に回す割合（%）", 0, 100, 70)
            core_return_rate = st.number_input(
                "コア戦略の想定利回り（年率%）", 0.0, 15.0, 5.0, 0.1, help="全世界株式インデックスファンドなどを想定。")
            satellite_return_rate = st.number_input(
                "サテライト戦略の想定利回り（年率%）", 0.0, 20.0, 7.0, 0.1, help="個別株やテーマ型ETFなどを想定。")

    # --- 診断ボタンが押されたら、計算して結果をSession Stateに保存 ---
    if st.button('📊 診断を開始する', type="primary", use_container_width=True):

        # --- 1. 年間収支の計算 ---
        net_monthly_a_info = calculate_net_income(income_a, pref_a, age_a)
        net_monthly_b_info = calculate_net_income(income_b, pref_b, age_b)
        bonus_to_calc_a = bonus_a if auto_tax_a else 0
        bonus_to_calc_b = bonus_b if auto_tax_b else 0
        net_bonus_a = calculate_bonus_net(
            bonus_to_calc_a, net_monthly_a_info['taxable_income_for_tax_rate'], pref_a, age_a)
        net_bonus_b = calculate_bonus_net(
            bonus_to_calc_b, net_monthly_b_info['taxable_income_for_tax_rate'], pref_b, age_b)

        final_resident_tax_a_val = resident_tax_a
        if auto_tax_a:
            annual_income_a_gross = income_a * 12 + bonus_to_calc_a
            annual_social_insurance_a = net_monthly_a_info['social_insurance_total'] * 12 + (
                bonus_to_calc_a - net_bonus_a)
            final_resident_tax_a_val = calculate_resident_tax(
                annual_income_a_gross, annual_social_insurance_a)
        else:
            final_resident_tax_a_val = 0

        final_resident_tax_b_val = resident_tax_b
        if auto_tax_b:
            annual_income_b_gross = income_b * 12 + bonus_to_calc_b
            annual_social_insurance_b = net_monthly_b_info['social_insurance_total'] * 12 + (
                bonus_to_calc_b - net_bonus_b)
            final_resident_tax_b_val = calculate_resident_tax(
                annual_income_b_gross, annual_social_insurance_b)
        else:
            final_resident_tax_b_val = 0

        annual_net_a = (
            net_monthly_a_info['net_income'] - final_resident_tax_a_val) * 12 + net_bonus_a
        annual_net_b = (
            net_monthly_b_info['net_income'] - final_resident_tax_b_val) * 12 + net_bonus_b
        household_annual_net_income = annual_net_a + annual_net_b

        total_monthly_expenses = rent + food + utilities + \
            communication + entertainment + other_expenses
        total_annual_expenses = total_monthly_expenses * 12
        annual_expenses_a = total_annual_expenses * (split_ratio_a / 100)
        annual_expenses_b = total_annual_expenses * \
            ((100 - split_ratio_a) / 100)
        annual_savings_a = annual_net_a - annual_expenses_a
        annual_savings_b = annual_net_b - annual_expenses_b
        household_annual_savings = annual_savings_a + annual_savings_b

        # --- 2. 全てのグラフオブジェクトを作成 ---
        fig_a, fig_b, fig_line, fig_expenses, fig_line_invest, details_df = None, None, None, None, None, pd.DataFrame()

        if annual_net_a > 0:
            pie_data_a = pd.DataFrame({'項目': ['支出', '貯金'], '金額': [max(
                0, annual_expenses_a), max(0, annual_savings_a)]})
            fig_a = px.pie(pie_data_a, values='金額', names='項目', hole=0.4, color_discrete_sequence=[
                           '#ff9999', '#66b3ff'], title='Aさんの年間収支')
            fig_a.update_traces(textposition='inside', textinfo='percent+label',
                                hovertemplate='%{label}: %{value:,.0f}円<br>(%{percent})<extra></extra>')
        if annual_net_b > 0:
            pie_data_b = pd.DataFrame({'項目': ['支出', '貯金'], '金額': [max(
                0, annual_expenses_b), max(0, annual_savings_b)]})
            fig_b = px.pie(pie_data_b, values='金額', names='項目', hole=0.4, color_discrete_sequence=[
                           '#ff9999', '#66b3ff'], title='Bさんの年間収支')
            fig_b.update_traces(textposition='inside', textinfo='percent+label',
                                hovertemplate='%{label}: %{value:,.0f}円<br>(%{percent})<extra></extra>')

        expenses_data = {'費目': ['家賃', '食費', '水道・光熱費', '通信費', '交際費・娯楽費', 'その他雑費'], '年間金額': [
            rent*12, food*12, utilities*12, communication*12, entertainment*12, other_expenses*12]}
        expenses_df = pd.DataFrame(expenses_data).sort_values(
            by='年間金額', ascending=True)
        fig_expenses = px.bar(expenses_df, x='年間金額', y='費目',
                              orientation='h', text='年間金額')
        fig_expenses.update_traces(
            texttemplate='%{x:,.0f}円', textposition='outside')
        fig_expenses.update_layout(
            uniformtext_minsize=8, uniformtext_mode='hide', yaxis_title=None, xaxis_title="金額（円）")

        if household_annual_savings > 0:
            years = list(range(1, sim_years + 1))
            yearly_details, history_a_invest, history_b_invest, history_a_savings, history_b_savings = [], [], [], [], []
            a_invest_asset, b_invest_asset, a_savings_only_asset, b_savings_only_asset = 0, 0, 0, 0
            a_assets = {'cash': 0, 'core': 0, 'satellite': 0}
            b_assets = {'cash': 0, 'core': 0, 'satellite': 0}

            for year in years:
                current_year_savings_a = annual_savings_a * \
                    (1 + annual_increase_rate / 100) ** (year - 1)
                current_year_savings_b = annual_savings_b * \
                    (1 + annual_increase_rate / 100) ** (year - 1)
                a_savings_only_asset += current_year_savings_a
                b_savings_only_asset += current_year_savings_b
                a_invest_asset = a_savings_only_asset
                b_invest_asset = b_savings_only_asset
                if st.session_state.unlocked:
                    if investor == 'Aさん':
                        core_inv = current_year_savings_a * \
                            (core_investment_ratio / 100)
                        satellite_inv = current_year_savings_a - core_inv
                        a_assets['core'] = (
                            a_assets['core'] * (1 + core_return_rate / 100)) + core_inv
                        a_assets['satellite'] = (
                            a_assets['satellite'] * (1 + satellite_return_rate / 100)) + satellite_inv
                        a_invest_asset = a_assets['core'] + \
                            a_assets['satellite'] + a_assets['cash']
                    elif investor == 'Bさん':
                        core_inv = current_year_savings_b * \
                            (core_investment_ratio / 100)
                        satellite_inv = current_year_savings_b - core_inv
                        b_assets['core'] = (
                            b_assets['core'] * (1 + core_return_rate / 100)) + core_inv
                        b_assets['satellite'] = (
                            b_assets['satellite'] * (1 + satellite_return_rate / 100)) + satellite_inv
                        b_invest_asset = b_assets['core'] + \
                            b_assets['satellite'] + b_assets['cash']
                history_a_invest.append(a_invest_asset)
                history_b_invest.append(b_invest_asset)
                history_a_savings.append(a_savings_only_asset)
                history_b_savings.append(b_savings_only_asset)
                detail = {'年': f'{year}年目',
                          '世帯資産合計': a_invest_asset + b_invest_asset}
                if st.session_state.unlocked and investor == 'Aさん':
                    detail.update({'Aさん資産合計（投資あり）': a_invest_asset,
                                  'Aさん資産（貯金のみ）': a_savings_only_asset, 'Bさん資産合計': b_invest_asset})
                elif st.session_state.unlocked and investor == 'Bさん':
                    detail.update(
                        {'Aさん資産合計': a_invest_asset, 'Bさん資産合計（投資あり）': b_invest_asset, 'Bさん資産（貯金のみ）': b_savings_only_asset})
                else:
                    detail.update({'Aさん累計貯金額': a_invest_asset,
                                  'Bさん累計貯金額': b_invest_asset})
                yearly_details.append(detail)
            details_df = pd.DataFrame(yearly_details)

            history_df_savings = pd.DataFrame(
                {'年': years, 'Aさん（貯金のみ）': history_a_savings, 'Bさん（貯金のみ）': history_b_savings})
            history_melted_savings = history_df_savings.melt(
                id_vars='年', var_name='シミュレーション', value_name='資産額')
            fig_line = px.line(history_melted_savings, x='年', y='資産額',
                               color='シミュレーション', markers=True, title=f'{sim_years}年後の資産額の推移')
            fig_line.update_layout(
                xaxis_title="経過年数", yaxis_title="累計資産額", yaxis_tickformat=',.0f')
            fig_line.update_traces(
                hovertemplate='<b>%{data.name}</b><br>%{x}年目: %{y:,.0f}円<extra></extra>')

            if st.session_state.unlocked:
                history_df_invest = pd.DataFrame({'年': years})
                if investor == 'Aさん':
                    history_df_invest['Aさん（投資あり）'] = history_a_invest
                    history_df_invest['Aさん（貯金のみ）'] = history_a_savings
                    history_df_invest['Bさん（貯金のみ）'] = history_b_invest
                else:
                    history_df_invest['Aさん（貯金のみ）'] = history_a_invest
                    history_df_invest['Bさん（投資あり）'] = history_b_invest
                    history_df_invest['Bさん（貯金のみ）'] = history_b_savings
                history_melted_invest = history_df_invest.melt(
                    id_vars='年', var_name='シミュレーション', value_name='資産額')
                fig_line_invest = px.line(history_melted_invest, x='年', y='資産額',
                                          color='シミュレーション', markers=True, title=f'{sim_years}年後の資産額 比較シミュレーション')
                fig_line_invest.update_layout(
                    xaxis_title="経過年数", yaxis_title="累計資産額", yaxis_tickformat=',.0f')
                fig_line_invest.update_traces(
                    hovertemplate='<b>%{data.name}</b><br>%{x}年目: %{y:,.0f}円<extra></extra>')

        # --- 3. ★★★ 計算結果を全てSession Stateに保存 ★★★ ---
        st.session_state.results = {
            "household_annual_net_income": household_annual_net_income, "total_annual_expenses": total_annual_expenses, "household_annual_savings": household_annual_savings,
            "annual_net_a": annual_net_a, "annual_expenses_a": annual_expenses_a, "annual_savings_a": annual_savings_a,
            "annual_net_b": annual_net_b, "annual_expenses_b": annual_expenses_b, "annual_savings_b": annual_savings_b,
            "rent": rent, "split_ratio_a": split_ratio_a, "sim_years": sim_years,
            "fig_a": fig_a, "fig_b": fig_b, "fig_expenses": fig_expenses, "fig_line": fig_line, "fig_line_invest": fig_line_invest,
            "details_df": details_df, "final_a_asset": a_invest_asset, "final_b_asset": b_invest_asset,
            "final_a_savings_only": a_savings_only_asset, "final_b_savings_only": b_savings_only_asset,
            # Excelダウンロード用に追加
            "inputs": {'income_a': income_a, 'bonus_a': bonus_a, 'pref_a': pref_a, 'age_a': age_a, 'auto_tax_a': auto_tax_a, 'income_b': income_b, 'bonus_b': bonus_b, 'pref_b': pref_b, 'age_b': age_b, 'auto_tax_b': auto_tax_b, 'rent': rent, 'food': food, 'utilities': utilities, 'communication': communication, 'entertainment': entertainment, 'other_expenses': other_expenses, 'split_ratio_a': split_ratio_a, 'sim_years': sim_years, 'annual_increase_rate': annual_increase_rate}
        }

# --- ★★★ Session Stateに結果があれば、タブを表示 ★★★ ---
if st.session_state.results:
    # --- 結果をローカル変数に展開 ---
    res = st.session_state.results

    # --- タブの定義 ---
    tab_options = ["ダッシュボード", "個人別の内訳", "将来の資産シミュレーション"]
    tab_icons = ['house', 'people', 'graph-up-arrow']
    if st.session_state.unlocked:
        tab_options.append("【裏】投資戦略シミュレーション")
        tab_icons.append('gem')

    selected_tab = option_menu(menu_title=None, options=tab_options, icons=tab_icons,
                               menu_icon="cast", default_index=0, orientation="horizontal")

    # --- タブ1: ダッシュボード ---
    if selected_tab == "ダッシュボード":
        st.subheader("📊 年間収支ダッシュボード")
        col1, col2, col3 = st.columns(3)
        col1.metric("世帯の合計手取り（年間）",
                    f"{res['household_annual_net_income']:,.0f} 円")
        col2.metric("合計支出（年間）", f"{res['total_annual_expenses']:,.0f} 円")
        col3.metric("貯金可能額（年間）", f"{res['household_annual_savings']:,.0f} 円")
        st.markdown("---")
        col1, col2 = st.columns([1, 1])
        with col1:
            st.subheader("💡 ワンポイントアドバイス")
            if res['household_annual_net_income'] > 0:
                annual_rent = res['rent'] * 12
                rent_ratio = annual_rent / res['household_annual_net_income']
                if rent_ratio > 0.3:
                    st.warning(f"**家賃が高いかも？** (手取りの {rent_ratio:.1%})")
                else:
                    st.success(f"**家賃は適切です。** (手取りの {rent_ratio:.1%})")
                savings_ratio = res['household_annual_savings'] / \
                    res['household_annual_net_income']
                if savings_ratio < 0:
                    st.error(
                        f"**家計が赤字です！** ({res['household_annual_savings']:,.0f}円)")
                elif savings_ratio < 0.1:
                    st.warning(f"**貯金が少なめかも？** (貯蓄率 {savings_ratio:.1%})")
                elif savings_ratio >= 0.2:
                    st.success(f"**素晴らしい貯蓄率です！** (貯蓄率 {savings_ratio:.1%})")
                else:
                    st.info(f"**着実に貯金できています。** (貯蓄率 {savings_ratio:.1%})")
        with col2:
            st.subheader("支出の内訳（年間）")
            if res['fig_expenses']:
                st.plotly_chart(res['fig_expenses'], use_container_width=True)

    # --- タブ2: 個人別の内訳 ---
    if selected_tab == "個人別の内訳":
        st.subheader("👥 個人の収支内訳（年間）")
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Aさんの年間収支")
            st.write(f"**年間手取り額:** {res['annual_net_a']:,.0f} 円")
            st.write(
                f"**年間負担支出額:** {res['annual_expenses_a']:,.0f} 円 ({res['split_ratio_a']}%)")
            st.metric("年間貯金額", f"{res['annual_savings_a']:,.0f} 円")
            if res['fig_a']:
                st.plotly_chart(res['fig_a'], use_container_width=True)
        with col2:
            st.subheader("Bさんの年間収支")
            st.write(f"**年間手取り額:** {res['annual_net_b']:,.0f} 円")
            st.write(
                f"**年間負担支出額:** {res['annual_expenses_b']:,.0f} 円 ({100 - res['split_ratio_a']}%)")
            st.metric("年間貯金額", f"{res['annual_savings_b']:,.0f} 円")
            if res['fig_b']:
                st.plotly_chart(res['fig_b'], use_container_width=True)

    # --- タブ3: 将来の資産シミュレーション ---
    if selected_tab == "将来の資産シミュレーション":
        st.subheader(f"📈 {res['sim_years']}年間の将来の資産推移（貯金のみ）")
        if res['fig_line']:
            st.plotly_chart(res['fig_line'], use_container_width=True)
            st.success(
                f"**{res['sim_years']}年後には...** 世帯合計で **{res['final_a_savings_only'] + res['final_b_savings_only']:,.0f}円** の資産になる見込みです。")
            st.subheader("資産シミュレーション詳細")
            format_dict = {
                col: '{:,.0f}円' for col in res['details_df'].columns if col != '年'}
            st.dataframe(res['details_df'].set_index('年').style.format(
                format_dict), use_container_width=True)
        else:
            st.warning("貯金額が0円以下の場合、シミュレーションは表示できません。")

    # --- タブ4: 【裏】投資戦略シミュレーション ---
    if selected_tab == "【裏】投資戦略シミュレーション":
        with st.expander("【裏機能】コア・サテライト戦略の始め方", expanded=True):
            st.info("この戦略は、まず生活防衛資金（生活費の6ヶ月〜1年分）を確保してから始めるのが鉄則です。")
            st.markdown(
                "1. **証券口座の開設**\n2. **コア戦略（つみたて投資枠）**\n3. **サテライト戦略（成長投資枠）**")
        st.subheader(f"📈 {res['sim_years']}年間の投資比較シミュレーション")
        if res['fig_line_invest']:
            st.plotly_chart(res['fig_line_invest'], use_container_width=True)
            st.success(
                f"**{res['sim_years']}年後には...** 世帯合計で **{res['final_a_asset'] + res['final_b_asset']:,.0f}円** の資産になる見込みです。")
            st.subheader("資産シミュレーション詳細")
            format_dict = {
                col: '{:,.0f}円' for col in res['details_df'].columns if col != '年'}
            st.dataframe(res['details_df'].set_index('年').style.format(
                format_dict), use_container_width=True)
        else:
            st.warning("貯金額が0円以下の場合、シミュレーションは表示できません。")

    # --- ダウンロード機能 ---
    st.markdown("---")
    st.subheader("📥 レポートのダウンロード")
    summary_list = [
        {'カテゴリ': '入力値: Aさん', '項目': '月給（総支給）',
            '金額・設定値': f"{res['inputs']['income_a']:,}円"},
        {'カテゴリ': '', '項目': 'ボーナス（年額）',
            '金額・設定値': f"{res['inputs']['bonus_a']:,}円"},
        # ... (他のサマリー項目)
    ]
    summary_df_for_excel = pd.DataFrame(summary_list)
    charts_to_export = {'pie_a': res['fig_a'], 'pie_b': res['fig_b'],
                        'line_trend': res['fig_line_invest'] if st.session_state.unlocked else res['fig_line']}
    df_xlsx = create_excel_report(
        summary_df_for_excel, res['details_df'], charts_to_export)
    st.download_button(label="📥 詳細レポートをExcelファイルでダウンロード", data=df_xlsx, file_name='家計診断レポート.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True)

else:
    st.info('サイドバーにすべての情報を入力し、「診断を開始する」ボタンを押してください。')
