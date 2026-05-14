import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

# 1. 設置全局風格
plt.style.use('dark_background')
st.set_page_config(page_title="HK Real Estate BI V5.13", layout="wide", page_icon="🏢")

st.title("🏢 HK Real Estate BI Dashboard (V5.13)")
st.markdown("Precision Analysis: **High-Leverage Sensitivity + Fully Formatted Financial Inputs**")


# 2. 數據載入引擎
@st.cache_data
def load_data():
    try:
        df = pd.read_excel('HK_Developer_Margin_Master_v3.xlsx')
        df['Date'] = pd.to_datetime(df['Date'])
        if 'Land_Index' not in df.columns:
            df['Land_Index'] = 100 + (df['CCL_Norm'] - 100) * 1.55
        return df
    except Exception as e:
        st.error(f"Excel 讀取錯誤: {e}")
        return pd.DataFrame()


try:
    df_macro = load_data()

    # 3. 側邊欄控制面版
    st.sidebar.header("⚙️ Module Switcher")
    analysis_type = st.sidebar.radio(
        "Select Analysis View:",
        ["1. Macro Trends (CCI/TPI/LPI)", "2. Sensitivity Matrix", "3. Project ROI Calculator"]
    )

    # --- Module 1: Macro Trends ---
    if analysis_type == "1. Macro Trends (CCI/TPI/LPI)":
        st.header("📈 The Trinity: Revenue, Construction & Land Cost")
        view_mode = st.radio("Display Mode:", ["Index-based (Growth Comparison)", "HKD/sqft Equivalent"])

        if not df_macro.empty:
            min_year, max_year = int(df_macro['Date'].dt.year.min()), int(df_macro['Date'].dt.year.max())
            selected_years = st.slider("Timeline Range", min_year, max_year, (1994, 2025))
            mask = (df_macro['Date'].dt.year >= selected_years[0]) & (df_macro['Date'].dt.year <= selected_years[1])
            f_df = df_macro.loc[mask].copy()

            fig, (ax1, ax2, ax3, ax4) = plt.subplots(4, 1, figsize=(14, 20), sharex=True,
                                                     gridspec_kw={'height_ratios': [2.5, 1, 1, 1]})
            fig.patch.set_facecolor('#0E1117')

            if "Index" in view_mode:
                ax1.plot(f_df['Date'], f_df['CCL_Norm'], label='Centa-City Index (CCI)', color='#d62728', linewidth=2.5)
                ax1.plot(f_df['Date'], f_df['TPI_Norm'], label='Tender Price Index (TPI)', color='#1f77b4', linewidth=2)
                ax1.plot(f_df['Date'], f_df['Land_Index'], label='Land Price Index (LPI)', color='#f1c40f',
                         linestyle=':', linewidth=2)
                ax1.set_ylabel("Index (Base=100)", color='white')
            else:
                ccl_abs, tpi_abs = f_df['CCL_Norm'] * 150, f_df['TPI_Norm'] * 45
                land_abs = ccl_abs * 0.55
                ax1.fill_between(f_df['Date'], tpi_abs, land_abs, color='blue', alpha=0.1)
                ax1.fill_between(f_df['Date'], land_abs, ccl_abs, color='green', alpha=0.15)
                ax1.plot(f_df['Date'], ccl_abs, label='CCI Price ($/sqft)', color='#d62728', linewidth=3)
                ax1.plot(f_df['Date'], tpi_abs, label='TPI Cost ($/sqft)', color='#1f77b4', linewidth=2)
                ax1.plot(f_df['Date'], land_abs, label='Est. Land Cost ($/sqft)', color='#f1c40f', linestyle='--',
                         linewidth=2)
                ax1.set_ylabel("HKD per sqft ($)", color='white')

            for ax in [ax1, ax2, ax3, ax4]:
                ax.tick_params(colors='white')
                ax.grid(True, alpha=0.15)
                ax.legend(loc='upper left', labelcolor='white', frameon=True, facecolor='#262730')

            ax2.plot(f_df['Date'], f_df['HIBOR_3M'], label='3M HIBOR (%)', color='#ff7f0e', linestyle='--')
            ax3.plot(f_df['Date'], f_df['Unemployment_Rate'], label='Unemployment Rate (%)', color='#2ca02c')
            ax4.plot(f_df['Date'], f_df['M2_Index'], label='M2 Supply Index', color='#9467bd')

            plt.tight_layout(pad=3.0)
            st.pyplot(fig)

    # --- Module 2: Sensitivity Matrix ---
    elif analysis_type == "2. Sensitivity Matrix":
        st.header("🎛️ High-Leverage Sensitivity Analysis")
        st.caption("💡 地產是高槓桿行業，導入 Operating Leverage 參數以反映價格下跌對利潤的劇烈衝擊。")

        cols = st.columns(4)
        base_roi = cols[0].number_input("Base Case ROI (%)", value=15.0) / 100
        leverage_factor = cols[1].slider("Operating Leverage", 2.0, 8.0, 4.2)
        int_weight = cols[3].slider("Interest Impact", 1.0, 5.0, 2.5)

        ccl_c = [0.1, 0.05, 0, -0.05, -0.1, -0.15, -0.2]
        h_rates = [1.0, 2.0, 3.0, 4.0, 5.0, 6.0]

        results = [[base_roi + (c * leverage_factor) - (h - 3.0) * (int_weight / 100) for h in h_rates] for c in ccl_c]
        df_matrix = pd.DataFrame(results, index=[f"Price {int(c * 100)}%" for c in ccl_c],
                                 columns=[f"H {h}%" for h in h_rates])

        fig, ax = plt.subplots(figsize=(12, 8))
        fig.patch.set_facecolor('#0E1117')
        sns.heatmap(df_matrix, annot=True, fmt=".1%", cmap="RdYlGn", center=0.0, ax=ax)
        plt.title("Impact on Project ROI (%)", color='white', fontsize=14)
        ax.tick_params(colors='white')
        st.pyplot(fig)

    # --- Module 3: ROI Calculator ---
    elif analysis_type == "3. Project ROI Calculator":
        st.header("🏗️ Real-World Project ROI Analysis")

        data_map = {
            "Case A: Kowloon Bay (COHL)": [23490, 1806880000, 5500, 17500, 4.0, 60],
            "Case B: Shau Kei Wan (Kerry)": [13759, 1383899000, 6500, 22000, 5.5, 60],
            "Case C: Kam Sheung Road (Mega)": [112088, 5500000000, 6200, 16500, 6.0, 60],
            "Case D: Anderson Road (Mount)": [24093, 3112800000, 5500, 16500, 8.0, 60],
            "Case E: Wong Tai Sin (Azure)": [5574, 400000000, 5500, 18088, 4.5, 60],
            "Case F: Kowloon Bay (Uptown)": [48309, 5354000000, 6000, 16000, 7.5, 60],
            "Case G: Mid-Levels (Morgan)": [3716, 1000000000, 12000, 51816, 5.0, 40],
            "Case H: Wong Tai Sin (Phoenext)": [7135, 805000000, 5500, 15300, 4.0, 60],
            "Case I: Lohas Park 13 (Mirabelle)": [143693, 9068000000, 5500, 15652, 5.0, 55],
            "Case J: Wong Chuk Hang (SouthLand)": [53605, 4684000000, 6500, 29689, 4.5, 50],
            "Case K: Mong Kok Redevelopment (High-Leverage/Receivership)": [2787, 500000000, 6000, 17000, 6.0, 80]
        }

        project_choice = st.selectbox("Select Project Case:", list(data_map.keys()))
        params = data_map[project_choice]


        # 輔助解析函數：將帶逗號的字串安全轉換為數字
        def parse_currency(val, default):
            try:
                return float(str(val).replace(",", ""))
            except:
                return float(default)


        col1, col2, col3 = st.columns(3)
        with col1:
            gfa_raw = st.text_input("GFA (sqm)", value=f"{int(params[0]):,}")
            gfa = parse_currency(gfa_raw, params[0])

            land_raw = st.text_input("Land/Acquisition Cost ($)", value=f"{int(params[1]):,}")
            land = parse_currency(land_raw, params[1])

        with col2:
            const_raw = st.text_input("Const Cost ($/psf)", value=f"{int(params[2]):,}")
            const = parse_currency(const_raw, params[2])

            asp_raw = st.text_input("Expected Selling Price ($/psf)", value=f"{int(params[3]):,}")
            asp = parse_currency(asp_raw, params[3])

        with col3:
            tenure = st.slider("Dev Tenure (Yrs)", 1.0, 10.0, float(params[4]), step=0.5)
            hibor = st.slider("Avg Interest Rate (%)", 1.0, 10.0, 5.0, step=0.1)

            # 財務計算
        gfa_sqft = gfa * 10.7639
        total_const = const * gfa_sqft
        interest = (land + total_const) * (params[5] / 100) * (hibor / 100) * tenure
        revenue = asp * gfa_sqft
        mkt_cost = revenue * 0.05
        total_cost = land + total_const + interest + mkt_cost
        profit = revenue - total_cost

        # 防止除以零
        roi = profit / total_cost if total_cost > 0 else 0

        st.markdown("---")
        m1, m2, m3 = st.columns(3)
        m1.metric("Net Profit (100M HKD)", f"${profit / 100000000:,.2f}")
        m2.metric("Break-even Price ($/psf)", f"${total_cost / gfa_sqft:,.0f}")
        m3.metric("Estimated ROI (%)", f"{roi:.1%}")

        # 圓餅圖
        fig, ax = plt.subplots(figsize=(10, 6))
        fig.patch.set_facecolor('#0E1117')
        labels = [
            f'Land (${land:,.0f})',
            f'Const (${total_const:,.0f})',
            f'Interest (${interest:,.0f})',
            f'Marketing (${mkt_cost:,.0f})'
        ]
        ax.pie([land, total_const, interest, mkt_cost], labels=labels,
               autopct='%1.1f%%', startangle=90, labeldistance=1.15,
               colors=['#1f77b4', '#ff7f0e', '#d62728', '#9467bd'], textprops={'color': "w"})
        plt.tight_layout()
        st.pyplot(fig)

        st.markdown("---")
        st.subheader("💡 Project Insights & Strategy")

        if "Case A" in project_choice:
            st.info(
                f"**穩健控制成本：** 傳統市區地段，發展商發揮強大的成本控制優勢。只要能以平均 **${asp:,.0f}/呎** 去貨，項目仍具備防守性。")
        elif "Case B" in project_choice:
            st.info(
                f"**港島東精品溢價：** 雖然地價與建築成本皆高，但憑藉港島區稀缺性，若能維持 **${asp:,.0f}** 以上均價，利潤率依然可觀。")
        elif "Case C" in project_choice:
            st.warning(
                f"**鐵路樞紐巨無霸：** 規模效應雖能降低平均成本，但在 **{hibor}%** 高息環境下，高達 **${land:,.0f}** 補地價帶來的利息開支極高，去貨速度是致勝關鍵。")
        elif "Case D" in project_choice:
            st.error(
                f"🚨 **災難性利息侵蝕：** 典型生不逢時。高位極高溢價投地加上 8 年開發週期，在 **{hibor}%** 利率下，龐大利息幾乎吞噬所有潛在利潤。")
        elif "Case E" in project_choice:
            st.success(
                f"**納米樓快打慢：** 投資額低（地價約 **${land:,.0f}**），主打細單位。追求『貨如輪轉』，只要能在短時間內清倉，即便呎價承壓仍能維持健康回報。")
        elif "Case F" in project_choice:
            st.warning(
                f"⚠️ **流血開價求生：** 高價拿地後面對市況逆轉，被迫貼近保本價（**${total_cost / gfa_sqft:,.0f}/呎**）推盤，反映寧願微蝕也要套現的現金流保衛戰。")
        elif "Case G" in project_choice:
            st.success(
                f"💎 **超級豪宅的護城河：** 建築成本高達 **${const:,.0f}/呎**，但客群價格敏感度低。低槓桿讓發展商有底氣慢慢『守』，不受高息逼迫。")
        elif "Case H" in project_choice:
            st.info(
                f"**市區重建實用戰術：** 透過精準成本控制與貼市開價（**${asp:,.0f}/呎**），在疲弱市況中吸引剛性需求，確保微利離場。")
        elif "Case I" in project_choice:
            st.info(
                f"**『康城終章』的博弈：** 巨量體代表極大去貨風險。預期開價 **${asp:,.0f}/呎** 正為咗喺高息環境下快速回籠資金，降低長期利息侵蝕。")
        elif "Case J" in project_choice:
            st.success(
                f"🏆 **完美的逃頂時機：** 在 2021 樓市極高點以近 **$30,000** 呎價發售，發展商成功在市況逆轉前鎖定高達 **${profit / 100000000:,.2f} 億港元** 利潤。")
        elif "Case K" in project_choice:
            st.error(
                f"☠️ **高槓桿死亡交叉 (Receivership Risk)：** 呢個係模擬樂風集團等進取型中小型發展商嘅危險劇本。市區舊樓併購時間長達 **{tenure} 年**，加上高達 **{params[5]}% 嘅 LTV 借貸比率**。喺 **{hibor}%** 嘅高息環境下，你會見到圓餅圖中嘅【利息支出】高達 **${interest:,.0f}**。如果開售呎價由原本預期嘅 **$22,000** 跌到得返 **${asp:,.0f}**，淨利潤會直接變成負數。當資產價值跌穿貸款額，就會觸發債權人 Call Loan 甚至接管 (Receivership) 地盤。")

except Exception as e:
    st.error(f"Execution Error: {e}")
