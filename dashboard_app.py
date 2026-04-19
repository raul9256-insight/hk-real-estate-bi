import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

# 1. 設定全局風格
plt.style.use('dark_background')
st.set_page_config(page_title="HK Real Estate BI V5.4", layout="wide", page_icon="🏢")

st.title("🏢 HK Real Estate BI Dashboard (V5.4)")
st.markdown("當前模型：**宏觀多因子分析 + 微觀地皮 ROI 實戰對比 (UI/字體最終優化版)**")


# 2. 讀取數據
@st.cache_data
def load_data():
    try:
        df = pd.read_excel('HK_Developer_Margin_Master_v3.xlsx')
        df['Date'] = pd.to_datetime(df['Date'])
        return df
    except:
        st.error("找不到數據庫，請確保 Excel 存在且格式正確。")
        return pd.DataFrame()


try:
    df_macro = load_data()

    # 3. 側邊欄控制面版
    st.sidebar.header("⚙️ Module Switcher")
    analysis_type = st.sidebar.radio(
        "Select Function:",
        ["1. Macro Trends (1994-2025)", "2. Sensitivity Matrix", "3. Project ROI Calculator"]
    )

    # --- 模組 1: 宏觀趨勢圖 (全英文修復版) ---
    if analysis_type == "1. Macro Trends (1994-2025)":
        st.header("📈 Macroeconomic Indicator Overview")

        # 外部中文解釋
        st.info("💡 數據對齊提示：'Index-based' 看漲幅趨勢；'HKD/sqft' 還原真實金額以觀察售價與成本的利潤利差。")
        view_mode = st.radio("Display Format:",
                             ["Index-based (Compare Growth)", "HKD/sqft Equivalent (Compare Spread)"])

        if not df_macro.empty:
            min_year, max_year = int(df_macro['Date'].dt.year.min()), int(df_macro['Date'].dt.year.max())
            selected_years = st.slider("Select Observation Period", min_year, max_year, (1994, 2025))
            mask = (df_macro['Date'].dt.year >= selected_years[0]) & (df_macro['Date'].dt.year <= selected_years[1])
            f_df = df_macro.loc[mask].copy()

            fig, (ax1, ax2, ax3, ax4) = plt.subplots(4, 1, figsize=(14, 16), sharex=True,
                                                     gridspec_kw={'height_ratios': [2, 1, 1, 1]})
            fig.patch.set_facecolor('#0E1117')

            if "Index-based" in view_mode:
                ax1.plot(f_df['Date'], f_df['CCL_Norm'], label='Price Index (CCL)', color='#d62728', linewidth=2.5)
                ax1.plot(f_df['Date'], f_df['TPI_Norm'], label='Cost Index (TPI)', color='#1f77b4', linewidth=2)
                ax1.set_ylabel("Index (Base=100)", color='white', fontsize=12)
                title_suffix = "(Index-based)"
            else:
                ccl_abs = f_df['CCL_Norm'] * 150
                tpi_abs = f_df['TPI_Norm'] * 45
                ax1.fill_between(f_df['Date'], tpi_abs, ccl_abs, color='gray', alpha=0.15, label='Gross Margin Space')
                ax1.plot(f_df['Date'], ccl_abs, label='Est. Sales Price ($/sqft)', color='#d62728', linewidth=2.5)
                ax1.plot(f_df['Date'], tpi_abs, label='Est. Const. Cost ($/sqft)', color='#1f77b4', linewidth=2)
                ax1.set_ylabel("HKD per sqft ($)", color='white', fontsize=12)
                title_suffix = "(HKD/sqft Equivalent)"

            ax1.set_title(f"Market Reality: {title_suffix}", fontsize=20, color='white', pad=20)

            ax2.plot(f_df['Date'], f_df['HIBOR_3M'], label='3M HIBOR (%)', color='#ff7f0e', linestyle='--')
            ax2.fill_between(f_df['Date'], 0, f_df['HIBOR_3M'], color='#ff7f0e', alpha=0.1)
            ax2.set_ylabel("Percent (%)", color='white')

            ax3.plot(f_df['Date'], f_df['Unemployment_Rate'], label='Unemployment Rate (%)', color='#2ca02c')
            ax3.set_ylabel("Percent (%)", color='white')

            ax4.plot(f_df['Date'], f_df['M2_Index'], label='M2 Supply Index', color='#9467bd', linewidth=2)
            ax4.set_ylabel("Index Value", color='white')
            ax4.set_xlabel("Timeline", color='white', fontsize=14)

            # 強制標籤與顏色
            for ax in [ax1, ax2, ax3, ax4]:
                ax.tick_params(axis='x', colors='white', labelsize=11)
                ax.tick_params(axis='y', colors='white', labelsize=11)
                ax.grid(True, alpha=0.15, linestyle='--')
                ax.legend(loc='upper left', labelcolor='white', facecolor='#262730')

            plt.tight_layout()
            st.pyplot(fig)

    # --- 模組 2: 盈虧平衡模擬器 ---
    elif analysis_type == "2. Sensitivity Matrix":
        st.header("🎛️ Project Profitability Sensitivity Analysis")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            base_margin = st.number_input("Base Ratio", value=0.89)
        with col2:
            sim_unemployment = st.slider("Unemployment Simulation (%)", 1.0, 10.0, 3.0)
        with col3:
            sim_m2_delta = st.slider("M2 Delta (%)", -5.0, 15.0, 0.0)
        with col4:
            cost_multiplier = st.slider("Interest Multiplier", 1.0, 5.0, 2.1)

        unemp_impact = (sim_unemployment - 3.0) * -0.03
        m2_impact = sim_m2_delta * 0.005
        ccl_changes = [0.10, 0.05, 0.00, -0.05, -0.10, -0.15, -0.20]
        hibor_rates = [1.0, 2.0, 3.0, 4.0, 5.0, 6.0]

        results = []
        for ccl_c in ccl_changes:
            row = []
            for h in hibor_rates:
                total_delta = ccl_c + unemp_impact + m2_impact
                cost_impact = 1 + ((h - 3.0) * (cost_multiplier / 100))
                new_ratio = base_margin * (1 + total_delta) / cost_impact
                row.append(new_ratio)
            results.append(row)

        df_matrix = pd.DataFrame(results, index=[f"{int(c * 100)}%" for c in ccl_changes],
                                 columns=[f"{h}%" for h in hibor_rates])

        st.subheader("🚩 Key Indicators")
        be_hibor = 3.0 + ((base_margin * (1 + (unemp_impact + m2_impact))) - 1) / (cost_multiplier / 100)
        c1, c2 = st.columns(2)
        with c1:
            st.metric("Break-even HIBOR", f"{be_hibor:.2f}%")
        with c2:
            st.write("✅ Safe Buffer" if be_hibor > 4.0 else "⚠️ Margin Risk")

        fig, ax = plt.subplots(figsize=(12, 8))
        fig.patch.set_facecolor('#0E1117')
        sns.heatmap(df_matrix, annot=True, fmt=".2f", cmap="RdYlGn", center=0.95, ax=ax)
        plt.title("Profit Sensitivity Matrix", color='white')
        ax.tick_params(colors='white')
        st.pyplot(fig)

    # --- 模組 3: 單項地皮 ROI 計算機 ---
    elif analysis_type == "3. Project ROI Calculator":
        st.header("🏗️ Real Estate Case Study (ROI)")

        project_choice = st.selectbox(
            "Select Real-World Case:",
            [
                "Case A: Kowloon Bay (COHL)",
                "Case B: Shau Kei Wan (Kerry)",
                "Case C: Kam Sheung Road (Mega Project)",
                "Case D: Anderson Road (Mount Anderson)",
                "Case E: Wong Tai Sin (The Met. Azure)",
                "Case F: Kowloon Bay (Uptown East)",
                "Case G: Mid-Levels (The Morgan)",
                "Case H: Wong Tai Sin (Phoenext)",
                "Case I: Lohas Park 13 (La Mirabelle)",
                "Case J: Wong Chuk Hang (The SouthLand)"
            ]
        )

        # 數據與之前一致，僅將顯示改為英文
        if "SouthLand" in project_choice:
            def_gfa, def_land, def_const, def_asp, def_years, def_ltv = 53605, 4684000000, 6500, 29689, 4.5, 50
        elif "La Mirabelle" in project_choice:
            def_gfa, def_land, def_const, def_asp, def_years, def_ltv = 143693, 9068000000, 5500, 15652, 5.0, 55
        elif "Phoenext" in project_choice:
            def_gfa, def_land, def_const, def_asp, def_years, def_ltv = 7135, 805000000, 5500, 15300, 4.0, 60
        elif "The Morgan" in project_choice:
            def_gfa, def_land, def_const, def_asp, def_years, def_ltv = 3716, 1000000000, 12000, 51816, 5.0, 40
        elif "Uptown East" in project_choice:
            def_gfa, def_land, def_const, def_asp, def_years, def_ltv = 48309, 5354000000, 6000, 16000, 7.5, 60
        elif "The Met. Azure" in project_choice:
            def_gfa, def_land, def_const, def_asp, def_years, def_ltv = 5574, 400000000, 5500, 18088, 4.5, 60
        elif "Mount Anderson" in project_choice:
            def_gfa, def_land, def_const, def_asp, def_years, def_ltv = 24093, 3112800000, 5500, 16500, 8.0, 60
        elif "Kam Sheung Road" in project_choice:
            def_gfa, def_land, def_const, def_asp, def_years, def_ltv = 112088, 5500000000, 6200, 16500, 6.0, 60
        elif "Kerry" in project_choice:
            def_gfa, def_land, def_const, def_asp, def_years, def_ltv = 13759, 1383899000, 6500, 22000, 5.5, 60
        else:
            def_gfa, def_land, def_const, def_asp, def_years, def_ltv = 23490, 1806880000, 5500, 17500, 4.0, 60

        st.subheader("📝 Project Assumptions")
        c1, c2, c3 = st.columns(3)
        with c1:
            gfa_raw = st.text_input("GFA (sqm)", value="{:,}".format(def_gfa))
            gfa_sqm = float(gfa_raw.replace(",", ""))
            land_raw = st.text_input("Land Cost ($)", value="{:,}".format(def_land))
            land_premium = float(land_raw.replace(",", ""))
        with c2:
            const_raw = st.text_input("Const. Cost ($/sqft)", value="{:,}".format(def_const))
            const_cost_psf = float(const_raw.replace(",", ""))
            asp_raw = st.text_input("ASP ($/sqft)", value="{:,}".format(def_asp))
            target_asp = float(asp_raw.replace(",", ""))
        with c3:
            ltv_ratio = st.slider("LTV (%)", 0, 100, def_ltv)
            dev_years = st.slider("Tenure (Years)", 1.0, 10.0, def_years, step=0.5)
            hibor_rate = st.slider("Avg HIBOR (%)", 1.0, 8.0, 4.0, step=0.1)

        gfa_sqft = gfa_sqm * 10.7639
        land_psf = land_premium / gfa_sqft
        total_land_cost = land_premium
        total_const_cost = const_cost_psf * gfa_sqft
        base_cost = total_land_cost + total_const_cost
        loan_amount = base_cost * (ltv_ratio / 100)
        total_interest = loan_amount * (hibor_rate / 100) * dev_years
        gross_revenue = target_asp * gfa_sqft
        admin_marketing_cost = gross_revenue * 0.05
        total_cost = base_cost + total_interest + admin_marketing_cost
        net_profit = gross_revenue - total_cost
        roi = (net_profit / total_cost) * 100
        break_even_asp = total_cost / gfa_sqft

        st.markdown("---")
        st.subheader("📊 Financial Viability Result")
        r1, r2, r3, r4 = st.columns(4)
        r1.metric("Land AV", "${:,.0f}/psf".format(land_psf))
        r2.metric("Break-even", "${:,.0f}/psf".format(break_even_asp))
        roi_color = "normal" if roi > 0 else "inverse"
        r3.metric("Net Profit (HKD 100M)", "${:,.2f}".format(net_profit / 100000000), delta_color=roi_color)
        r4.metric("ROI (%)", "{:.1f}%".format(roi), delta="{:.1f}% vs 15%".format(roi - 15), delta_color=roi_color)

        st.write("### 💰 Cost Structure Breakdown (Unit: 100M)")
        fig, ax = plt.subplots(figsize=(10, 6))
        fig.patch.set_facecolor('#0E1117')
        labels = ['Land', 'Const.', 'Interest', 'Marketing']
        sizes = [total_land_cost, total_const_cost, total_interest, admin_marketing_cost]
        colors = ['#1f77b4', '#ff7f0e', '#d62728', '#9467bd']


        def label_formatter(pct, allvals):
            absolute = pct / 100. * np.sum(allvals)
            return "${:.2f}\n({:.1f}%)".format(absolute / 100000000, pct)


        ax.pie(sizes, labels=labels, autopct=lambda pct: label_formatter(pct, sizes), startangle=90, colors=colors,
               textprops=dict(color="w", fontsize=10), pctdistance=0.75, labeldistance=1.1)
        ax.axis('equal')
        plt.tight_layout()
        st.pyplot(fig)

except Exception as e:
    st.error(f"Error: {e}")