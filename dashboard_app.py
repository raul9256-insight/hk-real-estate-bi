"""HK Real Estate BI V5.14 — corrected, explicitly illustrative analysis.

Launch with: streamlit run app.py
The macro module needs HK_Developer_Margin_Master_v3.xlsx beside this file.
"""

from pathlib import Path
from math import isfinite

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
import streamlit as st

from finance_model import ProjectInputs, calculate_project, sensitivity_grid


st.set_page_config(page_title="HK Real Estate BI V5.14", layout="wide", page_icon="🏢")
plt.style.use("dark_background")


# Original illustrative case inputs retained; they are NOT independently verified
# observations of realized project costs, prices, debt, or profits.
# Tuple: GFA sqm, land HKD, construction HKD/gross sqft,
#        sales HKD/saleable sqft, years, assumed LTV percentage.
PROJECT_CASES = {
    "Case A: Kowloon Bay (COHL)": (23490, 1806880000, 5500, 17500, 4.0, 60),
    "Case B: Shau Kei Wan (Kerry)": (13759, 1383899000, 6500, 22000, 5.5, 60),
    "Case C: Kam Sheung Road (Mega)": (112088, 5500000000, 6200, 16500, 6.0, 60),
    "Case D: Anderson Road (Mount)": (24093, 3112800000, 5500, 16500, 8.0, 60),
    "Case E: Wong Tai Sin (Azure)": (5574, 400000000, 5500, 18088, 4.5, 60),
    "Case F: Kowloon Bay (Uptown)": (48309, 5354000000, 6000, 16000, 7.5, 60),
    "Case G: Mid-Levels (Morgan)": (3716, 1000000000, 12000, 51816, 5.0, 40),
    "Case H: Wong Tai Sin (Phoenext)": (7135, 805000000, 5500, 15300, 4.0, 60),
    "Case I: Lohas Park 13 (Mirabelle)": (143693, 9068000000, 5500, 15652, 5.0, 55),
    "Case J: Wong Chuk Hang (SouthLand)": (53605, 4684000000, 6500, 29689, 4.5, 50),
    "Case K: Mong Kok Redevelopment (High-Leverage/Receivership)": (
        2787, 500000000, 6000, 17000, 6.0, 80
    ),
}


@st.cache_data(show_spinner=False)
def load_macro_data(filepath: str, modified_ns: int) -> pd.DataFrame:
    """Cache Excel by filename AND modification time; never fabricate Land_Index."""
    _ = modified_ns
    df = pd.read_excel(filepath)
    if "Date" not in df.columns:
        raise ValueError("Excel 缺少 Date 欄位。")
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    for field in ("CCL_Norm", "TPI_Norm", "Land_Index", "HIBOR_3M", "Unemployment_Rate", "M2_Index"):
        if field in df.columns:
            df[field] = pd.to_numeric(df[field], errors="coerce")
    return df.dropna(subset=["Date"]).sort_values("Date")


def read_amount(raw: str, label: str, *, positive: bool = False) -> float:
    """Do not silently replace a mistyped number with the project's default."""
    try:
        amount = float(str(raw).strip().replace(",", ""))
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{label} 必須是有效數字（可用逗號分隔千位）。") from exc
    if not isfinite(amount) or amount < 0 or (positive and amount == 0):
        condition = "大於零" if positive else "非負"
        raise ValueError(f"{label} 必須是有限的{condition}數值。")
    return amount


def project_inputs_form(prefix: str) -> tuple[str, ProjectInputs]:
    """Use the same editable assumptions in the matrix and ROI modules."""
    project_choice = st.selectbox(
        "Select Illustrative Project Case:", list(PROJECT_CASES), key=f"{prefix}_case"
    )
    gfa_default, land_default, construction_default, asp_default, years_default, ltv_default = PROJECT_CASES[project_choice]
    case_key = f"{prefix}_{list(PROJECT_CASES).index(project_choice)}"
    left, mid, right = st.columns(3)

    with left:
        gfa_raw = st.text_input("GFA (sqm)", f"{gfa_default:,}", key=f"{case_key}_gfa")
        land_raw = st.text_input("Land/Acquisition Cost (HK$)", f"{land_default:,}", key=f"{case_key}_land")
    with mid:
        construction_raw = st.text_input(
            "Construction Cost (HK$/gross sqft)", f"{construction_default:,}",
            key=f"{case_key}_const",
        )
        asp_raw = st.text_input(
            "Expected Selling Price (HK$/saleable sqft)", f"{asp_default:,}",
            key=f"{case_key}_asp",
        )
    with right:
        tenure = st.slider(
            "Development Tenure (Years)", 0.5, 15.0, float(years_default),
            step=0.5, key=f"{case_key}_tenure",
        )
        rate = st.slider(
            "Average Annual Loan Rate (%)", 0.0, 15.0, 5.0,
            step=0.1, key=f"{case_key}_rate",
        )

    with st.expander("⚙️ Area, leverage and cost assumptions (important)", expanded=True):
        a, b, c, d = st.columns(4)
        with a:
            ltv = st.slider(
                "LTV on land + building cost (%)", 0.0, 100.0, float(ltv_default),
                step=1.0, key=f"{case_key}_ltv",
            )
        with b:
            efficiency = st.slider(
                "Saleable / GFA area (%)", 50.0, 100.0, 100.0,
                step=1.0, key=f"{case_key}_eff",
            )
        with c:
            marketing = st.slider(
                "Marketing (% of sale revenue)", 0.0, 25.0, 5.0,
                step=0.5, key=f"{case_key}_marketing",
            )
        with d:
            utilization = st.slider(
                "Average outstanding debt (% of peak)", 0.0, 100.0, 100.0,
                step=5.0, key=f"{case_key}_utilization",
            )
        st.caption(
            "施工費按總樓面面積 GFA；售價按實用面積計。原始案例沒有提供實用率，"
            "故以 100% 作兼容性預設，使用市場實用呎價前務必改為有證據的實用率。"
        )

    try:
        inputs = ProjectInputs(
            gfa_sqm=read_amount(gfa_raw, "GFA", positive=True),
            land_cost_hkd=read_amount(land_raw, "土地成本"),
            construction_cost_per_gfa_sqft=read_amount(construction_raw, "建造成本"),
            selling_price_per_saleable_sqft=read_amount(asp_raw, "預期售價"),
            tenure_years=tenure,
            interest_rate_pct=rate,
            ltv_pct=ltv,
            saleable_efficiency_pct=efficiency,
            marketing_rate_pct=marketing,
            average_debt_utilization_pct=utilization,
        )
        calculate_project(inputs)  # validates before either module renders
    except ValueError as exc:
        st.error(str(exc))
        st.stop()
    return project_choice, inputs


def macro_module() -> None:
    st.header("📈 Revenue, Construction & Land Cost — Macro Trends")
    workbook = Path(__file__).resolve().parent / "HK_Developer_Margin_Master_v3.xlsx"
    if not workbook.is_file():
        st.warning(
            f"未找到 {workbook.name}。請將原本的 Excel 放在 app.py 同一資料夾。"
            "其餘兩個模組無需 Excel，仍可使用。"
        )
        return
    try:
        df = load_macro_data(str(workbook), workbook.stat().st_mtime_ns)
    except (OSError, ValueError, ImportError) as exc:
        st.error(f"讀取 Excel 失敗：{exc}")
        return

    if not {"CCL_Norm", "TPI_Norm"}.issubset(df.columns):
        st.error("Excel 必須具有 CCL_Norm 及 TPI_Norm 欄位。")
        return

    # Rebase on the FIRST COMMON, POSITIVE observation in the full dataset;
    # never rebase again merely because the user adjusts the timeline.
    valid = df.dropna(subset=["CCL_Norm", "TPI_Norm"]).copy()
    valid = valid[(valid["CCL_Norm"] > 0) & (valid["TPI_Norm"] > 0)]
    if valid.empty:
        st.error("沒有同時為正數的樓價／投標價格指數觀察值。")
        return
    base_date = valid.iloc[0]["Date"]
    valid["CCI_Rebased"] = valid["CCL_Norm"] / valid.iloc[0]["CCL_Norm"] * 100
    valid["TPI_Rebased"] = valid["TPI_Norm"] / valid.iloc[0]["TPI_Norm"] * 100

    land_available = (
        "Land_Index" in valid.columns
        and valid["Land_Index"].notna().all()
        and (valid["Land_Index"] > 0).all()
    )
    if land_available:
        valid["Land_Rebased"] = valid["Land_Index"] / valid.iloc[0]["Land_Index"] * 100
        land_label = "Land_Index from workbook (source unverified)"
    else:
        st.warning(
            "沒有完整、正數且可用的 Land_Index。原程式會用線性公式推造地價指數，"
            "可能出現負值；新版本不會將這個推算冒充實際市場數據。"
        )
        proxy_enabled = st.checkbox("顯示明確標示的示意地價代理指數", value=False)
        if proxy_enabled:
            valid["Land_Rebased"] = 100 * (valid["CCI_Rebased"] / 100) ** 1.55
            land_label = "Illustrative land proxy (CCI elasticity 1.55; NOT data)"
        else:
            land_label = None

    view_mode = st.radio(
        "Display Mode:",
        ["Rebased index (common starting date = 100)", "Illustrative HKD/sqft anchors"],
        horizontal=True,
    )
    start_year, end_year = int(valid.Date.dt.year.min()), int(valid.Date.dt.year.max())
    if start_year == end_year:
        selected_years = (start_year, end_year)
        st.caption(f"可用數據年份：{start_year}")
    else:
        selected_years = st.slider(
            "Timeline Range", start_year, end_year, (start_year, end_year)
        )
    displayed = valid[
        valid["Date"].dt.year.between(selected_years[0], selected_years[1])
    ].copy()
    st.caption(
        f"共同基準日期：{base_date:%Y-%m-%d}，兩項指數均設為 100。"
        "此處的指數／案例數值尚未附原始出處，請勿視作已驗證研究數據。"
    )

    chart_cols = [
        ("CCI_Rebased", "Centa-City (CCL_Norm)", "#e74c3c"),
        ("TPI_Rebased", "Tender Price (TPI_Norm)", "#3498db"),
    ]
    if land_label:
        chart_cols.append(("Land_Rebased", land_label, "#f1c40f"))

    if view_mode.startswith("Illustrative"):
        st.warning(
            "下列每呎港元曲線只按使用者輸入的『基準單價 × 指數變化』示意換算，"
            "不是指數供應者發布的歷史實際單價。銷售／施工／土地的面積及成本口徑"
            "未經統一前不可將差額解讀為實際發展利潤。"
        )
        a, b, c = st.columns(3)
        cci_anchor = a.number_input("Base sale price (HK$/saleable sqft)", min_value=0.0, value=15000.0, step=500.0)
        tpi_anchor = b.number_input("Base construction cost (HK$/GFA sqft)", min_value=0.0, value=4500.0, step=250.0)
        land_anchor = c.number_input("Base land cost (HK$/GFA sqft)", min_value=0.0, value=7500.0, step=250.0)
        anchors = {"CCI_Rebased": cci_anchor, "TPI_Rebased": tpi_anchor, "Land_Rebased": land_anchor}
        ylabel = "Illustrative HKD / sqft (different area bases)"
    else:
        anchors = {name: 100.0 for name, _, _ in chart_cols}
        ylabel = f"Index ({base_date:%Y-%m-%d} = 100)"

    indicators = [
        ("HIBOR_3M", "3-month HIBOR (%)", "#ff7f0e"),
        ("Unemployment_Rate", "Unemployment (%)", "#2ca02c"),
        ("M2_Index", "M2 index (as supplied)", "#9467bd"),
    ]
    indicators = [row for row in indicators if row[0] in displayed.columns]
    fig, axes = plt.subplots(
        1 + len(indicators), 1, figsize=(12, 5 + 2.3 * len(indicators)),
        sharex=True, gridspec_kw={"height_ratios": [2.5] + [1] * len(indicators)},
    )
    axes = np.atleast_1d(axes)
    fig.patch.set_facecolor("#0E1117")
    for name, label, color in chart_cols:
        axes[0].plot(
            displayed.Date, displayed[name] * anchors[name] / 100,
            label=label, color=color, linewidth=2.3,
        )
    axes[0].set_ylabel(ylabel)
    for ax, (field, label, color) in zip(axes[1:], indicators):
        ax.plot(displayed.Date, displayed[field], label=label, color=color)
        ax.set_ylabel(label)
    for ax in axes:
        ax.grid(True, alpha=0.16)
        ax.legend(loc="best", fontsize=9, facecolor="#262730", labelcolor="white")
        ax.tick_params(colors="white")
    fig.tight_layout(pad=2)
    st.pyplot(fig, clear_figure=True)
    plt.close(fig)


def sensitivity_module() -> None:
    st.header("🎛️ Project-linked Sensitivity Matrix")
    st.caption("全部格子由下方相同的項目模型重新計算；不再使用任意 Operating Leverage 及 Interest Impact 係數。")
    _, inputs = project_inputs_form("sensitivity")
    base = calculate_project(inputs)
    st.metric("Base Case: Cumulative ROI on Cost", f"{base.roi_on_cost:.1%}")
    price_shocks = [10, 5, 0, -5, -10, -15, -20]
    # A nonnegative, strictly increasing rate grid avoids duplicate headings
    # when a project's base interest rate is below 3%.
    rate_start = max(0.0, inputs.interest_rate_pct - 3.0)
    rates = [round(rate_start + offset, 2) for offset in range(7)]
    grid = sensitivity_grid(inputs, price_shocks, rates)
    result_df = pd.DataFrame(
        grid,
        index=[f"Sale price {shock:+d}%" for shock in price_shocks],
        columns=[f"Rate {rate:.1f}%" for rate in rates],
    )
    fig, ax = plt.subplots(figsize=(12, 7))
    fig.patch.set_facecolor("#0E1117")
    sns.heatmap(result_df, annot=True, fmt=".1%", cmap="RdYlGn", center=0, ax=ax)
    ax.set_title("Cumulative project ROI on cost (NOT IRR)")
    ax.tick_params(colors="white")
    fig.tight_layout()
    st.pyplot(fig, clear_figure=True)
    plt.close(fig)
    st.caption(
        "橫軸是絕對年利率、縱軸是相對預設售價的變化；"
        "固定成本、LTV、實用率、營銷費及平均借款使用率均沿用上述輸入。"
        "這只是 deterministic scenarios，並非預測概率。"
    )


def roi_module() -> None:
    st.header("🏗️ Illustrative Project ROI Calculator")
    chosen, inputs = project_inputs_form("roi")
    result = calculate_project(inputs)
    st.divider()
    a, b, c = st.columns(3)
    a.metric("Estimated Net Profit (HK$100m)", f"{result.profit_hkd / 1e8:,.2f}")
    b.metric("Break-even (HK$/saleable sqft)", f"{result.break_even_per_saleable_sqft:,.0f}")
    c.metric("Cumulative ROI on Cost (%)", f"{result.roi_on_cost:.1%}")
    st.caption(
        f"Revenue HK${result.revenue_hkd:,.0f} · Fixed loan principal HK${result.loan_principal_hkd:,.0f} · "
        f"Interest HK${result.interest_hkd:,.0f} · Saleable area {result.saleable_sqft:,.0f} sqft. "
        "ROI = 利潤 ÷ 包含營銷及利息的總成本，屬整段發展期累計回報，非 IRR／股本回報。"
    )

    labels = [
        f"Land (HK${result.land_hkd:,.0f})",
        f"Construction (HK${result.construction_hkd:,.0f})",
        f"Interest (HK${result.interest_hkd:,.0f})",
        f"Marketing (HK${result.marketing_hkd:,.0f})",
    ]
    values = [result.land_hkd, result.construction_hkd, result.interest_hkd, result.marketing_hkd]
    fig, ax = plt.subplots(figsize=(10, 6))
    fig.patch.set_facecolor("#0E1117")
    ax.pie(
        values, labels=labels, autopct="%1.1f%%", startangle=90,
        labeldistance=1.15,
        colors=["#1f77b4", "#ff7f0e", "#d62728", "#9467bd"],
        textprops={"color": "white"},
    )
    fig.tight_layout()
    st.pyplot(fig, clear_figure=True)
    plt.close(fig)

    st.subheader("💡 Scenario interpretation")
    cushion = (
        inputs.selling_price_per_saleable_sqft / result.break_even_per_saleable_sqft - 1
    ) * 100
    summary = (
        f"**{chosen}**：預設假設不代表已核實的真實交易／利潤。"
        f"在售價 HK${inputs.selling_price_per_saleable_sqft:,.0f}/實用呎、"
        f"LTV {inputs.ltv_pct:.0f}%、年息 {inputs.interest_rate_pct:.1f}%、"
        f"平均欠款使用率 {inputs.average_debt_utilization_pct:.0f}% 的條件下，"
        f"模型計算的保本價約 HK${result.break_even_per_saleable_sqft:,.0f}/實用呎，"
        f"售價相對保本價差距為 {cushion:+.1f}%。"
    )
    if result.profit_hkd < 0:
        st.error(summary + " 這個情景下項目虧損；需再檢視融資及現金流壓力。")
    elif cushion < 10:
        st.warning(summary + " 安全邊際偏薄；小幅成本或售價偏差可能令盈利轉負。")
    else:
        st.success(summary + " 在所列假設下仍有正盈利緩衝，但不代表實際項目已獲利。")
    if inputs.ltv_pct >= 70:
        st.info(
            "高 LTV 代表財務風險值得留意，但資產下跌不會自動等於收回貸款或接管。"
            "是否違約須查看貸款契約、估值、償債安排及債權人行動；本模型不能預測 Receivership。"
        )


def main() -> None:
    st.title("🏢 HK Real Estate BI Dashboard (V5.14)")
    st.markdown("**Auditable assumptions · Correct break-even · Project-linked scenarios**")
    st.sidebar.header("⚙️ Module Switcher")
    analysis_type = st.sidebar.radio(
        "Select Analysis View:",
        ["1. Macro Trends (CCI/TPI/LPI)", "2. Sensitivity Matrix", "3. Project ROI Calculator"],
    )
    with st.expander("📚 Data provenance & methodology (read before interpretation)"):
        st.markdown(
            "- **宏觀數據：** 從 `HK_Developer_Margin_Master_v3.xlsx` 讀入。"
            "未提供原始數據、官方來源連結、抽樣口徑及更新日期，故本程式不能獨立核實數據。\n"
            "- **Land_Index：** Excel 未提供時預設不顯示。可選的代理線只作示意，"
            "是 CCI 的 1.55 次方變換，**不是**實際地價指數。\n"
            "- **每呎價格圖：** 按可更改的基準價格換算的 illustrative series，並非歷史成交價。\n"
            "- **項目案例：** 11 個原始案例輸入保留作假設情景，未核實是否符合真實項目。\n"
            "- **財務：** 建造成本 × GFA；售價 × 實用面積；營銷費是收入比例。"
            "貸款利息採用簡單利息及平均借款比例，**未**處理分期付款、稅項、"
            "預售回款、銷售速度、IRR、違約契約或接管程序。\n"
            "- **研究用途：** 使用前應補齊 Data Dictionary、原始來源、"
            "面積口徑、估算依據、基準情景及版本日期。"
        )
    if analysis_type.startswith("1."):
        macro_module()
    elif analysis_type.startswith("2."):
        sensitivity_module()
    else:
        roi_module()


if __name__ == "__main__":
    main()
