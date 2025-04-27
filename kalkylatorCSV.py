import os
import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
from openpyxl import load_workbook

# ----- Konfig för sparande till Excel -----
EXCEL_FILE = "Resultat.xlsx"
SHEET_NAME = "Resultat"

def append_df_to_excel(df, filename=EXCEL_FILE, sheet_name=SHEET_NAME):
    """Appendar en DataFrame till ett Excel-blad, skapar fil/blad vid behov."""
    if not os.path.exists(filename):
        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        wb = load_workbook(filename)
        if sheet_name not in wb.sheetnames:
            ws = wb.create_sheet(sheet_name)
            ws.append(list(df.columns))
        else:
            ws = wb[sheet_name]
        for row in df.itertuples(index=False):
            ws.append(list(row))
        wb.save(filename)

# ----- Layout & Styling -----
st.set_page_config(page_title="CO₂ Massflödes Kalkylator v2.10", layout="wide")
st.markdown("""
    <style>
    .block-container { padding:2rem 3rem; }
    .big-title { font-size:2.5rem; font-weight:700; margin-bottom:1rem; text-align:center; }
    .stHeader, .stSubheader { margin-top:2rem; }
    </style>
""", unsafe_allow_html=True)
st.markdown('<div class="big-title">CO₂ Massflödes Kalkylator v2.10</div>', unsafe_allow_html=True)

# ----- Inmatningsläge -----
mode = st.radio("Välj inmatningsmetod:", ("Manuell inmatning", "Ladda upp CSV"))

# ----- Sidebar-filter -----
with st.sidebar:
    st.header("📐 Filterinställningar")
    rotor_diameter   = st.number_input("Rotor diameter (mm)",    min_value=1,  value=350)
    rotor_depth      = st.number_input("Rotor depth (mm)",       min_value=1,  value=100)
    active_pct       = st.number_input("Active area (%)",         min_value=1,  max_value=100, value=95)
    sector_deg_proc  = st.number_input("Absorbering sector (°)",  min_value=1,  max_value=360, value=270)
    sector_deg_regen = st.number_input("Regenerering sector (°)", min_value=1,  max_value=360, value=90)

# ----- Beräkna areor -----
rotor_m2      = np.pi * (rotor_diameter / 1000)**2 / 4
area_proc_m2  = rotor_m2 * (active_pct / 100) * (sector_deg_proc / 360)
area_regen_m2 = rotor_m2 * (active_pct / 100) * (sector_deg_regen / 360)

# ----- Hjälpfunktioner -----
def calc_density(T):
    return 1.293 * 273.15 / (273.15 + T)

def calc_abs_humidity(T, RH):
    es = 6.112 * np.exp((17.62 * T) / (243.12 + T))
    e  = RH / 100 * es
    w  = 0.622 * e / (1013.25 - e)
    return w * 1000  # g vatten/kg torr luft

# ----- Manuellt läge (ingen sparfunktion) -----
if mode == "Manuell inmatning":
    st.header("⚙️ Manuella parametrar")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Absorbering IN/UT")
        flow_in_proc  = st.number_input("Flöde IN (l/s)", 0.0, 500.0, 73.0, key="m_flow_in_proc")
        T_in_proc     = st.number_input("Temp IN (°C)", -20.0, 200.0, 20.0, key="m_T_in_proc")
        RH_in_proc    = st.number_input("RH IN (%)",    0.0, 100.0, 30.0, key="m_RH_in_proc")
        T_out_proc    = st.number_input("Temp UT (°C)", -20.0, 200.0, 25.0, key="m_T_out_proc")
        RH_out_proc   = st.number_input("RH UT (%)",    0.0, 100.0, 20.0, key="m_RH_out_proc")
    with c2:
        st.subheader("Regenerering IN/UT")
        flow_in_reg   = st.number_input("Flöde IN (l/s)", 0.0, 500.0, 30.0, key="m_flow_in_reg")
        T_in_reg      = st.number_input("Temp IN (°C)", -20.0, 200.0, 23.0, key="m_T_in_reg")
        RH_in_reg     = st.number_input("RH IN (%)",    0.0, 100.0, 60.0, key="m_RH_in_reg")
        T_out_reg     = st.number_input("Temp UT (°C)", -20.0, 200.0, 40.0, key="m_T_out_reg")
        RH_out_reg    = st.number_input("RH UT (%)",    0.0, 100.0, 50.0, key="m_RH_out_reg")

    # Beräkningar
    rho_in_proc   = calc_density(T_in_proc)
    rho_out_proc  = calc_density(T_out_proc)
    rho_in_reg    = calc_density(T_in_reg)
    rho_out_reg   = calc_density(T_out_reg)

    ah_in_proc    = calc_abs_humidity(T_in_proc, RH_in_proc)
    ah_out_proc   = calc_abs_humidity(T_out_proc, RH_out_proc)
    ah_in_reg     = calc_abs_humidity(T_in_reg,   RH_in_reg)
    ah_out_reg    = calc_abs_humidity(T_out_reg,  RH_out_reg)

    mf_in_proc    = rho_in_proc  * (flow_in_proc / 1000) / area_proc_m2
    mf_out_proc   = mf_in_proc
    mf_in_reg     = rho_in_reg   * (flow_in_reg  / 1000) / area_regen_m2
    mf_out_reg    = mf_in_reg

    vol_out_proc  = mf_out_proc * area_proc_m2 / rho_out_proc * 1000
    vol_out_reg   = mf_out_reg  * area_regen_m2 / rho_out_reg  * 1000

    ct_proc       = (rotor_depth / 1000) / ((flow_in_proc / 1000) / area_proc_m2)
    ct_reg        = (rotor_depth / 1000) / ((flow_in_reg  / 1000) / area_regen_m2)

    with st.expander("📊 Resultat (Manuellt)", expanded=True):
        st.markdown("### Absorbering")
        st.write(f"• Massflöde IN:    {mf_in_proc:.3f} kg/m²/s")
        st.write(f"• Massflöde UT:    {mf_out_proc:.3f} kg/m²/s")
        st.write(f"• Fukt IN:         {ah_in_proc:.1f} g/kg")
        st.write(f"• Fukt UT:         {ah_out_proc:.1f} g/kg")
        st.write(f"• Volymflöde IN:   {flow_in_proc:.1f} l/s")
        st.write(f"• Volymflöde UT:   {vol_out_proc:.1f} l/s")
        st.write(f"• Kontakttid:      {ct_proc:.2f} s")

        st.markdown("### Regenerering")
        st.write(f"• Massflöde IN:    {mf_in_reg:.3f} kg/m²/s")
        st.write(f"• Massflöde UT:    {mf_out_reg:.3f} kg/m²/s")
        st.write(f"• Fukt IN:         {ah_in_reg:.1f} g/kg")
        st.write(f"• Fukt UT:         {ah_out_reg:.1f} g/kg")
        st.write(f"• Volymflöde IN:   {flow_in_reg:.1f} l/s")
        st.write(f"• Volymflöde UT:   {vol_out_reg:.1f} l/s")
        st.write(f"• Kontakttid:      {ct_reg:.2f} s")


# ----- CSV-läge med “Spara”-knapp -----
else:
    st.header("📂 Ladda upp CSV för automatisk beräkning")
    uploaded = st.file_uploader("Välj CSV", type="csv", key="csvup")
    if uploaded:
        df = pd.read_csv(uploaded)
        df.rename(columns={
            "GX1_Temp":    "GX1_TEMP",
            "capacity_reg":"CAPACITY_REG",
            "capacity_abs":"CAPACITY_ABS"
        }, inplace=True)

        # Radvisa beräkningar
        df["rho_in_proc"]   = calc_density(df["GX3_TEMP"])
        df["rho_out_proc"]  = calc_density(df["GX4_TEMP"])
        df["rho_in_reg"]    = calc_density(df["GX2_TEMP"])
        df["rho_out_reg"]   = calc_density(df["GX1_TEMP"])

        df["ah_in_proc"]    = calc_abs_humidity(df["GX3_TEMP"], df["GX3_RH"])
        df["ah_out_proc"]   = calc_abs_humidity(df["GX4_TEMP"], df["GX4_RH"])
        df["ah_in_reg"]     = calc_abs_humidity(df["GX2_TEMP"], df["GX2_RH"])
        df["ah_out_reg"]    = calc_abs_humidity(df["GX1_TEMP"], df["GX1_RH"])

        df["flow_in_proc"]  = df["FLOW_Q2"]
        df["flow_in_reg"]   = df["FLOW_Q1"]

        df["mf_in_proc"]    = df["rho_in_proc"]  * (df["flow_in_proc"] / 1000) / area_proc_m2
        df["mf_out_proc"]   = df["mf_in_proc"]
        df["mf_in_reg"]     = df["rho_in_reg"]   * (df["flow_in_reg"] / 1000) / area_regen_m2
        df["mf_out_reg"]    = df["mf_in_reg"]

        df["vol_out_proc"]  = df["mf_out_proc"] * area_proc_m2 / df["rho_out_proc"] * 1000
        df["vol_out_reg"]   = df["mf_out_reg"]  * area_regen_m2 / df["rho_out_reg"]  * 1000

        df["ct_proc"]       = (rotor_depth / 1000) / ((df["flow_in_proc"] / 1000) / area_proc_m2)
        df["ct_reg"]        = (rotor_depth / 1000) / ((df["flow_in_reg"]  / 1000) / area_regen_m2)
        df["water_added_g_h"]= (df["ah_out_reg"] - df["ah_in_reg"]) * (df["rho_in_reg"]*(df["flow_in_reg"] / 1000)) * 3600

        # Aggrera medelvärden
        res = {
            "Abs IN mf (kg/m²/s)":    df["mf_in_proc"].mean(),
            "Reg IN mf (kg/m²/s)":    df["mf_in_reg"].mean(),
            "Diff mf (kg/m²/s)":      df["mf_in_proc"].mean() - df["mf_in_reg"].mean(),
            "Abs IN vol (l/s)":       df["flow_in_proc"].mean(),
            "Abs UT vol (l/s)":       df["vol_out_proc"].mean(),
            "Reg IN vol (l/s)":       df["flow_in_reg"].mean(),
            "Reg UT vol (l/s)":       df["vol_out_reg"].mean(),
            "Abs IN ah (g/kg)":       df["ah_in_proc"].mean(),
            "Abs UT ah (g/kg)":       df["ah_out_proc"].mean(),
            "Reg IN ah (g/kg)":       df["ah_in_reg"].mean(),
            "Reg UT ah (g/kg)":       df["ah_out_reg"].mean(),
            "Kontakttid Abs (s)":     df["ct_proc"].mean(),
            "Kontakttid Reg (s)":     df["ct_reg"].mean(),
            "Vatten tillsatt (g/h)":  df["water_added_g_h"].mean(),
            "CO₂-upptag regen":       df["CAPACITY_REG"].mean(),
            "CO₂-upptag abs":         df["CAPACITY_ABS"].mean(),
        }
        result = pd.Series(res, name="Mean Value")
        st.dataframe(result.to_frame().T, use_container_width=True)

        # Förbered DataFrame med metadata
        df_res = result.to_frame().T.reset_index(drop=True)
        df_res["Mode"]       = "CSV"
        df_res["SourceFile"] = uploaded.name

        # Ladda ner som CSV
        csv_out = df_res.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Ladda ner resultat som CSV", csv_out, "resultat.csv", "text/csv")

        # Spara till Excel först när knapp trycks
        if st.button("💾 Spara resultat till Excel"):
             append_df_to_excel(df_res)  # Spara på servern
             st.success(f"Resultatet från `{uploaded.name}` har sparats i `{EXCEL_FILE}` på servern")

             # Läs in Excelfilen igen
             with open(EXCEL_FILE, "rb") as f:
             excel_bytes = f.read()

    # Gör nedladdningsknapp
    st.download_button(
        label="⬇️ Ladda ner hela Resultat.xlsx",
        data=excel_bytes,
        file_name=EXCEL_FILE,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


        # … dina grafer som tidigare …
        chart_col, _ = st.columns([3, 1])
        with chart_col:
            # (oförändrade grafer) …
            pass
# … efter st.download_button …
        # Här lägger vi graferna i en smal kolumn (≈75% bredd)
        chart_col, _ = st.columns([3, 1])
        with chart_col:

            # Massflöde ABS vs REG (kg/m²/s)
            st.subheader("📦 Massflöde ABS vs REG (kg/m²/s)")
            abs_mf  = res["Abs IN mf (kg/m²/s)"]
            reg_mf  = res["Reg IN mf (kg/m²/s)"]
            diff_mf = res["Diff mf (kg/m²/s)"]
            mf_df = pd.DataFrame({
                "Kategori": ["ABS", "REG", "DIFF"],
                "Värde":    [abs_mf, reg_mf, diff_mf]
            })
            mf_chart = (
                alt.Chart(mf_df)
                   .mark_bar(size=80)
                   .encode(
                       x=alt.X("Kategori:N", scale=alt.Scale(paddingInner=0.2)),
                       y=alt.Y("Värde:Q", title="kg/m²/s"),
                       color="Kategori:N"
                   )
                   .properties(width=500, height=250)
            )
            st.altair_chart(mf_chart, use_container_width=False)

            # Volymflöden
            st.subheader("🌬️ Volymflöden (l/s)")
            cats = ["Abs IN","Abs UT","Reg IN","Reg UT"]
            df_v = pd.DataFrame({
                "Kategori": cats,
                "Värde": [
                    res["Abs IN vol (l/s)"],
                    res["Abs UT vol (l/s)"],
                    res["Reg IN vol (l/s)"],
                    res["Reg UT vol (l/s)"],
                ]
            })
            vol_chart = (
                alt.Chart(df_v)
                   .mark_bar(size=80)
                   .encode(
                       x="Kategori:N",
                       y="Värde:Q",
                       color="Kategori:N"
                   )
                   .properties(width=600, height=250)
            )
            st.altair_chart(vol_chart, use_container_width=False)

            # Absolut fukt
            st.subheader("💧 Absolut fukt (g/kg)")
            df_h = pd.DataFrame({
                "Kategori": cats,
                "Värde": [
                    res["Abs IN ah (g/kg)"],
                    res["Abs UT ah (g/kg)"],
                    res["Reg IN ah (g/kg)"],
                    res["Reg UT ah (g/kg)"],
                ]
            })
            hum_chart = (
                alt.Chart(df_h)
                   .mark_bar(size=80)
                   .encode(
                       x="Kategori:N",
                       y="Värde:Q",
                       color="Kategori:N"
                   )
                   .properties(width=600, height=250)
            )
            st.altair_chart(hum_chart, use_container_width=False)

            # Vatten tillsatt
            st.subheader("💦 Tillsatt vatten (g/h)")
            df_w = pd.DataFrame({
                "Kategori": ["Vatten tillsatt"],
                "Värde":    [res["Vatten tillsatt (g/h)"]],
            })
            water_chart = (
                alt.Chart(df_w)
                   .mark_bar(size=60, color="#1f77b4")
                   .encode(
                       x="Kategori:N",
                       y="Värde:Q"
                   )
                   .properties(width=200, height=250)
            )
            st.altair_chart(water_chart, use_container_width=False)
