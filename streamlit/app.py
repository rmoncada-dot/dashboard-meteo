import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import json
import os
from pathlib import Path
from io import StringIO

# ============================================================
#  CONFIG
# ============================================================
st.set_page_config(
    page_title="Dashboard Meteo Stazioni",
    page_icon="🌬️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
#  STILE CSS
# ============================================================
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1F4E79, #2E75B6);
        color: white;
        padding: 1.2rem 1.5rem;
        border-radius: 8px;
        margin-bottom: 1rem;
    }
    .kpi-card {
        background: #f0f5fc;
        border-radius: 8px;
        padding: 1rem;
        text-align: center;
        border-left: 4px solid #2E75B6;
    }
    .kpi-value { font-size: 1.8rem; font-weight: 700; color: #1F4E79; }
    .kpi-label { font-size: 0.75rem; color: #666; margin-bottom: 4px; }
    .kpi-sub   { font-size: 0.7rem; color: #999; margin-top: 2px; }
    .warn-box {
        background: #fff8e1;
        border-left: 3px solid #f0b429;
        padding: 0.6rem 1rem;
        border-radius: 0 6px 6px 0;
        font-size: 0.85rem;
        color: #7f5c00;
    }
    .p50-pill  { background:#d4edda; color:#155724; padding:3px 10px; border-radius:10px; font-weight:700; font-size:0.8rem; }
    .p75-pill  { background:#fff0b3; color:#5c4400; padding:3px 10px; border-radius:10px; font-weight:700; font-size:0.8rem; }
    .p90-pill  { background:#ffe0b2; color:#6b3a00; padding:3px 10px; border-radius:10px; font-weight:700; font-size:0.8rem; }
</style>
""", unsafe_allow_html=True)

# ============================================================
#  COSTANTI
# ============================================================
WIND_COLS = {
    "TOP 92m": "TOP 92;wind_speed;Avg (m/s)",
    "RIF 88m": "RIF 88;wind_speed;Avg (m/s)",
    "RIF 70m": "RIF 70;wind_speed;Avg (m/s)",
    "RIF 50m": "RIF 50;wind_speed;Avg (m/s)",
}
HEIGHT_COLORS = {
    "TOP 92m": "#1F4E79",
    "RIF 88m": "#2E75B6",
    "RIF 70m": "#1D9E75",
    "RIF 50m": "#BA7517",
}
ANN_SEQ = [f"2025-{m:02d}" for m in range(4, 13)] + ["2026-01", "2026-02", "2026-03"]
CLR_MONTHS = px.colors.qualitative.Set2

# ============================================================
#  PARSING CSV
# ============================================================
@st.cache_data(show_spinner=False)
def parse_csv(content: str, fname: str) -> pd.DataFrame:
    sep = ";" if content.count(";") > content.count(",") else ","
    try:
        df = pd.read_csv(StringIO(content), sep=sep, low_memory=False)
    except Exception:
        return pd.DataFrame()

    if "datetime" not in df.columns:
        return pd.DataFrame()

    df["datetime"] = pd.to_datetime(df["datetime"], errors="coerce")
    df = df.dropna(subset=["datetime"])
    df["month"]  = df["datetime"].dt.strftime("%Y-%m")
    df["hour"]   = df["datetime"].dt.hour
    df["source"] = fname
    return df

@st.cache_data(show_spinner=False)
def load_json(content: str) -> pd.DataFrame:
    data = json.loads(content)
    return pd.DataFrame(data)

# ============================================================
#  STATISTICHE
# ============================================================
def exceedance_percentile(series: pd.Series, p: float) -> float:
    """Eccedenza: P50=50° pct stat, P75=25° pct stat, P90=10° pct stat"""
    clean = series.dropna()
    if len(clean) == 0:
        return np.nan
    stat_p = 100 - p
    return float(np.percentile(clean, stat_p))

def compute_monthly_stats(df: pd.DataFrame) -> pd.DataFrame:
    wind_col = WIND_COLS["TOP 92m"]
    temp_col = "TEMP-UMID;temperature;Avg (°C)"
    hum_col  = "TEMP-UMID;humidity;Avg (%)"
    pres_col = "GEOVES BOX;air_pressure;Avg (hPa)"

    rows = []
    for month, g in df.groupby("month"):
        wind = g[wind_col].dropna() if wind_col in g else pd.Series(dtype=float)
        wind = wind[wind > 0]
        temp = g[temp_col].dropna()[(g[temp_col] > -5) & (g[temp_col] < 60)] if temp_col in g else pd.Series(dtype=float)
        hum  = g[hum_col].dropna()  if hum_col  in g else pd.Series(dtype=float)
        pres = g[pres_col].dropna()[(g[pres_col] > 900) & (g[pres_col] < 1100)] if pres_col in g else pd.Series(dtype=float)

        # Multi-height averages
        h_avgs = {}
        for lbl, col in WIND_COLS.items():
            if col in g.columns:
                v = g[col].dropna()
                v = v[v > 0]
                h_avgs[lbl] = round(v.mean(), 3) if len(v) else np.nan
            else:
                h_avgs[lbl] = np.nan

        # Shear alpha
        v92 = h_avgs.get("TOP 92m", np.nan)
        v50 = h_avgs.get("RIF 50m", np.nan)
        shear = round(np.log(v92 / v50) / np.log(92 / 50), 4) if v92 and v50 and v92 > 0 and v50 > 0 else np.nan

        # Availability
        expected = 144 * 30  # approx
        avail = min(100, round(len(g) / expected * 100, 1))

        rows.append({
            "month": month,
            "misurazioni": len(g),
            "wind_avg": round(wind.mean(), 3) if len(wind) else np.nan,
            "wind_max": round(wind.max(), 3)  if len(wind) else np.nan,
            "p50": exceedance_percentile(wind, 50),
            "p75": exceedance_percentile(wind, 75),
            "p90": exceedance_percentile(wind, 90),
            **{f"wind_{k.replace(' ','').lower()}": v for k, v in h_avgs.items()},
            "shear_alpha": shear,
            "temp_avg": round(temp.mean(), 2) if len(temp) else np.nan,
            "temp_max": round(temp.max(), 2)  if len(temp) else np.nan,
            "temp_min": round(temp.min(), 2)  if len(temp) else np.nan,
            "hum_avg":  round(hum.mean(), 1)  if len(hum)  else np.nan,
            "pres_avg": round(pres.mean(), 1) if len(pres) else np.nan,
            "avail_pct": avail,
            "anomalous": month.endswith("-03") and len(wind) > 0 and wind.mean() > 8,
        })
    return pd.DataFrame(rows).sort_values("month").reset_index(drop=True)

# ============================================================
#  SIDEBAR
# ============================================================
with st.sidebar:
    st.markdown("## 🌬️ Dashboard Meteo")
    st.markdown("---")

    data_source = st.radio("Sorgente dati", ["📂 Carica CSV", "📊 Carica JSON (da Excel)", "📡 Demo (dati G243043)"])

    df_raw = pd.DataFrame()
    df_stats = pd.DataFrame()
    station_name = "G243043"
    station_loc  = "Durrà"

    if data_source == "📂 Carica CSV":
        uploaded = st.file_uploader("Carica CSV mensili", type="csv", accept_multiple_files=True)
        station_name = st.text_input("Codice stazione", "G243043")
        station_loc  = st.text_input("Località", "Durrà")
        if uploaded:
            dfs = []
            for f in uploaded:
                content = f.read().decode("utf-8", errors="ignore")
                parsed = parse_csv(content, f.name)
                if not parsed.empty:
                    dfs.append(parsed)
            if dfs:
                df_raw = pd.concat(dfs).sort_values("datetime").reset_index(drop=True)
                df_stats = compute_monthly_stats(df_raw)
                st.success(f"✅ {len(dfs)} file · {len(df_raw):,} righe")

    elif data_source == "📊 Carica JSON (da Excel)":
        json_file = st.file_uploader("Carica export_streamlit.json", type="json")
        station_name = st.text_input("Codice stazione", "G243043")
        station_loc  = st.text_input("Località", "Durrà")
        if json_file:
            content = json_file.read().decode("utf-8")
            df_stats = load_json(content)
            # Rename columns to match internal names
            col_map = {
                "mese": "month", "vento_top92_avg": "wind_avg",
                "vento_top92_max": "wind_max", "disponibilita_pct": "avail_pct",
                "vento_88m": "wind_rif88m", "vento_70m": "wind_rif70m", "vento_50m": "wind_rif50m",
                "vento_top92m": "wind_top92m", "shear_alpha": "shear_alpha",
                "temp_avg": "temp_avg", "temp_max": "temp_max", "temp_min": "temp_min",
                "umidita_avg": "hum_avg", "pressione_avg": "pres_avg",
            }
            df_stats = df_stats.rename(columns={k: v for k, v in col_map.items() if k in df_stats.columns})
            st.success(f"✅ JSON caricato · {len(df_stats)} mesi")

    else:  # Demo
        # Dati embedded G243043
        demo_data = {
            "month":      ["2025-03","2025-04","2025-05","2025-06"],
            "misurazioni":[1528, 4320, 4464, 4320],
            "wind_avg":   [9.166, 5.864, 5.854, 4.675],
            "wind_max":   [17.48, 17.78, 18.02, 14.23],
            "p50":        [9.098, 5.259, 5.393, 4.538],
            "p75":        [6.599, 3.247, 3.750, 2.950],
            "p90":        [3.906, 1.981, 2.483, 1.789],
            "wind_top92m":[9.166, 5.864, 5.854, 4.675],
            "wind_rif88m":[8.796, 5.640, 5.612, 4.535],
            "wind_rif70m":[7.976, 5.218, 5.275, 4.292],
            "wind_rif50m":[7.536, 4.993, 5.007, 4.075],
            "shear_alpha":[None, 0.2638, 0.2564, 0.2250],
            "temp_avg":   [None, 13.89, 18.54, 24.44],
            "temp_max":   [None, 21.71, 27.48, 34.95],
            "temp_min":   [None, 6.45,  12.65, 14.30],
            "hum_avg":    [75.2, 70.3, 60.1, 49.9],
            "pres_avg":   [None, 969.8, 969.4, 972.3],
            "avail_pct":  [100, 100, 100, 100],
            "anomalous":  [True, False, False, False],
        }
        df_stats = pd.DataFrame(demo_data)

    st.markdown("---")
    if not df_stats.empty:
        valid = df_stats[~df_stats.get("anomalous", pd.Series(False, index=df_stats.index))]
        ann_loaded = [m for m in ANN_SEQ if m in df_stats["month"].values]
        ann_valid  = [m for m in ann_loaded if m in valid["month"].values]
        pct = int(len(ann_valid) / 12 * 100)
        st.markdown(f"**Anno Apr 2025 – Mar 2026**")
        st.progress(pct / 100)
        st.caption(f"{len(ann_valid)}/12 mesi · {pct}% completato")

        if ann_valid:
            p50_ann = round(valid[valid["month"].isin(ann_valid)]["p50"].mean(), 2)
            p75_ann = round(valid[valid["month"].isin(ann_valid)]["p75"].mean(), 2)
            p90_ann = round(valid[valid["month"].isin(ann_valid)]["p90"].mean(), 2)
            st.markdown(f"""
            <span class='p50-pill'>P50: {p50_ann} m/s</span>&nbsp;
            <span class='p75-pill'>P75: {p75_ann} m/s</span>&nbsp;
            <span class='p90-pill'>P90: {p90_ann} m/s</span>
            """, unsafe_allow_html=True)

# ============================================================
#  HEADER
# ============================================================
st.markdown(f"""
<div class='main-header'>
  <h2 style='margin:0'>🌬️ Dashboard Meteo — {station_name}</h2>
  <p style='margin:0;opacity:.8;font-size:.9rem'>
    {station_loc} &nbsp;·&nbsp; Disponibilità · P50/P75/P90 · Profilo Vento · AEP
  </p>
</div>
""", unsafe_allow_html=True)

if df_stats.empty:
    st.info("👈 Carica i CSV o il JSON dalla barra laterale per avviare l'analisi.")
    st.stop()

# ============================================================
#  TABS
# ============================================================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Disponibilità",
    "📈 P50 / P75 / P90",
    "📐 Profilo Vento",
    "🌡️ Meteo",
    "⚡ AEP"
])

valid_stats = df_stats[~df_stats.get("anomalous", pd.Series(False, index=df_stats.index))]
months_labels = df_stats["month"].tolist()

# ---- TAB 1: DISPONIBILITÀ ----
with tab1:
    if df_stats["avail_pct"].notna().any():
        c1, c2, c3, c4 = st.columns(4)
        total_rec = int(df_stats["misurazioni"].sum())
        avg_avail = round(df_stats["avail_pct"].mean(), 1)
        n_months  = len(df_stats)
        gaps      = int((df_stats["avail_pct"] < 98).sum())

        for col, lbl, val, sub in [
            (c1, "Misurazioni totali", f"{total_rec:,}", "record a 10 min"),
            (c2, "Disponibilità media", f"{avg_avail}%", "su tutti i mesi"),
            (c3, "Mesi analizzati", str(n_months), "totale"),
            (c4, "Mesi sotto 98%", str(gaps), "buchi rilevati"),
        ]:
            col.markdown(f"""
            <div class='kpi-card'>
              <div class='kpi-label'>{lbl}</div>
              <div class='kpi-value'>{val}</div>
              <div class='kpi-sub'>{sub}</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        colors = ["#1D9E75" if v >= 98 else "#e6a817" if v >= 90 else "#c0392b"
                  for v in df_stats["avail_pct"]]
        fig = go.Figure(go.Bar(
            x=df_stats["month"], y=df_stats["avail_pct"],
            marker_color=colors, text=df_stats["avail_pct"].apply(lambda x: f"{x}%"),
            textposition="outside"
        ))
        fig.update_layout(title="Disponibilità mensile (%)", yaxis_range=[0, 105],
                          yaxis_title="%", height=350, plot_bgcolor="#f8fbff")
        st.plotly_chart(fig, use_container_width=True)

        st.dataframe(
            df_stats[["month","misurazioni","avail_pct"]].rename(columns={
                "month":"Mese","misurazioni":"Misurazioni","avail_pct":"Disponibilità (%)"}),
            use_container_width=True, hide_index=True
        )

# ---- TAB 2: P50 / P75 / P90 ----
with tab2:
    if df_stats["anomalous"].any():
        st.markdown("<div class='warn-box'>⚠️ <b>Marzo 2025</b>: dati esclusi dall'analisi (calibrazione sensore). Convenzione eccedenza: P50 &gt; P75 &gt; P90.</div>", unsafe_allow_html=True)
        st.markdown("")

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=df_stats["month"], y=df_stats["p50"], name="P50 (50% eccedenza)",
        line=dict(color="#1D9E75", width=3), mode="lines+markers", marker=dict(size=8)))
    fig.add_trace(go.Scatter(x=df_stats["month"], y=df_stats["p75"], name="P75 (75% eccedenza)",
        line=dict(color="#d4a017", width=2, dash="dot"), mode="lines+markers", marker=dict(size=7)))
    fig.add_trace(go.Scatter(x=df_stats["month"], y=df_stats["p90"], name="P90 (90% eccedenza)",
        line=dict(color="#BA7517", width=2, dash="dash"), mode="lines+markers", marker=dict(size=7)))
    fig.update_layout(title="P50 / P75 / P90 — Velocità Vento TOP 92m (m/s)",
                      yaxis_title="m/s", height=380, plot_bgcolor="#f8fbff",
                      legend=dict(orientation="h", yanchor="bottom", y=1.02))
    st.plotly_chart(fig, use_container_width=True)

    # Cards mensili
    cols = st.columns(min(len(df_stats), 4))
    for i, (_, row) in enumerate(df_stats.iterrows()):
        with cols[i % 4]:
            p50v = f"{row['p50']:.2f}" if pd.notna(row['p50']) else "—"
            p75v = f"{row['p75']:.2f}" if pd.notna(row['p75']) else "—"
            p90v = f"{row['p90']:.2f}" if pd.notna(row['p90']) else "—"
            st.markdown(f"""
            <div class='kpi-card'>
              <b>{row['month']}</b><br>
              <span class='p50-pill'>P50: {p50v}</span><br>
              <span class='p75-pill'>P75: {p75v}</span><br>
              <span class='p90-pill'>P90: {p90v}</span>
            </div><br>""", unsafe_allow_html=True)

# ---- TAB 3: PROFILO VENTO ----
with tab3:
    h_cols = ["wind_top92m","wind_rif88m","wind_rif70m","wind_rif50m"]
    h_labels = ["TOP 92m","RIF 88m","RIF 70m","RIF 50m"]
    h_heights = [92, 88, 70, 50]

    available = [c for c in h_cols if c in df_stats.columns]
    if available:
        col1, col2 = st.columns(2)

        with col1:
            # Grouped bar per altezza
            fig = go.Figure()
            for lbl, col_name in zip(h_labels, h_cols):
                if col_name in df_stats.columns:
                    fig.add_trace(go.Bar(
                        name=lbl, x=df_stats["month"], y=df_stats[col_name],
                        marker_color=HEIGHT_COLORS[lbl]
                    ))
            fig.update_layout(barmode="group", title="Velocità media per altezza (m/s)",
                              yaxis_title="m/s", height=360, plot_bgcolor="#f8fbff",
                              legend=dict(orientation="h", yanchor="bottom", y=1.02))
            st.plotly_chart(fig, use_container_width=True)

        with col2:
            # Profilo verticale per mese
            fig2 = go.Figure()
            for i, (_, row) in enumerate(df_stats.iterrows()):
                speeds = [row.get(c, np.nan) for c in h_cols]
                if any(pd.notna(s) for s in speeds):
                    fig2.add_trace(go.Scatter(
                        x=speeds, y=h_heights,
                        name=row["month"], mode="lines+markers",
                        line=dict(color=CLR_MONTHS[i % len(CLR_MONTHS)], width=2),
                        marker=dict(size=8)
                    ))
            fig2.update_layout(
                title="Profilo verticale velocità (m/s vs altezza)",
                xaxis_title="Velocità (m/s)", yaxis_title="Altezza (m)",
                height=360, plot_bgcolor="#f8fbff",
                legend=dict(orientation="h", yanchor="bottom", y=1.02)
            )
            st.plotly_chart(fig2, use_container_width=True)

        # Shear alpha
        if "shear_alpha" in df_stats.columns:
            shear_data = df_stats[df_stats["shear_alpha"].notna()]
            fig3 = go.Figure(go.Bar(
                x=shear_data["month"], y=shear_data["shear_alpha"],
                marker_color=[CLR_MONTHS[i % len(CLR_MONTHS)] for i in range(len(shear_data))],
                text=shear_data["shear_alpha"].apply(lambda x: f"α={x:.3f}"),
                textposition="outside"
            ))
            fig3.add_hline(y=0.143, line_dash="dash", line_color="gray",
                           annotation_text="α=0.143 (1/7 std)", annotation_position="right")
            fig3.update_layout(
                title="Coefficiente Wind Shear α — legge della potenza",
                yaxis_title="α", yaxis_range=[0, 0.5], height=300,
                plot_bgcolor="#f8fbff"
            )
            st.plotly_chart(fig3, use_container_width=True)
    else:
        st.info("Dati multi-altezza non disponibili. Assicurati che il CSV contenga le colonne TOP 92, RIF 88, RIF 70, RIF 50.")

# ---- TAB 4: METEO ----
with tab4:
    col1, col2 = st.columns(2)
    with col1:
        if "temp_avg" in df_stats.columns:
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=valid_stats["month"], y=valid_stats["temp_avg"],
                name="Media", line=dict(color="#D85A30", width=3), fill="tozeroy",
                fillcolor="rgba(216,90,48,0.1)"))
            fig.add_trace(go.Scatter(x=valid_stats["month"], y=valid_stats["temp_max"],
                name="Max", line=dict(color="#c0392b", dash="dot", width=2)))
            fig.add_trace(go.Scatter(x=valid_stats["month"], y=valid_stats["temp_min"],
                name="Min", line=dict(color="#2980b9", dash="dot", width=2)))
            fig.update_layout(title="Temperatura (°C)", yaxis_title="°C",
                              height=320, plot_bgcolor="#f8fbff",
                              legend=dict(orientation="h", y=1.1))
            st.plotly_chart(fig, use_container_width=True)

    with col2:
        if "hum_avg" in df_stats.columns:
            fig2 = go.Figure(go.Scatter(
                x=df_stats["month"], y=df_stats["hum_avg"],
                fill="tozeroy", line=dict(color="#2980b9", width=3),
                fillcolor="rgba(41,128,185,0.1)"
            ))
            fig2.update_layout(title="Umidità media (%)", yaxis_title="%",
                               yaxis_range=[0, 100], height=320, plot_bgcolor="#f8fbff")
            st.plotly_chart(fig2, use_container_width=True)

    if "pres_avg" in df_stats.columns:
        fig3 = go.Figure(go.Scatter(
            x=valid_stats["month"], y=valid_stats["pres_avg"],
            fill="tozeroy", line=dict(color="#8e44ad", width=3),
            fillcolor="rgba(142,68,173,0.1)"
        ))
        fig3.update_layout(title="Pressione atmosferica (hPa)", yaxis_title="hPa",
                           height=280, plot_bgcolor="#f8fbff")
        st.plotly_chart(fig3, use_container_width=True)

# ---- TAB 5: AEP ----
with tab5:
    st.markdown("### ⚡ Calcolo AEP — Annual Energy Production")
    st.markdown("<div class='warn-box'>Carica la power curve della turbina per calcolare l'AEP stimato. Convezione eccedenza: P50 > P75 > P90.</div>", unsafe_allow_html=True)
    st.markdown("")

    pc_file = st.file_uploader("Power Curve (CSV: wind_speed_ms, power_kw)", type="csv")

    colA, colB, colC, colD = st.columns(4)
    n_turbines = colA.number_input("Turbine", 1, 500, 1)
    avail_pct  = colB.number_input("Disponibilità (%)", 1, 100, 95)
    losses_pct = colC.number_input("Perdite wake+elec. (%)", 0, 50, 10)
    hub_height = colD.selectbox("Altezza cubo (m)", [92, 80, 100, 120, 150], index=0)

    if pc_file and not valid_stats.empty:
        pc_content = pc_file.read().decode("utf-8", errors="ignore")
        pc_sep = ";" if pc_content.count(";") > pc_content.count(",") else ","
        pc_df = pd.read_csv(StringIO(pc_content), sep=pc_sep, comment="#")
        pc_df.columns = [c.strip().lower() for c in pc_df.columns]

        ws_col  = next((c for c in pc_df.columns if "speed" in c or c in ["ws","v"]), pc_df.columns[0])
        kw_col  = next((c for c in pc_df.columns if "power" in c or "kw" in c or "potenza" in c), pc_df.columns[1])
        pc_df = pc_df[[ws_col, kw_col]].dropna()
        pc_df.columns = ["ws","kw"]
        pc_df = pc_df.astype(float).sort_values("ws").reset_index(drop=True)

        def interp_power(ws_val):
            if ws_val <= pc_df["ws"].iloc[0]:  return pc_df["kw"].iloc[0]
            if ws_val >= pc_df["ws"].iloc[-1]: return pc_df["kw"].iloc[-1]
            return float(np.interp(ws_val, pc_df["ws"], pc_df["kw"]))

        def shear_correct(v, h_meas=92, h_hub=hub_height, alpha=0.143):
            return v * (h_hub / h_meas) ** alpha

        # Monthly AEP
        avail_f = avail_pct / 100
        loss_f  = 1 - losses_pct / 100
        HOURS   = 8760
        rated_kw = pc_df["kw"].max()

        monthly_mwh = []
        for _, row in valid_stats.iterrows():
            v_avg = row["wind_avg"]
            if pd.isna(v_avg): monthly_mwh.append(np.nan); continue
            v_hub = shear_correct(v_avg)
            kw    = interp_power(v_hub)
            mwh   = kw * (30 * 24) * n_turbines * avail_f * loss_f / 1000
            monthly_mwh.append(round(mwh))

        ann_ext  = round(np.nanmean(monthly_mwh) * 12) if monthly_mwh else 0
        cf       = ann_ext / (rated_kw * HOURS * n_turbines / 1000) if rated_kw > 0 else 0

        # P50/P75/P90 AEP
        avg_wind = valid_stats["wind_avg"].mean()
        def aep_at_scale(scale):
            total = sum([interp_power(shear_correct(v * scale)) * (30*24) / len(valid_stats) * 12
                         for v in valid_stats["wind_avg"].dropna()])
            return round(total * n_turbines * avail_f * loss_f / 1000)

        aep_p50 = ann_ext
        aep_p75 = aep_at_scale(valid_stats["p75"].mean() / avg_wind) if avg_wind else None
        aep_p90 = aep_at_scale(valid_stats["p90"].mean() / avg_wind) if avg_wind else None

        # KPI row
        k1, k2, k3, k4 = st.columns(4)
        for col, lbl, val, sub in [
            (k1, "AEP P50", f"{ann_ext/1000:.2f} GWh/anno", f"{len(valid_stats)}/12 mesi"),
            (k2, "Capacity Factor", f"{cf*100:.1f}%", "AEP / (Pnom × 8760 h)"),
            (k3, "Potenza nominale", f"{rated_kw*n_turbines/1000:.1f} MW", f"{n_turbines} × {rated_kw:.0f} kW"),
            (k4, "Hub height", f"{hub_height} m", "con shear α=0.143"),
        ]:
            col.markdown(f"""<div class='kpi-card'>
              <div class='kpi-label'>{lbl}</div>
              <div class='kpi-value'>{val}</div>
              <div class='kpi-sub'>{sub}</div></div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        col_a, col_b = st.columns(2)

        with col_a:
            fig_pc = go.Figure(go.Scatter(x=pc_df["ws"], y=pc_df["kw"],
                fill="tozeroy", line=dict(color="#2E75B6", width=3),
                fillcolor="rgba(46,117,182,0.1)"))
            fig_pc.update_layout(title="Power Curve (kW)", xaxis_title="Velocità (m/s)",
                                 yaxis_title="kW", height=300, plot_bgcolor="#f8fbff")
            st.plotly_chart(fig_pc, use_container_width=True)

        with col_b:
            fig_sens = go.Figure(go.Bar(
                x=["P50 — mediano","P75 — 75% eccedenza","P90 — 90% eccedenza"],
                y=[aep_p50, aep_p75, aep_p90],
                marker_color=["#1D9E75","#d4a017","#BA7517"],
                text=[f"{v:,.0f} MWh" for v in [aep_p50, aep_p75, aep_p90]],
                textposition="outside"
            ))
            fig_sens.update_layout(title="Sensitività AEP (MWh/anno)",
                                   yaxis_title="MWh", height=300, plot_bgcolor="#f8fbff")
            st.plotly_chart(fig_sens, use_container_width=True)

        fig_monthly = go.Figure(go.Bar(
            x=valid_stats["month"], y=monthly_mwh,
            marker_color=CLR_MONTHS[:len(monthly_mwh)],
            text=[f"{v:,}" for v in monthly_mwh], textposition="outside"
        ))
        fig_monthly.update_layout(title="Produzione mensile stimata (MWh)",
                                  yaxis_title="MWh", height=320, plot_bgcolor="#f8fbff")
        st.plotly_chart(fig_monthly, use_container_width=True)

    else:
        st.info("Carica la power curve della turbina (CSV: wind_speed_ms, power_kw) per avviare il calcolo.")
        with st.expander("📥 Scarica template power curve"):
            template = "wind_speed_ms,power_kw\n# Sostituire i valori kW con quelli reali della turbina\n" + \
                       "\n".join([f"{ws},{kw}" for ws, kw in
                                  [(0,0),(1,0),(2,0),(3,30),(4,105),(5,235),(6,430),(7,695),
                                   (8,1010),(9,1360),(10,1690),(11,1900),(12,1990),(13,2000),
                                   (14,2000),(15,2000),(16,2000),(17,2000),(18,2000),(19,2000),
                                   (20,2000),(21,2000),(22,2000),(23,2000),(24,2000),(25,0)]])
            st.download_button("⬇️ Scarica template", template, "power_curve_template.csv", "text/csv")

# ============================================================
#  FOOTER
# ============================================================
st.markdown("---")
st.markdown(
    "<p style='text-align:center;font-size:0.75rem;color:#aaa;'>"
    "Dashboard Meteo Stazioni · Dati stazione G243043 Durrà · "
    "Powered by Streamlit + Plotly"
    "</p>", unsafe_allow_html=True
)
