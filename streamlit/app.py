import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import json
import os
from io import StringIO
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io

st.set_page_config(page_title="Dashboard Meteo", page_icon="🌬️", layout="wide")

# ============================================================
#  GOOGLE DRIVE
# ============================================================
DRIVE_FOLDER_ID = "1BjMd963hCVPlpvBQxuJJL7qkxMgOCiD0"
SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

@st.cache_resource
def get_drive_service():
    try:
        sa_key = os.environ.get("GCP_SA_KEY", "")
        if not sa_key:
            return None
        creds = service_account.Credentials.from_service_account_info(
            json.loads(sa_key), scopes=SCOPES)
        return build("drive", "v3", credentials=creds, cache_discovery=False)
    except Exception as e:
        st.warning(f"Drive non disponibile: {e}")
        return None

@st.cache_data(ttl=300)
def list_csv_files():
    svc = get_drive_service()
    if not svc:
        return []
    try:
        res = svc.files().list(
            q=f"'{DRIVE_FOLDER_ID}' in parents and mimeType='text/csv' and trashed=false",
            fields="files(id,name,modifiedTime)", orderBy="name"
        ).execute()
        return res.get("files", [])
    except Exception as e:
        st.error(f"Errore Drive: {e}")
        return []

@st.cache_data(ttl=300)
def read_csv_drive(file_id, file_name):
    svc = get_drive_service()
    if not svc:
        return pd.DataFrame()
    try:
        req = svc.files().get_media(fileId=file_id)
        buf = io.BytesIO()
        dl = MediaIoBaseDownload(buf, req)
        done = False
        while not done:
            _, done = dl.next_chunk()
        buf.seek(0)
        content = buf.read().decode("utf-8", errors="ignore")
        return parse_csv(content, file_name)
    except Exception as e:
        st.error(f"Errore lettura {file_name}: {e}")
        return pd.DataFrame()

# ============================================================
#  CSS
# ============================================================
st.markdown("""<style>
.kpi-card{background:#f0f5fc;border-radius:8px;padding:1rem;text-align:center;border-left:4px solid #2E75B6;margin-bottom:8px}
.kpi-value{font-size:1.6rem;font-weight:700;color:#1F4E79}
.kpi-label{font-size:0.75rem;color:#666;margin-bottom:4px}
.kpi-sub{font-size:0.7rem;color:#999;margin-top:2px}
.warn-box{background:#fff8e1;border-left:3px solid #f0b429;padding:0.6rem 1rem;border-radius:0 6px 6px 0;font-size:0.85rem;color:#7f5c00}
.p50{background:#d4edda;color:#155724;padding:2px 8px;border-radius:8px;font-weight:700;font-size:0.8rem}
.p75{background:#fff0b3;color:#5c4400;padding:2px 8px;border-radius:8px;font-weight:700;font-size:0.8rem}
.p90{background:#ffe0b2;color:#6b3a00;padding:2px 8px;border-radius:8px;font-weight:700;font-size:0.8rem}
</style>""", unsafe_allow_html=True)

# ============================================================
#  COSTANTI
# ============================================================
WIND_COLS = {
    "TOP 92m": "TOP 92;wind_speed;Avg (m/s)",
    "RIF 88m": "RIF 88;wind_speed;Avg (m/s)",
    "RIF 70m": "RIF 70;wind_speed;Avg (m/s)",
    "RIF 50m": "RIF 50;wind_speed;Avg (m/s)",
}
ANN_SEQ = [f"2025-{m:02d}" for m in range(4,13)] + ["2026-01","2026-02","2026-03"]
CLR = px.colors.qualitative.Set2

# ============================================================
#  PARSING
# ============================================================
@st.cache_data(show_spinner=False)
def parse_csv(content: str, fname: str) -> pd.DataFrame:
    lines = content.split("\n")
    first = lines[0] if lines else ""
    sep = "," if len(first.split(",")) >= len(first.split(";")) else ";"
    if len(first.split(",")) >= 10:
        sep = ","
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

# ============================================================
#  STATISTICHE
# ============================================================
def exc_pct(series, p):
    clean = series.dropna()
    clean = clean[clean > 0]
    return float(np.percentile(clean, 100-p)) if len(clean) else np.nan

def compute_stats(df):
    wc = WIND_COLS["TOP 92m"]
    tc = "TEMP-UMID;temperature;Avg (°C)"
    hc = "TEMP-UMID;humidity;Avg (%)"
    pc = "GEOVES BOX;air_pressure;Avg (hPa)"
    rows = []
    for month, g in df.groupby("month"):
        wind = g[wc].dropna()[g[wc].dropna() > 0] if wc in g else pd.Series(dtype=float)
        temp = g[tc][(g[tc] > -5)&(g[tc] < 60)].dropna() if tc in g else pd.Series(dtype=float)
        hum  = g[hc].dropna() if hc in g else pd.Series(dtype=float)
        pres = g[pc][(g[pc] > 900)&(g[pc] < 1100)].dropna() if pc in g else pd.Series(dtype=float)
        h = {}
        for lbl, col in WIND_COLS.items():
            if col in g.columns:
                v = g[col].dropna(); v = v[v > 0]
                h[lbl] = round(v.mean(),3) if len(v) else np.nan
            else:
                h[lbl] = np.nan
        v92 = h.get("TOP 92m",np.nan); v50 = h.get("RIF 50m",np.nan)
        shear = round(np.log(v92/v50)/np.log(92/50),4) if v92 and v50 and v92>0 and v50>0 else np.nan
        rows.append({
            "month": month, "misurazioni": len(g),
            "wind_avg": round(wind.mean(),3) if len(wind) else np.nan,
            "wind_max": round(wind.max(),3)  if len(wind) else np.nan,
            "p50": exc_pct(wind,50), "p75": exc_pct(wind,75), "p90": exc_pct(wind,90),
            "wind_top92m": h.get("TOP 92m"), "wind_rif88m": h.get("RIF 88m"),
            "wind_rif70m": h.get("RIF 70m"), "wind_rif50m": h.get("RIF 50m"),
            "shear_alpha": shear,
            "temp_avg": round(temp.mean(),2) if len(temp) else np.nan,
            "temp_max": round(temp.max(),2)  if len(temp) else np.nan,
            "temp_min": round(temp.min(),2)  if len(temp) else np.nan,
            "hum_avg":  round(hum.mean(),1)  if len(hum)  else np.nan,
            "pres_avg": round(pres.mean(),1) if len(pres) else np.nan,
            "avail_pct": min(100, round(len(g)/144/30*100,1)),
            "anomalous": month.endswith("-03") and len(wind)>0 and wind.mean()>8,
        })
    return pd.DataFrame(rows).sort_values("month").reset_index(drop=True)

# ============================================================
#  SIDEBAR
# ============================================================
with st.sidebar:
    st.markdown("## 🌬️ Dashboard Meteo")
    st.markdown("---")
    st.markdown("### ☁️ Google Drive")

    drive_files = list_csv_files()
    df_raw = pd.DataFrame()
    df_stats = pd.DataFrame()
    station_name = "G243043"
    station_loc  = "Durrà"

    if drive_files:
        st.success(f"✅ {len(drive_files)} CSV su Drive")
        names = [f["name"] for f in drive_files]
        selected = st.multiselect("Mesi da caricare", names, default=names)

        col_btn1, col_btn2 = st.columns(2)
        refresh = col_btn1.button("🔄 Aggiorna", type="primary")
        if refresh:
            st.cache_data.clear()
            st.rerun()

        if selected:
            with st.spinner("Caricamento da Drive..."):
                dfs = []
                bar = st.progress(0)
                for idx, f in enumerate(drive_files):
                    if f["name"] in selected:
                        parsed = read_csv_drive(f["id"], f["name"])
                        if not parsed.empty:
                            dfs.append(parsed)
                    bar.progress((idx+1)/len(drive_files))
                bar.empty()
                if dfs:
                    df_raw   = pd.concat(dfs).sort_values("datetime").reset_index(drop=True)
                    df_stats = compute_stats(df_raw)
                    st.success(f"✅ {len(df_raw):,} righe caricate")
    else:
        st.warning("Nessun CSV trovato su Drive.")
        st.info("Verifica che la cartella sia condivisa con il Service Account.")
        st.markdown("### 📂 Carica manualmente")
        uploaded = st.file_uploader("CSV", type="csv", accept_multiple_files=True)
        if uploaded:
            dfs = []
            for f in uploaded:
                parsed = parse_csv(f.read().decode("utf-8","ignore"), f.name)
                if not parsed.empty:
                    dfs.append(parsed)
            if dfs:
                df_raw   = pd.concat(dfs).sort_values("datetime").reset_index(drop=True)
                df_stats = compute_stats(df_raw)

    station_name = st.text_input("Codice stazione", "G243043")
    station_loc  = st.text_input("Località", "Durrà")

    st.markdown("---")
    if not df_stats.empty:
        valid = df_stats[~df_stats.get("anomalous", pd.Series(False,index=df_stats.index))]
        ann = [m for m in ANN_SEQ if m in valid["month"].values]
        pct = int(len(ann)/12*100)
        st.markdown("**Anno Apr 2025–Mar 2026**")
        st.progress(pct/100)
        st.caption(f"{len(ann)}/12 mesi · {pct}%")
        if ann:
            p50a = round(valid[valid["month"].isin(ann)]["p50"].mean(),2)
            p75a = round(valid[valid["month"].isin(ann)]["p75"].mean(),2)
            p90a = round(valid[valid["month"].isin(ann)]["p90"].mean(),2)
            st.markdown(f"<span class='p50'>P50:{p50a}</span> <span class='p75'>P75:{p75a}</span> <span class='p90'>P90:{p90a}</span> m/s", unsafe_allow_html=True)

# ============================================================
#  HEADER
# ============================================================
st.markdown(f"""<div style="background:linear-gradient(135deg,#1F4E79,#2E75B6);color:#fff;padding:1.2rem 1.5rem;border-radius:8px;margin-bottom:1rem">
<h2 style="margin:0">🌬️ {station_name} — {station_loc}</h2>
<p style="margin:0;opacity:.8;font-size:.9rem">Dati aggiornati automaticamente da Google Drive · Cloud Run</p>
</div>""", unsafe_allow_html=True)

if df_stats.empty:
    st.info("👈 I dati si caricano da Google Drive. Se non appaiono, clicca 🔄 Aggiorna.")
    st.stop()

def kpi(col, lbl, val, sub):
    col.markdown(f"<div class='kpi-card'><div class='kpi-label'>{lbl}</div><div class='kpi-value'>{val}</div><div class='kpi-sub'>{sub}</div></div>", unsafe_allow_html=True)

valid_stats = df_stats[~df_stats.get("anomalous", pd.Series(False,index=df_stats.index))]

tab1,tab2,tab3,tab4,tab5 = st.tabs(["📊 Disponibilità","📈 P50/P75/P90","📐 Profilo Vento","🌡️ Meteo","⚡ AEP"])

# TAB 1
with tab1:
    c1,c2,c3,c4 = st.columns(4)
    kpi(c1,"Misurazioni totali",f"{int(df_stats['misurazioni'].sum()):,}","record 10 min")
    kpi(c2,"Disponibilità media",f"{round(df_stats['avail_pct'].mean(),1)}%","su tutti i mesi")
    kpi(c3,"Mesi analizzati",str(len(df_stats)),"totale")
    kpi(c4,"CSV su Drive",str(len(drive_files)),"aggiornamento auto")
    st.markdown("<br>",unsafe_allow_html=True)
    avail_c = ["#1D9E75" if v>=98 else "#e6a817" if v>=90 else "#c0392b" for v in df_stats["avail_pct"]]
    fig = go.Figure(go.Bar(x=df_stats["month"],y=df_stats["avail_pct"],
        marker_color=avail_c,text=df_stats["avail_pct"].apply(lambda x:f"{x}%"),textposition="outside"))
    fig.update_layout(title="Disponibilità mensile (%)",yaxis_range=[0,105],height=350,plot_bgcolor="#f8fbff")
    st.plotly_chart(fig,use_container_width=True)

# TAB 2
with tab2:
    st.markdown("<div class='warn-box'>⚠️ Eccedenza: P50 &gt; P75 &gt; P90 — P90 = scenario conservativo</div>",unsafe_allow_html=True)
    st.markdown("")
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=df_stats["month"],y=df_stats["p50"],name="P50",
        line=dict(color="#1D9E75",width=3),mode="lines+markers",marker=dict(size=8)))
    fig.add_trace(go.Scatter(x=df_stats["month"],y=df_stats["p75"],name="P75",
        line=dict(color="#d4a017",width=2,dash="dot"),mode="lines+markers",marker=dict(size=7)))
    fig.add_trace(go.Scatter(x=df_stats["month"],y=df_stats["p90"],name="P90",
        line=dict(color="#BA7517",width=2,dash="dash"),mode="lines+markers",marker=dict(size=7)))
    fig.update_layout(title="P50/P75/P90 — Vento TOP 92m (m/s)",height=380,plot_bgcolor="#f8fbff")
    st.plotly_chart(fig,use_container_width=True)
    cols = st.columns(min(len(df_stats),4))
    for i,(_,row) in enumerate(df_stats.iterrows()):
        with cols[i%4]:
            p50v = f"{row['p50']:.2f}" if pd.notna(row['p50']) else "—"
            p75v = f"{row['p75']:.2f}" if pd.notna(row['p75']) else "—"
            p90v = f"{row['p90']:.2f}" if pd.notna(row['p90']) else "—"
            st.markdown(f"<div class='kpi-card'><b>{row['month']}</b><br><span class='p50'>P50:{p50v}</span> <span class='p75'>P75:{p75v}</span> <span class='p90'>P90:{p90v}</span></div>",unsafe_allow_html=True)

# TAB 3
with tab3:
    h_cols=["wind_top92m","wind_rif88m","wind_rif70m","wind_rif50m"]
    h_lbls=["TOP 92m","RIF 88m","RIF 70m","RIF 50m"]
    h_clrs={"TOP 92m":"#1F4E79","RIF 88m":"#2E75B6","RIF 70m":"#1D9E75","RIF 50m":"#BA7517"}
    col1,col2 = st.columns(2)
    with col1:
        fig=go.Figure()
        for lbl,cn in zip(h_lbls,h_cols):
            if cn in df_stats.columns:
                fig.add_trace(go.Bar(name=lbl,x=df_stats["month"],y=df_stats[cn],marker_color=h_clrs[lbl]))
        fig.update_layout(barmode="group",title="Velocità per altezza (m/s)",height=360,plot_bgcolor="#f8fbff")
        st.plotly_chart(fig,use_container_width=True)
    with col2:
        fig2=go.Figure()
        for i,(_,row) in enumerate(df_stats.iterrows()):
            speeds=[row.get(c,np.nan) for c in h_cols]
            if any(pd.notna(s) for s in speeds):
                fig2.add_trace(go.Scatter(x=speeds,y=[92,88,70,50],name=row["month"],
                    mode="lines+markers",line=dict(color=CLR[i%len(CLR)],width=2),marker=dict(size=8)))
        fig2.update_layout(title="Profilo verticale",xaxis_title="m/s",yaxis_title="m",height=360,plot_bgcolor="#f8fbff")
        st.plotly_chart(fig2,use_container_width=True)
    if "shear_alpha" in df_stats.columns:
        sd=df_stats[df_stats["shear_alpha"].notna()]
        fig3=go.Figure(go.Bar(x=sd["month"],y=sd["shear_alpha"],
            marker_color=[CLR[i%len(CLR)] for i in range(len(sd))],
            text=sd["shear_alpha"].apply(lambda x:f"α={x:.3f}"),textposition="outside"))
        fig3.add_hline(y=0.143,line_dash="dash",line_color="gray",annotation_text="α=0.143 std")
        fig3.update_layout(title="Wind Shear α",yaxis_range=[0,0.5],height=280,plot_bgcolor="#f8fbff")
        st.plotly_chart(fig3,use_container_width=True)

# TAB 4
with tab4:
    col1,col2=st.columns(2)
    with col1:
        if "temp_avg" in df_stats.columns and not valid_stats.empty:
            fig=go.Figure()
            fig.add_trace(go.Scatter(x=valid_stats["month"],y=valid_stats["temp_avg"],name="Media",
                line=dict(color="#D85A30",width=3),fill="tozeroy",fillcolor="rgba(216,90,48,0.1)"))
            fig.add_trace(go.Scatter(x=valid_stats["month"],y=valid_stats["temp_max"],name="Max",
                line=dict(color="#c0392b",dash="dot",width=2)))
            fig.add_trace(go.Scatter(x=valid_stats["month"],y=valid_stats["temp_min"],name="Min",
                line=dict(color="#2980b9",dash="dot",width=2)))
            fig.update_layout(title="Temperatura (°C)",height=320,plot_bgcolor="#f8fbff")
            st.plotly_chart(fig,use_container_width=True)
    with col2:
        if "hum_avg" in df_stats.columns:
            fig2=go.Figure(go.Scatter(x=df_stats["month"],y=df_stats["hum_avg"],
                fill="tozeroy",line=dict(color="#2980b9",width=3),fillcolor="rgba(41,128,185,0.1)"))
            fig2.update_layout(title="Umidità (%)",yaxis_range=[0,100],height=320,plot_bgcolor="#f8fbff")
            st.plotly_chart(fig2,use_container_width=True)
    if "pres_avg" in df_stats.columns and not valid_stats.empty:
        fig3=go.Figure(go.Scatter(x=valid_stats["month"],y=valid_stats["pres_avg"],
            fill="tozeroy",line=dict(color="#8e44ad",width=3),fillcolor="rgba(142,68,173,0.1)"))
        fig3.update_layout(title="Pressione (hPa)",height=280,plot_bgcolor="#f8fbff")
        st.plotly_chart(fig3,use_container_width=True)

# TAB 5
with tab5:
    st.markdown("### ⚡ AEP — Annual Energy Production")
    pc_file=st.file_uploader("Power Curve CSV (wind_speed_ms, power_kw)",type="csv")
    cA,cB,cC,cD=st.columns(4)
    n_t=cA.number_input("Turbine",1,500,1)
    av=cB.number_input("Disponibilità (%)",1,100,95)
    ls=cC.number_input("Perdite (%)",0,50,10)
    hh=cD.selectbox("Hub height (m)",[92,80,100,120,150])
    if pc_file and not valid_stats.empty:
        content=pc_file.read().decode("utf-8","ignore")
        sep=";" if content.count(";")>content.count(",") else ","
        pc=pd.read_csv(StringIO(content),sep=sep,comment="#")
        pc.columns=[c.strip().lower() for c in pc.columns]
        wsc=next((c for c in pc.columns if "speed" in c or c in["ws","v"]),pc.columns[0])
        kwc=next((c for c in pc.columns if "power" in c or "kw" in c),pc.columns[1])
        pc=pc[[wsc,kwc]].dropna().astype(float).sort_values(wsc).reset_index(drop=True)
        pc.columns=["ws","kw"]
        def ip(v): return float(np.interp(v,pc["ws"],pc["kw"]))
        def sc(v): return v*(hh/92)**0.143
        af=av/100; lf=1-ls/100; rk=pc["kw"].max(); aw=valid_stats["wind_avg"].mean()
        mwh=[round(ip(sc(r["wind_avg"]))*30*24*n_t*af*lf/1000) if pd.notna(r["wind_avg"]) else np.nan for _,r in valid_stats.iterrows()]
        ann=round(np.nanmean(mwh)*12)
        cf=ann/(rk*8760*n_t/1000)
        k1,k2,k3,k4=st.columns(4)
        kpi(k1,"AEP P50",f"{ann/1000:.2f} GWh/anno",f"{len(valid_stats)}/12 mesi")
        kpi(k2,"Capacity Factor",f"{cf*100:.1f}%","AEP/(Pnom×8760h)")
        kpi(k3,"Potenza nominale",f"{rk*n_t/1000:.1f} MW",f"{n_t}×{rk:.0f} kW")
        kpi(k4,"Hub height",f"{hh} m","α=0.143")
        p75a=round(np.nanmean([ip(sc(v*valid_stats["p75"].mean()/aw))*30*24/len(valid_stats)*12 for v in valid_stats["wind_avg"].dropna()])*n_t*af*lf/1000)
        p90a=round(np.nanmean([ip(sc(v*valid_stats["p90"].mean()/aw))*30*24/len(valid_stats)*12 for v in valid_stats["wind_avg"].dropna()])*n_t*af*lf/1000)
        ca,cb=st.columns(2)
        with ca:
            fp=go.Figure(go.Scatter(x=pc["ws"],y=pc["kw"],fill="tozeroy",line=dict(color="#2E75B6",width=3),fillcolor="rgba(46,117,182,0.1)"))
            fp.update_layout(title="Power Curve (kW)",xaxis_title="m/s",yaxis_title="kW",height=300,plot_bgcolor="#f8fbff")
            st.plotly_chart(fp,use_container_width=True)
        with cb:
            fs=go.Figure(go.Bar(x=["P50","P75","P90"],y=[ann,p75a,p90a],
                marker_color=["#1D9E75","#d4a017","#BA7517"],
                text=[f"{v:,} MWh" for v in [ann,p75a,p90a]],textposition="outside"))
            fs.update_layout(title="AEP per scenario (MWh/anno)",height=300,plot_bgcolor="#f8fbff")
            st.plotly_chart(fs,use_container_width=True)
    else:
        st.info("Carica la power curve per calcolare l'AEP.")
        with st.expander("📥 Template"):
            t="wind_speed_ms,power_kw\n"+"\n".join([f"{w},{k}" for w,k in[(0,0),(3,30),(5,235),(7,695),(9,1360),(11,1900),(13,2000),(25,0)]])
            st.download_button("⬇️ Scarica",t,"power_curve_template.csv","text/csv")

st.markdown("---")
st.markdown(f"<p style='text-align:center;font-size:0.75rem;color:#aaa'>Dashboard Meteo · {station_name} {station_loc} · Google Drive · Cloud Run</p>",unsafe_allow_html=True)
