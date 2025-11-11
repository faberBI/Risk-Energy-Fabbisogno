import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from io import BytesIO

# Configurazione pagina
st.set_page_config(page_title="Calcolo Open Position", layout="wide")

st.title("📊 Analisi Fabbisogno Year by Year")

st.markdown("""
Carica un file Excel con la seguente struttura (colonne):
**Anno, Fabbisogno Adjusted, Fabbisogno, PPA ERG secure, PPA ERG baseline, PPA ERG Top, FRW, Solar**
""")

uploaded_file = st.file_uploader("📂 Carica file Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    required_cols = ["Anno", "Fabbisogno Adjusted", "Fabbisogno", 
                     "PPA ERG secure", "PPA ERG baseline", "PPA ERG Top", "FRW", "Solar"]
    missing = [c for c in required_cols if c not in df.columns]

    if missing:
        st.error(f"Mancano le colonne obbligatorie: {', '.join(missing)}")
    else:
        # --- Calcoli principali ---
        df["Open Position w Solar (Adjusted) secure"] = (
            df["Fabbisogno Adjusted"] - (df["PPA ERG secure"] + df["FRW"] + df["Solar"])
        )
        df["Open Position w Solar (Adjusted) top"] = (
            df["Fabbisogno Adjusted"] - (df["PPA ERG Top"] + df["FRW"] + df["Solar"])
        )

        df["PPA_cum_secure"] = df["PPA ERG secure"]
        df["FRW_cum_secure"] = df["PPA ERG secure"] + df["FRW"]
        df["Solar_cum_secure"] = df["PPA ERG secure"] + df["FRW"] + df["Solar"]

        df["PPA_cum_top"] = df["PPA ERG Top"]
        df["FRW_cum_top"] = df["PPA ERG Top"] + df["FRW"]
        df["Solar_cum_top"] = df["PPA ERG Top"] + df["FRW"] + df["Solar"]

        st.success("✅ Calcolo completato!")

        st.subheader("📈 Tabella con Open Position calcolati")
        st.dataframe(df.style.format("{:.0f}"))

        df["Anno"] = pd.to_datetime(df["Anno"], format="%Y")

        # --- Costruzione grafico ---
        fig = go.Figure()

        # 1️⃣ Fabbisogno Adjusted (area blu)
        fig.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["Fabbisogno Adjusted"],
            name="Fabbisogno Adjusted",
            mode="lines",
            line=dict(color="rgba(0, 25, 108, 1)", width=3),
            fill="tozeroy",
            fillcolor="rgba(0, 25, 108, 0.3)",
            hovertemplate="Fabbisogno Adjusted: %{y} GWh<extra></extra>"
        ))

        # 2️⃣ Coperture SCENARIO SECURE (verde scuro)
        fig.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["Solar_cum_secure"],
            name="Copertura Secure",
            mode="lines",
            line=dict(color="rgba(0, 100, 0, 1)", width=2),
            fill="tonexty",
            fillcolor="rgba(0, 100, 0, 0.5)",
            hovertemplate="Copertura Secure: %{y} GWh<extra></extra>"
        ))

        # 3️⃣ Coperture SCENARIO TOP (verde chiaro)
        fig.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["Solar_cum_top"],
            name="Copertura Top",
            mode="lines",
            line=dict(color="rgba(0, 180, 80, 1)", width=2),
            fill="tonexty",
            fillcolor="rgba(0, 180, 80, 0.5)",
            hovertemplate="Copertura Top: %{y} GWh<extra></extra>"
        ))

        # 4️⃣ Linee tratteggiate per PPA ERG Secure e Top
        fig.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["PPA ERG secure"],
            name="PPA ERG Secure (Linea)",
            mode="lines",
            line=dict(color="rgba(0, 100, 0, 1)", width=3, dash="dash"),
            hovertemplate="PPA ERG Secure: %{y} GWh<extra></extra>"
        ))
        fig.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["PPA ERG Top"],
            name="PPA ERG Top (Linea)",
            mode="lines",
            line=dict(color="rgba(0, 200, 0, 1)", width=3, dash="dash"),
            hovertemplate="PPA ERG Top: %{y} GWh<extra></extra>"
        ))

        # 5️⃣ Open Position Secure (bianca)
        fig.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["Fabbisogno Adjusted"],
            name="Open Position Secure",
            mode="lines",
            line=dict(color="white", width=1.5),
            fill="tonexty",
            fillcolor="rgba(255,255,255,1)",
            customdata=df["Open Position w Solar (Adjusted) secure"],
            hovertemplate="Open Position Secure: %{customdata} GWh<extra></extra>"
        ))

        # 6️⃣ Open Position Top (grigio chiaro)
        fig.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["Fabbisogno Adjusted"],
            name="Open Position Top",
            mode="lines",
            line=dict(color="rgba(200,200,200,1)", width=1.5),
            fill="tonexty",
            fillcolor="rgba(220,220,220,0.6)",
            customdata=df["Open Position w Solar (Adjusted) top"],
            hovertemplate="Open Position Top: %{customdata} GWh<extra></extra>"
        ))

        # 7️⃣ Fabbisogno Reale → in primo piano
        fig.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["Fabbisogno"],
            name="Fabbisogno (Reale)",
            mode="lines",
            line=dict(color="rgba(0, 60, 170, 1)", width=3, dash="dash"),
            hovertemplate="Fabbisogno: %{y} GWh<extra></extra>"
        ))

        # Layout
        fig.update_layout(
            title="Confronto Scenari Secure vs Top - Coperture e PPA ERG",
            yaxis_title="GWh",
            xaxis_title="Anno",
            legend_title="Legenda",
            hovermode="x unified",
            xaxis=dict(tickformat="%Y", dtick="M12"),
            plot_bgcolor="white"
        )

        st.plotly_chart(fig, use_container_width=True)

        # --- Download Excel ---
        output_buffer = BytesIO()
        df.to_excel(output_buffer, index=False, engine="openpyxl")
        output_buffer.seek(0)

        st.download_button(
            label="📥 Scarica Excel con risultati",
            data=output_buffer,
            file_name="open_position_secure_vs_top.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("⬆️ Carica un file Excel per iniziare.")
