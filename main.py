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
        # Scenario Secure with solar
        df["Open Position w Solar (Adjusted) secure"] = (
            df["Fabbisogno Adjusted"] - (df["PPA ERG secure"] + df["FRW"] + df["Solar"])
        )

        # Scenario Top with solar
        df["Open Position w Solar (Adjusted) top"] = (
            df["Fabbisogno Adjusted"] - (df["PPA ERG Top"] + df["FRW"] + df["Solar"])
        )
        # Scenario secure w/o solar
        df["Open Position w/o Solar (Adjusted) secure"] = (
            df["Fabbisogno Adjusted"] - (df["PPA ERG secure"] + df["FRW"])
        )
        # Scenario top w/o solar
        df["Open Position w/o Solar (Adjusted) top"] = (
            df["Fabbisogno Adjusted"] - (df["PPA ERG Top"] + df["FRW"])
        )

        # Cumulativi
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
# --- Costruzione grafico “fig” ---
        fig = go.Figure()
        
        # Calcolo coperture totali
        df["Coperture_secure"] = df["PPA ERG secure"] + df["FRW"] + df["Solar"]
        df["Coperture_top"] = df["PPA ERG Top"] + df["FRW"] + df["Solar"]
        
        # Calcolo Open Position
        df["OpenPosition_secure"] = df["Fabbisogno Adjusted"] - df["Coperture_secure"]
        df["OpenPosition_top"] = df["Fabbisogno Adjusted"] - df["Coperture_top"]
        
        # 1️⃣ Fabbisogno Adjusted (linea blu piena)
        fig.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["Fabbisogno Adjusted"],
            name="Fabbisogno Adjusted",
            mode="lines",
            line=dict(color="rgba(0,25,108,1)", width=3),
            hovertemplate="Fabbisogno Adjusted: %{y} GWh<extra></extra>"
        ))
        
        # 2️⃣ Fabbisogno Reale (blu medio tratteggiato)
        fig.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["Fabbisogno"],
            name="Fabbisogno (Reale)",
            mode="lines",
            line=dict(color="rgba(0,60,170,1)", width=3, dash="dash"),
            hovertemplate="Fabbisogno: %{y} GWh<extra></extra>"
        ))
        
        # 3️⃣ Coperture Secure (area verde media con tratteggio)
        fig.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["Coperture_secure"],
            name="Coperture Totali (Secure)",
            mode="lines",
            line=dict(color="rgba(0,150,0,1)", width=2, dash="dot"),
            fill="tozeroy",
            fillcolor="rgba(0,150,0,0.4)",
            hovertemplate="Coperture Secure: %{y} GWh<extra></extra>"
        ))
        
        # 4️⃣ Coperture Top (area verde chiaro con tratteggio)
        fig.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["Coperture_top"],
            name="Coperture Totali (Top)",
            mode="lines",
            line=dict(color="rgba(0,220,100,1)", width=2, dash="dash"),
            fill="tozeroy",
            fillcolor="rgba(0,220,100,0.3)",
            hovertemplate="Coperture Top: %{y} GWh<extra></extra>"
        ))
        
        # 5️⃣ Open Position Secure (bianco sopra coperture secure)
        fig.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["Fabbisogno Adjusted"],
            name="Open Position Secure",
            mode="lines",
            line=dict(color="white", width=2),
            fill="tonexty",
            fillcolor="rgba(255,255,255,1)",
            customdata=df["OpenPosition_secure"],
            hovertemplate="Open Position Secure: %{customdata} GWh<extra></extra>"
        ))
        
        # 6️⃣ Open Position Top (grigio chiaro sopra coperture top)
        fig.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["Fabbisogno Adjusted"],
            name="Open Position Top",
            mode="lines",
            line=dict(color="rgba(220,220,220,1)", width=2),
            fill="tonexty",
            fillcolor="rgba(230,230,230,0.8)",
            customdata=df["OpenPosition_top"],
            hovertemplate="Open Position Top: %{customdata} GWh<extra></extra>"
        ))
        
        # Layout
        fig.update_layout(
            title="📈 Fabbisogno vs Coperture Totali - Scenari Secure & Top",
            yaxis_title="GWh",
            xaxis_title="Anno",
            legend_title="Legenda",
            hovermode="x unified",
            xaxis=dict(tickformat="%Y", dtick="M12"),
            plot_bgcolor="white",
            paper_bgcolor="white"
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
