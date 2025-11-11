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

        # --- Calcolo coperture complessive e componenti ---
        df["PPA_cum_secure"] = df["PPA ERG secure"]
        df["PPA_cum_top"] = df["PPA ERG Top"]

        df["Coperture secure"] = df["PPA ERG secure"] + df["FRW"] + df["Solar"]
        df["Coperture top"] = df["PPA ERG Top"] + df["FRW"] + df["Solar"]

        # Open position
        df["Open Position Secure"] = df["Fabbisogno Adjusted"] - df["Coperture secure"]
        df["Open Position Top"] = df["Fabbisogno Adjusted"] - df["Coperture top"]

        st.success("✅ Calcolo completato!")

        st.subheader("📈 Tabella con Open Position calcolati")
        st.dataframe(df.style.format("{:.0f}"))

        df["Anno"] = pd.to_datetime(df["Anno"], format="%Y")

        fig = go.Figure()

        # 🎯 1️⃣ Fabbisogno Adjusted
        fig.add_trace(go.Scatter(
            x=df["Anno"], y=df["Fabbisogno Adjusted"],
            name="Fabbisogno Adjusted",
            mode="lines",
            line=dict(color="rgba(0,25,108,1)", width=3),
            hovertemplate="Fabbisogno Adjusted: %{y} GWh<extra></extra>"
        ))
        
        # 🎯 2️⃣ Fabbisogno Reale
        fig.add_trace(go.Scatter(
            x=df["Anno"], y=df["Fabbisogno"],
            name="Fabbisogno (Reale)",
            mode="lines",
            line=dict(color="rgba(0,60,170,1)", width=3, dash="dash"),
            hovertemplate="Fabbisogno: %{y} GWh<extra></extra>"
        ))
        
        # 🎯 3️⃣ Coperture Secure (stack: PPA + FRW + Solar)
        fig.add_trace(go.Scatter(
            x=df["Anno"], y=df["PPA ERG secure"],
            name="PPA Secure",
            mode="lines",
            line=dict(color="rgba(255,200,0,1)", width=2),
            fill="tozeroy",
            fillcolor="rgba(255,200,0,0.5)",
            hovertemplate="PPA Secure: %{y} GWh<extra></extra>"
        ))
        fig.add_trace(go.Scatter(
            x=df["Anno"], y=df["PPA ERG secure"] + df["FRW"],
            name="FRW Secure",
            mode="lines",
            line=dict(color="rgba(0,100,255,1)", width=2),
            fill="tonexty",
            fillcolor="rgba(0,100,255,0.4)",
            hovertemplate="FRW Secure: %{y} GWh<extra></extra>"
        ))
        fig.add_trace(go.Scatter(
            x=df["Anno"], y=df["Coperture secure"],
            name="Solar Secure",
            mode="lines",
            line=dict(color="rgba(255,140,0,1)", width=2),
            fill="tonexty",
            fillcolor="rgba(255,140,0,0.4)",
            hovertemplate="Solar Secure: %{y} GWh<extra></extra>"
        ))
        
        # 🎯 4️⃣ Coperture Top (stack: PPA + FRW + Solar) tratteggiate
        fig.add_trace(go.Scatter(
            x=df["Anno"], y=df["PPA ERG Top"],
            name="PPA Top",
            mode="lines",
            line=dict(color="rgba(255,220,80,1)", width=2, dash="dot"),
            fill="tozeroy",
            fillcolor="rgba(255,220,80,0.3)",
            hovertemplate="PPA Top: %{y} GWh<extra></extra>"
        ))
        fig.add_trace(go.Scatter(
            x=df["Anno"], y=df["PPA ERG Top"] + df["FRW"],
            name="FRW Top",
            mode="lines",
            line=dict(color="rgba(100,160,255,1)", width=2, dash="dot"),
            fill="tonexty",
            fillcolor="rgba(100,160,255,0.3)",
            hovertemplate="FRW Top: %{y} GWh<extra></extra>"
        ))
        fig.add_trace(go.Scatter(
            x=df["Anno"], y=df["Coperture top"],
            name="Solar Top",
            mode="lines",
            line=dict(color="rgba(255,180,80,1)", width=2, dash="dot"),
            fill="tonexty",
            fillcolor="rgba(255,180,80,0.3)",
            hovertemplate="Solar Top: %{y} GWh<extra></extra>"
        ))
        
        # 🎯 5️⃣ Open Position Secure
        fig.add_trace(go.Scatter(
            x=df["Anno"], y=df["Fabbisogno Adjusted"],
            name="Open Position Secure",
            mode="lines",
            line=dict(color="white", width=2),
            fill="tonexty",
            fillcolor="rgba(255,255,255,1)",
            customdata=df["Open Position Secure"],
            hovertemplate="Open Position Secure: %{customdata} GWh<extra></extra>"
        ))
        
        # 🎯 6️⃣ Open Position Top
        fig.add_trace(go.Scatter(
            x=df["Anno"], y=df["Fabbisogno Adjusted"],
            name="Open Position Top",
            mode="lines",
            line=dict(color="rgba(220,220,220,1)", width=2),
            fill="tonexty",
            fillcolor="rgba(230,230,230,0.9)",
            customdata=df["Open Position Top"],
            hovertemplate="Open Position Top: %{customdata} GWh<extra></extra>"
        ))
        
        # --- Layout ---
        fig.update_layout(
            title="📊 Fabbisogno, Coperture (PPA + FRW + Solar) e Open Position – Secure vs Top",
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

else:
    st.info("⬆️ Carica un file Excel per iniziare.")
