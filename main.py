import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from io import BytesIO

# Configurazione pagina
st.set_page_config(page_title="Calcolo Open Position", layout="wide")

st.title("📊 Analisi Fabbisogno Year by Year")

st.markdown("""
Carica un file Excel con la seguente struttura (colonne):
**Anno, Fabbisogno Adjusted, Fabbisogno, PPA ERG, PPA ERG cuscinetto, FRW, Solar**
""")

# Upload file
uploaded_file = st.file_uploader("📂 Carica file Excel", type=["xlsx"])

if uploaded_file:
    # Legge il file Excel
    df = pd.read_excel(uploaded_file)

    # Controllo colonne richieste
    required_cols = ["Anno", "Fabbisogno Adjusted", "Fabbisogno", 
                     "PPA ERG", "PPA ERG cuscinetto", "FRW", "Solar"]
    missing = [c for c in required_cols if c not in df.columns]
    
    if missing:
        st.error(f"Mancano le colonne obbligatorie: {', '.join(missing)}")
    else:
        # PPA effettivo = usa PPA ERG cuscinetto se presente, altrimenti PPA ERG
        df["PPA_effettivo"] = df["PPA ERG cuscinetto"].fillna(df["PPA ERG"])

        # Calcolo Open Position (Adjusted e No Adjusted)
        df["Open Position w Solar (Adjusted)"] = (
            df["Fabbisogno Adjusted"] - (df["PPA_effettivo"] + df["FRW"] + df["Solar"])
        )
        df["Open Position w/o Solar (Adjusted)"] = (
            df["Fabbisogno Adjusted"] - (df["PPA_effettivo"] + df["FRW"])
        )
        df["Open Position w Solar (No Adjusted)"] = (
            df["Fabbisogno"] - (df["PPA_effettivo"] + df["FRW"] + df["Solar"])
        )
        df["Open Position w/o Solar (No Adjusted)"] = (
            df["Fabbisogno"] - (df["PPA_effettivo"] + df["FRW"])
        )

        st.success("✅ Calcolo completato!")

        # Mostra tabella con risultati
        st.subheader("📈 Tabella con Open Position calcolati")
        st.dataframe(df.style.format("{:.0f}"))

        # Conversione "Anno" in datetime (solo anno)
        df["Anno"] = pd.to_datetime(df["Anno"], format="%Y")

        # Calcolo cumulativi per stacking
        df["PPA_cum"] = df["PPA_effettivo"]
        df["FRW_cum"] = df["PPA_effettivo"] + df["FRW"]
        df["Solar_cum"] = df["PPA_effettivo"] + df["FRW"] + df["Solar"]

        # --- Grafico con Plotly ---
        fig_solar = go.Figure()

        # 1️⃣ Fabbisogno Adjusted → blu scuro
        fig_solar.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["Fabbisogno Adjusted"],
            name="Fabbisogno Adjusted",
            mode='lines',
            line=dict(color='rgba(0, 25, 108, 1)', width=3),
            fill='tozeroy',
            fillcolor='rgba(0, 25, 108, 0.3)',
            hovertemplate='Fabbisogno Adjusted: %{y} GWh<extra></extra>'
        ))

        # 1️⃣.5️⃣ Fabbisogno → blu medio (nuovo layer)
        fig_solar.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["Fabbisogno"],
            name="Fabbisogno",
            mode='lines',
            line=dict(color='rgba(0, 60, 170, 1)', width=2, dash='dash'),
            fill=None,
            hovertemplate='Fabbisogno: %{y} GWh<extra></extra>'
        ))

        # 2️⃣ Coperture stacked (verde)
        # PPA
        fig_solar.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["PPA_cum"],
            name="PPA ERG/Cuscinetto",
            mode='lines',
            line=dict(color='rgba(0, 176, 80, 0.6)', width=2),
            fill='tonexty',
            fillcolor='rgba(0, 176, 80, 0.3)',
            hovertemplate='PPA: %{y} GWh<extra></extra>'
        ))
        # FRW
        fig_solar.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["FRW_cum"],
            name="FRW",
            mode='lines',
            line=dict(color='rgba(0, 176, 80, 0.8)', width=2),
            fill='tonexty',
            fillcolor='rgba(0, 176, 80, 0.5)',
            hovertemplate='FRW: %{y} GWh<extra></extra>'
        ))
        # Solar
        fig_solar.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["Solar_cum"],
            name="Solar",
            mode='lines',
            line=dict(color='rgba(0, 176, 80, 1)', width=2),
            fill='tonexty',
            fillcolor='rgba(0, 176, 80, 0.7)',
            hovertemplate='Solar: %{y} GWh<extra></extra>'
        ))

        # 3️⃣ Open Position → bianco
        fig_solar.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["Fabbisogno Adjusted"],
            name="Open Position",
            mode='lines',
            line=dict(color='white', width=2),
            fill='tonexty',
            fillcolor='rgba(255,255,255,1)',
            customdata=df["Open Position w Solar (Adjusted)"],
            hovertemplate='Open Position: %{customdata} GWh<extra></extra>'
        ))

        # Layout grafico
        fig_solar.update_layout(
            title="Scenario con Solar - Grafico ad Aree",
            yaxis_title="GWh",
            xaxis_title="Anno",
            legend_title="Legenda",
            hovermode="x unified",
            xaxis=dict(
                tickformat="%Y",
                dtick="M12"
            ),
            plot_bgcolor="white"
        )

        # Mostra grafico
        st.plotly_chart(fig_solar, use_container_width=True)

        # --- Download Excel risultato ---
        output_buffer = BytesIO()
        df.to_excel(output_buffer, index=False, engine="openpyxl")
        output_buffer.seek(0)

        st.download_button(
            label="📥 Scarica Excel con risultati",
            data=output_buffer,
            file_name="open_position_calcolato.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("⬆️ Carica un file Excel per iniziare.")
