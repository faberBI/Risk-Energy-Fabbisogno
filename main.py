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

        df["Coperture Secure"] = df["PPA ERG secure"] + df["FRW"] + df["Solar"]
        df["Coperture Top"] = df["PPA ERG Top"] + df["FRW"] + df["Solar"]

        # Open position
        df["Open Position Secure"] = df["Fabbisogno Adjusted"] - df["Coperture Secure"]
        df["Open Position Top"] = df["Fabbisogno Adjusted"] - df["Coperture Top"]

        st.success("✅ Calcolo completato!")

        st.subheader("📈 Tabella con Open Position calcolati")
        st.dataframe(df.style.format("{:.0f}"))

        df["Anno"] = pd.to_datetime(df["Anno"], format="%Y")
        
        fig = go.Figure()
        
        # Stack coperture Secure
        fig.add_trace(go.Scatter(
            x=df['Anno'], y=df['Solar'],
            name='Solar',
            mode='lines',
            line=dict(width=0.5, color='#FFD580'),
            stackgroup='secure',  # <-- qui
            hovertemplate='Solar: %{y} GWh<extra></extra>'
        ))
        
        fig.add_trace(go.Scatter(
            x=df['Anno'], y=df['PPA ERG secure'],
            name='PPA ERG Secure',
            mode='lines',
            line=dict(width=0.5, color='#A3C4DC'),
            stackgroup='secure',  # <-- qui
            hovertemplate='PPA ERG Secure: %{y} GWh<extra></extra>'
        ))
       
        fig.add_trace(go.Scatter(
            x=df['Anno'], y=df['FRW'],
            name='FRW',
            mode='lines',
            line=dict(width=0.5, color='#C2EABD'),
            stackgroup='secure',  # <-- qui
            hovertemplate='FRW: %{y} GWh<extra></extra>'
        ))
        
        # Open Position Secure
        fig.add_trace(go.Scatter(
            x=df['Anno'], y=df['Open Position Secure'],
            name='Open Position Secure',
            mode='lines',
            line=dict(color='white', width=1.5),
            stackgroup='secure',  # <-- qui
            hovertemplate='Open Position Secure: %{y} GWh<extra></extra>'
        ))
        
        # Linee Fabbisogno e Top
        fig.add_trace(go.Scatter(
            x=df['Anno'], y=df['Fabbisogno Adjusted'],
            name='Fabbisogno Adjusted',
            mode='lines',
            line=dict(color='black', width=2, dash='dashdot'),
            hovertemplate='Fabbisogno Totale: %{y} GWh<extra></extra>'
        ))
                # Linee Fabbisogno e Top
        fig.add_trace(go.Scatter(
            x=df['Anno'], 
            y=df['Fabbisogno'],
            name='Fabbisogno',
            mode='lines',
            line=dict(color='blue', width=2, dash='solid'),  # linea blu tratteggiata
            hovertemplate='Fabbisogno Totale: %{y} GWh<extra></extra>'
            ))
        
        fig.add_trace(go.Scatter(
            x=df['Anno'], y=df['Coperture Top'],
            name='Copertura Totale Top',
            mode='lines',
            line=dict(color='green', width=2, dash='dash'),
            hovertemplate='Copertura Top: %{y} GWh<extra></extra>'
        ))

        fig.add_trace(go.Scatter(
            x=df['Anno'], y=df['Coperture Secure'],
            name='Copertura Totale Secure',
            mode='lines',
            line=dict(color='#4CAF50', width=2, dash='dashdot'),
            hovertemplate='Copertura Secure: %{y} GWh<extra></extra>'
        ))
        
        df['Anno'] = df['Anno'].dt.year        
        # Layout

        fig.update_layout(
            title="Confronto Scenari Secure vs Top - Coperture e Open Position",
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
