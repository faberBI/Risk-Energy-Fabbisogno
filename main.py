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
        
        # --- Stack coperture Secure ---
        fig.add_trace(go.Scatter(
            x=df['Anno'], y=df['Solar'],
            name='Solar',
            mode='lines',
            line=dict(width=0.5, color='#FFD580'),
            fill='tozeroy',
            hovertemplate='Solar: %{y} GWh<extra></extra>'
        ))
        
        fig.add_trace(go.Scatter(
            x=df['Anno'], y=df['PPA ERG secure']+df['Solar'],
            name='PPA ERG Secure',
            mode='lines',
            line=dict(width=0.5, color='#A3C4DC'),
            fill='tonexty',
            customdata=df['Solar']
            hovertemplate='PPA ERG Secure: %{customdata} GWh<extra></extra>'
        ))
        
        fig.add_trace(go.Scatter(
            x=df['Anno'], y=df['FRW']+df['PPA ERG secure']+df['Solar'],
            name='FRW',
            mode='lines',
            line=dict(width=0.5, color='#C2EABD'),
            fill='tonexty',
            customdata=df['FWD'],
            hovertemplate='FRW: %{customdata} GWh<extra></extra>'
        ))
        
        # --- Area bianca Open Position Secure ---
        fig.add_trace(go.Scatter(
            x=df['Anno'], y=df['Open Position Secure'],
            name='Open Position Secure',
            mode='lines',
            line=dict(color='white', width=1.5),
            fill='tonexty',
            fillcolor='rgba(255,255,255,0.8)',
            hovertemplate='Open Position Secure: %{y} GWh<extra></extra>'
        ))
        
        # --- Linea Fabbisogno Totale ---
        fig.add_trace(go.Scatter(
            x=df['Anno'], y=df['Fabbisogno Adjusted'],
            name='Fabbisogno Totale',
            mode='lines',
            line=dict(color='black', width=2),
            hovertemplate='Fabbisogno Totale: %{y} GWh<extra></extra>'
        ))
        
        # --- Linea Copertura Totale Top (tratteggiata) ---
        fig.add_trace(go.Scatter(
            x=df['Anno'], y=df['Coperture Top'],
            name='Copertura Totale Top',
            mode='lines',
            line=dict(color='green', width=2, dash='dash'),
            hovertemplate='Copertura Top: %{y} GWh<extra></extra>'
        ))
        
        # --- Linea Open Position Top (tratteggiata) ---
        fig.add_trace(go.Scatter(
            x=df['Anno'], y=df['Open Position Top'],
            name='Open Position Top',
            mode='lines',
            line=dict(color='red', width=2, dash='dash'),
            hovertemplate='Open Position Top: %{y} GWh<extra></extra>'
        ))
        
        # Layout
        fig.update_layout(
            title='Fabbisogno vs Coperture Totali',
            xaxis_title='Anno',
            yaxis_title='Energia (GWh)',
            hovermode='x unified',
            plot_bgcolor='white',
            paper_bgcolor='white',
            legend_title='Legenda'
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
