import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import plotly.graph_objects as go


st.set_page_config(page_title="Calcolo Open Position", layout="wide")

st.title("📊 Analisi Fabbisogno Year by Year")

st.markdown("""
Carica un file Excel con la seguente struttura (colonne):
**Anno, Fabbisogno Adjusted, Fabbisogno, PPA ERG, PPA ERG cuscinetto, FRW, Solar**
""")

uploaded_file = st.file_uploader("📂 Carica file Excel", type=["xlsx"])

if uploaded_file:
    # Legge il file Excel
    df = pd.read_excel(uploaded_file)

    # Controllo colonne
    required_cols = ["Anno", "Fabbisogno Adjusted", "Fabbisogno", 
                     "PPA ERG", "PPA ERG cuscinetto", "FRW", "Solar"]
    
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"Mancano le colonne obbligatorie: {', '.join(missing)}")
    else:
        # Usa PPA ERG cuscinetto se presente (non NaN)
        df["PPA_effettivo"] = df["PPA ERG cuscinetto"].fillna(df["PPA ERG"])

        # Calcolo Open Position (Adjusted)
        df["Open Position w Solar (Adjusted)"] = (
            df["Fabbisogno Adjusted"] - (df["PPA_effettivo"] + df["FRW"] + df["Solar"])
        )
        df["Open Position w/o Solar (Adjusted)"] = (
            df["Fabbisogno Adjusted"] - (df["PPA_effettivo"] + df["FRW"])
        )

        # Calcolo Open Position (No Adjusted)
        df["Open Position w Solar (No Adjusted)"] = (
            df["Fabbisogno"] - (df["PPA_effettivo"] + df["FRW"] + df["Solar"])
        )
        df["Open Position w/o Solar (No Adjusted)"] = (
            df["Fabbisogno"] - (df["PPA_effettivo"] + df["FRW"])
        )

        st.success("✅ Calcolo completato!")

        st.subheader("📈 Tabella con Open Position calcolati")
        st.dataframe(df.style.format("{:.0f}"))

        # -------------------------
        # Grafico con Solar
        st.subheader("📊 Scenario: con Solar")
        
        # Calcolo copertura totale
        df["Copertura_totale_con_solar"] = df["PPA_effettivo"] + df["FRW"] + df["Solar"]
        
        fig_solar = go.Figure()
        
        # Aree stacked: PPA, FRW, Solar
        fig_solar.add_trace(go.Scatter(
            x=df["Anno"], y=df["PPA_effettivo"], 
            name="PPA ERG/Cuscinetto", stackgroup='one', mode='none', fillcolor='lightblue'
        ))
        fig_solar.add_trace(go.Scatter(
            x=df["Anno"], y=df["FRW"], 
            name="FRW", stackgroup='one', mode='none', fillcolor='orange'
        ))
        fig_solar.add_trace(go.Scatter(
            x=df["Anno"], y=df["Solar"], 
            name="Solar", stackgroup='one', mode='none', fillcolor='yellow'
        ))
        
        # Open Position come area tratteggiata rossa
        fig_solar.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["Fabbisogno Adjusted"],
            name="Open Position",
            mode='lines',
            line=dict(color='red', dash='dash'),
            fill='tonexty',  # riempie l'area tra questa trace e la precedente
            fillcolor='rgba(255,0,0,0.3)'  # rosso trasparente
        ))
        
        # Linea di riferimento fabbisogno Adjusted (opzionale)
        fig_solar.add_trace(go.Scatter(
            x=df["Anno"], y=df["Fabbisogno Adjusted"],
            name="Fabbisogno Adjusted",
            mode='lines',
            line=dict(color='black', dash='dash')
        ))
        
        fig_solar.update_layout(
            title="Copertura con Solar",
            yaxis_title="MW", xaxis_title="Anno",
            legend_title="Legenda",
            hovermode="x unified"
        )
        
        st.plotly_chart(fig_solar, use_container_width=True)
        
        
        # -------------------------
        # Grafico senza Solar
        st.subheader("📊 Scenario: senza Solar")
        
        # Calcolo copertura totale senza Solar
        df["Copertura_totale_senza_solar"] = df["PPA_effettivo"] + df["FRW"]
        
        fig_no_solar = go.Figure()
        
        # Aree stacked: PPA, FRW
        fig_no_solar.add_trace(go.Scatter(
            x=df["Anno"], y=df["PPA_effettivo"], 
            name="PPA ERG/Cuscinetto", stackgroup='one', mode='none', fillcolor='lightblue'
        ))
        fig_no_solar.add_trace(go.Scatter(
            x=df["Anno"], y=df["FRW"], 
            name="FRW", stackgroup='one', mode='none', fillcolor='orange'
        ))
        
        # Open Position come area tratteggiata rossa
        fig_no_solar.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["Fabbisogno Adjusted"],
            name="Open Position",
            mode='lines',
            line=dict(color='red', dash='dash'),
            fill='tonexty',  # riempie l'area tra questa trace e la precedente
            fillcolor='rgba(255,0,0,0.3)'  # rosso trasparente
        ))
        
        # Linea di riferimento fabbisogno Adjusted
        fig_no_solar.add_trace(go.Scatter(
            x=df["Anno"], y=df["Fabbisogno Adjusted"],
            name="Fabbisogno Adjusted",
            mode='lines',
            line=dict(color='black', dash='dash')
        ))
        
        fig_no_solar.update_layout(
            title="Copertura senza Solar",
            yaxis_title="MW", xaxis_title="Anno",
            legend_title="Legenda",
            hovermode="x unified"
        )
        
        st.plotly_chart(fig_no_solar, use_container_width=True)
        
        # Download del risultato
        from io import BytesIO

        # Creiamo un buffer in memoria
        output_buffer = BytesIO()
        df.to_excel(output_buffer, index=False, engine="openpyxl")
        output_buffer.seek(0)  # Torniamo all'inizio del buffer

        st.download_button(
            label="📥 Scarica Excel con risultati",
            data=output_buffer,
            file_name="open_position_calcolato.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("⬆️ Carica un file Excel per iniziare.")
