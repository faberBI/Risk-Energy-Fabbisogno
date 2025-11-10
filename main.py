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

        # Calcolo copertura totale
        df["Copertura_totale_con_solar"] = df["PPA_effettivo"] + df["FRW"] + df["Solar"]
        df["OpenPosition_con_solar"] = (df["Fabbisogno Adjusted"] - df["Copertura_totale_con_solar"]).clip(lower=0)
        
        fig_solar = go.Figure()
        
        # 1️⃣ Fabbisogno Adjusted → sfondo blu
        fig_solar.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["Fabbisogno Adjusted"],
            name="Fabbisogno Adjusted",
            mode='lines',
            line=dict(color='blue'),
            fill='tozeroy',
            fillcolor='rgba(0,0,255,0.2)',
            hovertemplate='Fabbisogno Adjusted: %{y} MW<extra></extra>'
        ))
        
        # 2️⃣ Coperture stacked (PPA + FRW + Solar)
        # PPA
        fig_solar.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["PPA_effettivo"],
            name="PPA ERG/Cuscinetto",
            mode='lines',
            line=dict(color='green'),
            fill='tonexty',
            fillcolor='rgba(0,128,0,0.7)',
            hovertemplate='PPA: %{y} MW<extra></extra>'
        ))
        # PPA + FRW
        fig_solar.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["PPA_effettivo"] + df["FRW"],
            name="FRW",
            mode='lines',
            line=dict(color='deepskyblue'),
            fill='tonexty',
            fillcolor='rgba(135,206,250,0.7)',
            hovertemplate='FRW: %{y} MW<extra></extra>'
        ))
        # PPA + FRW + Solar
        fig_solar.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["Copertura_totale_con_solar"],
            name="Solar",
            mode='lines',
            line=dict(color='gold'),
            fill='tonexty',
            fillcolor='rgba(255,215,0,0.7)',
            hovertemplate='Solar: %{y} MW<extra></extra>'
        ))
        
        # 3️⃣ Open Position → area bianca sopra copertura fino al fabbisogno
        fig_solar.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["Fabbisogno Adjusted"],  # in alto layer = fabbisogno totale
            name="Open Position",
            mode='lines',
            line=dict(color='white'),
            fill='tonexty',
            fillcolor='rgba(255,255,255,1)',
            hovertemplate='Open Position: %{y - df.Copertura_totale_con_solar} MW<extra></extra>'
        ))
        
        fig_solar.update_layout(
            title="Scenario con Solar - Grafico ad Aree",
            yaxis_title="MW",
            xaxis_title="Anno",
            legend_title="Legenda",
            hovermode="x unified"
        )
        
        st.plotly_chart(fig_solar, use_container_width=True)


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
