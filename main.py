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

        # Grafico ad aree stacked
        st.subheader("📈 Andamento Fabbisogno Year by Year")

        fig = go.Figure()
        
        # PPA
        fig.add_trace(go.Scatter(
            x=df["Anno"], y=df["PPA_effettivo"], 
            name="PPA ERG/Cuscinetto", stackgroup='one', mode='none',
            fillcolor='lightblue'
        ))
        # FRW
        fig.add_trace(go.Scatter(
            x=df["Anno"], y=df["FRW"],
            name="FRW", stackgroup='one', mode='none',
            fillcolor='orange'
        ))
        # Solar
        fig.add_trace(go.Scatter(
            x=df["Anno"], y=df["Solar"],
            name="Solar", stackgroup='one', mode='none',
            fillcolor='yellow'
        ))
        # Open Position (residuo)
        fig.add_trace(go.Scatter(
            x=df["Anno"], y=df["Open Position"],
            name="Open Position", stackgroup='one', mode='none',
            fillcolor='red'
        ))
        
        # Linea fabbisogno adjusted
        fig.add_trace(go.Scatter(
            x=df["Anno"], y=df["Fabbisogno Adjusted"], 
            name="Fabbisogno Adjusted", mode='lines+markers',
            line=dict(color='black', dash='dash')
        ))
        
        fig.update_layout(
            title="Coperture vs Fabbisogno",
            yaxis_title="MW", xaxis_title="Anno",
            legend_title="Legenda",
            hovermode="x unified"
        )
        st.plotly_chart(fig, use_container_width=True)
        
        # Download del risultato
        output = df.to_excel(index=False, engine="openpyxl")
        st.download_button(
            label="📥 Scarica Excel con risultati",
            data=output,
            file_name="open_position_calcolato.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("⬆️ Carica un file Excel per iniziare.")
