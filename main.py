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

        # -----------------------------
        # Calcoli per scenario con Solar
        # -----------------------------
        df["Copertura_totale_con_solar"] = df["PPA_effettivo"] + df["FRW"] + df["Solar"]
        df["OpenPosition_con_solar"] = (df["Fabbisogno Adjusted"] - df["Copertura_totale_con_solar"]).clip(lower=0)
        
        fig_solar = go.Figure()
        
        # Prima traccia: Open Position (in basso o sopra? dipende dal colore, qui la mettiamo in basso come stack residuo)
        fig_solar.add_trace(go.Bar(
            x=df["Anno"],
            y=df["OpenPosition_con_solar"],
            name="Open Position",
            marker=dict(color="red", pattern=dict(shape="\\"))
        ))
        
        # Coperture stacked: PPA, FRW, Solar
        fig_solar.add_trace(go.Bar(
            x=df["Anno"],
            y=df["PPA_effettivo"],
            name="PPA ERG/Cuscinetto",
            marker=dict(color="lightblue")
        ))
        fig_solar.add_trace(go.Bar(
            x=df["Anno"],
            y=df["FRW"],
            name="FRW",
            marker=dict(color="orange")
        ))
        fig_solar.add_trace(go.Bar(
            x=df["Anno"],
            y=df["Solar"],
            name="Solar",
            marker=dict(color="yellow")
        ))
        
        # Area totale Fabbisogno (sfondo)
        fig_solar.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["Fabbisogno Adjusted"],
            name="Fabbisogno Adjusted",
            mode="lines",
            line=dict(color="black", dash="dash")
        ))
        
        fig_solar.update_layout(
            barmode='stack',
            title="Scenario con Solar",
            yaxis_title="MW",
            xaxis_title="Anno",
            legend_title="Legenda",
            hovermode="x unified"
        )
        
        st.plotly_chart(fig_solar, use_container_width=True)
        
        
        # -----------------------------
        # Scenario senza Solar
        # -----------------------------
        df["Copertura_totale_senza_solar"] = df["PPA_effettivo"] + df["FRW"]
        df["OpenPosition_senza_solar"] = (df["Fabbisogno Adjusted"] - df["Copertura_totale_senza_solar"]).clip(lower=0)
        
        fig_no_solar = go.Figure()
        
        # Open Position
        fig_no_solar.add_trace(go.Bar(
            x=df["Anno"],
            y=df["OpenPosition_senza_solar"],
            name="Open Position",
            marker=dict(color="red", pattern=dict(shape="\\"))
        ))
        
        # Coperture stacked
        fig_no_solar.add_trace(go.Bar(
            x=df["Anno"],
            y=df["PPA_effettivo"],
            name="PPA ERG/Cuscinetto",
            marker=dict(color="lightblue")
        ))
        fig_no_solar.add_trace(go.Bar(
            x=df["Anno"],
            y=df["FRW"],
            name="FRW",
            marker=dict(color="orange")
        ))
        
        # Linea Fabbisogno Adjusted
        fig_no_solar.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["Fabbisogno Adjusted"],
            name="Fabbisogno Adjusted",
            mode="lines",
            line=dict(color="black", dash="dash")
        ))
        
        fig_no_solar.update_layout(
            barmode='stack',
            title="Scenario senza Solar",
            yaxis_title="MW",
            xaxis_title="Anno",
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
