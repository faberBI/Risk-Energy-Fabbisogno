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
        df["PPA_cum_secure"] = df["PPA ERG secure"] + df["FRW"] + df["Solar"]
        df["PPA_cum_top"] = df["PPA ERG Top"] + df["FRW"] + df["Solar"]

        df["Open Secure"] = df["Fabbisogno Adjusted"] - df["PPA_cum_secure"]
        df["Open Top"] = df["Fabbisogno Adjusted"] - df["PPA_cum_top"]

        st.success("✅ Calcolo completato!")

        st.subheader("📈 Tabella con Open Position calcolati")
        st.dataframe(df.style.format("{:.0f}"))

        df["Anno"] = pd.to_datetime(df["Anno"], format="%Y")

        # --- Grafico Plotly ---
        fig = go.Figure()

        # 1️⃣ Fabbisogno Adjusted
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

        # 2️⃣ Area Copertura Secure (verde scuro con tratteggio simulato)
        fig.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["PPA_cum_secure"],
            name="Copertura Secure",
            mode="lines",
            line=dict(color="rgba(0, 100, 0, 1)", width=2, dash="dash"),
            fill="tozeroy",
            fillcolor="rgba(0, 100, 0, 0.4)",
            hovertemplate="Copertura Secure: %{y} GWh<extra></extra>"
        ))

        # 3️⃣ Area Copertura Top (verde chiaro con tratteggio simulato)
        fig.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["PPA_cum_top"],
            name="Copertura Top",
            mode="lines",
            line=dict(color="rgba(0, 200, 80, 1)", width=2, dash="dash"),
            fill="tozeroy",
            fillcolor="rgba(0, 200, 80, 0.35)",
            hovertemplate="Copertura Top: %{y} GWh<extra></extra>"
        ))

        # 4️⃣ Open Position Secure
        fig.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["Fabbisogno Adjusted"],
            name="Open Position Secure",
            mode="lines",
            line=dict(color="white", width=1.5),
            fill="tonexty",
            fillcolor="rgba(255,255,255,0.9)",
            hovertemplate="Open Position Secure: %{y} GWh<extra></extra>"
        ))

        # 5️⃣ Open Position Top
        fig.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["Fabbisogno Adjusted"],
            name="Open Position Top",
            mode="lines",
            line=dict(color="rgba(200,200,200,1)", width=1.5),
            fill="tonexty",
            fillcolor="rgba(220,220,220,0.8)",
            hovertemplate="Open Position Top: %{y} GWh<extra></extra>"
        ))

        # 6️⃣ Fabbisogno reale
        fig.add_trace(go.Scatter(
            x=df["Anno"],
            y=df["Fabbisogno"],
            name="Fabbisogno (Reale)",
            mode="lines",
            line=dict(color="rgba(0, 60, 170, 1)", width=3, dash="dash"),
            hovertemplate="Fabbisogno: %{y} GWh<extra></extra>"
        ))

        # Layout finale
        fig.update_layout(
            title="Coperture e Open Position - Scenari Secure vs Top",
            yaxis_title="GWh",
            xaxis_title="Anno",
            legend_title="Legenda",
            hovermode="x unified",
            plot_bgcolor="white",
            xaxis=dict(tickformat="%Y", dtick="M12")
        )

        st.plotly_chart(fig, use_container_width=True)

        # Download Excel
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
