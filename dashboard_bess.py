import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os

# Caminho do arquivo Excel (ajuste se precisar)
pasta = r"C:\Users\vinic\OneDrive - DSolar\DSOLAR\Projetos RCT\Projetos\Teste"
arquivo_excel = os.path.join(pasta, "Memória_de_Massa_-_Hiper_Monlevade(1).xlsx")

@st.cache_data
def carregar_dados():
    df_demanda = pd.read_excel(arquivo_excel, sheet_name=1, engine='openpyxl')
    df_energia = pd.read_excel(arquivo_excel, sheet_name=2, engine='openpyxl')

    df_demanda['Data'] = pd.to_datetime(df_demanda['Data'])
    df_demanda = df_demanda[df_demanda['Posto'] == 'HP'].copy()
    df_demanda['Demanda (kW)'] = df_demanda['DemandaAtivaConsumokW']
    demanda_diaria = df_demanda.groupby(df_demanda['Data'].dt.date)['Demanda (kW)'].max().reset_index()
    demanda_diaria['Data'] = pd.to_datetime(demanda_diaria['Data'])

    df_energia['Data'] = pd.to_datetime(df_energia['Data'])
    df_energia = df_energia[df_energia['Posto'] == 'HP'].copy()
    df_energia['Consumo (kWh)'] = df_energia['EnergiaAtivaConsumokWh']
    consumo_diario = df_energia.groupby(df_energia['Data'].dt.date)['Consumo (kWh)'].sum().reset_index()
    consumo_diario['Data'] = pd.to_datetime(consumo_diario['Data'])

    df_merged = pd.merge(consumo_diario, demanda_diaria, on='Data', how='outer').sort_values('Data')
    df_merged['AnoMes'] = df_merged['Data'].dt.strftime('%Y-%m')

    return df_merged

df = carregar_dados()

st.title("Dashboard de Consumo e Demanda - BESS")
st.markdown("Visualize o consumo de energia diária (kWh) e a demanda máxima (kW) por mês.")

# Seletor de mês
meses = df['AnoMes'].unique()
mes_selecionado = st.selectbox("Selecione o mês para visualização", meses)

df_mes = df[df['AnoMes'] == mes_selecionado].copy()

# Plotagem do gráfico de barras lado a lado
def plot_barras(df_mes):
    dias = df_mes['Data'].dt.day
    consumo = df_mes['Consumo (kWh)']
    demanda = df_mes['Demanda (kW)']

    largura = 0.4
    indices = np.arange(len(dias))

    fig, ax = plt.subplots(figsize=(14, 7))
    fig.patch.set_facecolor('black')
    ax.set_facecolor('black')

    barras_consumo = ax.bar(indices - largura/2, consumo, largura, label='Consumo (kWh)', color='#00FFFF', edgecolor='#00CED1', linewidth=1.5)
    barras_demanda = ax.bar(indices + largura/2, demanda, largura, label='Demanda Máx (kW)', color='#FFA500', edgecolor='#FF8C00', linewidth=1.5)

    ax.set_title(f"Consumo e Demanda - {mes_selecionado}", color='cyan', fontsize=16)
    ax.set_xlabel('Dia do Mês', color='cyan', fontsize=14)
    ax.set_ylabel('kWh / kW', color='cyan', fontsize=14)
    ax.set_xticks(indices)
    ax.set_xticklabels(dias, rotation=45, color='cyan')
    ax.tick_params(axis='y', colors='cyan')

    ax.legend(facecolor='black', edgecolor='white', fontsize=12)

    # Adiciona os valores em cima das barras
    def adicionar_labels(barras, valores):
        for barra, valor in zip(barras, valores):
            altura = barra.get_height()
            ax.text(
                barra.get_x() + barra.get_width()/2,
                altura + max(valores)*0.01,
                f'{valor:.0f}',
                ha='center', va='bottom', color='white', fontsize=9, fontweight='bold'
            )

    adicionar_labels(barras_consumo, consumo)
    adicionar_labels(barras_demanda, demanda)

    st.pyplot(fig)

plot_barras(df_mes)

# Estatísticas resumidas do mês
st.markdown(f"### Resumo - {mes_selecionado}")
st.write(f"**Total consumido:** {df_mes['Consumo (kWh)'].sum():,.2f} kWh")
st.write(f"**Demanda máxima registrada:** {df_mes['Demanda (kW)'].max():,.2f} kW")
st.write(f"**Demanda média:** {df_mes['Demanda (kW)'].mean():,.2f} kW")
