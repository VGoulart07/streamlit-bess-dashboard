import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os

st.set_page_config(page_title="Dashboard Consumo com BESS", layout="wide")
plt.style.use('dark_background')

pasta = r"C:\Users\vinic\OneDrive - DSolar\DSOLAR\Projetos RCT\Projetos\Teste"
arquivo = "analise_diaria_bess_corrigida.xlsx"

@st.cache_data
def carregar_dados(caminho):
    df = pd.read_excel(caminho, engine='openpyxl')
    df.columns = [col.strip() for col in df.columns]
    df['Data'] = pd.to_datetime(df['Data'])
    df['AnoMes'] = df['Data'].dt.strftime('%Y-%m')
    df['Dia'] = df['Data'].dt.day
    df.rename(columns={
        df.columns[1]: 'Consumo (kWh)',
        df.columns[3]: 'Energia BESS (kWh)'
    }, inplace=True)
    return df

df = carregar_dados(arquivo)

st.title("Dashboard Consumo com BESS")
st.markdown("Visualização do consumo com e sem BESS, além da redução percentual mensal.")

capacidade_bateria = st.slider("Capacidade total útil da bateria (kWh):", 
                              min_value=0.0, max_value=5000.0, value=1258.2, step=10.0)

# --- Cálculo da redução percentual total considerando a capacidade atual ---
consumo_total = df['Consumo (kWh)'].sum()
consumo_com_bess_total = consumo_total - capacidade_bateria
if consumo_com_bess_total < 0:
    consumo_com_bess_total = 0

reducao_total_percentual = 100 * (consumo_total - consumo_com_bess_total) / consumo_total

st.markdown(f"### Redução total estimada do consumo considerando capacidade de **{capacidade_bateria:.2f} kWh**: **{reducao_total_percentual:.2f}%**")
# ---

def plot_grafico_mensal(df, capacidade_bateria):
    meses = df['AnoMes'].unique()
    for mes in meses:
        df_mes = df[df['AnoMes'] == mes]
        dias = df_mes['Dia']
        consumo = df_mes['Consumo (kWh)']
        consumo_com_bess = consumo - capacidade_bateria
        consumo_com_bess[consumo_com_bess < 0] = 0

        indices = np.arange(len(dias))
        largura = 0.4

        fig, ax = plt.subplots(figsize=(12,6))
        fig.patch.set_facecolor('black')

        barras_sem = ax.bar(indices - largura/2, consumo, largura, label='Consumo Sem BESS', color='#00FFFF', edgecolor='#00CED1', linewidth=1.5)
        barras_com = ax.bar(indices + largura/2, consumo_com_bess, largura, label='Consumo Com BESS', color='#FFA500', edgecolor='#FF8C00', linewidth=1.5)

        ax.set_facecolor('black')
        ax.set_title(f"Consumo com e sem BESS - {mes}", color='cyan', fontsize=16)
        ax.set_xlabel('Dia do Mês', color='cyan', fontsize=14)
        ax.set_ylabel('Consumo (kWh)', color='cyan', fontsize=14)
        ax.set_xticks(indices)
        ax.set_xticklabels(dias, rotation=45, color='cyan')
        ax.tick_params(axis='y', colors='cyan')
        ax.legend(facecolor='black', edgecolor='white', fontsize=12)

        def add_labels(barras, valores):
            for barra, valor in zip(barras, valores):
                altura = barra.get_height()
                ax.text(barra.get_x() + barra.get_width()/2, altura + max(valores)*0.01,
                        f'{valor:.0f}', ha='center', va='bottom', color='white', fontsize=9, fontweight='bold')

        add_labels(barras_sem, consumo)
        add_labels(barras_com, consumo_com_bess)

        st.pyplot(fig)
        plt.close()

# --- Novo: cálculo da redução percentual diária ---
def calcula_reducao_diaria(df):
    # % redução = 100 * (1 - energia BESS / consumo)
    df['Reducao Percentual Diaria (%)'] = 100 * (1 - df['Energia BESS (kWh)'] / df['Consumo (kWh)'])
    df['Reducao Percentual Diaria (%)'] = df['Reducao Percentual Diaria (%)'].clip(lower=0, upper=100)
    return df

# --- Gráfico de redução percentual mensal ajustado ---
def plot_reducao_percentual_mensal(df):
    # calcula média da redução diária por mês
    resumo_mensal = df.groupby('AnoMes')['Reducao Percentual Diaria (%)'].mean().reset_index()

    fig, ax = plt.subplots(figsize=(12,6))
    fig.patch.set_facecolor('black')

    ax.bar(resumo_mensal['AnoMes'], resumo_mensal['Reducao Percentual Diaria (%)'],
           color='#32CD32', edgecolor='#228B22', linewidth=1.5)

    ax.set_facecolor('black')
    ax.set_title('Média Mensal da Redução Percentual Diária com BESS', color='lime', fontsize=16)
    ax.set_xlabel('Mês', color='lime', fontsize=14)
    ax.set_ylabel('Redução Percentual Média (%)', color='lime', fontsize=14)
    ax.tick_params(axis='x', rotation=45, colors='lime')
    ax.tick_params(axis='y', colors='lime')

    for i, valor in enumerate(resumo_mensal['Reducao Percentual Diaria (%)']):
        ax.text(i, valor + 0.5, f'{valor:.1f}%', ha='center', va='bottom', color='white', fontsize=10, fontweight='bold')

    st.pyplot(fig)
    plt.close()

# --- Aplica cálculo da redução diária no dataframe ---
df = calcula_reducao_diaria(df)

st.header("Gráficos Mensais de Consumo")
plot_grafico_mensal(df, capacidade_bateria)

st.header("Redução Percentual Mensal")
plot_reducao_percentual_mensal(df)
