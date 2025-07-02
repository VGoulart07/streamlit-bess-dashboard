import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os

plt.style.use('dark_background')

pasta = r"C:\Users\vinic\OneDrive - DSolar\DSOLAR\Projetos RCT\Projetos\Teste"
arquivo = os.path.join(pasta, "analise_diaria_bess_corrigida.xlsx")

df = pd.read_excel(arquivo, engine='openpyxl')
df.columns = [col.strip() for col in df.columns]

df['Data'] = pd.to_datetime(df['Data'])
df['AnoMes'] = df['Data'].dt.strftime('%Y-%m')
df['Dia'] = df['Data'].dt.day

df.rename(columns={
    df.columns[1]: 'Consumo (kWh)',
    df.columns[3]: 'Energia BESS (kWh)'
}, inplace=True)

capacidade_bateria = 1258.20  # Capacidade total útil
eficiencia = 0.98             # Eficiência de descarga (98%)
descarga_util = capacidade_bateria * eficiencia  # Energia efetivamente disponível por dia

for mes in df['AnoMes'].unique():
    df_mes = df[df['AnoMes'] == mes]

    dias = df_mes['Dia']
    consumo = df_mes['Consumo (kWh)']

    consumo_com_bess = consumo - descarga_util
    consumo_com_bess[consumo_com_bess < 0] = 0

    indices = np.arange(len(dias))
    largura = 0.4

    fig, ax = plt.subplots(figsize=(14, 7))
    fig.patch.set_facecolor('black')

    barras_sem = ax.bar(indices - largura/2, consumo, largura, label='Consumo Sem BESS', color='#00FFFF', edgecolor='#00CED1', linewidth=1.5)
    barras_com = ax.bar(indices + largura/2, consumo_com_bess, largura, label='Consumo Com BESS (eficiência 98%)', color='#FFA500', edgecolor='#FF8C00', linewidth=1.5)

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

    plt.tight_layout()
    nome_arquivo = os.path.join(pasta, f"grafico_comparativo_bess_{mes}.png")
    plt.savefig(nome_arquivo, dpi=150, facecolor=fig.get_facecolor())
    plt.close()
    print(f"Gráfico salvo: {nome_arquivo}")
