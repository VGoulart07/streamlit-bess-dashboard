import pandas as pd
import matplotlib.pyplot as plt
import os
import numpy as np

# Configuração geral do estilo dos gráficos
plt.style.use('dark_background')  # fundo preto

# Caminho da pasta de entrada e saída
pasta = r"C:\Users\vinic\OneDrive - DSolar\DSOLAR\Projetos RCT\Projetos\Teste"
arquivo_entrada = os.path.join(pasta, "Memória_de_Massa_-_Hiper_Monlevade(1).xlsx")
arquivo_saida_excel = os.path.join(pasta, "consumo_demanda_mensal.xlsx")  # mantém o excel com abas, se quiser

# --- Leitura dos dados ---
df_demanda = pd.read_excel(arquivo_entrada, sheet_name=1, engine='openpyxl')
df_energia = pd.read_excel(arquivo_entrada, sheet_name=2, engine='openpyxl')

# Preparação e filtros
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

# Juntando dados
df_merged = pd.merge(consumo_diario, demanda_diaria, on='Data', how='outer').sort_values('Data')
df_merged['AnoMes'] = df_merged['Data'].dt.strftime('%Y-%m')

# Função para criar gráfico de barras com labels neon
def criar_grafico_barras(df_mes, mes, pasta_saida):
    dias = df_mes['Data'].dt.day
    consumo = df_mes['Consumo (kWh)']
    demanda = df_mes['Demanda (kW)']

    largura = 0.4
    indices = np.arange(len(dias))

    fig, ax = plt.subplots(figsize=(14, 7))
    fig.patch.set_facecolor('black')  # fundo preto

    barras_consumo = ax.bar(indices - largura/2, consumo, largura, label='Consumo (kWh)', color='#00FFFF', edgecolor='#00CED1', linewidth=1.5)
    barras_demanda = ax.bar(indices + largura/2, demanda, largura, label='Demanda Máx (kW)', color='#FFA500', edgecolor='#FF8C00', linewidth=1.5)

    ax.set_facecolor('black')
    ax.set_title(f"Consumo e Demanda - {mes}", color='cyan', fontsize=16)
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

    plt.tight_layout()

    # Salvar PNG
    arquivo_png = os.path.join(pasta_saida, f"grafico_barras_consumo_demanda_{mes}.png")
    plt.savefig(arquivo_png, dpi=150, facecolor=fig.get_facecolor())
    plt.close()

    print(f"Gráfico salvo: {arquivo_png}")

# Gerar gráfico PNG por mês
for mes in df_merged['AnoMes'].unique():
    df_mes = df_merged[df_merged['AnoMes'] == mes].copy()
    criar_grafico_barras(df_mes, mes, pasta)

print("Todos os gráficos PNG foram gerados com sucesso!")
