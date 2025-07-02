import pandas as pd
import matplotlib.pyplot as plt
import os
import numpy as np

plt.style.use('dark_background')

# Caminhos
pasta = r"C:\Users\vinic\OneDrive - DSolar\DSOLAR\Projetos RCT\Projetos\Teste"
arquivo_entrada = os.path.join(pasta, "Memória_de_Massa_-_Hiper_Monlevade(1).xlsx")
arquivo_saida_excel = os.path.join(pasta, "analise_diaria_bess_corrigida.xlsx")

# Parâmetros BESS
num_baterias = 2
capacidade_unitaria_kWh = 699
potencia_unitaria_kW = 375
dod = 0.90
eficiencia = 0.98

capacidade_total_kWh = capacidade_unitaria_kWh * num_baterias * dod
potencia_total_kW = potencia_unitaria_kW * num_baterias

print(f"Capacidade útil total do BESS: {capacidade_total_kWh:.2f} kWh")
print(f"Potência total do BESS: {potencia_total_kW} kW")
print(f"Eficiência de descarga: {eficiencia*100:.1f}%\n")

# Leitura dados
df_demanda = pd.read_excel(arquivo_entrada, sheet_name=1, engine='openpyxl')
df_energia = pd.read_excel(arquivo_entrada, sheet_name=2, engine='openpyxl')

# Preparação
df_demanda['Data'] = pd.to_datetime(df_demanda['Data'])
df_demanda = df_demanda[df_demanda['Posto'] == 'HP'].copy()
df_demanda['Demanda (kW)'] = df_demanda['DemandaAtivaConsumokW']

df_energia['Data'] = pd.to_datetime(df_energia['Data'])
df_energia = df_energia[df_energia['Posto'] == 'HP'].copy()
df_energia['Consumo (kWh)'] = df_energia['EnergiaAtivaConsumokWh']

# Agregação diária
demanda_diaria = df_demanda.groupby(df_demanda['Data'].dt.date)['Demanda (kW)'].max().reset_index()
demanda_diaria['Data'] = pd.to_datetime(demanda_diaria['Data'])

consumo_diario = df_energia.groupby(df_energia['Data'].dt.date)['Consumo (kWh)'].sum().reset_index()
consumo_diario['Data'] = pd.to_datetime(consumo_diario['Data'])

df_merged = pd.merge(consumo_diario, demanda_diaria, on='Data', how='outer').sort_values('Data')

# Excedente em potência que ultrapassa o limite do BESS
df_merged['Excedente Potência (kW)'] = df_merged['Demanda (kW)'] - potencia_total_kW
df_merged['Excedente Potência (kW)'] = df_merged['Excedente Potência (kW)'].apply(lambda x: x if x > 0 else 0)

# Convertendo pico excedente em energia (kWh) - assumindo consumo horário médio (1h)
# Se dados são de intervalo variável, ajustar aqui
energia_excedente_potencia_kWh = df_merged['Excedente Potência (kW)']  # simplificação 1 hora

# Energia disponível na bateria por dia (considerando DoD e eficiência)
energia_disponivel_kWh = capacidade_total_kWh * eficiencia

# Energia que pode ser "suprida" pelo BESS no dia (considera demanda e capacidade)
# Primeiro cobrir o pico excedente (em energia), depois o restante do consumo excedente até limite da bateria
# Energia excedente total diária = consumo - energia disponível do BESS - pico excedente (não coberto)

# Energia excedente = consumo - energia útil do BESS - energia que o BESS pode entregar para o pico
# Na prática, o BESS reduz o pico até o limite de potência e capacidade; o excedente é o que sobra

# Definimos energia excedente diária como:
# energia_excedente = max(0, consumo - energia_disponivel_kWh)

# Porém, o pico excedente deve ser somado pois o BESS não cobre além da potência
# Logo, energia excedente real = energia_excedente + energia equivalente ao pico excedente (em kWh)

energia_excedente_kWh = df_merged['Consumo (kWh)'] - energia_disponivel_kWh
energia_excedente_kWh = energia_excedente_kWh.apply(lambda x: x if x > 0 else 0)

df_merged['Energia Excedente (kWh)'] = energia_excedente_kWh + energia_excedente_potencia_kWh

# Percentual de redução da energia excedente considerando o BESS
energia_sem_bess = df_merged['Consumo (kWh)']
energia_com_bess = energia_sem_bess - df_merged['Energia Excedente (kWh)']
energia_com_bess = energia_com_bess.apply(lambda x: x if x > 0 else 0)

df_merged['% Redução Energia Excedente'] = 100 * (energia_com_bess / energia_sem_bess)
df_merged['% Redução Energia Excedente'] = df_merged['% Redução Energia Excedente'].fillna(0)

# Salvar Excel com resultados
df_merged.to_excel(arquivo_saida_excel, index=False)
print(f"Arquivo Excel salvo em: {arquivo_saida_excel}")

# --- Gráfico adicional de redução percentual diária ---
plt.figure(figsize=(15,6))
plt.plot(df_merged['Data'], df_merged['% Redução Energia Excedente'], color='#00FF00', linewidth=2)
plt.title('Redução Percentual Diária de Energia Excedente Com BESS', fontsize=16, color='cyan')
plt.xlabel('Data', color='cyan', fontsize=14)
plt.ylabel('% Redução', color='cyan', fontsize=14)
plt.grid(True, color='gray')
plt.ylim(0, 110)
plt.xticks(rotation=45)
plt.tight_layout()

arquivo_grafico_reducao = os.path.join(pasta, 'grafico_reducao_percentual_diaria.png')
plt.savefig(arquivo_grafico_reducao, dpi=150, facecolor='black')
plt.close()
print(f"Gráfico de redução percentual salvo: {arquivo_grafico_reducao}")
