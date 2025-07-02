import pandas as pd
import matplotlib.pyplot as plt

# Caminho do arquivo Excel
arquivo = r"C:\Users\vinic\OneDrive - DSolar\DSOLAR\Projetos RCT\Projetos\Teste\Memória_de_Massa_-_Hiper_Monlevade(1).xlsx"

# Leitura das abas (1: Demanda, 2: Consumo)
df_demanda = pd.read_excel(arquivo, sheet_name=1, engine='openpyxl')
df_energia = pd.read_excel(arquivo, sheet_name=2, engine='openpyxl')

# === DEMANDA ===
# Conversões e filtros
df_demanda['Data'] = pd.to_datetime(df_demanda['Data'])
df_demanda = df_demanda[df_demanda['Posto'] == 'HP']

# Seleciona somente a demanda ativa consumida
df_demanda['Demanda (kW)'] = df_demanda['DemandaAtivaConsumokW']

# Agrupa por dia, pegando o maior valor de demanda HP por dia
demanda_diaria = df_demanda.groupby(df_demanda['Data'].dt.date)['Demanda (kW)'].max().reset_index()

# === ENERGIA ===
df_energia['Data'] = pd.to_datetime(df_energia['Data'])
df_energia = df_energia[df_energia['Posto'] == 'HP']

# Soma diária da energia consumida ativa
df_energia['Consumo (kWh)'] = df_energia['EnergiaAtivaConsumokWh']
consumo_diario = df_energia.groupby(df_energia['Data'].dt.date)['Consumo (kWh)'].sum().reset_index()

# === JUNÇÃO ===
df_merged = pd.merge(consumo_diario, demanda_diaria, on='Data', how='outer')
df_merged['Data'] = pd.to_datetime(df_merged['Data'])
df_merged = df_merged.sort_values(by='Data')

# === RESULTADOS ===
print("\nResumo do período:")
print(f"Total consumido: {df_merged['Consumo (kWh)'].sum():,.2f} kWh")
print(f"Demanda máxima registrada (HP): {df_merged['Demanda (kW)'].max():,.2f} kW")
print(f"Demanda média (HP): {df_merged['Demanda (kW)'].mean():,.2f} kW")

# === GRÁFICO ===
plt.figure(figsize=(12, 6))
plt.plot(df_merged['Data'], df_merged['Consumo (kWh)'], label='Consumo Diário (kWh)', color='blue')
plt.plot(df_merged['Data'], df_merged['Demanda (kW)'], label='Demanda HP Máxima (kW)', color='orange')
plt.title("Histórico Diário - Consumo e Demanda (HP)")
plt.xlabel("Data")
plt.ylabel("kWh / kW")
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.show()

# === BESS ESTIMATIVA ===
autonomia_horas = 3
demanda_media = df_merged['Demanda (kW)'].mean()
capacidade_bateria_kWh = demanda_media * autonomia_horas

print(f"\n🔋 Estimativa Inicial de BESS (Autonomia {autonomia_horas}h):")
print(f"Capacidade estimada: {capacidade_bateria_kWh:.2f} kWh")
