import pandas as pd

try:
    df = pd.read_excel("data/compras.xlsx")
    print("\n📝 Histórico de Compras:\n")
    print(df)
except FileNotFoundError:
    print("⚠️ Nenhuma planilha encontrada. Faça uma compra primeiro.")
