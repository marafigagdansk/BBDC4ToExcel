import yfinance as yf
from openpyxl import Workbook, load_workbook

SYMBOL = 'BBDC4.SA'  
file_path = 'C:\\Users\\Pc\\Documents\\Teste\\media ponderada B3.xlsx'

# Baixa os dados
stock = yf.Ticker(SYMBOL)
data = stock.history(period='1d', interval='1m')

# Obtém o preço de fechamento mais recente
closing_price = data['Close'].iloc[-1]

# Carrega o arquivo Excel existente
try:
    book = load_workbook(file_path)
    sheet = book.active
except FileNotFoundError:
    print("Arquivo Excel não encontrado. Por favor, verifique o caminho e crie o arquivo primeiro.")
    exit()

# Adiciona o preço de fechamento na célula C1
sheet['C1'] = closing_price

# Salva o arquivo
book.save(file_path)

# Exibe o preço de fechamento mais recente no terminal
print(f'Último preço de fechamento de {SYMBOL}: {closing_price}')