import yfinance as yf
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook

class ExcelUpdaterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Atualizador de Excel")

        # Layout
        self.label = tk.Label(root, text="Escolha o arquivo Excel:")
        self.label.pack(pady=10)

        self.select_button = tk.Button(root, text="Selecionar Arquivo", command=self.select_file)
        self.select_button.pack(pady=5)

        self.cell_label = tk.Label(root, text="Celula para atualizar (ex: C1):")
        self.cell_label.pack(pady=5)

        self.cell_entry = tk.Entry(root)
        self.cell_entry.pack(pady=5)

        self.update_button = tk.Button(root, text="Atualizar Preço", command=self.update_price, state=tk.DISABLED)
        self.update_button.pack(pady=5)

        self.file_path = ""

    def select_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.file_path:
            self.update_button.config(state=tk.NORMAL)

    def update_price(self):
        if not self.file_path:
            messagebox.showerror("Erro", "Nenhum arquivo selecionado.")
            return

        cell = self.cell_entry.get().strip()
        if not cell:
            messagebox.showerror("Erro", "Nenhuma célula especificada.")
            return

        SYMBOL = 'BBDC4.SA'  # Substitua pelo símbolo desejado

        try:
            # Baixa os dados
            stock = yf.Ticker(SYMBOL)
            data = stock.history(period='1d', interval='1m')
            closing_price = data['Close'].iloc[-1]

            # Carrega o arquivo Excel
            book = load_workbook(self.file_path)
            sheet = book.active

            # Atualiza a célula especificada com o preço de fechamento
            sheet[cell] = closing_price

            # Salva o arquivo
            book.save(self.file_path)

            messagebox.showinfo("Sucesso", f"Preço atualizado na célula {cell}: {closing_price}")

        except Exception as e:
            messagebox.showerror("Erro", str(e))