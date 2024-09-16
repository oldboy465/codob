import PyPDF2
import openpyxl
import tkinter as tk
from tkinter import filedialog, messagebox
import re

class PDFExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Philipão Desenrolado no PDF")
        self.root.configure(bg="black")  # Fundo preto
        self.pdf_files = []
        
        # Interface gráfica
        self.label = tk.Label(root, text="Philipão Desenrolado no PDF", fg="lime", font=("Courier", 16), bg="black")
        self.label.pack(pady=10)

        self.select_button = tk.Button(root, text="Selecionar PDFs", command=self.select_pdfs, fg="lime", bg="black")
        self.select_button.pack(pady=5)
        
        self.extract_button = tk.Button(root, text="Extrair Dados", command=self.extract_data, fg="lime", bg="black")
        self.extract_button.pack(pady=5)

        self.save_button = tk.Button(root, text="Salvar Dados", command=self.save_data, state=tk.DISABLED, fg="lime", bg="black")
        self.save_button.pack(pady=5)
        
        self.data = []

    def select_pdfs(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if files:
            self.pdf_files.extend(files)
            messagebox.showinfo("Informação", f"{len(files)} arquivos selecionados.")

    def extract_data(self):
        if not self.pdf_files:
            messagebox.showwarning("Aviso", "Nenhum arquivo PDF selecionado.")
            return
        
        for pdf_file in self.pdf_files:
            with open(pdf_file, "rb") as file:
                reader = PyPDF2.PdfReader(file)
                num_pages = len(reader.pages)
                
                for page_num in range(num_pages):
                    page = reader.pages[page_num]
                    text = page.extract_text().replace("\n", " ")
                    
                    # Extrair os dados específicos
                    data_referencia = self.extract_data_referencia(text)
                    numero = self.extract_numero(text)
                    favorecido = self.extract_favorecido(text)
                    observacao = self.extract_observacao(text)
                    valor_total = self.extract_valor_total(text)
                    
                    self.data.append([
                        data_referencia, "Descentralizada", numero, data_referencia, "Diversos",
                        "240201 Universidade Estadual do Maranhão", observacao, valor_total,
                        "Não", "Pagamento Consolidado", favorecido
                    ])
        
        messagebox.showinfo("Informação", "Dados extraídos com sucesso.")
        self.save_button.config(state=tk.NORMAL)

    def extract_data_referencia(self, text):
        pattern = r'\d{2}/\d{2}/\d{4}'  # padrão para capturar datas no formato dd/mm/yyyy
        match = re.search(pattern, text)
        if match:
            return match.group(0)
        return ""

    def extract_numero(self, text):
        pattern = r'2023OB\d{6}'  # padrão para capturar números que começam com '2023OB' e têm 12 dígitos
        match = re.search(pattern, text)
        if match:
            return match.group(0)
        return ""

    def extract_favorecido(self, text):
        pattern = r'\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b'  # padrão para capturar números de CNPJ no formato xx.xxx.xxx/xxxx-xx
        match = re.search(pattern, text)
        if match:
            return match.group(0)
        return ""

    def extract_observacao(self, text):
        pattern = r'Observação (.*?)(Domicílio Bancário)'  # padrão para capturar o texto entre "Observação" e "Domicílio Bancário"
        match = re.search(pattern, text, re.DOTALL)
        if match:
            return match.group(1).strip()
        return ""

    def extract_valor_total(self, text):
        pattern = r'Domicílio Bancário Origem\s*([\d,.]+)'  # padrão para capturar o valor total após "Domicílio Bancário Origem"
        match = re.search(pattern, text)
        if match:
            return match.group(1)
        return ""

    def save_data(self):
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not save_path:
            return
        
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Dados Extraídos"
        
        headers = [
            "Data Referência", "Tipo", "Número", "Data Lançamento", "Pagamento Tipo", "Unidade Gestora",
            "Observação", "Valor Total", "Repasse Recursos Federais", 
            "Pagamento Consolidado", "Favorecido"
        ]
        sheet.append(headers)
        
        for row in self.data:
            sheet.append(row)
        
        workbook.save(save_path)
        messagebox.showinfo("Informação", f"Dados salvos em {save_path}")
        self.pdf_files = []
        self.data = []
        self.save_button.config(state=tk.DISABLED)

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFExtractorApp(root)
    root.mainloop()
