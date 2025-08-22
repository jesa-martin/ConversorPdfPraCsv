import pdfplumber
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import logging
import re
import csv

# Classe auxiliar para redirecionar os logs para o widget de texto com cores
class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        logging.Handler.__init__(self)
        self.text_widget = text_widget
        # Define as "tags" de estilo para cada nível de log
        self.text_widget.tag_config("INFO", foreground="#cccccc") # Cinza claro para info
        self.text_widget.tag_config("WARNING", foreground="#ffc700") # Cor primária para avisos
        self.text_widget.tag_config("ERROR", foreground="#ff6b6b") # Vermelho para erros
        self.text_widget.tag_config("CRITICAL", foreground="#ff6b6b", underline=1)

    def emit(self, record):
        msg = self.format(record)
        level_tag = record.levelname
        def append():
            self.text_widget.configure(state='normal')
            self.text_widget.insert(tk.END, msg + '\n', level_tag)
            self.text_widget.configure(state='disabled')
            self.text_widget.yview(tk.END)
        self.text_widget.after(0, append)

class PDFtoCSVConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Conversor de Taxas PDF para CSV")
        self.root.geometry("700x700") # Janela um pouco maior para melhor espaçamento
        
        self.config_file = 'config.ini'
        self.logger = logging.getLogger(__name__)
        
        # Variáveis
        self.pdf_path = tk.StringVar()
        self.auto_save = tk.BooleanVar(value=True)
        self.csv_path = tk.StringVar()
        self.status = tk.StringVar(value="Pronto")
        self.debug_mode = tk.BooleanVar(value=True)
        self.last_dir = tk.StringVar()
        
        # Interface
        self.create_widgets()
        
        # Configura o handler de log para a caixa de texto
        text_handler = TextHandler(self.log_widget)
        log_format = logging.Formatter('%(asctime)s %(message)s', '%H:%M:%S')
        text_handler.setFormatter(log_format)
        self.logger.addHandler(text_handler)
        self.logger.setLevel(logging.INFO)

        self.load_config()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def create_widgets(self):
        # --- Seção de Estilização ---
        PRIMARY_COLOR = '#ffc700'
        SECONDARY_COLOR = '#1e1e1e' # Um preto mais suave
        TERTIARY_COLOR = '#ffffff'
        DISABLED_COLOR = '#3a3a3a'
        BASE_FONT = ("Segoe UI", 10)
        TITLE_FONT = ("Segoe UI", 11, "bold")
        
        style = ttk.Style(self.root)
        style.theme_use('clam') # Tema que permite mais customização

        # Configuração da janela principal
        self.root.configure(background=SECONDARY_COLOR)

        # Estilos Globais para widgets ttk
        style.configure('.', background=SECONDARY_COLOR, foreground=TERTIARY_COLOR, fieldbackground=SECONDARY_COLOR, borderwidth=0, focusthickness=0)
        style.map('.', foreground=[('disabled', '#6a6a6a')], fieldbackground=[('disabled', DISABLED_COLOR)])

        # Estilo do Botão Principal (Accent)
        style.configure('Accent.TButton', foreground='black', background=PRIMARY_COLOR, font=("Segoe UI", 11, "bold"), padding=(10, 8))
        style.map('Accent.TButton', background=[('active', '#e6b300')], relief=[('pressed', 'sunken')])

        # Estilo dos Botões Padrão
        style.configure('TButton', foreground=TERTIARY_COLOR, background='#333333', font=BASE_FONT, padding=5)
        style.map('TButton', background=[('active', '#4a4a4a')], relief=[('pressed', 'sunken')])

        # Estilo do Checkbutton
        style.configure('TCheckbutton', font=BASE_FONT)
        style.map('TCheckbutton', indicatorcolor=[('selected', PRIMARY_COLOR), ('!selected', '#555555')], background=[('active', SECONDARY_COLOR)])

        # Estilo da Caixa de Entrada (Entry)
        style.configure('TEntry', insertcolor=TERTIARY_COLOR, fieldbackground='#333333', borderwidth=2, relief='flat')
        style.map('TEntry', bordercolor=[('focus', PRIMARY_COLOR), ('!focus', '#333333')])
        
        # Estilo do Frame de Log (LabelFrame)
        style.configure('TLabelframe', background=SECONDARY_COLOR, bordercolor='#444444')
        style.configure('TLabelframe.Label', foreground=PRIMARY_COLOR, background=SECONDARY_COLOR, font=("Segoe UI", 10, "bold"))
        
        # --- Criação dos Widgets ---
        main_frame = ttk.Frame(self.root, padding="30 20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.grid_rowconfigure(6, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)

        # Seção de seleção do PDF
        ttk.Label(main_frame, text="Arquivo PDF:", font=TITLE_FONT).grid(row=0, column=0, padx=(0, 10), pady=10, sticky="e")
        ttk.Entry(main_frame, textvariable=self.pdf_path, width=50).grid(row=0, column=1, pady=10, sticky="ew")
        ttk.Button(main_frame, text="Procurar", command=self.select_pdf).grid(row=0, column=2, padx=(10, 0), pady=10)
        
        # Opção de salvamento automático
        ttk.Checkbutton(main_frame, text="Salvar automaticamente no mesmo local do PDF", variable=self.auto_save, command=self.toggle_auto_save).grid(row=1, column=1, columnspan=2, pady=10, sticky="w", padx=5)
        
        # Seção de seleção do CSV
        self.csv_label = ttk.Label(main_frame, text="Pasta de Saída:", font=TITLE_FONT)
        self.csv_entry = ttk.Entry(main_frame, textvariable=self.csv_path, width=50, state='disabled')
        self.csv_button = ttk.Button(main_frame, text="Procurar", command=self.select_output, state='disabled')
        self.csv_label.grid(row=2, column=0, padx=(0, 10), pady=10, sticky="e")
        self.csv_entry.grid(row=2, column=1, pady=10, sticky="ew")
        self.csv_button.grid(row=2, column=2, padx=(10, 0), pady=10)
        
        # Modo debug
        ttk.Checkbutton(main_frame, text="Modo debug (mostrar mais informações)", variable=self.debug_mode).grid(row=3, column=1, columnspan=2, pady=10, sticky="w", padx=5)
        
        # Botão de conversão
        ttk.Button(main_frame, text="Converter para CSV", command=self.convert, style='Accent.TButton').grid(row=4, column=1, pady=25)
        
        # Barra de status
        ttk.Label(main_frame, textvariable=self.status, font=("Segoe UI", 9, "italic")).grid(row=5, column=0, columnspan=3, sticky="w", pady=(10,0))
        
        # Área de log
        log_frame = ttk.LabelFrame(main_frame, text="Log de Depuração", padding=10)
        log_frame.grid(row=6, column=0, columnspan=3, sticky="nsew", pady=(10, 0))
        log_frame.grid_rowconfigure(0, weight=1)
        log_frame.grid_columnconfigure(0, weight=1)

        self.log_widget = scrolledtext.ScrolledText(log_frame, state='disabled', height=10, wrap=tk.WORD, background='#2a2a2a', foreground=TERTIARY_COLOR, bd=0, relief='flat', insertbackground=TERTIARY_COLOR, font=("Consolas", 9))
        self.log_widget.grid(row=0, column=0, sticky="nsew")
    
    def on_closing(self):
        self.save_config()
        self.root.destroy()

    def load_config(self):
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    last_dir = f.readline().strip()
                    if os.path.isdir(last_dir):
                        self.last_dir.set(last_dir)
                    else:
                        self.last_dir.set(os.path.expanduser("~"))
            else:
                self.last_dir.set(os.path.expanduser("~"))
        except Exception as e:
            self.logger.error(f"Erro ao carregar configuração: {e}")
            self.last_dir.set(os.path.expanduser("~"))
            
    def save_config(self):
        try:
            with open(self.config_file, 'w') as f:
                f.write(self.last_dir.get())
        except Exception as e:
            self.logger.error(f"Erro ao salvar configuração: {e}")

    def toggle_auto_save(self):
        if self.auto_save.get():
            self.csv_entry.config(state='disabled')
            self.csv_button.config(state='disabled')
            self.csv_path.set("")
        else:
            self.csv_entry.config(state='normal')
            self.csv_button.config(state='normal')
            if self.pdf_path.get():
                self.suggest_output_dir()
    
    def select_pdf(self):
        initial_dir = self.last_dir.get()
        filename = filedialog.askopenfilename(initialdir=initial_dir, filetypes=[("PDF Files", "*.pdf")], title="Selecione o arquivo PDF com as taxas")
        if filename:
            self.pdf_path.set(filename)
            new_dir = os.path.dirname(filename)
            self.last_dir.set(new_dir)
            self.save_config() 
            if not self.auto_save.get():
                self.suggest_output_dir()
    
    def suggest_output_dir(self):
        pdf_path = self.pdf_path.get()
        if pdf_path:
            output_dir = os.path.dirname(pdf_path)
            self.csv_path.set(output_dir)
    
    def select_output(self):
        initial_dir = self.csv_path.get() or self.last_dir.get()
        folder = filedialog.askdirectory(title="Selecione a pasta para salvar os arquivos CSV", initialdir=initial_dir)
        if folder:
            self.csv_path.set(folder)
    
    def extract_data(self, pdf_path):
        all_plans_data = {}
        bandeira_map = ["VISA", "Master Card", "Elo", "Hipercard", "American Express", "Outros", "Markup", "PIX"]
        
        with pdfplumber.open(pdf_path) as pdf:
            current_plan_name = None
            data_rows = []
            headers_with_boundaries = []

            for page_num, page in enumerate(pdf.pages):
                words = page.extract_words(x_tolerance=1, y_tolerance=3, keep_blank_chars=False)
                if not words: continue

                lines = {}
                for word in words:
                    y0 = round(word['top'] / 5.0) * 5.0 
                    if y0 not in lines: lines[y0] = []
                    lines[y0].append(word)

                sorted_lines = sorted(lines.items())
                
                for y, line_words in sorted_lines:
                    line_words.sort(key=lambda w: w['x0'])
                    line_text = ' '.join(w['text'] for w in line_words)
                    
                    if "PAYTIME" in line_text.upper().replace("Ν", "N"):
                        if current_plan_name and headers_with_boundaries and data_rows:
                            header_texts = [h['text'] for h in headers_with_boundaries]
                            all_plans_data[current_plan_name] = {'headers': header_texts, 'rows': data_rows}
                        
                        data_rows = []
                        headers_with_boundaries = []
                        
                        match = re.search(r'^(.*?PAYTIME.*?)(Débito.*)$', line_text, re.IGNORECASE)
                        if match:
                            current_plan_name = match.group(1).strip()
                            header_text_part = match.group(2)
                            header_words = [w for w in line_words if header_text_part.find(w['text']) != -1]
                        else:
                            current_plan_name = line_text
                            header_words = []

                        if self.debug_mode.get(): self.logger.info(f"Plano encontrado: {current_plan_name}")
                        
                        if header_words:
                            for i, word in enumerate(header_words):
                                left = word['x0']
                                right = header_words[i+1]['x0'] if i + 1 < len(header_words) else page.width
                                headers_with_boundaries.append({'text': word['text'], 'left': left, 'right': right})
                            if self.debug_mode.get(): self.logger.info(f"Taxas registradas")
                        continue
                    
                    if headers_with_boundaries:
                        bandeira_idx = len(data_rows)
                        if bandeira_idx < len(bandeira_map):
                            bandeira = bandeira_map[bandeira_idx]
                            row_data = {'Bandeira': bandeira}
                            
                            for header_info in headers_with_boundaries:
                                value_in_column = ""
                                for word in line_words:
                                    word_center = (word['x0'] + word['x1']) / 2
                                    if header_info['left'] <= word_center < header_info['right']:
                                        value_in_column += word['text'] + " "
                                row_data[header_info['text']] = value_in_column.strip() if value_in_column else "-"
                            
                            if any(val not in ["", "-"] for key, val in row_data.items() if key != 'Bandeira'):
                                data_rows.append(row_data)

            if current_plan_name and headers_with_boundaries and data_rows:
                header_texts = [h['text'] for h in headers_with_boundaries]
                all_plans_data[current_plan_name] = {'headers': header_texts, 'rows': data_rows}

        return all_plans_data
    
    def convert(self):
        try:
            pdf_path = self.pdf_path.get()
            if not pdf_path:
                messagebox.showerror("Erro", "Por favor, selecione um arquivo PDF")
                return
            
            output_dir = os.path.dirname(pdf_path) if self.auto_save.get() else self.csv_path.get()
            if not output_dir:
                messagebox.showerror("Erro", "Por favor, especifique uma pasta de saída")
                return
            
            self.log_widget.configure(state='normal')
            self.log_widget.delete(1.0, tk.END)
            self.log_widget.configure(state='disabled')

            self.status.set("Processando...")
            self.root.update()
            
            all_plans = self.extract_data(pdf_path)
            
            if not all_plans:
                self.status.set("Nenhum dado encontrado")
                messagebox.showwarning("Aviso", "Nenhum dado estruturado foi encontrado no PDF.")
                return

            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            csv_file = os.path.join(output_dir, f"{base_name}_unificado.csv")
            
            with open(csv_file, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f, delimiter=';', quoting=csv.QUOTE_ALL)
                
                for plan_name, data in all_plans.items():
                    rows_data = data.get('rows', [])
                    
                    if not rows_data:
                        self.logger.warning(f"Nenhum dado de linha encontrado para o plano '{plan_name}'. Ignorando.")
                        continue
                    
                    # --- Lógica para Unificar Colunas Iguais ---
                    consolidated_rows_data = []
                    brand_names_to_keep = []
                    
                    brand_dict = {row['Bandeira']: row for row in rows_data}
                    
                    # Verifica e consolida VISA e Master Card
                    visa_data = brand_dict.get('VISA')
                    master_data = brand_dict.get('Master Card')
                    
                    should_unify_visa_master = False
                    if visa_data and master_data:
                        visa_rates = [visa_data.get(h, '-') for h in data.get('headers', [])]
                        master_rates = [master_data.get(h, '-') for h in data.get('headers', [])]
                        if visa_rates == master_rates:
                            should_unify_visa_master = True

                    # Verifica e consolida Hipercard, American Express e Outros
                    hipercard_data = brand_dict.get('Hipercard')
                    amex_data = brand_dict.get('American Express')
                    outros_data = brand_dict.get('Outros')

                    should_unify_others = False
                    if hipercard_data and amex_data and outros_data:
                        others_match = True
                        for h in data.get('headers', []):
                            # Ignora a verificação da taxa de 'Débito'
                            if h == 'Débito':
                                continue
                            if hipercard_data.get(h) != amex_data.get(h) or amex_data.get(h) != outros_data.get(h):
                                others_match = False
                                break
                        if others_match:
                            should_unify_others = True

                    # Constrói a nova lista de dados e nomes de bandeira
                    processed_brands = set()
                    for row_dict in rows_data:
                        brand = row_dict.get('Bandeira', '-')
                        if brand in processed_brands:
                            continue
                        
                        if should_unify_visa_master and brand in ['VISA', 'Master Card']:
                            brand_names_to_keep.append('Visa/Master')
                            new_row_dict = {'Bandeira': 'Visa/Master'}
                            new_row_dict.update({h: visa_data.get(h, '-') for h in data.get('headers', [])})
                            consolidated_rows_data.append(new_row_dict)
                            processed_brands.add('VISA')
                            processed_brands.add('Master Card')
                        
                        elif should_unify_others and brand in ['Hipercard', 'American Express', 'Outros']:
                            brand_names_to_keep.append('Outros')
                            new_row_dict = {'Bandeira': 'Outros'}
                            for h in data.get('headers', []):
                                if h == 'Débito':
                                    new_row_dict[h] = outros_data.get(h, '-')
                                else:
                                    new_row_dict[h] = outros_data.get(h, '-')
                            consolidated_rows_data.append(new_row_dict)
                            processed_brands.add('Hipercard')
                            processed_brands.add('American Express')
                            processed_brands.add('Outros')
                        
                        else:
                            brand_names_to_keep.append(brand)
                            consolidated_rows_data.append(row_dict)
                            processed_brands.add(brand)

                    # Constrói o cabeçalho para a tabela do plano atual
                    header_row = [plan_name] + brand_names_to_keep
                    
                    # Constrói os dados das linhas transpostas a partir da lista consolidada
                    payment_options = data.get('headers', [])
                    transposed_rows = []
                    
                    for option in payment_options:
                        new_row = [option] # Começa a linha com a opção de pagamento
                        for row_dict in consolidated_rows_data:
                            value = row_dict.get(option, '-')
                            new_row.append(value)
                        transposed_rows.append(new_row)
                    
                    # Escreve o cabeçalho da tabela atual
                    writer.writerow(header_row)
                    
                    # Escreve as linhas de dados da tabela atual
                    writer.writerows(transposed_rows)
                    
                    # Adiciona duas linhas vazias como separador entre os planos
                    writer.writerow([])
                    writer.writerow([])

                self.status.set("Conversão concluída!")
                self.logger.info(f"Sucesso! Todos os planos foram salvos em: {csv_file}")
                messagebox.showinfo("Sucesso", f"A conversão foi concluída. Os dados unificados foram salvos em:\n{csv_file}")
            
        except Exception as e:
            self.status.set("Erro durante a conversão")
            self.logger.error(f"Erro: {str(e)}", exc_info=True)
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{str(e)}")
        finally:
            if self.status.get() != "Conversão concluída!":
                self.status.set("Pronto")

if __name__ == "__main__":
    root = tk.Tk()
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass
    app = PDFtoCSVConverter(root)
    root.mainloop()
