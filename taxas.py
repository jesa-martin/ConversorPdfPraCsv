import pdfplumber
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import logging
import re
import csv
import configparser

class PDFtoCSVConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Conversor de Taxas PDF para CSV")
        self.root.geometry("700x650")
        
        self.config_file = 'config.ini'
        
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.INFO)
        if self.logger.hasHandlers():
            self.logger.handlers.clear()
        
        file_handler = logging.FileHandler('conversor_log.txt', encoding='utf-8')
        log_format = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', '%Y-%m-%d %H:%M:%S')
        file_handler.setFormatter(log_format)
        self.logger.addHandler(file_handler)
        
        self.pdf_path = tk.StringVar()
        self.auto_save = tk.BooleanVar(value=True)
        self.csv_path = tk.StringVar()
        self.status = tk.StringVar(value="Pronto")
        self.debug_mode = tk.BooleanVar(value=True)
        self.last_dir = tk.StringVar()
        
        self.find_text = tk.StringVar()
        self.replace_text = tk.StringVar()
        self.plan_replacements = {}
        
        self.create_widgets()
        
        self.load_config()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        PRIMARY_COLOR = '#ffc700'
        SECONDARY_COLOR = '#1e1e1e'
        TERTIARY_COLOR = '#ffffff'
        DISABLED_COLOR = '#3a3a3a'
        BASE_FONT = ("Segoe UI", 10)
        TITLE_FONT = ("Segoe UI", 11, "bold")
        
        style = ttk.Style(self.root)
        style.theme_use('clam')

        self.root.configure(background=SECONDARY_COLOR)

        style.configure('.', background=SECONDARY_COLOR, foreground=TERTIARY_COLOR, fieldbackground=SECONDARY_COLOR, borderwidth=0, focusthickness=0)
        style.map('.', foreground=[('disabled', '#6a6a6a')], fieldbackground=[('disabled', DISABLED_COLOR)])

        style.configure('Accent.TButton', foreground='black', background=PRIMARY_COLOR, font=("Segoe UI", 11, "bold"), padding=(10, 8))
        style.map('Accent.TButton', background=[('active', '#e6b300')], relief=[('pressed', 'sunken')])

        style.configure('TButton', foreground=TERTIARY_COLOR, background='#333333', font=BASE_FONT, padding=5)
        style.map('TButton', background=[('active', '#4a4a4a')], relief=[('pressed', 'sunken')])

        style.configure('TCheckbutton', font=BASE_FONT)
        style.map('TCheckbutton', indicatorcolor=[('selected', PRIMARY_COLOR), ('!selected', '#555555')], background=[('active', SECONDARY_COLOR)])

        style.configure('TEntry', insertcolor=TERTIARY_COLOR, fieldbackground='#333333', borderwidth=2, relief='flat')
        style.map('TEntry', bordercolor=[('focus', PRIMARY_COLOR), ('!focus', '#333333')])
        
        style.configure('TLabelframe', background=SECONDARY_COLOR, bordercolor='#444444')
        style.configure('TLabelframe.Label', foreground=PRIMARY_COLOR, background=SECONDARY_COLOR, font=("Segoe UI", 10, "bold"))
        
        main_frame = ttk.Frame(self.root, padding="30 20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.grid_columnconfigure(0, weight=1)
        
        main_frame.grid_rowconfigure(5, weight=1)

        pdf_frame = ttk.LabelFrame(main_frame, text="Arquivo PDF", padding=10)
        pdf_frame.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        pdf_frame.grid_columnconfigure(0, weight=1)

        ttk.Entry(pdf_frame, textvariable=self.pdf_path, width=50).grid(row=0, column=0, pady=5, sticky="ew")
        ttk.Button(pdf_frame, text="Procurar", command=self.select_pdf).grid(row=0, column=1, padx=(10, 0), pady=5)
        ttk.Checkbutton(pdf_frame, text="Salvar automaticamente no mesmo local", variable=self.auto_save, command=self.toggle_auto_save).grid(row=1, column=0, columnspan=2, pady=5, sticky="w")
        
        output_frame = ttk.LabelFrame(main_frame, text="Pasta de Saída", padding=10)
        output_frame.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        output_frame.grid_columnconfigure(0, weight=1)
        
        self.csv_entry = ttk.Entry(output_frame, textvariable=self.csv_path, width=50, state='disabled')
        self.csv_entry.grid(row=0, column=0, pady=5, sticky="ew")
        self.csv_button = ttk.Button(output_frame, text="Procurar", command=self.select_output, state='disabled')
        self.csv_button.grid(row=0, column=1, padx=(10, 0), pady=5)
        
        replace_frame = ttk.LabelFrame(main_frame, text="Alteração de nome de Plano", padding=10)
        replace_frame.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(20, 0))
        replace_frame.grid_columnconfigure(0, weight=1)
        replace_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(replace_frame, text="Nome original:").grid(row=0, column=0, padx=(0, 5), sticky="w")
        self.find_entry = ttk.Entry(replace_frame, textvariable=self.find_text, width=40)
        self.find_entry.grid(row=0, column=1, padx=(5, 0), sticky="ew")

        ttk.Label(replace_frame, text="Substituir por:").grid(row=1, column=0, padx=(0, 5), sticky="w")
        self.replace_entry = ttk.Entry(replace_frame, textvariable=self.replace_text, width=40)
        self.replace_entry.grid(row=1, column=1, padx=(5, 0), sticky="ew")

        button_frame = ttk.Frame(replace_frame)
        button_frame.grid(row=0, column=2, rowspan=2, padx=(10, 0), pady=5, sticky="n")
        
        ttk.Button(button_frame, text="Adicionar Regra", command=self.add_replacement).pack(fill=tk.X, pady=2)
        ttk.Button(button_frame, text="Remover Selecionado", command=self.remove_replacement).pack(fill=tk.X, pady=2)
        
        self.replacement_listbox = tk.Listbox(replace_frame, height=4, background='#2a2a2a', foreground=TERTIARY_COLOR, bd=0, relief='flat', highlightbackground=SECONDARY_COLOR, selectbackground=PRIMARY_COLOR, selectforeground='black')
        self.replacement_listbox.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        
        ttk.Checkbutton(main_frame, text="Modo debug (mostrar mais informações)", variable=self.debug_mode).grid(row=3, column=0, columnspan=3, pady=10, sticky="w", padx=5)
        
        ttk.Button(main_frame, text="Converter para CSV", command=self.convert, style='Accent.TButton').grid(row=4, column=0, columnspan=3, pady=25)
        
        ttk.Label(main_frame, textvariable=self.status, font=("Segoe UI", 9, "italic")).grid(row=5, column=0, columnspan=3, sticky="w", pady=(10,0))
    
    def add_replacement(self):
        find = self.find_text.get().strip()
        replace = self.replace_text.get().strip()
        if find and replace:
            self.plan_replacements[find.lower()] = replace
            self.update_listbox()
            self.find_text.set("")
            self.replace_text.set("")
            self.save_config()
            self.logger.info(f"Regra de substituição adicionada: '{find.lower()}' -> '{replace}'")
        else:
            self.logger.warning("Campos de substituição não podem estar vazios.")

    def remove_replacement(self):
        selected_indices = self.replacement_listbox.curselection()
        if selected_indices:
            index = selected_indices[0]
            item_text = self.replacement_listbox.get(index)
            find_text = item_text.split("' -> '")[0].replace("'", "").strip()
            
            if find_text in self.plan_replacements:
                del self.plan_replacements[find_text]
                self.update_listbox()
                self.save_config()
                self.logger.info(f"Regra de substituição removida: '{find_text}'")

    def load_config(self):
        config = configparser.ConfigParser()
        if os.path.exists(self.config_file):
            config.read(self.config_file, encoding='utf-8')
            if 'SETTINGS' in config:
                if 'last_dir' in config['SETTINGS']:
                    self.last_dir.set(config['SETTINGS']['last_dir'])
                if 'auto_save' in config['SETTINGS']:
                    self.auto_save.set(config['SETTINGS'].getboolean('auto_save'))
                if 'debug_mode' in config['SETTINGS']:
                    self.debug_mode.set(config['SETTINGS'].getboolean('debug_mode'))
                if 'csv_path' in config['SETTINGS']:
                    self.csv_path.set(config['SETTINGS']['csv_path'])
            
            if 'REPLACEMENTS' in config:
                self.plan_replacements = dict(config['REPLACEMENTS'])
                self.update_listbox()
        else:
            self.logger.warning("Arquivo de configuração não encontrado, usando configurações padrão.")

    def save_config(self):
        config = configparser.ConfigParser()
        config['SETTINGS'] = {
            'last_dir': self.last_dir.get(),
            'auto_save': self.auto_save.get(),
            'debug_mode': self.debug_mode.get(),
            'csv_path': self.csv_path.get()
        }
        config['REPLACEMENTS'] = self.plan_replacements
        with open(self.config_file, 'w', encoding='utf-8') as configfile:
            config.write(configfile)
            self.logger.info("Configurações salvas.")

    def update_listbox(self):
        self.replacement_listbox.delete(0, tk.END)
        for find, replace in self.plan_replacements.items():
            self.replacement_listbox.insert(tk.END, f"'{find}' -> '{replace}'")

    def on_closing(self):
        self.save_config()
        self.root.destroy()
        
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
    
    def prompt_for_rule_application(self, find_name, replace_name, found_plans):
        prompt_win = tk.Toplevel(self.root)
        prompt_win.title("Regra Não Aplicada")
        prompt_win.geometry("500x350")
        prompt_win.resizable(False, False)
        
        x = self.root.winfo_x() + (self.root.winfo_width() / 2) - 250
        y = self.root.winfo_y() + (self.root.winfo_height() / 2) - 175
        prompt_win.geometry(f"+{int(x)}+{int(y)}")

        prompt_win.transient(self.root)
        prompt_win.grab_set()

        ttk.Label(prompt_win, text=f"A regra de substituição '{find_name}' -> '{replace_name}'\n não foi encontrada em nenhum plano.", font=("Segoe UI", 10)).pack(pady=10, padx=10)
        
        ttk.Label(prompt_win, text="Selecione um plano para aplicar esta regra:").pack(pady=(0, 5), padx=10)

        list_frame = ttk.Frame(prompt_win)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10)
        
        scrollbar_v = ttk.Scrollbar(list_frame, orient=tk.VERTICAL)
        scrollbar_h = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL)
        
        plan_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar_v.set, xscrollcommand=scrollbar_h.set, height=10, background='#2a2a2a', foreground='white', bd=0, relief='flat', highlightbackground='gray', selectbackground='#ffc700', selectforeground='black')
        
        scrollbar_v.config(command=plan_listbox.yview)
        scrollbar_h.config(command=plan_listbox.xview)
        
        plan_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_v.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_h.pack(side=tk.BOTTOM, fill=tk.X)

        for plan in found_plans:
            plan_listbox.insert(tk.END, plan)
        
        result_container = {"selected_plan": None}

        def on_ok():
            selected_index = plan_listbox.curselection()
            if selected_index:
                selected_plan = plan_listbox.get(selected_index[0])
                result_container["selected_plan"] = selected_plan
            prompt_win.destroy()
        
        def on_cancel():
            result_container["selected_plan"] = None
            prompt_win.destroy()
        
        def on_close():
            result_container["selected_plan"] = None
            prompt_win.destroy()
        
        button_frame = ttk.Frame(prompt_win)
        button_frame.pack(pady=10)

        ttk.Button(button_frame, text="Aplicar Seleção", command=on_ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Manter Original", command=on_cancel).pack(side=tk.RIGHT, padx=5)

        prompt_win.protocol("WM_DELETE_WINDOW", on_close)
        
        self.root.wait_window(prompt_win)
        
        return result_container["selected_plan"]
    
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
            
            self.status.set("Processando...")
            self.root.update()
            
            all_plans = self.extract_data(pdf_path)
            
            if not all_plans:
                self.status.set("Nenhum dado encontrado")
                self.logger.warning("Nenhum dado estruturado foi encontrado no PDF.")
                messagebox.showwarning("Aviso", "Nenhum dado estruturado foi encontrado no PDF.")
                return

            found_plans_names = list(all_plans.keys())
            
            for find_str, replace_str in list(self.plan_replacements.items()):
                # Verifica se a regra de substituição já foi aplicada ou se o plano existe no PDF
                if find_str.lower() not in [name.lower() for name in found_plans_names] and \
                   replace_str.lower() not in [name.lower() for name in found_plans_names]:
                    
                    self.logger.warning(f"A regra '{find_str}' -> '{replace_str}' não foi encontrada. Solicitando ação do usuário.")
                    
                    selected_plan_to_replace = self.prompt_for_rule_application(find_str, replace_str, found_plans_names)
                    
                    if selected_plan_to_replace is None:
                        # O usuário fechou a janela ou clicou em cancelar
                        self.status.set("Conversão cancelada.")
                        self.logger.info("Operação de conversão cancelada pelo usuário.")
                        return
                    
                    if selected_plan_to_replace != "":
                        # O usuário selecionou um novo plano.
                        original_data = all_plans[selected_plan_to_replace]
                        del all_plans[selected_plan_to_replace]
                        all_plans[replace_str] = original_data
                        
                        # Remove a regra antiga e adiciona a nova, salvando no config.
                        del self.plan_replacements[find_str] 
                        self.plan_replacements[selected_plan_to_replace] = replace_str
                        self.save_config()
                        self.update_listbox()
                        
                        self.logger.info(f"Regra '{find_str}' -> '{replace_str}' aplicada a '{selected_plan_to_replace}'.")
                        
                        found_plans_names = list(all_plans.keys())
                    else:
                        # O usuário escolheu "Manter Original", não faz nada com a regra, apenas continua.
                        self.logger.info(f"O nome original da regra '{find_str}' será mantido no resultado.")
                        # A regra original é removida, pois não foi aplicada.
                        del self.plan_replacements[find_str]
                        self.save_config()
                        self.update_listbox()
                        
            # O restante do código de conversão (que já tínhamos)
            # ...
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            csv_file = os.path.join(output_dir, f"{base_name}_unificado.csv")
            
            with open(csv_file, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f, delimiter=';', quoting=csv.QUOTE_ALL)
                
                for plan_name, data in all_plans.items():
                    rows_data = data.get('rows', [])
                    
                    if not rows_data:
                        self.logger.warning(f"Nenhum dado de linha encontrado para o plano '{plan_name}'. Ignorando.")
                        continue
                    
                    consolidated_rows_data = []
                    brand_names_to_keep = []
                    
                    brand_dict = {row['Bandeira']: row for row in rows_data}
                    
                    visa_data = brand_dict.get('VISA')
                    master_data = brand_dict.get('Master Card')
                    
                    should_unify_visa_master = False
                    if visa_data and master_data:
                        visa_rates = [visa_data.get(h, '-') for h in data.get('headers', [])]
                        master_rates = [master_data.get(h, '-') for h in data.get('headers', [])]
                        if visa_rates == master_rates:
                            should_unify_visa_master = True

                    hipercard_data = brand_dict.get('Hipercard')
                    amex_data = brand_dict.get('American Express')
                    outros_data = brand_dict.get('Outros')

                    should_unify_others = False
                    if hipercard_data and amex_data and outros_data:
                        others_match = True
                        for h in data.get('headers', []):
                            if h == 'Débito':
                                continue
                            if hipercard_data.get(h) != amex_data.get(h) or amex_data.get(h) != outros_data.get(h):
                                others_match = False
                                break
                        if others_match:
                            should_unify_others = True

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

                    header_row = [plan_name] + brand_names_to_keep
                    
                    payment_options = data.get('headers', [])
                    transposed_rows = []
                    
                    for option in payment_options:
                        new_row = [option]
                        for row_dict in consolidated_rows_data:
                            value = row_dict.get(option, '-')
                            new_row.append(value)
                        transposed_rows.append(new_row)
                    
                    writer.writerow(header_row)
                    writer.writerows(transposed_rows)
                    writer.writerow([])
                    writer.writerow([])

            self.status.set("Conversão concluída!")
            self.logger.info(f"Sucesso! Todos os planos foram salvos em: {csv_file}")
            
        except Exception as e:
            self.status.set("Erro durante a conversão")
            self.logger.error(f"Erro: {str(e)}", exc_info=True)
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{str(e)}")
        finally:
            if self.status.get() != "Conversão concluída!":
                self.status.set("Pronto")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFtoCSVConverter(root)
    root.mainloop()
