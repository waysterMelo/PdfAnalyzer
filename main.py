import hashlib
import platform
import hmac
import io
import json
import os
import queue
import re
import shutil
import subprocess
import tempfile
import threading
import tkinter as tk
from tkinter import ttk, filedialog
from openpyxl.worksheet.table import Table, TableStyleInfo
import cv2
import fitz
import numpy as np
import pandas as pd
from PIL import Image, ImageTk, ImageFilter, ImageEnhance
from datetime import datetime, timedelta
import sys
import pytesseract
from tkinter import messagebox
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook
from pytesseract import pytesseract
from ttkthemes.themed_tk import ThemedTk

SECRET_KEY = b"waystermelo@"
LICENSE_FILE = "license.txt"


image_path = os.path.join(os.path.dirname(__file__), 'img', 'logo.webp')

def create_signature(data):
    """Cria uma assinatura HMAC-SHA256 para os dados fornecidos."""
    return hmac.new(SECRET_KEY, data.encode('utf-8'), hashlib.sha256).hexdigest()

def verify_signature(data, signature):
    """Verifica se a assinatura fornecida corresponde ao conteúdo dos dados."""
    return hmac.compare_digest(create_signature(data), signature)

def check_license():
    """Verifica se a licença de teste ainda é válida e inicializa a data de instalação se necessário."""
    global activation_time

    try:
        # Verificar se o arquivo de licença existe
        if os.path.exists(LICENSE_FILE):
            with open(LICENSE_FILE, "r") as file:
                content = json.load(file)
                data_str = f"{content['activation_time']}|{content['duracao']}"
                saved_signature = content.get("signature")
                duracao_em_minutos = content.get("duracao")  # Carregar a duração diretamente do arquivo de licença

                # Verificar se a duração foi corretamente especificada
                if not isinstance(duracao_em_minutos, int) or duracao_em_minutos <= 0:
                    raise ValueError("Duração da licença inválida ou ausente no arquivo de licença.")

                # Verificar a integridade do arquivo com a assinatura
                if not verify_signature(data_str, saved_signature):
                    messagebox.showerror("Erro de Licença", "A licença foi manipulada! O programa será encerrado.")
                    return False

                # Carregar a data de ativação a partir do arquivo
                activation_time = datetime.fromisoformat(content["activation_time"])
        else:
            messagebox.showerror("Erro de Licença", "Arquivo de licença não encontrado! O programa será encerrado.")
            return False

        # Verificar o tempo de expiração
        expiration_time = activation_time + timedelta(minutes=duracao_em_minutos)
        if datetime.now() > expiration_time:
            # Se o período de teste tiver expirado
            messagebox.showwarning("Licença Expirada", "Seu período de teste expirou. O programa será encerrado.")
            return False
        else:
            remaining_time = expiration_time - datetime.now()
            remaining_minutes = int(remaining_time.total_seconds() // 60)
            messagebox.showinfo("Licença de Teste", f"Você tem {remaining_minutes} minutos restantes de teste.")
            return True

    except Exception as e:
        messagebox.showerror("Erro de Licença", f"Ocorreu um erro ao verificar a licença: {str(e)}")
        return False

def configurar_tesseract():
    """Configuração do Tesseract OCR."""
    tessdata_prefix = r'C:/Program Files/Tesseract-OCR/tessdata/'
    tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'
    tesseract_config = TesseractConfig(tessdata_prefix, tesseract_cmd)
    tesseract_config.test_setup()

def iniciar_interface_principal():
    """Inicia a interface principal da aplicação."""
    global root, bg_image
    if not check_license():
        # Se a licença estiver expirada ou com erro, o programa é encerrado
        return

    root = tk.Tk()
    root.title("PDF Analyzer Blank")
    root.configure(bg="#1C2833")
    root.resizable(True, True)  # Permitir redimensionamento da janela

    # Centralizando a janela na tela
    window_width, window_height = 900, 600
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    position_top = int((screen_height - window_height) / 2)
    position_left = int((screen_width - window_width) / 2)
    root.geometry(f"{window_width}x{window_height}+{position_left}+{position_top}")

    # Carregando a imagem de fundo
    try:
        image = Image.open(image_path)
        image = image.resize((window_width, window_height))  # Redimensiona a imagem para o tamanho da janela
        bg_image = ImageTk.PhotoImage(image)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar a imagem de fundo: {e}")
        return

    # Rótulo para a imagem de fundo
    background_label = tk.Label(root, image=bg_image)
    background_label.place(relwidth=1, relheight=1)

    # Nome da aplicação no centro
    app_name = tk.Label(root, text="PDF Analyzer Blank", font=("Helvetica", 40, "bold"),
                        fg="#FFFFFF", bg="#1E3D59", padx=20, pady=10, relief="raised", bd=10)
    app_name.place(relx=0.5, rely=0.25, anchor='center')

    # Função para iniciar a análise e carregar GUI
    def iniciar_analise():
        root.destroy()
        configurar_tesseract()
        PDFAnalyzerGUI()

    # Botão para iniciar a análise
    start_button = tk.Button(root, text="Iniciar Análise", font=("Helvetica", 18, "bold"),
                             fg="#FFFFFF", bg="#1E3D59", activebackground="#34495E",
                             activeforeground="#FFFFFF", padx=20, pady=10, relief="raised",
                             bd=5, command=iniciar_analise)
    start_button.place(relx=0.5, rely=0.55, anchor='center')

    # Direitos autorais no final da tela
    copyright_label = tk.Label(root, text="Direitos Autorais © Arquindex.",
                               font=("Helvetica", 12, "bold"), fg="#FFFFFF", bg="#1C2833", padx=5, pady=5)
    copyright_label.place(relx=0.01, rely=0.95, anchor='w')

    # Iniciar o loop da interface gráfica
    root.mainloop()

class CircularProgressBar(ttk.Frame):
    def __init__(self, parent, size=100, thickness=10, max_value=100, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.size = size
        self.thickness = thickness
        self.max_value = max_value
        self.value = 0

        # Configura o Canvas para desenhar o círculo
        self.canvas = tk.Canvas(self, width=size, height=size, bg="black", highlightthickness=0)
        self.canvas.pack()

        self.text = self.canvas.create_text(size/2, size/2, text="0%", font=("Helvetica", 24), fill="white")
        self.arc = self.canvas.create_arc(
            self.thickness, self.thickness,
            self.size - self.thickness, self.size - self.thickness,
            start=90, extent=0, outline="#00FF00", width=self.thickness, style="arc"
        )

    def set_value(self, value):
        """Define o valor atual da barra de progresso"""
        self.value = min(self.max_value, max(0, value))
        extent = (self.value / self.max_value) * 360
        self.canvas.itemconfig(self.arc, extent=-extent)  # Extent negativo para sentido horário
        self.canvas.itemconfig(self.text, text=f"{int((self.value / self.max_value) * 100)}%")

class PDFAnalyzerGUI:
    def __init__(self):
        # Inicializa os componentes da interface gráfica e outros atributos
        self.canvas = None
        self.open_folder_button = None
        self.open_analise_button = None
        self.pages_blank_after_ocr_label = None
        self.pages_total_checked_label = None
        self.circular_progress = None
        self.progress_label = None
        self.progress_var = None
        self.analyze_button = None
        self.select_label = None
        self.pages_low_info_label = None
        self.window = ThemedTk(theme="arc")
        self.window.title("Analisador de PDFs - Digitalizados")
        self.window.state("zoomed")

        # Diretório selecionado e instâncias de classes auxiliares
        self.directory = None
        self.analyzer = PDFAnalyzer()
        self.report_generator = ReportGenerator()

        # Fila para gerenciar progresso de processamento
        self.progress_queue = queue.Queue()

        # Configurações de estilo e criação dos componentes da interface
        self.setup_style()
        self.create_widgets()

        # Inicia o processamento da fila de progresso
        self.process_queue()

        # Inicia a interface gráfica
        self.window.mainloop()

    def reset_for_new_analysis(self):
        """Reinicia a interface para uma nova análise."""
        # Limpa o diretório selecionado e redefine os componentes da interface
        self.directory = None
        self.reset_labels()
        self.select_label.config(text="Nenhum diretório selecionado")
        self.analyze_button.config(state="disabled")
        self.canvas.delete("all")
        self.circular_progress.set_value(0)
        self.progress_queue.queue.clear()  # Limpa a fila de progresso
        self.open_analise_button.config(state="disabled")  # Desabilita o botão

    def setup_style(self):
        # Configura o estilo dos componentes da interface gráfica
        style = ttk.Style(self.window)
        style.configure("TFrame", background="#0D1B2A")  # Azul mais escuro para melhor contraste
        style.configure("TLabel", background="#121212", foreground="white", font=("Segoe UI", 10))
        style.configure("Header.TLabel", foreground="#F0F0F0", font=("Segoe UI", 10, "bold"))
        style.configure("TButton", background="#374956", foreground="black", font=("Segoe UI", 10, "bold"), padding=10)
        style.map("TButton", background=[("active", "#516B7F")], relief=[("pressed", "ridge"), ("!pressed", "flat")])

    def create_widgets(self):
        # Criação dos dois frames para dividir a interface em colunas
        main_frame_left = ttk.Frame(self.window, padding="20")
        main_frame_left.pack(side='left', fill='y', expand=False)

        main_frame_right = ttk.Frame(self.window, padding="20")
        main_frame_right.pack(side='right', fill='both', expand=True)

        # Componentes da coluna esquerda (controles e progresso)
        header_label = ttk.Label(main_frame_left, text="Selecione um diretório para análise de PDFs:",
                                 style="Header.TLabel")
        header_label.pack(pady=(0, 20))

        select_button = ttk.Button(main_frame_left, text="Selecionar Diretório", command=self.select_directory,
                                   width=30)
        select_button.pack(pady=10)

        self.select_label = ttk.Label(main_frame_left, text="Nenhum diretório selecionado", foreground="#a6a6a6")
        self.select_label.pack(pady=(5, 15))

        self.analyze_button = ttk.Button(main_frame_left, text="Iniciar Análise", state="disabled",
                                         command=self.start_analysis, width=30)
        self.analyze_button.pack(pady=10)

        self.circular_progress = CircularProgressBar(main_frame_left, size=200, thickness=20, max_value=100)
        self.circular_progress.pack(pady=20)

        self.pages_blank_after_ocr_label = ttk.Label(main_frame_left, text="Página em Branco / Pouca Informação: 0")
        self.pages_blank_after_ocr_label.pack(pady=5)
        self.pages_low_info_label = ttk.Label(main_frame_left, text="Necessidade de verificar: 0")
        self.pages_low_info_label.pack(pady=5)
        self.pages_total_checked_label = ttk.Label(main_frame_left, text="Total de Páginas Verificadas: 0")
        self.pages_total_checked_label.pack(pady=5)

        # Frame para organizar os botões na mesma linha
        button_frame = ttk.Frame(main_frame_left)
        button_frame.pack(pady=15)

        self.open_folder_button = ttk.Button(button_frame, text="Abrir Pasta do Relatório", command=self.open_folder,
                                             width=25)
        self.open_folder_button.grid(row=0, column=0, padx=5)

        self.open_analise_button = ttk.Button(button_frame, text="Abrir Análises", command=self.open_analysis_screen,
                                              width=25, state="disabled")
        self.open_analise_button.grid(row=0, column=1, padx=5)

        # Adicione o botão "Iniciar Nova Análise" ao lado dos outros
        self.restart_button = ttk.Button(button_frame, text="Iniciar Nova Análise", command=self.reset_for_new_analysis,
                                         width=25)
        self.restart_button.grid(row=0, column=2, padx=5)

        # Frame do canvas para exibir imagens na coluna direita
        canvas_frame = ttk.Frame(main_frame_right, padding="10", borderwidth=2, relief="ridge", style="TFrame")
        canvas_frame.pack(expand=True, fill="both")

        self.canvas = tk.Canvas(canvas_frame, bg="white")
        self.canvas.pack(expand=True, fill="both")

    def select_directory(self):
        # Abre um diálogo para selecionar o diretório contendo os arquivos PDF
        self.directory = filedialog.askdirectory()
        if self.directory:
            self.select_label.config(text=f"Diretório Selecionado: {self.directory}")
            self.analyze_button.config(state="normal")  # Habilita o botão de análise

    def start_analysis(self):
        # Inicia a análise dos PDFs no diretório selecionado em uma nova thread
        if self.directory:
            self.analyze_button.config(state="disabled")  # Desabilita o botão durante a análise
            threading.Thread(target=self.run_analysis_thread, daemon=True).start()

    def run_analysis_thread(self):
        # Executa a análise dos PDFs e gera um relatório
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_xlsx = os.path.join(self.directory, f"analysis_report_{timestamp}.xlsx")
        self.analyze_pdfs_in_directory(output_xlsx)

    def analyze_pdfs_in_directory(self, output_xlsx):
        # Analisa todos os PDFs no diretório selecionado e gera um relatório
        self.reset_labels();
        pdf_files = [os.path.join(self.directory, f) for f in os.listdir(self.directory) if f.lower().endswith('.pdf')]
        total_pages = sum(fitz.open(pdf_file).page_count for pdf_file in pdf_files)  # Calcula o total de páginas
        total_pages_processed = 0

        # Itera sobre cada arquivo PDF no diretório
        for pdf_file in pdf_files:
            pdf_name = os.path.basename(pdf_file)
            with fitz.open(pdf_file) as pdf_document:
                for page_num in range(pdf_document.page_count):
                    # Carrega a página e converte para imagem
                    page = pdf_document.load_page(page_num)
                    pix = page.get_pixmap()
                    img = Image.open(io.BytesIO(pix.tobytes("png")))

                    # Realiza análise na página
                    status, white_pixel_percentage, ocr_performed, extracted_text = self.analyzer.analyze_page(img)
                    # Adiciona os resultados ao gerador de relatórios
                    self.report_generator.add_record(pdf_name, page_num + 1, status, white_pixel_percentage,
                                                     ocr_performed, extracted_text)

                    # Atualizar labels e progresso
                    total_pages_processed += 1
                    self.update_labels(total_pages_processed)
                    progress_percentage = (total_pages_processed / total_pages) * 100
                    self.progress_queue.put(progress_percentage)

                    # Enviar a imagem para ser exibida no canvas
                    self.progress_queue.put(("image", img))

        # Finaliza o relatório após processar todas as páginas
        self.report_generator.finalize(output_xlsx)
        if not os.path.exists(output_xlsx):
            print("Erro: O relatório não foi criado.")
            return
        self.progress_queue.put("DONE")  # Indica que a análise foi concluída

    def process_queue(self):
        # Processa os itens da fila para atualizar a interface em tempo real
        try:
            while True:
                message = self.progress_queue.get_nowait()
                if isinstance(message, tuple) and message[0] == "image":
                    image = message[1]
                    self.display_image_on_canvas(image)
                elif isinstance(message, float) or isinstance(message, int):
                    self.update_progress(message)
                elif message == "DONE":
                    self.analyze_button.config(state="normal")
                    # Verifica se há páginas que necessitam de revisão
                    if self.analyzer.pages_blank_after_ocr_count == 0 and self.analyzer.pages_low_info_count == 0:
                        # Nenhuma página para revisar, desabilita o botão
                        self.open_analise_button.config(state="disabled")
                    else:
                        # Há páginas para revisar, habilita o botão
                        self.open_analise_button.config(state="normal")
                    messagebox.showinfo("Análise Concluída",
                                        "A análise foi concluída e o relatório foi gerado com sucesso!")
                    break  # Sai do loop ao concluir a análise
        except queue.Empty:
            pass
        # Verifica a fila novamente após 100 ms
        self.window.after(100, self.process_queue)

    def update_progress(self, value):
        """Atualiza o progresso circular"""
        self.circular_progress.set_value(value)

    def update_labels(self, total_pages_checked):
        # Atualiza as labels que exibem informações sobre a análise
        self.pages_blank_after_ocr_label.config(
            text=f"Página em Branco / Pouca Informação: {self.analyzer.pages_blank_after_ocr_count}")
        self.pages_total_checked_label.config(
            text=f"Total de Páginas Verificadas: {total_pages_checked}")
        self.pages_low_info_label.config(
            text=f"Necessário revisar: {self.analyzer.pages_low_info_count}")

    def display_image_on_canvas(self, image):
        print("Exibindo imagem no canvas...")

        # Obtenha dimensões do canvas e da imagem
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        img_width, img_height = image.size

        # Define dimensões padrão se o canvas não estiver inicializado corretamente
        if canvas_width == 1 or canvas_height == 1:
            canvas_width, canvas_height = 400, 600
            self.canvas.config(width=canvas_width, height=canvas_height)

        # Define o filtro de reamostragem adequado
        try:
            resample_filter = Image.Resampling.LANCZOS
        except AttributeError:
            resample_filter = Image.LANCZOS

        # Calcula a nova escala da imagem para caber no canvas
        ratio = min(canvas_width / img_width, canvas_height / img_height)
        new_size = (int(img_width * ratio), int(img_height * ratio))
        image = image.resize(new_size, resample=resample_filter)

        # Calcula a posição para centralizar a imagem no canvas
        x = (canvas_width - new_size[0]) // 2
        y = (canvas_height - new_size[1]) // 2
        print(f"Posicionando imagem no canvas: ({x}, {y})")

        # Exibe a imagem no canvas
        tk_image = ImageTk.PhotoImage(image)
        self.canvas.delete("all")
        self.canvas.create_image(x, y, anchor="nw", image=tk_image)
        self.canvas.image = tk_image  # Evita que o Python faça coleta de lixo da imagem

    def open_folder(self):
        # Abre o diretório contendo os PDFs ou relatórios
        if self.directory:
            try:
                # Use platform.system() para verificar o sistema operacional corretamente
                current_system = platform.system()
                if current_system == "Windows":
                    os.startfile(self.directory)
                elif current_system == "Darwin":
                    subprocess.Popen(["open", self.directory])
                elif current_system == "Linux":
                    subprocess.Popen(["xdg-open", self.directory])
                else:
                    raise EnvironmentError("Sistema operacional não suportado")
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível abrir a pasta: {str(e)}")

    def open_analysis_screen(self):
        # Verifica se um diretório foi selecionado
        if not self.directory:
            messagebox.showerror("Erro", "Nenhum diretório selecionado.")
            return

        # Encontra o arquivo de relatório mais recente com o padrão analysis_report_YYYYMMDD_HHMMSS.xlsx
        report_files = [f for f in os.listdir(self.directory) if
                        f.startswith("analysis_report_") and f.endswith(".xlsx")]

        if not report_files:
            messagebox.showerror("Erro", "Nenhum relatório de análise encontrado.")
            return

        # Ordena os relatórios por data de criação em ordem decrescente para obter o mais recente
        report_files.sort(
            key=lambda x: datetime.strptime(x.replace("analysis_report_", "").replace(".xlsx", ""), "%Y%m%d_%H%M%S"),
            reverse=True)
        analysis_report_path = os.path.join(self.directory, report_files[0])

        # Minimiza a janela principal antes de abrir a nova tela
        self.window.iconify()

        # Cria a tela de análises pendentes
        AnalysisScreen(self.window, analysis_report_path)

    def reset_labels(self):
        """Zera as contagens de labels para uma nova análise."""
        self.pages_blank_after_ocr_label.config(text="Página em Branco / Pouca Informação: 0")
        self.pages_total_checked_label.config(text="Total de Páginas Verificadas: 0")
        self.pages_low_info_label.config(text="Necessidade de verificar: 0")

        # Resetando contadores da instância do PDFAnalyzer
        self.analyzer.pages_blank_after_ocr_count = 0
        self.analyzer.pages_low_info_count = 0
        self.analyzer.pages_blank_count = 0

class AnalysisScreen:
    def __init__(self, master, analysis_report_path):
        print("Inicializando AnalysisScreen...")
        self.window = tk.Toplevel(master)
        self.window.title("Análises Pendentes")
        self.window.state('zoomed')
        self.window.configure(bg="#2B3E50")

        # Configurações de estilo
        style = ttk.Style(self.window)
        style.configure("TFrame", background="#2B3E50")
        style.configure("TLabel", background="#2B3E50", foreground="white", font=("Segoe UI", 10))
        style.configure("Header.TLabel", foreground="#ECECEC", font=("Segoe UI", 12, "bold"))

        self.analysis_report_path = analysis_report_path
        self.selected_directory = os.path.dirname(analysis_report_path)

        # Configuração da interface
        self.header_frame = ttk.Frame(self.window, style="TFrame")
        self.header_frame.pack(pady=10, padx=20, fill='x')

        self.report_label = ttk.Label(
            self.header_frame,
            text=f"Relatório: {os.path.basename(self.analysis_report_path)}",
            style="Header.TLabel"
        )
        self.report_label.pack(side='left', padx=5)

        self.date_label = ttk.Label(
            self.header_frame,
            text=f"Data: {pd.to_datetime('today').strftime('%d/%m/%Y')}",
            style="Header.TLabel"
        )
        self.date_label.pack(side='right', padx=5)

        # Treeview para exibir arquivos pendentes
        columns = ("Arquivo PDF", "Página", "Status")
        self.pending_files_tree = ttk.Treeview(
            self.window, columns=columns, show="headings", height=25
        )
        self.pending_files_tree.pack(pady=20, padx=10, fill='y', side='left')

        self.pending_files_tree.heading("Arquivo PDF", text="Arquivo PDF")
        self.pending_files_tree.heading("Página", text="Página")
        self.pending_files_tree.heading("Status", text="Status")
        self.pending_files_tree.column("Arquivo PDF", width=200)
        self.pending_files_tree.column("Página", width=50, anchor='center')
        self.pending_files_tree.column("Status", width=190, anchor='center')

        # Frame para visualização de PDF
        self.pdf_view_frame = ttk.Frame(
            self.window, padding="10", borderwidth=2, relief="ridge", style="TFrame"
        )
        self.pdf_view_frame.pack(pady=10, padx=10, side='right', fill='both', expand=True)

        self.pdf_canvas = tk.Canvas(self.pdf_view_frame, bg="#2B3E50")
        self.pdf_canvas.pack(expand=True, fill='both')
        self.pdf_canvas.bind('<Configure>', self.center_image)

        self.window.bind("<Delete>", self.on_delete_key_press)

        # Botão de deletar página
        delete_button = tk.Button(
            self.pdf_view_frame,
            text="Deletar Página Selecionada",
            command=self.delete_selected_pdf,
            bg="#f44336",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            activebackground="#e53935",
            padx=10,
            pady=5
        )
        delete_button.pack(side='bottom', pady=5)

        # Botão para abrir diretório
        open_button = tk.Button(
            self.window,
            text="Abrir Diretório dos PDFs",
            command=self.open_pdf_directory
        )
        open_button.pack(pady=10)

        # Label de loading
        self.loading_label = tk.Label(
            self.window,
            text="",
            bg="#2B3E50",
            fg="white",
            font=("Segoe UI", 10, "bold")
        )
        self.loading_label.pack(pady=5)

        self.load_pending_files()
        self.pending_files_tree.bind('<<TreeviewSelect>>', self.on_pdf_select)

    def on_delete_key_press(self, event):
        self.delete_selected_pdf()

    def load_pending_files(self):
        try:
            df = pd.read_excel(self.analysis_report_path)
            pending_pages = df[(df['Status'] != 'OK') & (df['Status'] != 'Identificado conteúdo após reanálise')][['Arquivo PDF', 'Página', 'Status']].drop_duplicates()
            for _, row in pending_pages.iterrows():
                pdf_name = row['Arquivo PDF']
                page_number = row['Página']
                status = row['Status']
                self.pending_files_tree.insert("", 'end', values=(pdf_name, page_number, status))
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar o relatório: {str(e)}")

    def on_pdf_select(self, event):
        selected_item = self.pending_files_tree.selection()

        if not selected_item:
            return

        try:
            selected_entry = self.pending_files_tree.item(selected_item, "values")
            if not selected_entry:
                raise ValueError("O item selecionado não está mais disponível.")

            pdf_name, page_info, status = selected_entry
            self.selected_pdf = os.path.join(self.selected_directory, pdf_name)
            page_number = int(page_info)  # Número da página (1-based)
            self.selected_page_index = page_number - 1  # Índice da página (0-based)
            print(
                f"PDF selecionado: {self.selected_pdf}, Página: {page_number}, Índice real: {self.selected_page_index}, Status: {status}")

            if os.path.exists(self.selected_pdf):
                try:
                    self.render_pdf_page(self.selected_pdf, page_number)
                except Exception as e:
                    print(f"Erro ao abrir o PDF: {str(e)}")
                    messagebox.showerror("Erro", f"Não foi possível abrir o PDF: {str(e)}")
            else:
                print("Erro: Arquivo PDF não encontrado!")
                messagebox.showerror("Erro", "Arquivo PDF não encontrado!")
        except (tk.TclError, ValueError) as e:
            print(f"Erro ao acessar o item: {str(e)}")

    def open_pdf_directory(self):
        if os.path.exists(self.selected_directory):
            os.startfile(self.selected_directory)
        else:
            messagebox.showerror("Erro", "Diretório dos PDFs não encontrado!")

    def delete_selected_pdf(self):
        try:
            selected_items = self.pending_files_tree.selection()
            if not selected_items:
                messagebox.showwarning("Aviso", "Nenhum PDF selecionado!")
                return

            pages_to_delete = {}
            for item in selected_items:
                pdf_name, page_info, status = self.pending_files_tree.item(item, "values")
                page_index = int(page_info) - 1  # Índice zero-based

                if pdf_name not in pages_to_delete:
                    pages_to_delete[pdf_name] = set()
                pages_to_delete[pdf_name].add(page_index)

            for pdf_name, page_indices in pages_to_delete.items():
                pdf_path = os.path.join(self.selected_directory, pdf_name)
                if os.path.exists(pdf_path):
                    # Abrir o PDF
                    pdf_document = fitz.open(pdf_path)

                    # Deletar páginas em ordem reversa
                    for page_index in sorted(page_indices, reverse=True):
                        if 0 <= page_index < pdf_document.page_count:
                            pdf_document.delete_page(page_index)

                    # Salvar o PDF modificado em um arquivo temporário
                    temp_pdf_path = pdf_path + ".tmp"
                    pdf_document.save(temp_pdf_path)
                    pdf_document.close()

                    # Substituir o PDF original pelo modificado
                    os.remove(pdf_path)
                    os.rename(temp_pdf_path, pdf_path)

                    # Atualizar o relatório e a Treeview
                    self.update_report_and_treeview(pdf_name, page_indices)
                    self.clear_canvas()
                else:
                    messagebox.showerror("Erro", f"O arquivo PDF {pdf_name} não foi encontrado.")

            messagebox.showinfo("Sucesso", "As páginas selecionadas foram deletadas com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao deletar as páginas: {str(e)}")

    def update_report_and_treeview(self, pdf_name, deleted_page_indices):
        report_path = self.analysis_report_path
        df = pd.read_excel(report_path)

        # Obter os números das páginas deletadas (1-based)
        deleted_page_numbers = [p + 1 for p in deleted_page_indices]

        # Remover as páginas deletadas do DataFrame
        df = df[~((df['Arquivo PDF'] == pdf_name) & (df['Página'].isin(deleted_page_numbers)))]

        # Ajustar os números das páginas remanescentes
        for deleted_page in sorted(deleted_page_numbers):
            df.loc[(df['Arquivo PDF'] == pdf_name) & (df['Página'] > deleted_page), 'Página'] -= 1

        # Salvar o DataFrame atualizado
        df.to_excel(report_path, index=False)

        # Recarregar a Treeview
        self.pending_files_tree.delete(*self.pending_files_tree.get_children())
        self.load_pending_files()

    def render_pdf_page(self, pdf_path, page_number):
        try:
            os.environ["PDF2IMAGE_PDFIUM_PATH"] = r"C:/pdfium/pdfium.dll"
            from pdf2image import convert_from_path

            images = convert_from_path(
                pdf_path,
                first_page=page_number,
                last_page=page_number,
                dpi=200
            )
            if not images:
                messagebox.showerror("Erro", f"Número de página {page_number} está fora do intervalo.")
                return

            pil_image = images[0]
            canvas_width = self.pdf_canvas.winfo_width()
            canvas_height = self.pdf_canvas.winfo_height()
            image_ratio = pil_image.width / pil_image.height
            canvas_ratio = canvas_width / canvas_height

            if image_ratio > canvas_ratio:
                new_width = canvas_width
                new_height = int(canvas_width / image_ratio)
            else:
                new_height = canvas_height
                new_width = int(canvas_height * image_ratio)

            pil_image = pil_image.resize((new_width, new_height), Image.LANCZOS)
            self.pdf_image = ImageTk.PhotoImage(pil_image)

            self.pdf_canvas.delete("all")
            x = (canvas_width - new_width) // 2
            y = (canvas_height - new_height) // 2
            self.pdf_canvas.create_image(x, y, anchor="nw", image=self.pdf_image)
            self.pdf_canvas.image = self.pdf_image
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao renderizar a página do PDF: {str(e)}")

    def clear_canvas(self):
        self.pdf_canvas.delete("all")
        self.pdf_canvas.image = None

    def center_image(self, event):
        """Centraliza a imagem no Canvas quando ele é redimensionado."""
        self.pdf_canvas.delete('all')
        if hasattr(self, 'pdf_image'):
            # Centraliza a imagem no meio do Canvas
            self.pdf_canvas.create_image(
                self.pdf_canvas.winfo_width() // 2,
                self.pdf_canvas.winfo_height() // 2,
                anchor='center',
                image=self.pdf_image,
                tags='pdf_image'
            )

class PDFAnalyzer:
    def __init__(self, min_text_length=10, pixel_threshold=0.989, language='eng+por'):
        print("Inicializando PDFAnalyzer...")
        self.min_text_length = min_text_length
        self.pixel_threshold = pixel_threshold
        self.language = language
        # Contadores para rastrear vários status de página
        self.pages_blank_count = 0
        self.pages_blank_after_ocr_count = 0
        self.pages_ocr_analyzed_count = 0
        self.total_characters = 0
        self.correct_characters = 0
        self.total_words = 0
        self.correct_words = 0
        self.pages_low_info_count = 0
        print("PDFAnalyzer inicializado com sucesso.")

    def is_blank_or_noisy(self, image):

        print("Verificando se a imagem é em branco ou ruidosa...")

        # Recorta 5% de cada lado para remover bordas potencialmente ruidosas
        width, height = image.size
        crop_percent = 0.10
        left = int(width * crop_percent)
        right = int(width * (1 - crop_percent))
        cropped_image = image.crop((left, 0, right, height))
        print(f"Imagem cortada para remover bordas: {left}px à {right}px")

        # Converte a imagem recortada para escala de cinza
        gray_image = cv2.cvtColor(np.array(cropped_image), cv2.COLOR_RGB2GRAY)
        print("Imagem convertida para escala de cinza.")

        # Aplica limiarização adaptativa para binarizar a imagem
        binary_image = cv2.adaptiveThreshold(
            gray_image, 255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY,
            blockSize=15,
            C=10
        )
        print("Binarização adaptativa aplicada.")

        # Calcula a porcentagem de pixels brancos na imagem binarizada
        white_pixel_percentage = np.mean(binary_image == 255)
        print(f"Proporção de pixels brancos: {white_pixel_percentage:.2%}")

        # Determina se a página é considerada em branco com base na porcentagem de pixels brancos
        is_blank = white_pixel_percentage >= self.pixel_threshold
        print(f"Imagem é em branco: {is_blank}")

        return is_blank, white_pixel_percentage, cropped_image

    def perform_ocr_and_reclassify(self, cropped_image):

        print("Iniciando o processo de OCR e reclassificação...")

        try:
            # Aplica filtro mediano para reduzir o ruído na imagem
            cropped_image = cropped_image.filter(ImageFilter.MedianFilter(size=3))
            print("Filtro mediano aplicado para reduzir ruído.")

            # Aumenta o contraste e a nitidez para melhorar a precisão do OCR
            cropped_image = ImageEnhance.Contrast(cropped_image).enhance(3.0)
            print("Contraste da imagem aumentado.")
            cropped_image = ImageEnhance.Sharpness(cropped_image).enhance(2.5)
            print("Nitidez da imagem aumentada.")

            # Ajusta o DPI e converte para preto e branco para melhores resultados de OCR
            with io.BytesIO() as output:
                cropped_image.save(output, format="PNG", dpi=(300, 300))
                output.seek(0)
                with Image.open(output) as image_dpi:
                    image_bw = image_dpi.convert('L')
                    image_bw = ImageEnhance.Contrast(image_bw).enhance(2.0)
                    # Converte a imagem para binária (preto e branco) usando um limiar
                    image_bw = image_bw.point(lambda x: 0 if x < 140 else 255, '1')
                    print("Imagem convertida para preto e branco para OCR.")
                    # Configuração do Tesseract para OCR
                    custom_config = r'--oem 3 --psm 6'
                    text = pytesseract.image_to_string(image_bw, lang=self.language, config=custom_config)
                    print("OCR realizado com Tesseract.")

            # Limpa o texto extraído removendo caracteres indesejados, mas preserva espaços para correção
            text = re.sub(r'[^A-Za-z0-9À-ÿ\s]', ' ', text)
            # Substitui múltiplos espaços por um único espaço
            text = re.sub(r'\s+', ' ', text).strip()
            print(f"Texto extraído pelo OCR: {text[:50]}... (truncado)" if len(
                text) > 50 else f"Texto extraído pelo OCR: {text}")

            # Determina se o OCR foi bem-sucedido com base no comprimento do texto limpo
            ocr_successful = len(text) >= self.min_text_length
            print(f"OCR foi bem-sucedido: {ocr_successful}")
            return ocr_successful, text

        except pytesseract.TesseractError as e:
            print(f"Erro no OCR: {e}")
            return False, ""

    def analyze_page(self, img):
        # Analisa a imagem de uma única página do PDF
        is_blank, white_pixel_percentage, cropped_img = self.is_blank_or_noisy(img)
        ocr_performed = False
        status = "OK"

        quantidade_caracteres = 0

        if is_blank:
            self.pages_blank_count += 1
            ocr_successful, extracted_text = self.perform_ocr_and_reclassify(cropped_img)
            ocr_performed = True

            if extracted_text:
                quantidade_caracteres = len(extracted_text.replace(" ", ""))

            if (ocr_successful or quantidade_caracteres >= 20) and white_pixel_percentage <= 0.995:
                status = "Identificado conteúdo após reanálise"
                self.pages_ocr_analyzed_count += 1
            elif quantidade_caracteres >= 2 and white_pixel_percentage >= self.pixel_threshold:
                status = "Necessidade de revisão"
                self.pages_low_info_count += 1  # Atualiza o contador para esse status
            else:
                status = "Página em branco ou pouca info."
                self.pages_blank_after_ocr_count += 1

        return status, white_pixel_percentage, ocr_performed, quantidade_caracteres

class ReportGenerator:
        def __init__(self):
            print("Inicializando ReportGenerator...")
            self.wb = Workbook()
            self.ws = self.wb.active
            self.ws.title = "PDF Analysis Report"
            print("Workbook e Worksheet inicializados.")
            # Adicionando as novas colunas 'OCR Feito' e 'Texto Extraído'
            self.headers = ["Arquivo PDF", "Página", "Status", "Porcentagem de Pixels Brancos", "OCR Feito", "Texto Extraído"]
            self.ws.append(self.headers)
            print(f"Worksheet inicializada com cabeçalhos: {self.headers}")

        def add_record(self, pdf_name, page_num, status, white_pixel_percentage, ocr_performed, extracted_text):
            try:
                print(f"Adicionando registro: PDF Name={pdf_name}, Página={page_num}, Status={status}, Porcentagem de Pixels Brancos={white_pixel_percentage}")
                # Converter ocr_performed para 'Sim' ou 'Não'
                ocr_feito = 'Sim' if ocr_performed else 'Não'
                row = [
                    pdf_name,
                    page_num,
                    status,
                    f"{white_pixel_percentage:.2%}",
                    ocr_feito,
                    extracted_text
                ]
                self.ws.append(row)
                print(f"Linha adicionada na planilha: {row}")

                # Destacar linha em vermelho se o status for "Precisa de Atenção"
                if status == "Precisa de Atenção":
                    print("Status 'Precisa de Atenção' detectado. Destacando a linha em vermelho.")
                    fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                    for col_idx in range(1, len(row) + 1):
                        cell = self.ws.cell(row=self.ws.max_row, column=col_idx)
                        cell.fill = fill
                        print(f"Célula {cell.coordinate} destacada em vermelho.")

                print("Registro adicionado com sucesso.")
            except Exception as e:
                print(f"Erro ao adicionar registro: {e}")

        def finalize(self, output_path):
            print("Finalizando o relatório...")

            # Garantir que o diretório de destino existe
            dir_path = os.path.dirname(output_path)
            print(f"Verificando existência do diretório: {dir_path}")
            if not os.path.exists(dir_path):
                try:
                    os.makedirs(dir_path)
                    print(f"Diretório criado: {dir_path}")
                except OSError as e:
                    print(f"Erro ao criar o diretório: {e}")
                    return
            else:
                print(f"Diretório já existe: {dir_path}")

            # Estilo da tabela e ajuste das colunas
            try:
                max_row = self.ws.max_row
                max_col = self.ws.max_column
                table_ref = f"A1:{get_column_letter(max_col)}{max_row}"
                print(f"Criando tabela com referência: {table_ref}")

                # Gerar um nome de tabela único para evitar conflitos
                table_name = f"PDFAnalysisTable_{int(datetime.now().timestamp())}"

                tab = Table(displayName=table_name, ref=table_ref)
                style = TableStyleInfo(
                    name="TableStyleMedium9",
                    showFirstColumn=False,
                    showLastColumn=False,
                    showRowStripes=True,
                    showColumnStripes=True
                )
                tab.tableStyleInfo = style
                self.ws.add_table(tab)
                print("Tabela criada com estilo aplicado.")

                for col in self.ws.columns:
                    max_length = 0
                    column = col[0].column
                    column_letter = get_column_letter(column)
                    print(f"Ajustando largura da coluna {column_letter}...")
                    for cell in col:
                        cell_value_length = len(str(cell.value))
                        print(f"Célula {cell.coordinate} valor: '{cell.value}', comprimento: {cell_value_length}")
                        max_length = max(max_length, cell_value_length)
                    adjusted_width = max_length + 2
                    self.ws.column_dimensions[column_letter].width = adjusted_width
                    print(f"Largura da coluna {column_letter} ajustada para {adjusted_width}.")

                # Salvar o relatório
                print(f"Salvando o relatório no caminho: {output_path}")
                self.wb.save(output_path)
                print(f"Relatório salvo em: {output_path}")

            except Exception as e:
                print(f"Erro ao salvar o relatório: {e}")
                messagebox.showerror("Erro", f"Erro ao salvar o relatório: {str(e)}")

class TesseractConfig:
    def __init__(self, tessdata_path, tesseract_cmd):
        self.tessdata_prefix = tessdata_path
        self.tesseract_cmd = tesseract_cmd
        self.configure_tesseract()

    def configure_tesseract(self):
        os.environ['TESSDATA_PREFIX'] = self.tessdata_prefix
        pytesseract.tesseract_cmd = self.tesseract_cmd  # Atribuição correta

    def test_setup(self):
        try:
            print("Verificando a configuração do Tesseract OCR...")
            tessdata_prefix_env = os.environ.get('TESSDATA_PREFIX')
            if not tessdata_prefix_env or not os.path.isdir(tessdata_prefix_env):
                raise EnvironmentError("TESSDATA_PREFIX não está configurado corretamente.")
            tesseract_version = pytesseract.get_tesseract_version()
            print(f"Tesseract OCR instalado corretamente. Versão: {tesseract_version}")
            messagebox.showinfo("Tesseract OCR", f"Tesseract OCR instalado corretamente.\nVersão: {tesseract_version}")
        except Exception as e:
            print(f"Erro ao inicializar o Tesseract OCR: {e}")
            messagebox.showerror("Erro Tesseract OCR", f"Erro ao inicializar o Tesseract OCR.\n{str(e)}")
            sys.exit(1)


if __name__ == "__main__":
    print("Executando main.py como script principal")
    iniciar_interface_principal()
