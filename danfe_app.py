import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import os
import shutil
import threading
import webbrowser
import time
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor, as_completed
from numpy import size
import pandas as pd
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment


class DanfeAppMassa:
    def __init__(self):
        try:
            # Configuração de tema e cores 
            ctk.set_appearance_mode("light")  # Tema claro
            ctk.set_default_color_theme("blue")
            
            # Paleta de cores renamerPRO© (suavizada)
            self.cores = {
                'azul_primary': '#003D7A',     # Azul principal
                'azul_secondary': '#0056B3',   # Azul secundário
                'azul_light': '#E3F2FD',       # Azul claro suavizado
                'azul_accent': '#1976D2',      # Azul accent mais suave
                'cinza_text': '#37474F',       # Cinza textos mais suave
                'cinza_light': '#F5F7FA',      # Cinza claro suavizado
                'cinza_medium': '#ECEFF1',     # Cinza médio para cards
                'verde_success': '#2E7D32',    # Verde mais suave
                'laranja_warning': '#F57C00',  # Laranja mais suave
                'vermelho_error': '#C62828',   # Vermelho mais discreto
                'branco_suave': '#FAFBFC'      # Branco suavizado
            }
            
            # Criar janela principal primeiro
            self.root = ctk.CTk()
            self.root.title("⚕️ renamerPRO©")
            self.root.attributes('-fullscreen', True)
            self.root.minsize(800, 600)
            self.root.resizable(True, True)
            self.root.configure(fg_color=self.cores['cinza_medium'])
            
            # Função para sair do fullscreen
            def exit_fullscreen(event):
                self.root.attributes('-fullscreen', False)
            self.root.bind('<Escape>', exit_fullscreen)
            
            # Inicializar variáveis após criar a janela
            self.pasta_xml = tk.StringVar()
            self.status_texto = tk.StringVar(value="Sistema pronto para processamento")
            self.arquivos_xml = []
            self.processando = False
            self.chaves_xml = {}
            self.linhas_renomeacao = []
            
            # Configurar grid responsivo
            self.root.grid_columnconfigure(0, weight=1)
            self.root.grid_rowconfigure(0, weight=1)
            
            # Aguardar a janela estar pronta
            self.root.update_idletasks()
            
            # Criar interface
            self.criar_interface()
            
            print("✅ Aplicação inicializada com sucesso")
            
        except Exception as e:
            print(f"❌ Erro na inicialização: {e}")
            if hasattr(self, 'root'):
                try:
                    messagebox.showerror("Erro de Inicialização", f"Erro ao inicializar a aplicação:\n{str(e)}")
                except:
                    pass
            raise

    def criar_botao_profissional(self, parent, text, command, width=200, height=45, 
                                cor_principal=None, cor_hover=None, icone=""):
        """Cria botões com design profissional"""
        if cor_principal is None:
            cor_principal = self.cores['azul_primary']
        if cor_hover is None:
            cor_hover = self.cores['azul_secondary']
            
        botao = ctk.CTkButton(
            parent,
            text=f"{icone} {text}",
            command=command,
            width=width,
            height=height,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=cor_principal,
            hover_color=cor_hover,
            corner_radius=8,
            border_width=2,
            border_color=cor_principal,
            text_color=self.cores['branco_suave']
        )
        return botao

    def criar_card_profissional(self, parent, titulo, subtitle=""):
        """Cria cards profissionais com header"""
        card = ctk.CTkFrame(
            parent,
            fg_color=self.cores['branco_suave'],
            corner_radius=15,
            border_width=2,
            border_color=self.cores['azul_light']
        )
        
        # Header do card
        header = ctk.CTkFrame(
            card,
            fg_color=self.cores['azul_primary'],
            corner_radius=10,
            height=30
        )
        header.pack(fill="x", padx=5, pady=(2, 0))
        header.pack_propagate(False)
        
        # Título
        titulo_label = ctk.CTkLabel(
            header,
            text=titulo,
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=self.cores['branco_suave']
        )
        titulo_label.pack(pady=8)
        
        if subtitle:
            subtitle_label = ctk.CTkLabel(
                card,
                text=subtitle,
                font=ctk.CTkFont(size=12),
                text_color=self.cores['cinza_text']
            )
            subtitle_label.pack(pady=(8, 0))
        
        return card
        
    def criar_interface(self):
        # Frame principal responsivo
        main_container = ctk.CTkFrame(self.root, fg_color="transparent")
        main_container.pack(fill="both", expand=True, padx=0, pady=0)
        main_container.grid_columnconfigure(0, weight=1)
        
        # Header compacto
        header_frame = ctk.CTkFrame(
            main_container,
            fg_color=self.cores['azul_primary'],
            corner_radius=12,
            height=30
        )
        header_frame.pack(fill="x", pady=0)
        header_frame.pack_propagate(False)
        
        # Título compacto
        titulo_principal = ctk.CTkLabel(
            header_frame,
            text="⚕️ renamerPRO©",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=self.cores['branco_suave']
        )
        titulo_principal.pack(pady=4)
        
        # TabView responsivo
        self.tabview = ctk.CTkTabview(
            main_container,
            corner_radius=12,
            fg_color=self.cores['branco_suave'],
            segmented_button_fg_color=self.cores['azul_light'],
            segmented_button_selected_color=self.cores['azul_primary'],
            segmented_button_selected_hover_color=self.cores['azul_secondary']
        )
        self.tabview.place(x=0, y=30, relwidth=1.0, relheight=0.96)
        
        # Abas
        self.tab_principal = self.tabview.add("🏥 Processamento em Massa")
        self.tab_renomear = self.tabview.add("📋 Renomeação Inteligente")
        
        self.criar_aba_principal()
        self.criar_aba_renomeacao()

    def criar_aba_principal(self):
        # Container principal scrollável
        container = ctk.CTkScrollableFrame(self.tab_principal, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=0, pady=0)
        container.grid_columnconfigure(0, weight=1)
        
        # Card de Configuração
        config_card = self.criar_card_profissional(
            container, 
            "⚙️ Configuração de Processamento",
            "Configure as pastas de origem e destino dos documentos"
        )
        config_card.pack(fill="x", pady=0, padx=0)
        
        # Grid responsivo para inputs
        input_frame = ctk.CTkFrame(config_card, fg_color="transparent")
        input_frame.pack(fill="x", padx=8, pady=3)
        input_frame.grid_columnconfigure(0, weight=1)
        
        # Pasta XML
        xml_label = ctk.CTkLabel(
            input_frame,
            text="📁 Pasta com Arquivos XML:",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=self.cores['cinza_text'],
            anchor="w"
        )
        xml_label.grid(row=0, column=0, sticky="ew", pady=(0, 5))
        
        xml_container = ctk.CTkFrame(input_frame, fg_color="transparent")
        xml_container.grid(row=1, column=0, sticky="ew", pady=(0, 15))
        xml_container.grid_columnconfigure(0, weight=1)
        
        self.entrada_pasta_xml = ctk.CTkEntry(
            xml_container,
            textvariable=self.pasta_xml,
            placeholder_text="Selecione a pasta contendo os arquivos XML...",
            font=ctk.CTkFont(size=12),
            height=40,
            corner_radius=8,
            border_color=self.cores['azul_light']
        )
        self.entrada_pasta_xml.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        
        btn_xml = self.criar_botao_profissional(
            xml_container,
            "Selecionar",
            self.selecionar_pasta_xml,
            width=120,
            icone="📁"
        )
        btn_xml.grid(row=0, column=1)
               
        # Info e Controles
        controle_card = self.criar_card_profissional(
            container,
            "📊 Controle de Processamento"
        )
        controle_card.pack(fill="x", pady=(2, 0), padx=0)
        
        # Status dos arquivos
        self.label_arquivos = ctk.CTkLabel(
            controle_card,
            text="📋 Aguardando seleção de pasta...",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=self.cores['cinza_text']
        )
        self.label_arquivos.pack(pady=3)
        
        # Botões de ação
        botoes_frame = ctk.CTkFrame(controle_card, fg_color="transparent")
        botoes_frame.pack(pady=(0, 3))
        
        self.btn_escanear = self.criar_botao_profissional(
            botoes_frame,
            "ESCANEAR PASTA",
            self.escanear_pasta,
            width=170,
            height=45,
            cor_principal=self.cores['laranja_warning'],
            cor_hover="#E5A500",
            icone="🔍"
        )
        self.btn_escanear.pack(side="left", padx=8)
        
        self.btn_processar = self.criar_botao_profissional(
            botoes_frame,
            "PROCESSAR DOCUMENTOS",
            self.processar_massa_thread,
            width=200,
            height=45,
            cor_principal=self.cores['verde_success'],
            cor_hover="#218838",
            icone="⚡"
        )
        self.btn_processar.pack(side="left", padx=8)
        self.btn_processar.configure(state="disabled")
        
        # Progresso
        progresso_frame = ctk.CTkFrame(controle_card, fg_color="transparent")
        progresso_frame.pack(fill="x", padx=12, pady=(0, 3))
        
        prog_label = ctk.CTkLabel(
            progresso_frame,
            text="📈 Progresso do Processamento:",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=self.cores['cinza_text'],
            anchor="w"
        )
        prog_label.pack(fill="x", pady=(0, 6))
        
        self.progresso_geral = ctk.CTkProgressBar(
            progresso_frame,
            height=18,
            corner_radius=10,
            progress_color=self.cores['azul_primary']
        )
        self.progresso_geral.pack(fill="x", pady=(0, 6))
        self.progresso_geral.set(0)
        
        self.label_progresso = ctk.CTkLabel(
            progresso_frame,
            text="0 / 0 documentos processados",
            font=ctk.CTkFont(size=12),
            text_color=self.cores['cinza_text']
        )
        self.label_progresso.pack()
        
        # Log profissional
        log_card = self.criar_card_profissional(
            container,
            "📋 Log de Processamento",
            "Acompanhe em tempo real o processamento dos documentos"
        )
        log_card.pack(fill="both", expand=True, pady=(3, 0), padx=0)
        
        self.log_text = ctk.CTkTextbox(
            log_card,
            font=ctk.CTkFont(size=11, family="Consolas"),
            corner_radius=8,
            fg_color=self.cores['cinza_medium'],
            text_color=self.cores['cinza_text'],
            height=120
        )
        self.log_text.pack(fill="both", expand=True, padx=8, pady=8)
        
        # Status bar profissional
        status_frame = ctk.CTkFrame(
            container,
            fg_color=self.cores['azul_primary'],
            corner_radius=10,
            height=40
        )
        status_frame.pack(fill="x", pady=(2, 0))
        status_frame.pack_propagate(False)
        
        status_container = ctk.CTkFrame(status_frame, fg_color="transparent")
        status_container.pack(expand=True, fill="both")
        
        status_label = ctk.CTkLabel(
            status_container,
            text="💼 Status:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=self.cores['branco_suave']
        )
        status_label.pack(side="left", padx=(15, 5), pady=8)
        
        self.status_label = ctk.CTkLabel(
            status_container,
            textvariable=self.status_texto,
            font=ctk.CTkFont(size=12),
            text_color=self.cores['azul_light']
        )
        self.status_label.pack(side="left", padx=5, pady=8)
        
        self.carregar_log_inicial()

    def criar_aba_renomeacao(self):
        # Container principal scrollável
        container = ctk.CTkScrollableFrame(self.tab_renomear, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=0, pady=0)
        container.grid_columnconfigure(0, weight=1)
        
        # Header da aba
        header_card = self.criar_card_profissional(
            container,
            "📋 Renomeação Inteligente por Chave de Acesso",
            "Sistema avançado para renomeação e processamento de documentos fiscais"
        )
        header_card.pack(fill="x", pady=0, padx=5)
        
        # Configuração
        config_frame = ctk.CTkFrame(header_card, fg_color="transparent")
        config_frame.pack(fill="x", padx=12, pady=12)
        config_frame.grid_columnconfigure(0, weight=1)
        
        pasta_label = ctk.CTkLabel(
            config_frame,
            text="📁 Diretório de Documentos XML:",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=self.cores['cinza_text'],
            anchor="w"
        )
        pasta_label.grid(row=0, column=0, sticky="ew", pady=(0, 5))
        
        pasta_container = ctk.CTkFrame(config_frame, fg_color="transparent")
        pasta_container.grid(row=1, column=0, sticky="ew")
        pasta_container.grid_columnconfigure(0, weight=1)
        
        self.entrada_pasta_renomear = ctk.CTkEntry(
            pasta_container,
            placeholder_text="Selecione o diretório contendo os arquivos XML...",
            font=ctk.CTkFont(size=12),
            height=40,
            corner_radius=8,
            border_color=self.cores['azul_light']
        )
        self.entrada_pasta_renomear.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        
        btn_pasta_renomear = self.criar_botao_profissional(
            pasta_container,
            "Localizar",
            self.selecionar_pasta_renomear,
            width=120,
            icone="🔍"
        )
        btn_pasta_renomear.grid(row=0, column=1)
       
        # Controles avançados
        controles_card = self.criar_card_profissional(
            container,
            "🛠️ Controles de Operações"
        )
        controles_card.pack(fill="x", pady=(0, 0))
        
        # Botões organizados profissionalmente
        botoes_container = ctk.CTkFrame(controles_card, fg_color="transparent")
        botoes_container.pack(fill="x", padx=10, pady=(10, 10))
        botoes_container.grid_columnconfigure((0, 1, 2, 3), weight=1)
        botoes_container.grid_columnconfigure(4, weight=2)  # Botão de processar com mais peso

        self.btn_lote_dados = self.criar_botao_profissional(
            botoes_container,
            "LOTE DE DADOS",
            self.abrir_janela_lote,
            height=40,
            cor_principal="#6f42c1",
            cor_hover="#5a349b",
            icone="📋"
        )
        self.btn_lote_dados.grid(row=0, column=0, padx=(0, 8), sticky="ew")

        self.btn_adicionar_linha = self.criar_botao_profissional(
            botoes_container,
            "NOVA LINHA",
            self.adicionar_linha_renomeacao,
            height=40,
            cor_principal=self.cores['azul_accent'],
            cor_hover="#0066CC",
            icone="➕"
        )
        self.btn_adicionar_linha.grid(row=0, column=1, padx=8, sticky="ew")

        self.btn_limpar_dados = self.criar_botao_profissional(
            botoes_container,
            "LIMPAR",
            self.limpar_dados_massa,
            height=40,
            cor_principal="#6C757D",
            cor_hover="#5A6268",
            icone="🧹"
        )
        self.btn_limpar_dados.grid(row=0, column=2, padx=8, sticky="ew")

        self.btn_exportar_excel = self.criar_botao_profissional(
            botoes_container,
            "EXPORTAR EXCEL",
            self.exportar_para_excel,
            height=40,
            cor_principal="#28A745",
            cor_hover="#218838",
            icone="📊"
        )
        self.btn_exportar_excel.grid(row=0, column=3, padx=8, sticky="ew")

        self.btn_processar_completo = self.criar_botao_profissional(
            botoes_container,
            "🚀 PROCESSAR COMPLETO",
            self.processar_completo_thread,
            height=45,
            cor_principal="#FF6B35",
            cor_hover="#E55A2B",
        )
        self.btn_processar_completo.grid(row=0, column=4, padx=(8, 0), sticky="ew")
        

        
        # Tabela profissional
        tabela_card = self.criar_card_profissional(
            container,
            "📊 Tabela de Mapeamento",
        )
        tabela_card.pack(fill="x", pady=(0, 10))
        
        # Cabeçalho da tabela
        header_tabela = ctk.CTkFrame(
            tabela_card,
            fg_color=self.cores['azul_light'],
            corner_radius=8
        )
        header_tabela.pack(fill="x", padx=12, pady=(12, 5))
        
        # Grid responsivo para cabeçalhos
        header_tabela.grid_columnconfigure(0, weight=2)  # Chave de Acesso
        header_tabela.grid_columnconfigure(1, weight=2)  # Nome do Arquivo
        header_tabela.grid_columnconfigure(2, weight=1)  # Status
        header_tabela.grid_columnconfigure(3, weight=2)  # Ações

        # Cabeçalhos
        ctk.CTkLabel(
            header_tabela,
            text="🔑 Chave de Acesso",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=self.cores['azul_primary']
        ).grid(row=0, column=0, padx=15, pady=12, sticky="w")

        ctk.CTkLabel(
            header_tabela,
            text="📄 Nome do Arquivo",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=self.cores['azul_primary']
        ).grid(row=0, column=1, padx=15, pady=12, sticky="w")

        ctk.CTkLabel(
            header_tabela,
            text="📊 Status",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=self.cores['azul_primary']
        ).grid(row=0, column=2, padx=30, pady=12, sticky="w")

        ctk.CTkLabel(
            header_tabela,
            text="🛠️ Ações",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=self.cores['azul_primary']
        ).grid(row=0, column=3, padx=15, pady=12, sticky="w")
        
        # Área scrollável responsiva
        self.scroll_frame = ctk.CTkScrollableFrame(
            tabela_card,
            fg_color=self.cores['cinza_medium'],
            corner_radius=8,
            height=180
        )
        self.scroll_frame.pack(fill="x", padx=12, pady=(0, 12))
        
        # Configurar responsividade do scroll
        self.scroll_frame.grid_columnconfigure(0, weight=1)
        
        self.linhas_renomeacao = []
        
        # Log profissional da renomeação
        log_renomear_card = self.criar_card_profissional(
            container,
            "📋 Log de Renomeação",
            "Acompanhe o processo de validação e renomeação dos arquivos"
        )
        log_renomear_card.pack(fill="x", pady=(0, 10))
        
        self.log_renomeacao = ctk.CTkTextbox(
              log_renomear_card, height=200,
            font=ctk.CTkFont(size=11, family="Consolas"),
            
            corner_radius=8,
            fg_color=self.cores['cinza_medium'],
            text_color=self.cores['cinza_text']
        )
        self.log_renomeacao.pack(fill="both", expand=True, padx=12, pady=12)
        
        # Log inicial
        self.log_renomeacao.insert("0.0", """renamerPRO©
📋 Aguardando configuração de diretório...
💡 Selecione o diretório e escaneie as chaves para começar.""")
        
        # Adicionar uma linha inicial para garantir visibilidade da tabela
        self.adicionar_linha_renomeacao()

    def adicionar_linha_renomeacao(self):
        # Container responsivo para linha
        linha_frame = ctk.CTkFrame(
            self.scroll_frame,
            fg_color=self.cores['branco_suave'],
            corner_radius=8,
            border_width=1,
            border_color=self.cores['azul_light']
        )
        linha_frame.pack(fill="x", padx=5, pady=3)
        
        # Grid responsivo
        linha_frame.grid_columnconfigure(0, weight=2)  # Chave
        linha_frame.grid_columnconfigure(1, weight=2)  # Nome
        linha_frame.grid_columnconfigure(2, weight=1)  # Status
        linha_frame.grid_columnconfigure(3, weight=0)  # Botão

        # Entry para chave
        entry_chave = ctk.CTkEntry(
            linha_frame,
            placeholder_text="Chave de acesso (44 dígitos)...",
            font=ctk.CTkFont(size=11),
            height=35,
            corner_radius=6,
            border_color=self.cores['azul_light']
        )
        entry_chave.grid(row=0, column=0, padx=(10, 5), pady=8, sticky="ew")

        # Entry para nome
        entry_nome = ctk.CTkEntry(
            linha_frame,
            placeholder_text="Nome do arquivo (sem extensão)...",
            font=ctk.CTkFont(size=11),
            height=35,
            corner_radius=6,
            border_color=self.cores['azul_light']
        )
        entry_nome.grid(row=0, column=1, padx=5, pady=8, sticky="ew")

        # Status
        label_status = ctk.CTkLabel(
            linha_frame,
            text="⏳ Aguardando",
            font=ctk.CTkFont(size=11),
            text_color=self.cores['cinza_text']
        )
        label_status.grid(row=0, column=2, padx=5, pady=8, sticky="w")

        # Botão remover
        btn_remover = ctk.CTkButton(
            linha_frame,
            text="      🗑️",
            command=lambda: self.remover_linha_renomeacao(linha_frame),
            width=38,
            height=38,
            font=ctk.CTkFont(size=18),
            text_color=self.cores['branco_suave'],
            fg_color=self.cores['vermelho_error'],
            hover_color="#C82333",
            corner_radius=6,
            border_width=0,
        )
        btn_remover.grid(row=0, column=3, padx=(5, 10), pady=8)
        
        self.linhas_renomeacao.append({
            'frame': linha_frame,
            'chave': entry_chave,
            'nome': entry_nome,
            'status': label_status
        })

    def remover_linha_renomeacao(self, linha_frame):
        # Encontrar e remover a linha
        for i, linha in enumerate(self.linhas_renomeacao):
            if linha['frame'] == linha_frame:
                linha_frame.destroy()
                self.linhas_renomeacao.pop(i)
                break
                
    def selecionar_pasta_renomear(self):
        try:
            # Verificar se a janela principal está disponível
            if not self.root or not self.root.winfo_exists():
                messagebox.showerror("Erro", "Janela principal não está disponível")
                return
                
            # Forçar atualização da janela antes de abrir o diálogo
            self.root.update_idletasks()
            
            pasta = filedialog.askdirectory(
                title="Selecione a pasta com XMLs para renomear",
                parent=self.root,
                mustexist=True
            )
            
            if pasta and os.path.exists(pasta):
                self.entrada_pasta_renomear.delete(0, 'end')
                self.entrada_pasta_renomear.insert(0, pasta)
                self.log_renomeacao.delete("0.0", "end")
                self.log_renomeacao.insert("0.0", f"📁 Pasta selecionada: {pasta}\n💡 Use 'PROCESSAR COMPLETO' para iniciar.")
                
        except tk.TclError as e:
            print(f"Erro TclError no seletor de pasta: {e}")
            # Tentar método alternativo
            try:
                import tkinter.simpledialog as simpledialog
                pasta = simpledialog.askstring(
                    "Pasta XML", 
                    "Digite o caminho da pasta com XMLs:",
                    parent=self.root
                )
                if pasta and os.path.exists(pasta):
                    self.entrada_pasta_renomear.delete(0, 'end')
                    self.entrada_pasta_renomear.insert(0, pasta)
            except Exception:
                messagebox.showerror("Erro", "Não foi possível abrir o seletor de pasta.\nDigite o caminho manualmente no campo.")
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir seletor de pasta: {str(e)}")
            print(f"Erro no seletor de pasta: {e}")
        
    def extrair_chave_xml(self, caminho_arquivo):
        try:
            tree = ET.parse(caminho_arquivo)
            root = tree.getroot()
            
            # Buscar chave de acesso em diferentes locais possíveis
            namespaces = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
            
            # Tentar encontrar a chave
            chave_elem = root.find('.//nfe:chNFe', namespaces)
            if chave_elem is not None:
                return chave_elem.text
                
            # Tentar sem namespace
            chave_elem = root.find('.//chNFe')
            if chave_elem is not None:
                return chave_elem.text
                
            # Buscar no atributo Id
            id_elem = root.find('.//*[@Id]')
            if id_elem is not None:
                id_value = id_elem.get('Id')
                if id_value and len(id_value) >= 44:
                    return id_value[-44:]  # Pegar os últimos 44 caracteres
                    
            return None
            
        except Exception:
            return None
    
    def extrair_valor_total_xml(self, caminho_arquivo):
        """Extrai o valor total da NFe do XML"""
        try:
            tree = ET.parse(caminho_arquivo)
            root = tree.getroot()
            
            # Buscar valor total em diferentes locais possíveis
            namespaces = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
            
            # Tentar encontrar o valor total
            valor_elem = root.find('.//nfe:vNF', namespaces)
            if valor_elem is not None:
                return f"R$ {float(valor_elem.text):.2f}"
                
            # Tentar sem namespace
            valor_elem = root.find('.//vNF')
            if valor_elem is not None:
                return f"R$ {float(valor_elem.text):.2f}"
                
            return None
            
        except Exception:
            return None
    
    def extrair_numero_nf_xml(self, caminho_arquivo):
        """Extrai o número da NFe do XML"""
        try:
            tree = ET.parse(caminho_arquivo)
            root = tree.getroot()
            
            # Buscar número da NF em diferentes locais possíveis
            namespaces = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
            
            # Tentar encontrar o número da NF
            numero_elem = root.find('.//nfe:nNF', namespaces)
            if numero_elem is not None:
                return numero_elem.text
                
            # Tentar sem namespace
            numero_elem = root.find('.//nNF')
            if numero_elem is not None:
                return numero_elem.text
                
            return None
            
        except Exception:
            return None
    
    def extrair_numero_pedido_xml(self, caminho_arquivo):
        """Extrai o número do pedido (xPed) do XML"""
        try:
            tree = ET.parse(caminho_arquivo)
            root = tree.getroot()
            
            # Buscar xPed em diferentes locais possíveis
            namespaces = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
            
            # Tentar encontrar o xPed (pode estar em qualquer item da nota)
            pedido_elem = root.find('.//nfe:xPed', namespaces)
            if pedido_elem is not None:
                return pedido_elem.text
                
            # Tentar sem namespace
            pedido_elem = root.find('.//xPed')
            if pedido_elem is not None:
                return pedido_elem.text
                
            return None
            
        except Exception:
            return None
    
    def extrair_numero_fornecedor_xml(self, caminho_arquivo):
        """Extrai o número do fornecedor (CNPJ/CPF do emitente) do XML"""
        try:
            tree = ET.parse(caminho_arquivo)
            root = tree.getroot()
            
            # Buscar CNPJ/CPF do emitente em diferentes locais possíveis
            namespaces = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
            
            # Tentar encontrar CNPJ do emitente
            cnpj_elem = root.find('.//nfe:emit/nfe:CNPJ', namespaces)
            if cnpj_elem is not None:
                return cnpj_elem.text
                
            # Tentar encontrar CPF do emitente
            cpf_elem = root.find('.//nfe:emit/nfe:CPF', namespaces)
            if cpf_elem is not None:
                return cpf_elem.text
                
            # Tentar sem namespace - CNPJ
            cnpj_elem = root.find('.//emit/CNPJ')
            if cnpj_elem is not None:
                return cnpj_elem.text
                
            # Tentar sem namespace - CPF
            cpf_elem = root.find('.//emit/CPF')
            if cpf_elem is not None:
                return cpf_elem.text
                
            return None
            
        except Exception:
            return None
            
    def validar_e_renomear_thread(self):
        # Usar função auxiliar (elimina duplicação)
        self.executar_thread_segura(self.validar_e_renomear)
        
    def validar_e_renomear(self):
        pasta = self.entrada_pasta_renomear.get()
        if not pasta:
            messagebox.showerror("Erro", "Selecione a pasta primeiro!")
            return
            
        if not self.chaves_xml:
            messagebox.showerror("Erro", "Escaneie as chaves primeiro!")
            return
            
        # Criar pastas automaticamente no diretório dos XMLs
        pasta_xml = self.entrada_pasta_renomear.get()
        pasta_renomeados = os.path.join(pasta_xml, "Arquivos Renomeados")
        pasta_nao_renomeados = os.path.join(pasta_xml, "Arquivos Não Renomeados")
        
        # Criar as pastas se não existirem
        os.makedirs(pasta_renomeados, exist_ok=True)
        os.makedirs(pasta_nao_renomeados, exist_ok=True)
            
        sucessos = 0
        erros = 0
        
        self.root.after(0, lambda: self.log_renomeacao.insert("end", "\n🚀 INICIANDO VALIDAÇÃO E RENOMEAÇÃO...\n\n"))
        self.root.after(0, lambda: self.log_renomeacao.insert("end", f"📁 Pasta de origem: {pasta_xml}\n"))
        self.root.after(0, lambda: self.log_renomeacao.insert("end", f"📁 Arquivos renomeados: {pasta_renomeados}\n"))
        self.root.after(0, lambda: self.log_renomeacao.insert("end", f"📁 Arquivos não renomeados: {pasta_nao_renomeados}\n\n"))
        
        # Coletar chaves que serão convertidas
        chaves_para_converter = set()
        for linha in self.linhas_renomeacao:
            chave_original = linha['chave'].get().strip()
            nome_final = linha['nome'].get().strip()
            
            if chave_original and nome_final:
                chaves_para_converter.add(chave_original.strip())
        
        # Processar renomeação
        for linha in self.linhas_renomeacao:
            chave_original = linha['chave'].get().strip()
            nome_final = linha['nome'].get().strip()
            
            if not chave_original or not nome_final:
                continue
                
            chave = chave_original.strip()
            
            # Validar chave (usando função auxiliar - elimina duplicação)
            if not self.validar_chave_nfe(chave):
                self.root.after(0, lambda l=linha: l['status'].configure(text="❌ Chave"))
                self.root.after(0, lambda c=chave: self.log_renomeacao.insert("end", f"❌ Chave inválida: {c}\n"))
                erros += 1
                continue
                
            # Verificar se chave existe
            if chave not in self.chaves_xml:
                self.root.after(0, lambda l=linha: l['status'].configure(text="❌ N/Existe"))
                self.root.after(0, lambda c=chave: self.log_renomeacao.insert("end", f"❌ Chave não encontrada: {c}\n"))
                erros += 1
                continue
                
            # Renomear arquivo
            try:
                arquivo_original = self.chaves_xml[chave]
                novo_nome = os.path.join(pasta_renomeados, f"{nome_final}.xml")
                
                if os.path.exists(novo_nome):
                    self.root.after(0, lambda l=linha: l['status'].configure(text="❌ Existe"))
                    self.root.after(0, lambda n=nome_final: self.log_renomeacao.insert("end", f"❌ Arquivo já existe: {n}.xml\n"))
                    erros += 1
                    continue
                    
                # Copiar arquivo para pasta de arquivos renomeados com novo nome
                import shutil
                shutil.move(arquivo_original, novo_nome)
                
                self.root.after(0, lambda l=linha: l['status'].configure(text="✅ OK"))
                self.root.after(0, lambda o=os.path.basename(arquivo_original), n=nome_final: 
                              self.log_renomeacao.insert("end", f"✅ {o} → {n}.xml\n"))
                sucessos += 1
                
            except Exception as e:
                self.root.after(0, lambda l=linha: l['status'].configure(text="❌ Erro"))
                self.root.after(0, lambda e=str(e): self.log_renomeacao.insert("end", f"❌ Erro: {e}\n"))
                erros += 1
              
        self.root.after(0, lambda: self.log_renomeacao.insert("end", f"\n🎉 RENOMEAÇÃO CONCLUÍDA!\n"))
        self.root.after(0, lambda: self.log_renomeacao.insert("end", f"✅ XMLs renomeados: {sucessos}\n"))
        self.root.after(0, lambda: self.log_renomeacao.insert("end", f"❌ Erros: {erros}\n"))
        self.root.after(0, lambda: self.log_renomeacao.see("end"))
        
        if sucessos > 0:
            resposta = messagebox.askyesno(
                "Concluído!", 
                f"Renomeação finalizada!\n\n✅ {sucessos} XMLs renomeados\n❌ {erros} erros\n\n📁 Arquivos organizados em:\n   • Arquivos Renomeados\n\nDeseja abrir a pasta com os resultados?"
            )
            
            if resposta:
                try:
                    os.startfile(pasta_xml)
                except:
                    import webbrowser
                    webbrowser.open(pasta_xml)
    
    def exportar_para_excel(self):
        """Exporta os dados dos XMLs da pasta selecionada para um arquivo Excel"""
        try:
            # Verificar se há pasta selecionada
            pasta = self.entrada_pasta_renomear.get()
            if not pasta:
                messagebox.showerror("Erro", "Selecione a pasta com XMLs primeiro!")
                return
            
            # Escanear todos os XMLs da pasta
            arquivos_xml = self.escanear_xmls_pasta(pasta)
            
            if not arquivos_xml:
                messagebox.showwarning("Aviso", "Nenhum arquivo XML encontrado na pasta selecionada!")
                return
            
            # Coletar dados dos XMLs
            dados_tabela = []
            self.log_renomeacao.insert("end", f"\n🔍 ESCANEANDO XMLs PARA EXPORTAÇÃO...\n")
            self.log_renomeacao.insert("end", f"📁 Pasta: {pasta}\n")
            self.log_renomeacao.insert("end", f"📊 Total de XMLs: {len(arquivos_xml)}\n\n")
            
            for i, arquivo_xml in enumerate(arquivos_xml, 1):
                try:
                    nome_arquivo = os.path.basename(arquivo_xml)
                    chave = self.extrair_chave_xml(arquivo_xml)
                    valor_total = self.extrair_valor_total_xml(arquivo_xml)
                    numero_nf = self.extrair_numero_nf_xml(arquivo_xml)
                    numero_pedido = self.extrair_numero_pedido_xml(arquivo_xml)
                    numero_fornecedor = self.extrair_numero_fornecedor_xml(arquivo_xml)
                    
                    # Verificar se existe na tabela de mapeamento
                    status_conversao = "Não mapeado"
                    nome_personalizado = ""
                    
                    for linha in self.linhas_renomeacao:
                        if linha['chave'].get().strip() == chave:
                            status_conversao = linha['status'].cget('text')
                            nome_personalizado = linha['nome'].get().strip()
                            break
                    
                    dados_tabela.append({
                        'N° NF': numero_nf or i,
                        'CNPJ fornecedor': numero_fornecedor or 'Não encontrado',
                        'N° Pedido': numero_pedido or 'Não encontrado',
                        'V. TOTAL DA NOTA': valor_total or '',
                        'Nome do Arquivo': nome_personalizado or nome_arquivo,
                        'Data/Hora Exportação': datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
                        'Chave de acesso': chave or 'Não encontrada',
                    })
                    
                    self.log_renomeacao.insert("end", f"✅ {nome_arquivo}\n")
                    
                except Exception as e:
                    self.log_renomeacao.insert("end", f"❌ Erro em {os.path.basename(arquivo_xml)}: {str(e)}\n")
                    continue
            
            # Solicitar local para salvar
            try:
                arquivo_excel = filedialog.asksaveasfilename(
                    title="Salvar planilha Excel",
                    defaultextension=".xlsx",
                    filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")],
                    initialfile=f"renamerPRO_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    parent=self.root
                )
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao abrir seletor de arquivo: {str(e)}")
                return
            
            if not arquivo_excel:
                return
            
            # Criar DataFrame e exportar
            df = pd.DataFrame(dados_tabela)
            
            # Configurar o Excel com formatação
            with pd.ExcelWriter(arquivo_excel, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Mapeamento NFe', index=False)
                
                # Obter a planilha para formatação
                worksheet = writer.sheets['Mapeamento NFe']
                
                # Ajustar largura das colunas
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
                
                # Aplicar formatação ao cabeçalho
                header_font = Font(bold=True, color="FFFFFF")
                header_fill = PatternFill(start_color="003D7A", end_color="003D7A", fill_type="solid")
                header_alignment = Alignment(horizontal="center", vertical="center")
                
                for cell in worksheet[1]:
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = header_alignment
            
            # Log da exportação
            self.log_renomeacao.insert("end", f"\n📊 EXPORTAÇÃO PARA EXCEL:\n")
            self.log_renomeacao.insert("end", f"✅ {len(dados_tabela)} registros exportados\n")
            self.log_renomeacao.insert("end", f"📁 Arquivo: {os.path.basename(arquivo_excel)}\n")
            self.log_renomeacao.insert("end", f"📅 Data: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
            self.log_renomeacao.see("end")
            
            # Perguntar se deseja abrir o arquivo
            resposta = messagebox.askyesno(
                "Exportação Concluída!", 
                f"✅ Dados exportados com sucesso!\n\n📊 {len(dados_tabela)} registros salvos\n📁 {os.path.basename(arquivo_excel)}\n\nDeseja abrir o arquivo Excel?"
            )
            
            if resposta:
                try:
                    os.startfile(arquivo_excel)
                except:
                    webbrowser.open(arquivo_excel)
                    
        except Exception as e:
            error_msg = f"Erro ao exportar para Excel: {str(e)}"
            messagebox.showerror("Erro na Exportação", error_msg)
            self.log_renomeacao.insert("end", f"\n❌ ERRO NA EXPORTAÇÃO: {error_msg}\n")
            self.log_renomeacao.see("end")
    
    def limpar_dados_massa(self):
        resposta = messagebox.askyesno(
            "Confirmar Limpeza", 
            "🧹 Tem certeza que deseja limpar TODOS os dados da tabela?\n\nEsta ação não pode ser desfeita!"
        )
        
        if resposta:
            # Limpar todas as linhas existentes
            for linha in self.linhas_renomeacao:
                linha['frame'].destroy()
            self.linhas_renomeacao.clear()
            
            # Log
            self.log_renomeacao.insert("end", f"\n🧹 DADOS LIMPOS:\n")
            self.log_renomeacao.insert("end", f"✅ Tabela limpa com sucesso\n")
            self.log_renomeacao.insert("end", f"📝 Tabela pronta para novos dados\n")
            self.log_renomeacao.see("end")
            
            messagebox.showinfo("Limpeza Concluída!", "✅ Todos os dados foram removidos da tabela!")
    
    def processar_completo_thread(self):
        """Thread para processamento completo: escanear + validar + renomear + gerar PDFs"""
        if self.processando:
            return
        
        # Usar função auxiliar (elimina duplicação)
        self.executar_thread_segura(self.processar_completo)
    
    def processar_selecionados_thread(self):
        if self.processando:
            return
        
        # Usar função auxiliar (elimina duplicação)
        self.executar_thread_segura(self.processar_selecionados)
    
    def processar_selecionados(self):
        pasta_xml = self.entrada_pasta_renomear.get()
        
        if not pasta_xml:
            messagebox.showerror("Erro", "Selecione a pasta com XMLs primeiro!")
            return
        
        # Escanear TODOS os XMLs da pasta (usando função auxiliar - elimina duplicação)
        todos_xmls = self.escanear_xmls_pasta(pasta_xml)
        
        if not todos_xmls:
            messagebox.showerror("Erro", "Nenhum arquivo XML encontrado na pasta!")
            return
        
        # Determinar pasta de saída
        pasta_saida = self.entrada_pasta_renomear.get()  # Usar a mesma pasta por padrão
        
        self.processando = True
        total = len(todos_xmls)
        sucessos = 0
        erros = 0
        
        self.root.after(0, lambda: self.btn_processar_completo.configure(state="disabled", text="🔄 Processando..."))

        
        self.root.after(0, lambda: self.log_renomeacao.insert("end", f"\n🚀 PROCESSANDO TODOS OS XMLs DA PASTA:\n"))
        self.root.after(0, lambda: self.log_renomeacao.insert("end", f"📊 Total: {total} arquivos\n"))
        self.root.after(0, lambda: self.log_renomeacao.insert("end", f"📤 Pasta saída: {pasta_saida}\n\n"))
        
        inicio = time.time()
        
        # Processar com função auxiliar (elimina duplicação)
        def callback_sucesso(nome):
            self.log_renomeacao.insert("end", f"✅ {nome}\n")
            
        def callback_erro(nome):
            self.log_renomeacao.insert("end", f"❌ {nome}\n")
            
        sucessos, erros, tempo_total = self.processar_xmls_paralelo(
            todos_xmls, pasta_saida, callback_sucesso, callback_erro
        )
        
        self.root.after(0, lambda: self.log_renomeacao.insert("end", f"\n🎉 PROCESSAMENTO CONCLUÍDO!\n"))
        self.root.after(0, lambda: self.log_renomeacao.insert("end", f"✅ PDFs criados: {sucessos}\n"))
        self.root.after(0, lambda: self.log_renomeacao.insert("end", f"❌ Erros: {erros}\n"))
        self.root.after(0, lambda: self.log_renomeacao.insert("end", f"⏱️ Tempo total: {tempo_total:.1f} segundos\n"))
        
        self.root.after(0, lambda: self.btn_processar_completo.configure(state="normal", text="🚀 PROCESSAR COMPLETO"))

        self.root.after(0, lambda: self.log_renomeacao.see("end"))
        
        # Usar função auxiliar (elimina duplicação)
        self.mostrar_conclusao_processamento(sucessos, erros, tempo_total, pasta_saida)
        
        self.processando = False
    
    def processar_completo(self):
        """Processamento completo: escanear + validar + renomear + gerar PDFs em um único fluxo"""
        pasta_xml = self.entrada_pasta_renomear.get()
        
        if not pasta_xml:
            messagebox.showerror("Erro", "Selecione a pasta com XMLs primeiro!")
            return
        
        if not os.path.exists(pasta_xml):
            messagebox.showerror("Erro", "Pasta de XMLs não existe!")
            return
        
        self.processando = True
        inicio_total = time.time()
        
        # Desabilitar botões durante processamento
        self.root.after(0, lambda: self.btn_processar_completo.configure(state="disabled", text="🔄 Processando..."))


        
        try:
            # ETAPA 1: ESCANEAMENTO AUTOMÁTICO
            self.root.after(0, lambda: self.log_renomeacao.insert("end", "\n🚀 INICIANDO PROCESSAMENTO COMPLETO...\n\n"))
            self.root.after(0, lambda: self.log_renomeacao.insert("end", "📋 ETAPA 1: Escaneando XMLs da pasta...\n"))
            
            todos_xmls = self.escanear_xmls_pasta(pasta_xml)
            if not todos_xmls:
                messagebox.showerror("Erro", "Nenhum arquivo XML encontrado na pasta!")
                return
            
            # Mapear chaves automaticamente
            self.chaves_xml = {}
            arquivos_processados = 0
            
            for arquivo in todos_xmls:
                nome_arquivo = os.path.basename(arquivo)
                chave = self.extrair_chave_xml(arquivo)
                
                if chave:
                    self.chaves_xml[chave] = arquivo
                    arquivos_processados += 1
            
            self.root.after(0, lambda: self.log_renomeacao.insert("end", f"✅ {arquivos_processados} chaves mapeadas automaticamente\n\n"))
            
            # Criar pastas automaticamente no diretório dos XMLs
            pasta_renomeados = os.path.join(pasta_xml, "Arquivos Renomeados")
            pasta_nao_renomeados = os.path.join(pasta_xml, "Arquivos Não Renomeados")
            pasta_pdf_convertido = os.path.join(pasta_xml, "PDFs convertidos")
            
            os.makedirs(pasta_renomeados, exist_ok=True)
            os.makedirs(pasta_nao_renomeados, exist_ok=True)
            os.makedirs(pasta_pdf_convertido, exist_ok=True)
            
            self.root.after(0, lambda: self.log_renomeacao.insert("end", f"📁 Pasta criada: {os.path.basename(pasta_renomeados)}\n"))
            self.root.after(0, lambda: self.log_renomeacao.insert("end", f"📁 Pasta criada: {os.path.basename(pasta_nao_renomeados)}\n"))
            self.root.after(0, lambda: self.log_renomeacao.insert("end", f"📁 Pasta criada: {os.path.basename(pasta_pdf_convertido)}\n\n"))
            
            # ETAPA 2: VALIDAÇÃO E RENOMEAÇÃO (se há dados na tabela)
            linhas_com_dados = [linha for linha in self.linhas_renomeacao if linha['chave'].get().strip()]
            
            if linhas_com_dados:
                self.root.after(0, lambda: self.log_renomeacao.insert("end", "📋 ETAPA 2: Validando e renomeando arquivos...\n"))
                
                sucessos_renomeacao = 0
                erros_renomeacao = 0
                
                # Coletar chaves que serão convertidas
                chaves_para_converter = set()
                for linha in linhas_com_dados:
                    chave_original = linha['chave'].get().strip()
                    if self.validar_chave_nfe(chave_original) and chave_original in self.chaves_xml:
                        chaves_para_converter.add(chave_original)
                
                # Processar renomeação
                for linha in linhas_com_dados:
                    chave_original = linha['chave'].get().strip()
                    novo_nome = linha['nome'].get().strip()
                    
                    if not self.validar_chave_nfe(chave_original):
                        self.root.after(0, lambda l=linha: l['status'].configure(text="❌ Chave inválida"))
                        erros_renomeacao += 1
                        continue
                    
                    if chave_original not in self.chaves_xml:
                        self.root.after(0, lambda l=linha: l['status'].configure(text="❌ XML não encontrado"))
                        erros_renomeacao += 1
                        continue
                    
                    if not novo_nome:
                        self.root.after(0, lambda l=linha: l['status'].configure(text="❌ Nome vazio"))
                        erros_renomeacao += 1
                        continue
                    
                    # Copiar e renomear arquivo
                    try:
                        arquivo_original = self.chaves_xml[chave_original]
                        nome_arquivo_saida = f"{novo_nome}.xml"
                        caminho_saida = os.path.join(pasta_renomeados, nome_arquivo_saida)
                        
                        shutil.move(arquivo_original, caminho_saida)
                        self.root.after(0, lambda l=linha: l['status'].configure(text="✅ Renomeado"))
                        sucessos_renomeacao += 1
                        
                    except Exception as e:
                        self.root.after(0, lambda l=linha, err=str(e): l['status'].configure(text=f"❌ Erro: {err[:20]}"))
                        erros_renomeacao += 1
                
                # Mover XMLs não renomeados
                xmls_a_mover = {k: v for k, v in self.chaves_xml.items() if k not in chaves_para_converter}
                for chave, arquivo in xmls_a_mover.items():
                    try:
                        nome_arquivo = os.path.basename(arquivo)
                        destino = os.path.join(pasta_nao_renomeados, nome_arquivo)
                        shutil.move(arquivo, destino)
                        # Remover da lista original para não ser processado novamente
                        del self.chaves_xml[chave]
                    except:
                        pass
            else:
                # Se não há dados na tabela, mover todos os XMLs para pasta não renomeados
                sucessos_renomeacao = 0
                erros_renomeacao = 0
                self.root.after(0, lambda: self.log_renomeacao.insert("end", "📋 ETAPA 2: Movendo todos os XMLs para pasta 'Arquivos Não Renomeados'...\n"))
                xmls_a_mover = self.chaves_xml.copy()
                for chave, arquivo in xmls_a_mover.items():
                    try:
                        nome_arquivo = os.path.basename(arquivo)
                        destino = os.path.join(pasta_nao_renomeados, nome_arquivo)
                        shutil.move(arquivo, destino)
                        del self.chaves_xml[chave]
                    except:
                        pass
            
            self.root.after(0, lambda: self.log_renomeacao.insert("end", f"✅ Renomeação: {sucessos_renomeacao} sucessos, {erros_renomeacao} erros\n\n"))
            
            # ETAPA 3: GERAÇÃO DE PDFs
            self.root.after(0, lambda: self.log_renomeacao.insert("end", "📋 ETAPA 3: Gerando DANFEs (PDFs)...\n"))
            
            # Gerar PDFs para arquivos renomeados e não renomeados
            xmls_renomeados = self.escanear_xmls_pasta(pasta_renomeados)
            xmls_nao_renomeados = self.escanear_xmls_pasta(pasta_nao_renomeados)
            
            sucessos_pdf = 0
            erros_pdf = 0
            
            def callback_sucesso_pdf(nome):
                self.log_renomeacao.insert("end", f"✅ PDF: {nome}\n")
                
            def callback_erro_pdf(nome):
                self.log_renomeacao.insert("end", f"❌ PDF: {nome}\n")
            
            # Processar PDFs dos arquivos renomeados
            if xmls_renomeados:
                self.root.after(0, lambda: self.log_renomeacao.insert("end", "📄 Gerando PDFs dos arquivos renomeados...\n"))
                suc_ren, err_ren, _ = self.processar_xmls_paralelo(
                    xmls_renomeados, pasta_pdf_convertido, callback_sucesso_pdf, callback_erro_pdf
                )
                sucessos_pdf += suc_ren
                erros_pdf += err_ren
            

            
            self.root.after(0, lambda: self.log_renomeacao.insert("end", f"\n✅ PDFs gerados: {sucessos_pdf}, erros: {erros_pdf}\n"))
            
            # RESULTADO FINAL
            tempo_total = time.time() - inicio_total
            
            self.root.after(0, lambda: self.log_renomeacao.insert("end", f"\n🎉 PROCESSAMENTO COMPLETO FINALIZADO!\n"))
            self.root.after(0, lambda: self.log_renomeacao.insert("end", f"📊 XMLs escaneados: {len(todos_xmls)}\n"))
            if linhas_com_dados:
                self.root.after(0, lambda: self.log_renomeacao.insert("end", f"📝 Arquivos renomeados: {sucessos_renomeacao}\n"))
            if 'sucessos_pdf' in locals():
                self.root.after(0, lambda: self.log_renomeacao.insert("end", f"📄 PDFs gerados: {sucessos_pdf}\n"))
            self.root.after(0, lambda: self.log_renomeacao.insert("end", f"⏱️ Tempo total: {tempo_total:.1f} segundos\n"))
            
            # Mostrar resultado
            if sucessos_pdf > 0:
                resposta = messagebox.askyesno(
                    "Processamento Completo!", 
                    f"🎉 Processamento finalizado com sucesso!\n\n"
                    f"📊 XMLs processados: {len(todos_xmls)}\n"
                    f"📝 Renomeados: {sucessos_renomeacao if linhas_com_dados else 0}\n"
                    f"📄 PDFs gerados: {sucessos_pdf}\n"
                    f"⏱️ Tempo: {tempo_total:.1f}s\n\n"
                    f"📁 Arquivos organizados em:\n"
                    f"   • Arquivos Renomeados\n"

                    f"   • PDF convertido\n\n"
                    f"Deseja abrir a pasta com os resultados?"
                )
                
                if resposta:
                    try:
                        os.startfile(pasta_xml)
                    except:
                        webbrowser.open(pasta_xml)
            else:
                messagebox.showinfo(
                    "Processamento Concluído",
                    f"✅ Processamento finalizado!\n\n"
                    f"📊 XMLs processados: {len(todos_xmls)}\n"
                    f"📝 Renomeados: {sucessos_renomeacao if linhas_com_dados else 0}\n"
                    f"⏱️ Tempo: {tempo_total:.1f}s\n\n"
                    f"📁 Arquivos organizados em:\n"
                    f"   • Arquivos Renomeados\n"

                    f"   • PDF convertido"
                )
        
        except Exception as e:
            error_msg = f"Erro durante processamento completo: {str(e)}"
            messagebox.showerror("Erro no Processamento", error_msg)
            self.root.after(0, lambda: self.log_renomeacao.insert("end", f"\n❌ ERRO: {error_msg}\n"))
        
        finally:
            # Reabilitar botões
            self.root.after(0, lambda: self.btn_processar_completo.configure(state="normal", text="🚀 PROCESSAR COMPLETO"))

    
            self.root.after(0, lambda: self.log_renomeacao.see("end"))
            
            self.processando = False
        
    # ============= FUNÇÕES AUXILIARES (ELIMINAM DUPLICAÇÕES) =============
    
    def escanear_xmls_pasta(self, pasta):
        """Função auxiliar para escanear XMLs de uma pasta (elimina duplicação)"""
        arquivos_xml = []
        try:
            for arquivo in os.listdir(pasta):
                if arquivo.lower().endswith('.xml'):
                    caminho_completo = os.path.join(pasta, arquivo)
                    arquivos_xml.append(caminho_completo)
        except Exception as e:
            print(f"Erro ao escanear pasta: {e}")
        return arquivos_xml

    def validar_chave_nfe(self, chave):
        """Função auxiliar para validar chave NFe (elimina duplicação)"""
        chave = str(chave).strip()
        return len(chave) == 44 and chave.isdigit()

    def executar_thread_segura(self, target_func):
        """Função auxiliar para threading (elimina duplicação)"""
        thread = threading.Thread(target=target_func)
        thread.daemon = True
        thread.start()

    def processar_xmls_paralelo(self, arquivos_xml, pasta_saida, callback_sucesso=None, callback_erro=None):
        """Função auxiliar para processamento paralelo (elimina duplicação)"""
        sucessos = 0
        erros = 0
        inicio = time.time()
        
        with ThreadPoolExecutor(max_workers=5) as executor:
            futures = {
                executor.submit(self.processar_xml_individual, arquivo, pasta_saida): arquivo 
                for arquivo in arquivos_xml
            }
            
            for future in as_completed(futures):
                arquivo = futures[future]
                nome_arquivo = os.path.basename(arquivo)
                
                try:
                    resultado = future.result()
                    if resultado:
                        sucessos += 1
                        if callback_sucesso:
                            self.root.after(0, lambda n=nome_arquivo: callback_sucesso(n))
                    else:
                        erros += 1
                        if callback_erro:
                            self.root.after(0, lambda n=nome_arquivo: callback_erro(n))
                except Exception as e:
                    erros += 1
                    if callback_erro:
                        self.root.after(0, lambda n=nome_arquivo, e=str(e): callback_erro(f"{n} - ERRO: {e}"))
        
        tempo_total = time.time() - inicio
        return sucessos, erros, tempo_total

    def mostrar_conclusao_processamento(self, sucessos, erros, tempo_total, pasta_saida):
        """Função auxiliar para mostrar conclusão (elimina duplicação)"""
        if sucessos > 0:
            resposta = messagebox.askyesno(
                "Processamento Concluído!", 
                f"✅ {sucessos} DANFEs geradas\n❌ {erros} erros\n\nDeseja abrir a pasta com os PDFs?"
            )
            
            if resposta:
                try:
                    os.startfile(pasta_saida)
                except:
                    webbrowser.open(pasta_saida)

    # ============= FUNÇÕES ORIGINAIS (REFATORADAS) =============

    def carregar_log_inicial(self):
        log_inicial = """  renamerPRO©
        
🔹 Sistema inicializado com sucesso
🔹 Aguardando configuração de pastas..."""

        self.log_text.delete("0.0", "end")
        self.log_text.insert("0.0", log_inicial)
        
    def selecionar_pasta_xml(self):
        try:
            pasta = filedialog.askdirectory(
                title="Selecione a pasta com os XMLs",
                parent=self.root
            )
            if pasta:
                self.pasta_xml.set(pasta)
                self.status_texto.set(f"Pasta XML selecionada: {os.path.basename(pasta)}")
                self.btn_escanear.configure(state="normal")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir seletor de pasta: {str(e)}")
            print(f"Erro no seletor de pasta XML: {e}")
                        

            
    def escanear_pasta(self):
        if not self.pasta_xml.get():
            messagebox.showerror("Erro", "Selecione a pasta com XMLs primeiro!")
            return
            
        pasta = self.pasta_xml.get()
        
        # Usar função auxiliar (elimina duplicação)
        self.arquivos_xml = self.escanear_xmls_pasta(pasta)
                
        total = len(self.arquivos_xml)
        
        if total == 0:
            self.label_arquivos.configure(text="❌ Nenhum arquivo XML encontrado na pasta!")
            self.btn_processar.configure(state="disabled")
            messagebox.showwarning("Aviso", "Nenhum arquivo XML encontrado na pasta selecionada!")
        else:
            self.label_arquivos.configure(text=f"✅ {total} arquivo(s) XML encontrado(s)")
            self.btn_processar.configure(state="normal")
            
            self.adicionar_log(f"\n🔍 ESCANEAMENTO CONCLUÍDO:")
            self.adicionar_log(f"📁 Pasta: {pasta}")
            self.adicionar_log(f"📊 Total de XMLs: {total}")
            
            self.adicionar_log(f"\n📄 Arquivos encontrados:")
            for i, arquivo in enumerate(self.arquivos_xml[:5]):
                nome = os.path.basename(arquivo)
                self.adicionar_log(f"  {i+1}. {nome}")
            
            if total > 5:
                self.adicionar_log(f"  ... e mais {total-5} arquivo(s)")
                
            self.adicionar_log(f"\n✅ Pronto para processar! Clique em 'PROCESSAR TODOS'")
            
    def adicionar_log(self, texto):
        self.log_text.insert("end", texto + "\n")
        self.log_text.see("end")
        
    def processar_massa_thread(self):
        if self.processando:
            return
            
        # Usar função auxiliar (elimina duplicação)
        self.executar_thread_segura(self.processar_massa)
        
    def processar_massa(self):
        if not self.arquivos_xml:
            messagebox.showerror("Erro", "Escaneie a pasta primeiro!")
            return
            
        self.processando = True
        total = len(self.arquivos_xml)
        sucessos = 0
        erros = 0
        
        pasta_saida = self.pasta_xml.get()
        
        self.root.after(0, lambda: self.btn_processar.configure(state="disabled", text="🔄 Processando..."))
        self.root.after(0, lambda: self.btn_escanear.configure(state="disabled"))
        self.root.after(0, lambda: self.status_texto.set("Processando XMLs em massa..."))
        
        self.root.after(0, lambda: self.adicionar_log(f"\n🚀 INICIANDO PROCESSAMENTO EM MASSA:"))
        self.root.after(0, lambda: self.adicionar_log(f"📊 Total: {total} arquivos"))
        self.root.after(0, lambda: self.adicionar_log(f"📤 Pasta saída: {pasta_saida}"))
        self.root.after(0, lambda: self.adicionar_log(f"⚡ Processamento paralelo ativado (5 simultâneos)\n"))
        
        inicio = time.time()
        
        # Processar com função auxiliar (elimina duplicação)
        def callback_sucesso(nome):
            nonlocal sucessos
            sucessos += 1
            self.adicionar_log(f"✅ {nome}")
            self.progresso_geral.set(sucessos / total)
            self.label_progresso.configure(text=f"{sucessos} / {total} arquivos processados")
            
        def callback_erro(nome):
            nonlocal erros
            erros += 1
            self.adicionar_log(f"❌ {nome}")
            processados = sucessos + erros
            self.progresso_geral.set(processados / total)
            self.label_progresso.configure(text=f"{processados} / {total} arquivos processados")
            
        sucessos, erros, tempo_total = self.processar_xmls_paralelo(
            self.arquivos_xml, pasta_saida, callback_sucesso, callback_erro
        )
        
        self.root.after(0, lambda: self.adicionar_log(f"\n🎉 PROCESSAMENTO CONCLUÍDO!"))
        self.root.after(0, lambda: self.adicionar_log(f"✅ Sucessos: {sucessos}"))
        self.root.after(0, lambda: self.adicionar_log(f"❌ Erros: {erros}"))
        self.root.after(0, lambda: self.adicionar_log(f"⏱️ Tempo total: {tempo_total:.1f} segundos"))
        self.root.after(0, lambda: self.adicionar_log(f"⚡ Média: {tempo_total/total:.1f}s por arquivo"))
        
        self.root.after(0, lambda: self.btn_processar.configure(state="normal", text="🎯 PROCESSAR TODOS"))
        self.root.after(0, lambda: self.btn_escanear.configure(state="normal"))
        self.root.after(0, lambda: self.status_texto.set(f"✅ Concluído: {sucessos} sucessos, {erros} erros"))
        
        # Usar função auxiliar (elimina duplicação)
        self.mostrar_conclusao_processamento(sucessos, erros, tempo_total, pasta_saida)
        
        self.processando = False

    def processar_xml_individual(self, arquivo_xml, pasta_saida):
        try:
            # Verificar se o PDF já existe para evitar reprocessamento
            nome_base_xml = os.path.splitext(os.path.basename(arquivo_xml))[0]
            caminho_pdf_final = os.path.join(pasta_saida, f"{nome_base_xml}.pdf")

            if os.path.exists(caminho_pdf_final):
                self.root.after(0, lambda: self.log_renomeacao.insert("end", f"✅ PDF já existe, pulando: {os.path.basename(caminho_pdf_final)}\n"))
                return True # Considerar como sucesso, pois o arquivo já está lá

            # Verificar se arquivo XML existe
            if not os.path.exists(arquivo_xml):
                self.adicionar_log(f"❌ Arquivo não encontrado: {arquivo_xml}")
                return False
            
            # Verificar se pasta de saída existe
            if not os.path.exists(pasta_saida):
                try:
                    os.makedirs(pasta_saida, exist_ok=True)
                except Exception as e:
                    self.adicionar_log(f"❌ Erro ao criar pasta: {pasta_saida} - {str(e)}")
                    return False
            
            # Usar PHP para gerar DANFE
            # Obter o diretório do script atual
            script_dir = os.path.dirname(os.path.abspath(__file__))
            php_full_path = os.path.join(script_dir, "php", "php.exe")
            script_php_full = os.path.join(script_dir, "gerador_danfe.php")
            
            # Verificar se arquivos existem
            if not os.path.exists(php_full_path):
                self.adicionar_log(f"❌ PHP não encontrado: {php_full_path}")
                return False
            
            if not os.path.exists(script_php_full):
                self.adicionar_log(f"❌ Script PHP não encontrado: {script_php_full}")
                return False
            
            # Comando para executar PHP (usar caminhos absolutos)
            cmd = [php_full_path, script_php_full, arquivo_xml]
            
                        # Executar PHP com melhor tratamento de erro
            try:
                # Executar do diretório php para carregar extensões
                php_dir = os.path.join(script_dir, "php")
                resultado = subprocess.run(
                    cmd, 
                    capture_output=True, 
                    text=True, 
                    timeout=120,
                    creationflags=subprocess.CREATE_NO_WINDOW,
                    cwd=php_dir  # Executar do diretório php para carregar extensões
                )
            except FileNotFoundError:
                self.adicionar_log(f"❌ PHP executável não encontrado: {php_full_path}")
                return False
            except Exception as e:
                self.adicionar_log(f"❌ Erro ao executar PHP: {str(e)}")
                return False
                        
            # Verificar resultado
            if resultado.returncode == 0 and "SUCCESS:" in resultado.stdout:
                arquivo_pdf = resultado.stdout.strip().replace("SUCCESS:", "")
                
                # Verificar se PDF foi criado
                if not os.path.exists(arquivo_pdf):
                    self.adicionar_log(f"❌ PDF não foi criado: {arquivo_pdf}")
                    return False
                
                # Mover PDF para pasta de saída se necessário
                if pasta_saida != os.path.dirname(arquivo_xml):
                    nome_pdf = os.path.basename(arquivo_pdf)
                    novo_caminho = os.path.join(pasta_saida, nome_pdf)
                    
                    try:
                        # Se arquivo já existe na pasta de destino, removê-lo
                        if os.path.exists(novo_caminho):
                            os.remove(novo_caminho)
                        
                        os.rename(arquivo_pdf, novo_caminho)
                        nome_pdf = os.path.basename(novo_caminho)
                    except Exception as e:
                        self.adicionar_log(f"❌ Erro ao mover PDF: {str(e)}")
                        return False
                else:
                    nome_pdf = os.path.basename(arquivo_pdf)
                
                self.adicionar_log(f"✅ {nome_pdf}")
                return True
            else:
                # Log do erro detalhado
                error_msg = resultado.stderr.strip() if resultado.stderr else "Erro desconhecido"
                stdout_msg = resultado.stdout.strip() if resultado.stdout else ""
                
                if "ERROR:" in stdout_msg:
                    error_msg = stdout_msg.replace("ERROR:", "").strip()
                
                # Se não há saída, pode ser problema com dependências
                if not stdout_msg and not error_msg:
                    error_msg = f"PHP erro código {resultado.returncode}"
                
                self.adicionar_log(f"❌ {os.path.basename(arquivo_xml)} - {error_msg}")
                return False
                
        except subprocess.TimeoutExpired:
            self.adicionar_log(f"❌ Timeout ao processar: {os.path.basename(arquivo_xml)}")
            return False
        except Exception as e:
            self.adicionar_log(f"❌ Erro inesperado: {str(e)}")
            return False
            
    def abrir_janela_lote(self):
        # Criar janela popup
        self.janela_lote = ctk.CTkToplevel(self.root)
        self.janela_lote.title("📋 Adicionar Lote de Dados")
        self.janela_lote.geometry("1200x900")
        self.janela_lote.minsize(1000, 600)
        self.janela_lote.transient(self.root)
        self.janela_lote.grab_set()
        
        # Instrução compacta
        instrucao = ctk.CTkLabel(
            self.janela_lote,
            text="Preencha os campos abaixo. Use uma linha por registro:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        instrucao.pack(pady=15)
        
        # Frame principal
        main_frame = ctk.CTkFrame(self.janela_lote)
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Frame das colunas
        colunas_frame = ctk.CTkFrame(main_frame)
        colunas_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Coluna esquerda - Chaves
        frame_chaves = ctk.CTkFrame(colunas_frame)
        frame_chaves.pack(side="left", fill="both", expand=True, padx=(10, 5), pady=10)
        
        ctk.CTkLabel(
            frame_chaves,
            text="🔑 Chave Acesso NF",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(pady=(10, 5))
        
        ctk.CTkLabel(
            frame_chaves,
            text="(44 dígitos numéricos)",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        ).pack(pady=(0, 10))
        
        self.textbox_chaves = ctk.CTkTextbox(
            frame_chaves,
            font=ctk.CTkFont(size=11),
            height=500
        )
        self.textbox_chaves.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # Coluna direita - Nomes
        frame_nomes = ctk.CTkFrame(colunas_frame)
        frame_nomes.pack(side="right", fill="both", expand=True, padx=(5, 10), pady=10)
        
        ctk.CTkLabel(
            frame_nomes,
            text="📄 Nome Arq. NF",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(pady=(10, 5))
        
        ctk.CTkLabel(
            frame_nomes,
            text="(Nome desejado para o arquivo)",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        ).pack(pady=(0, 10))
        
        self.textbox_nomes = ctk.CTkTextbox(
            frame_nomes,
            font=ctk.CTkFont(size=11),
            height=500
        )
        self.textbox_nomes.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # Botões
        botoes_frame = ctk.CTkFrame(main_frame)
        botoes_frame.pack(fill="x", padx=10, pady=10)
        
        btn_limpar = ctk.CTkButton(
            botoes_frame,
            text="🧹 Limpar",
            command=self.limpar_lote,
            width=140,
            height=45,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#808080",
            hover_color="#696969"
        )
        btn_limpar.pack(side="left", padx=(20, 15), pady=15)
        
        btn_processar = ctk.CTkButton(
            botoes_frame,
            text="✅ Processar Lote",
            command=self.processar_lote_dados,
            width=170,
            height=45,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#2E8B57",
            hover_color="#228B22"
        )
        btn_processar.pack(side="right", padx=(15, 15), pady=15)
        
        btn_cancelar = ctk.CTkButton(
            botoes_frame,
            text="❌ Cancelar",
            command=self.fechar_janela_lote,
            width=140,
            height=45,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#DC143C",
            hover_color="#B22222"
        )
        btn_cancelar.pack(side="right", padx=15, pady=15)
    
    def limpar_lote(self):
        self.textbox_chaves.delete("0.0", "end")
        self.textbox_nomes.delete("0.0", "end")
        self.textbox_chaves.focus()
    
    def fechar_janela_lote(self):
        self.janela_lote.destroy()
    
    def processar_lote_dados(self):
        try:
            # Obter textos das textboxes
            texto_chaves = self.textbox_chaves.get("0.0", "end").strip()
            texto_nomes = self.textbox_nomes.get("0.0", "end").strip()
            
            if not texto_chaves or not texto_nomes:
                messagebox.showwarning("Aviso", "Preencha ambos os campos!")
                return
            
            # Dividir em linhas
            linhas_chaves = [linha.strip() for linha in texto_chaves.split('\n') if linha.strip()]
            linhas_nomes = [linha.strip() for linha in texto_nomes.split('\n') if linha.strip()]
            
            # Validar quantidade de linhas
            if len(linhas_chaves) != len(linhas_nomes):
                messagebox.showerror("Erro", f"Quantidade de linhas diferente!\n\nChaves: {len(linhas_chaves)} linhas\nNomes: {len(linhas_nomes)} linhas\n\nCada chave deve ter um nome correspondente.")
                return
            
            # Validar chaves (usando função auxiliar - elimina duplicação)
            chaves_validas = []
            for i, chave_original in enumerate(linhas_chaves):
                chave = chave_original.strip()
                if not self.validar_chave_nfe(chave):
                    messagebox.showerror("Erro", f"Chave inválida na linha {i+1}:\n{chave_original}\n\nChaves devem ter exatamente 44 dígitos numéricos.")
                    return
                chaves_validas.append(chave)
            
            # Limpar tabela atual
            for linha in self.linhas_renomeacao:
                linha['frame'].destroy()
            self.linhas_renomeacao.clear()
            
            # Adicionar dados processados
            for chave, nome in zip(chaves_validas, linhas_nomes):
                self.adicionar_linha_renomeacao()
                ultima_linha = self.linhas_renomeacao[-1]
                
                # Preencher campos
                ultima_linha['chave'].delete(0, 'end')
                ultima_linha['chave'].insert(0, chave)
                ultima_linha['nome'].delete(0, 'end')
                ultima_linha['nome'].insert(0, nome)
            
            # Log
            total_processado = len(chaves_validas)
            self.log_renomeacao.insert("end", f"\n📋 LOTE DE DADOS PROCESSADO:\n")
            self.log_renomeacao.insert("end", f"✅ {total_processado} registros adicionados\n")
            self.log_renomeacao.insert("end", "💡 Use 'PROCESSAR COMPLETO' para validar\n")
            self.log_renomeacao.see("end")
            
            # Fechar janela
            self.janela_lote.destroy()
            
            messagebox.showinfo("Sucesso!", f"✅ {total_processado} registros adicionados com sucesso!\n\nAgora escaneie as chaves para validar.")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar lote: {str(e)}")

    def executar(self):
        # Configurar controles de janela
        self.root.protocol("WM_DELETE_WINDOW", self.fechar_aplicacao)
        
        # Configurar título
        self.root.title("⚕️ renamerPRO©")
        
        print("🖥️ Aplicação iniciada")
        
        self.root.mainloop()
    
    def fechar_aplicacao(self):
        """Fecha a aplicação com cleanup adequado"""
        try:
            # Fechar aplicação
            self.root.quit()
            print("👋 Aplicação fechada com sucesso")
            
        except Exception as e:
            print(f"❌ Erro ao fechar aplicação: {e}")
            self.root.quit()

if __name__ == "__main__":
    app = DanfeAppMassa()
    app.executar()