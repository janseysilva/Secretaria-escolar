"""App de Secretaria Escolar - gera memorandos, ofícios e outros documentos."""

import json
import os
import shutil
import sys
import threading
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import documentos
import relatorio_bolsa_familia

PASTA_BASE = os.path.dirname(os.path.abspath(__file__))
PASTA_DADOS = os.path.join(PASTA_BASE, "dados_escola")
PASTA_IMAGENS = os.path.join(PASTA_DADOS, "imagens")
ARQUIVO_CONFIG = os.path.join(PASTA_DADOS, "config.json")

os.makedirs(PASTA_IMAGENS, exist_ok=True)

NOMES_MESES = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
]

COR_FUNDO = "#EEF1F7"
COR_CARTAO = "#FFFFFF"
COR_AZUL = "#1D9E75"
COR_AZUL_ESCURO = "#0F6E56"
COR_AZUL_CLARO = "#5DCAA5"
COR_TEXTO = "#1F2937"
COR_TEXTO_FRACO = "#6B7280"
COR_BORDA = "#D6DCE8"
COR_VERDE = "#1E8E4E"

FONTE_PADRAO = ("Segoe UI", 10)
FONTE_LABEL = ("Segoe UI", 10, "bold")


def carregar_config():
    if os.path.exists(ARQUIVO_CONFIG):
        with open(ARQUIVO_CONFIG, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def salvar_config(dados):
    with open(ARQUIVO_CONFIG, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)


MESES_PT = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
            "agosto", "setembro", "outubro", "novembro", "dezembro"]


def data_por_extenso(cidade="", data=None):
    d = data or datetime.date.today()
    texto = f"{d.day} de {MESES_PT[d.month - 1]} de {d.year}"
    return f"{cidade}, {texto}" if cidade else texto


def copiar_imagem_para_projeto(caminho_origem, nome_destino):
    if not caminho_origem:
        return None
    _, ext = os.path.splitext(caminho_origem)
    destino = os.path.join(PASTA_IMAGENS, f"{nome_destino}{ext}")
    shutil.copyfile(caminho_origem, destino)
    return destino


class CampoTexto:
    """Label + Entry, empacotados verticalmente."""

    def __init__(self, pai, rotulo, valor_inicial="", largura=50):
        self.frame = tk.Frame(pai, bg=COR_CARTAO)
        tk.Label(self.frame, text=rotulo, bg=COR_CARTAO, fg=COR_TEXTO, font=FONTE_LABEL,
                 justify="left").pack(anchor="w")
        self.var = tk.StringVar(value=valor_inicial)
        self.entry = tk.Entry(self.frame, textvariable=self.var, font=FONTE_PADRAO,
                               width=largura, relief="solid", bd=1,
                               highlightthickness=1, highlightbackground=COR_BORDA)
        self.entry.pack(anchor="w", fill="x", pady=(3, 10))

    def get(self):
        return self.var.get().strip()

    def set(self, valor):
        self.var.set(valor or "")

    def pack(self, **kwargs):
        self.frame.pack(**kwargs)


class CampoOpcao:
    """Label + lista de opcoes (dropdown). Cada opcao e (valor, rotulo) -
    get() retorna o valor (usado no codigo), o usuario ve o rotulo."""

    def __init__(self, pai, rotulo, opcoes, largura=35):
        self.frame = tk.Frame(pai, bg=COR_CARTAO)
        tk.Label(self.frame, text=rotulo, bg=COR_CARTAO, fg=COR_TEXTO, font=FONTE_LABEL,
                 justify="left").pack(anchor="w")
        self.opcoes = opcoes
        self.var = tk.StringVar(value=opcoes[0][1] if opcoes else "")
        self.combo = ttk.Combobox(self.frame, textvariable=self.var, font=FONTE_PADRAO,
                                   width=largura, state="readonly",
                                   values=[rotulo for _valor, rotulo in opcoes])
        self.combo.pack(anchor="w", fill="x", pady=(3, 10))

    def get(self):
        rotulo_atual = self.var.get()
        for valor, rotulo in self.opcoes:
            if rotulo == rotulo_atual:
                return valor
        return ""

    def pack(self, **kwargs):
        self.frame.pack(**kwargs)


class CampoTipoCodigo:
    """Label + Radiobuttons (SIGEAM / Matricula / Outros), com uma caixa de
    texto ao lado de "Outros" pra digitar um rotulo customizado. get()
    retorna o rotulo que deve aparecer entre parenteses no documento."""

    def __init__(self, pai, rotulo):
        self.frame = tk.Frame(pai, bg=COR_CARTAO)
        tk.Label(self.frame, text=rotulo, bg=COR_CARTAO, fg=COR_TEXTO, font=FONTE_LABEL,
                 justify="left").pack(anchor="w")
        linha = tk.Frame(self.frame, bg=COR_CARTAO)
        linha.pack(anchor="w", fill="x", pady=(3, 10))

        self.var = tk.StringVar(value="sigeam")
        for valor, texto in (("sigeam", "SIGEAM"), ("matricula", "Matrícula"), ("outros", "Outros")):
            tk.Radiobutton(linha, text=texto, variable=self.var, value=valor,
                           bg=COR_CARTAO, fg=COR_TEXTO, selectcolor=COR_CARTAO,
                           font=FONTE_PADRAO, activebackground=COR_CARTAO,
                           command=self._atualizar_estado_outros).pack(side="left", padx=(0, 12))

        self.var_outros = tk.StringVar()
        self.entry_outros = tk.Entry(linha, textvariable=self.var_outros, font=FONTE_PADRAO,
                                      width=18, relief="solid", bd=1,
                                      highlightthickness=1, highlightbackground=COR_BORDA,
                                      state="disabled")
        self.entry_outros.pack(side="left")
        self._atualizar_estado_outros()

    def _atualizar_estado_outros(self):
        self.entry_outros.config(state="normal" if self.var.get() == "outros" else "disabled")

    def get(self):
        valor = self.var.get()
        if valor == "sigeam":
            return "SIGEAM"
        if valor == "matricula":
            return "Matrícula"
        return self.var_outros.get().strip()

    def pack(self, **kwargs):
        self.frame.pack(**kwargs)


class CampoImagem:
    """Label + botao de escolher arquivo + nome do arquivo escolhido."""

    def __init__(self, pai, rotulo, valor_inicial=None):
        self.frame = tk.Frame(pai, bg=COR_CARTAO)
        tk.Label(self.frame, text=rotulo, bg=COR_CARTAO, fg=COR_TEXTO, font=FONTE_LABEL,
                 justify="left").pack(anchor="w")
        linha = tk.Frame(self.frame, bg=COR_CARTAO)
        linha.pack(anchor="w", fill="x", pady=(3, 10))
        self.caminho = valor_inicial
        self.label_arquivo = tk.Label(linha, text=self._texto_arquivo(), bg=COR_CARTAO,
                                       fg=COR_TEXTO_FRACO, font=FONTE_PADRAO)
        self.label_arquivo.pack(side="left", padx=(0, 10))
        btn = tk.Button(linha, text="Escolher imagem...", command=self._escolher,
                         bg="#E3F2ED", fg=COR_AZUL_ESCURO, font=("Segoe UI", 9),
                         activebackground="#CFEAE0", activeforeground=COR_AZUL_ESCURO,
                         relief="flat", padx=10, pady=3, cursor="hand2")
        btn.pack(side="left", padx=(0, 6))
        btn_remover = tk.Button(linha, text="Remover", command=self._remover,
                                 bg="#FBE5E5", fg="#B3261E", font=("Segoe UI", 9),
                                 activebackground="#F5C4C4", activeforeground="#B3261E",
                                 relief="flat", padx=10, pady=3, cursor="hand2")
        btn_remover.pack(side="left")

    def _texto_arquivo(self):
        if self.caminho and os.path.exists(self.caminho):
            return os.path.basename(self.caminho)
        return "Nenhuma imagem selecionada"

    def _escolher(self):
        caminho = filedialog.askopenfilename(
            title="Escolher imagem",
            filetypes=[("Imagens", "*.png *.jpg *.jpeg"), ("Todos os arquivos", "*.*")])
        if caminho:
            self.caminho = caminho
            self.label_arquivo.config(text=self._texto_arquivo())

    def _remover(self):
        self.caminho = None
        self.label_arquivo.config(text=self._texto_arquivo())

    def get(self):
        return self.caminho

    def pack(self, **kwargs):
        self.frame.pack(**kwargs)


DOCUMENTOS_DISPONIVEIS = [
    ("memorando", "📄", "Memorando", True, "#185FA5"),
    ("oficio", "📋", "Ofício", True, "#0F6E56"),
    ("declaracao", "📃", "Declaração Escolar", True, "#534AB7"),
    ("bolsa", "🎓", "Relatório do Bolsa Família", True, "#854F0B"),
    ("capa_livro", "📚", "Capa de Abertura de Livro", True, "#993C1D"),
    ("ficha_matricula", "📝", "Ficha de Matrícula", True, "#993556"),
    ("lista_reuniao", "👥", "Lista de Reunião de Pais e Alunos", True, "#3B6D11"),
    ("lista_frequencia", "🗓️", "Lista de Frequência Escolar", True, "#0C447C"),
    ("justificativa_faltas", "✍️", "Justificativa de Excesso de Faltas", True, "#A32D2D"),
    ("certificado", "🏅", "Certificado de Conclusão", True, "#BA7517"),
    ("historico_escolar", "📖", "Histórico Escolar", True, "#5F5E5A"),
]


def _clarear_cor(hex_cor, fator=0.15):
    """Clareia uma cor hex misturando com branco - usado pro efeito de
    hover nos cartoes coloridos (fill fica um pouco mais claro)."""
    hex_cor = hex_cor.lstrip("#")
    r, g, b = int(hex_cor[0:2], 16), int(hex_cor[2:4], 16), int(hex_cor[4:6], 16)
    r = int(r + (255 - r) * fator)
    g = int(g + (255 - g) * fator)
    b = int(b + (255 - b) * fator)
    return f"#{r:02x}{g:02x}{b:02x}"


def _retangulo_arredondado(canvas, x1, y1, x2, y2, raio=12, **kwargs):
    """Desenha um retangulo com cantos arredondados num Canvas - tkinter
    nao tem isso nativamente, esse e o jeito padrao (poligono suavizado)."""
    pontos = [
        x1 + raio, y1, x2 - raio, y1, x2, y1, x2, y1 + raio,
        x2, y2 - raio, x2, y2, x2 - raio, y2, x1 + raio, y2,
        x1, y2, x1, y2 - raio, x1, y1 + raio, x1, y1,
    ]
    return canvas.create_polygon(pontos, smooth=True, **kwargs)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Secretaria Escolar")
        self.geometry("880x720")
        self.configure(bg=COR_FUNDO)

        self.config_escola = carregar_config()

        self._montar_topo()
        self._montar_rodape()
        self._montar_paginas()

        if self.config_escola.get("nome_escola"):
            self._ir_para("home")
        else:
            self._ir_para("escola", primeira_vez=True)

    def _montar_topo(self):
        faixa = tk.Frame(self, bg=COR_AZUL_ESCURO, height=64)
        faixa.pack(fill="x", side="top")
        faixa.pack_propagate(False)
        tk.Label(faixa, text="Secretaria Escolar", bg=COR_AZUL_ESCURO, fg="white",
                 font=("Segoe UI", 15, "bold")).pack(side="left", padx=20)
        tk.Label(faixa, text="Documentos oficiais gerados na hora", bg=COR_AZUL_ESCURO,
                 fg="#BFE6D9", font=("Segoe UI", 9)).pack(side="left")

    @staticmethod
    def _botao_contorno(pai, texto, comando, cor, cor_hover_fundo):
        """Botão com contorno fino colorido (moldura de 1px ao redor de um
        botão sem borda própria - tkinter não tem "border-color" de verdade
        num Button só, então a moldura é um Frame colorido por baixo)."""
        moldura = tk.Frame(pai, bg=cor)
        botao = tk.Button(moldura, text=texto, command=comando,
                           bg=COR_CARTAO, fg=cor, font=("Segoe UI", 9, "bold"),
                           activebackground=cor_hover_fundo, activeforeground=cor,
                           relief="flat", bd=0, padx=13, pady=5, cursor="hand2")
        botao.pack(padx=1, pady=1)
        return moldura

    def _montar_rodape(self):
        rodape = tk.Frame(self, bg=COR_CARTAO, height=44)
        rodape.pack(fill="x", side="bottom")
        rodape.pack_propagate(False)

        borda = tk.Frame(rodape, bg=COR_BORDA, height=1)
        borda.pack(fill="x", side="top")

        self._botao_contorno(
            rodape, "← Voltar ao início", lambda: self._ir_para("home"),
            COR_AZUL_ESCURO, "#E3F2ED",
        ).pack(side="left", padx=20, pady=7)
        self._botao_contorno(
            rodape, "⚙ Dados da escola", lambda: self._ir_para("escola"),
            "#8A8880", "#F1EFE8",
        ).pack(side="right", padx=20, pady=7)

    def _montar_paginas(self):
        self.container = tk.Frame(self, bg=COR_FUNDO)
        self.container.pack(fill="both", expand=True, padx=14, pady=14)
        self.container.rowconfigure(0, weight=1)
        self.container.columnconfigure(0, weight=1)

        self.paginas = {}
        for nome in ("home", "escola", "memorando", "oficio", "declaracao", "capa_livro", "ficha_matricula",
                     "lista_reuniao", "lista_frequencia", "bolsa", "justificativa_faltas", "certificado",
                     "historico_escolar"):
            frame = tk.Frame(self.container, bg=COR_FUNDO)
            frame.grid(row=0, column=0, sticky="nsew")
            self.paginas[nome] = frame

        self._montar_pagina_home()
        self._montar_pagina_escola()
        self._montar_pagina_memorando()
        self._montar_pagina_oficio()
        self._montar_pagina_declaracao()
        self._montar_pagina_capa_livro()
        self._montar_pagina_ficha_matricula()
        self._montar_pagina_lista_reuniao()
        self._montar_pagina_lista_frequencia()
        self._montar_pagina_bolsa()
        self._montar_pagina_justificativa_faltas()
        self._montar_pagina_certificado()
        self._montar_pagina_historico_escolar()

    def _ir_para(self, nome, primeira_vez=False):
        if nome == "home":
            self._atualizar_saudacao_home()
        if nome == "escola":
            if primeira_vez:
                self._banner_boas_vindas.pack(fill="x", pady=(0, 10), before=self.aba_escola)
            else:
                self._banner_boas_vindas.pack_forget()
        if nome == "oficio":
            self._atualizar_data_auto(self.campo_data_oficio, "_ultima_data_oficio_auto")
        if nome == "declaracao":
            self._atualizar_data_auto(self.campo_data_declaracao, "_ultima_data_declaracao_auto")
        if nome == "capa_livro":
            self._atualizar_data_auto(self.campo_data_capa_livro, "_ultima_data_capa_livro_auto")
        if nome == "justificativa_faltas":
            self._atualizar_data_auto(self.campo_data_justificativa, "_ultima_data_justificativa_auto")
        if nome == "certificado":
            self._atualizar_data_auto(self.campo_data_certificado, "_ultima_data_certificado_auto")
        if nome == "historico_escolar":
            self._atualizar_data_auto(self.campo_data_historico, "_ultima_data_historico_auto")
        self.paginas[nome].tkraise()

    def _atualizar_data_auto(self, campo, atributo_ultimo):
        # So reescreve se o campo ainda estiver com o ultimo valor que o
        # proprio app preencheu sozinho - se o usuario editou a mao, nao mexe.
        if campo.get() == getattr(self, atributo_ultimo, None):
            cidade = (self.config_escola.get("cidade") or "").strip()
            novo_valor = data_por_extenso(cidade)
            campo.set(novo_valor)
            setattr(self, atributo_ultimo, novo_valor)

    # ---------------- Início (escolher documento) ----------------
    def _montar_pagina_home(self):
        pagina = self.paginas["home"]

        canvas = tk.Canvas(pagina, bg=COR_FUNDO, highlightthickness=0)
        scrollbar = ttk.Scrollbar(pagina, orient="vertical", command=canvas.yview)
        conteudo = tk.Frame(canvas, bg=COR_FUNDO)

        conteudo.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        janela_canvas = canvas.create_window((0, 0), window=conteudo, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(janela_canvas, width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        cabecalho = tk.Frame(conteudo, bg=COR_FUNDO)
        cabecalho.pack(fill="x", pady=(4, 18))
        self.label_saudacao_home = tk.Label(
            cabecalho, text="", bg=COR_FUNDO, fg=COR_AZUL_ESCURO, font=("Segoe UI", 15, "bold"))
        self.label_saudacao_home.pack(anchor="w")
        tk.Label(cabecalho, text="Qual documento você quer gerar?", bg=COR_FUNDO,
                 fg=COR_TEXTO_FRACO, font=("Segoe UI", 10)).pack(anchor="w", pady=(2, 0))

        grade = tk.Frame(conteudo, bg=COR_FUNDO)
        grade.pack(fill="both", expand=True)
        colunas = 3
        for col in range(colunas):
            grade.columnconfigure(col, weight=1)

        for i, (chave, emoji, titulo, disponivel, cor) in enumerate(DOCUMENTOS_DISPONIVEIS):
            linha, col = divmod(i, colunas)
            self._criar_cartao_documento(grade, chave, emoji, titulo, disponivel, cor).grid(
                row=linha, column=col, sticky="nsew", padx=6, pady=6)

    def _criar_cartao_documento(self, pai, chave, emoji, titulo, disponivel, cor):
        cor_base = cor if disponivel else "#8A8880"
        cor_fundo = _clarear_cor(cor_base, 0.88)
        cor_badge = _clarear_cor(cor_base, 0.65)
        cor_fundo_hover = cor_badge  # bem mais escuro que cor_fundo, pra dar pra notar o efeito
        cor_sombra = _clarear_cor(cor_base, 0.45)
        cor_texto = cor_base if disponivel else "#5F5E5A"
        cor_status = _clarear_cor(cor_base, 0.15) if disponivel else "#888780"

        DESLOC_SOMBRA = 4  # deslocamento da camada de sombra atras do cartao - efeito "flutuando"

        canvas = tk.Canvas(pai, bg=COR_FUNDO, highlightthickness=0, height=128,
                            cursor="hand2" if disponivel else "arrow")
        estado = {"rect": None}

        def redesenhar(_e=None):
            canvas.delete("all")
            w = max(canvas.winfo_width(), 10)
            h = max(canvas.winfo_height(), 10)

            # camada solida atras, deslocada pra baixo/direita - tkinter nao
            # tem sombra desfocada nativa, esse "duplo retangulo" e o jeito
            # de simular um cartao flutuando sobre o fundo.
            _retangulo_arredondado(
                canvas, 1 + DESLOC_SOMBRA, 1 + DESLOC_SOMBRA, w - 1, h - 1, 12,
                fill=cor_sombra, outline="")
            estado["rect"] = _retangulo_arredondado(
                canvas, 1, 1, w - 1 - DESLOC_SOMBRA, h - 1 - DESLOC_SOMBRA, 12,
                fill=cor_fundo, outline="")

            raio_badge = 17
            cx, cy = 1 + 14 + raio_badge, 1 + 14 + raio_badge
            canvas.create_oval(cx - raio_badge, cy - raio_badge, cx + raio_badge, cy + raio_badge,
                                fill=cor_badge, outline="")
            canvas.create_text(cx, cy, text=emoji, font=("Segoe UI Emoji", 15), anchor="center")

            largura_titulo = (w - 1 - DESLOC_SOMBRA) - 14 - 12
            canvas.create_text(14, 14 + raio_badge * 2 + 10, text=titulo, font=("Segoe UI", 12, "bold"),
                                anchor="nw", fill=cor_texto, width=largura_titulo)
            canvas.create_text(14, h - 1 - DESLOC_SOMBRA - 12, text="Disponível" if disponivel else "Em breve",
                                font=("Segoe UI", 9, "bold"), anchor="sw", fill=cor_status)

        canvas.bind("<Configure>", redesenhar)

        if disponivel:
            def abrir(_e=None):
                self._ir_para(chave)
            def entrar(_e=None):
                if estado["rect"] is not None:
                    canvas.itemconfig(estado["rect"], fill=cor_fundo_hover)
            def sair(_e=None):
                if estado["rect"] is not None:
                    canvas.itemconfig(estado["rect"], fill=cor_fundo)
            canvas.bind("<Button-1>", abrir)
            canvas.bind("<Enter>", entrar)
            canvas.bind("<Leave>", sair)
        else:
            def avisar(_e=None):
                messagebox.showinfo(titulo, f"{titulo} ainda não está disponível — em breve!")
            canvas.bind("<Button-1>", avisar)
        return canvas

    def _atualizar_saudacao_home(self):
        nome_escola = self.config_escola.get("nome_escola", "").strip()
        self.label_saudacao_home.config(
            text=f"Bem-vindo(a), {nome_escola}!" if nome_escola else "Bem-vindo(a)!")

    # ---------------- Dados da Escola ----------------
    def _montar_pagina_escola(self):
        pagina = self.paginas["escola"]

        self._banner_boas_vindas = tk.Frame(pagina, bg="#FFF4CE", padx=16, pady=10)
        tk.Label(self._banner_boas_vindas,
                 text="Bem-vindo(a)! Antes de gerar qualquer documento, preencha os dados da sua escola abaixo.",
                 bg="#FFF4CE", fg="#7A5B00", font=("Segoe UI", 9, "bold"), justify="left").pack(anchor="w")

        aba_escola = tk.Frame(pagina, bg=COR_FUNDO)
        aba_escola.pack(fill="both", expand=True)
        self.aba_escola = aba_escola

        canvas = tk.Canvas(self.aba_escola, bg=COR_FUNDO, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.aba_escola, orient="vertical", command=canvas.yview)
        cartao_externo = tk.Frame(canvas, bg=COR_FUNDO)

        cartao_externo.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=cartao_externo, anchor="nw", width=820)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        cartao = tk.Frame(cartao_externo, bg=COR_CARTAO, padx=24, pady=20)
        cartao.pack(fill="x", padx=4, pady=4)

        tk.Label(cartao, text="Dados da escola", bg=COR_CARTAO, fg=COR_AZUL_ESCURO,
                 font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(0, 4))
        tk.Label(cartao, text="Preenchido uma vez só. Esses dados entram automaticamente em todo documento gerado.",
                 bg=COR_CARTAO, fg=COR_TEXTO_FRACO, font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 16))

        c = self.config_escola
        self.campo_nome_escola = CampoTexto(cartao, "Nome da escola", c.get("nome_escola", ""))
        self.campo_nome_escola.pack(fill="x")

        self.campo_secretaria = CampoTexto(cartao, "Secretaria/órgão superior (ex: SEMED) - opcional",
                                            c.get("secretaria", ""))
        self.campo_secretaria.pack(fill="x")

        self.campo_cidade = CampoTexto(cartao, "Cidade (usada na data do Ofício e da Declaração, ex: Manaus)",
                                        c.get("cidade", ""))
        self.campo_cidade.pack(fill="x")

        self.campo_endereco = CampoTexto(cartao, "Endereço da escola (aparece no cabeçalho de todo documento)",
                                          c.get("endereco", ""))
        self.campo_endereco.pack(fill="x")

        self.campo_telefone = CampoTexto(cartao, "Telefone da escola (aparece no cabeçalho de todo documento)",
                                          c.get("telefone", ""), largura=25)
        self.campo_telefone.pack(fill="x")

        self.campo_email = CampoTexto(cartao, "Email da escola (aparece no cabeçalho de todo documento)",
                                       c.get("email", ""))
        self.campo_email.pack(fill="x")

        self.campo_diretor_nome = CampoTexto(cartao, "Nome completo do(a) diretor(a)", c.get("diretor_nome", ""))
        self.campo_diretor_nome.pack(fill="x")

        self.campo_diretor_cargo = CampoTexto(
            cartao, "Cargo (ex: Diretor(a), Gestor(a) Escolar)", c.get("diretor_cargo", ""))
        self.campo_diretor_cargo.pack(fill="x")

        self.campo_diretor_portaria = CampoTexto(
            cartao, "Portaria de nomeação (opcional)", c.get("diretor_portaria", ""))
        self.campo_diretor_portaria.pack(fill="x")

        self.campo_secretario_nome = CampoTexto(
            cartao, "Nome completo do(a) secretário(a) escolar (opcional)", c.get("secretario_nome", ""))
        self.campo_secretario_nome.pack(fill="x")

        tk.Label(cartao, text="Imagem (opcional, mas deixa o documento com a cara oficial da escola)",
                 bg=COR_CARTAO, fg=COR_AZUL_ESCURO, font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(10, 10))

        self.campo_logo = CampoImagem(cartao, "Brasão/logotipo (aparece no topo do documento)",
                                       c.get("logo_path"))
        self.campo_logo.pack(fill="x")

        self.campo_assinatura = CampoImagem(
            cartao, "Assinatura digitalizada do(a) diretor(a)/gestor(a) (opcional)", c.get("assinatura_path"))
        self.campo_assinatura.pack(fill="x")

        tk.Label(cartao,
                 text="Recebido por e a segunda Data ficam em branco no documento, para\n"
                      "preencher à mão na hora de entregar.",
                 bg=COR_CARTAO, fg=COR_TEXTO_FRACO, font=("Segoe UI", 9), justify="left").pack(anchor="w", pady=(0, 10))

        btn_salvar = tk.Button(cartao, text="Salvar dados da escola", command=self._salvar_dados_escola,
                                bg=COR_AZUL, fg="white", font=("Segoe UI", 10, "bold"),
                                activebackground=COR_AZUL_ESCURO, activeforeground="white",
                                relief="flat", padx=18, pady=8, cursor="hand2")
        btn_salvar.pack(anchor="w", pady=(10, 0))

        self.label_status_escola = tk.Label(cartao, text="", bg=COR_CARTAO, fg=COR_VERDE, font=FONTE_PADRAO)
        self.label_status_escola.pack(anchor="w", pady=(8, 0))

    def _salvar_dados_escola(self):
        if not self.campo_nome_escola.get():
            messagebox.showwarning("Campo obrigatório", "Preencha o nome da escola.")
            return

        logo_path = self.campo_logo.get()
        if logo_path and PASTA_IMAGENS not in logo_path:
            logo_path = copiar_imagem_para_projeto(logo_path, "logo")

        assinatura_path = self.campo_assinatura.get()
        if assinatura_path and PASTA_IMAGENS not in assinatura_path:
            assinatura_path = copiar_imagem_para_projeto(assinatura_path, "assinatura")

        dados = {
            "nome_escola": self.campo_nome_escola.get(),
            "secretaria": self.campo_secretaria.get(),
            "cidade": self.campo_cidade.get(),
            "endereco": self.campo_endereco.get(),
            "telefone": self.campo_telefone.get(),
            "email": self.campo_email.get(),
            "diretor_nome": self.campo_diretor_nome.get(),
            "diretor_cargo": self.campo_diretor_cargo.get(),
            "diretor_portaria": self.campo_diretor_portaria.get(),
            "secretario_nome": self.campo_secretario_nome.get(),
            "logo_path": logo_path,
            "assinatura_path": assinatura_path,
        }
        salvar_config(dados)
        self.config_escola = dados
        self.label_status_escola.config(text="Dados salvos com sucesso.")
        self.after(600, lambda: self._ir_para("home"))

    # ---------------- Memorando ----------------
    def _montar_pagina_memorando(self):
        pagina = self.paginas["memorando"]

        canvas = tk.Canvas(pagina, bg=COR_FUNDO, highlightthickness=0)
        scrollbar = ttk.Scrollbar(pagina, orient="vertical", command=canvas.yview)
        cartao_externo = tk.Frame(canvas, bg=COR_FUNDO)

        cartao_externo.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        janela_canvas = canvas.create_window((0, 0), window=cartao_externo, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(janela_canvas, width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        cartao = tk.Frame(cartao_externo, bg=COR_CARTAO, padx=24, pady=20)
        cartao.pack(fill="x", padx=4, pady=4)

        tk.Label(cartao, text="Novo memorando", bg=COR_CARTAO, fg=COR_AZUL_ESCURO,
                 font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(0, 16))

        linha_numero = tk.Frame(cartao, bg=COR_CARTAO)
        linha_numero.pack(fill="x")
        self.campo_numero = CampoTexto(linha_numero, "Número do memorando", largura=15)
        self.campo_numero.pack(side="left", padx=(0, 20))
        self.campo_ano = CampoTexto(linha_numero, "Ano", str(datetime.date.today().year), largura=10)
        self.campo_ano.pack(side="left")

        self.campo_protocolo = CampoTexto(
            cartao, "Protocolo (ex: SIGED Nº 12345) - opcional, deixe em branco se não usar", largura=40)
        self.campo_protocolo.pack(fill="x")

        self.campo_para = CampoTexto(cartao, "Para (destinatário)", largura=60)
        self.campo_para.pack(fill="x")

        self.campo_assunto = CampoTexto(cartao, "Assunto", largura=70)
        self.campo_assunto.pack(fill="x")

        self.campo_saudacao = CampoTexto(cartao, "Saudação", "Prezado Chefe,", largura=40)
        self.campo_saudacao.pack(fill="x")

        tk.Label(cartao, text="Corpo do memorando (cada parágrafo em uma linha)",
                 bg=COR_CARTAO, fg=COR_TEXTO, font=FONTE_LABEL).pack(anchor="w")
        self.texto_corpo = tk.Text(cartao, height=10, font=FONTE_PADRAO, relief="solid", bd=1,
                                    highlightthickness=1, highlightbackground=COR_BORDA, wrap="word")
        self.texto_corpo.pack(fill="x", pady=(3, 10))
        self.texto_corpo.insert("1.0", "Venho, por meio deste, ")

        self.campo_fecho = CampoTexto(cartao, "Fecho", "Atenciosamente,", largura=30)
        self.campo_fecho.pack(fill="x")

        self.campo_assinado_por = CampoTexto(
            cartao,
            "Assinado por (opcional - só aparece na caixa ASSINATURA se a escola não tiver\n"
            "cadastrado uma imagem de assinatura em Dados da Escola)",
            largura=50)
        self.campo_assinado_por.pack(fill="x")

        self.campo_data = CampoTexto(cartao, "Data", datetime.date.today().strftime("%d/%m/%Y"), largura=15)
        self.campo_data.pack(fill="x")

        tk.Label(cartao,
                 text="Recebido por e a segunda Data saem em branco no documento, para\n"
                      "preencher à mão na hora de entregar.",
                 bg=COR_CARTAO, fg=COR_TEXTO_FRACO, font=("Segoe UI", 9), justify="left").pack(anchor="w", pady=(0, 10))

        btn_gerar = tk.Button(cartao, text="Gerar Memorando (.docx)", command=self._gerar_memorando,
                               bg=COR_AZUL, fg="white", font=("Segoe UI", 11, "bold"),
                               activebackground=COR_AZUL_ESCURO, activeforeground="white",
                               relief="flat", padx=20, pady=10, cursor="hand2")
        btn_gerar.pack(anchor="w", pady=(10, 0))

        self.label_status_memo = tk.Label(cartao, text="", bg=COR_CARTAO, fg=COR_VERDE, font=FONTE_PADRAO)
        self.label_status_memo.pack(anchor="w", pady=(8, 0))

    def _gerar_memorando(self):
        if not self.config_escola.get("nome_escola"):
            messagebox.showwarning(
                "Dados da escola pendentes",
                "Preencha e salve os Dados da Escola antes de gerar um memorando.")
            self._ir_para("escola")
            return

        numero = self.campo_numero.get()
        para = self.campo_para.get()
        assunto = self.campo_assunto.get()
        corpo = self.texto_corpo.get("1.0", "end").strip()

        if not numero or not para or not assunto or not corpo:
            messagebox.showwarning(
                "Campos obrigatórios", "Preencha ao menos Número, Para, Assunto e o corpo do memorando.")
            return

        dados_memo = {
            "numero": numero,
            "ano": self.campo_ano.get(),
            "protocolo": self.campo_protocolo.get(),
            "para": para,
            "assunto": assunto,
            "saudacao": self.campo_saudacao.get(),
            "corpo": corpo,
            "fecho": self.campo_fecho.get(),
            "assinado_por": self.campo_assinado_por.get(),
            "data": self.campo_data.get(),
        }

        nome_sugerido = f"Memorando {numero}-{self.campo_ano.get()}.docx".replace("/", "-")
        caminho = filedialog.asksaveasfilename(
            title="Salvar memorando",
            initialfile=nome_sugerido,
            defaultextension=".docx",
            filetypes=[("Documento Word", "*.docx")])
        if not caminho:
            return

        try:
            documentos.gerar_memorando(self.config_escola, dados_memo, caminho)
        except Exception as e:
            messagebox.showerror("Erro ao gerar memorando", str(e))
            return

        self.label_status_memo.config(text=f"Memorando gerado: {os.path.basename(caminho)}")
        if messagebox.askyesno("Memorando gerado", "Memorando gerado com sucesso. Deseja abrir o arquivo agora?"):
            try:
                os.startfile(caminho)
            except Exception:
                pass

    # ---------------- Ofício ----------------
    def _montar_pagina_oficio(self):
        pagina = self.paginas["oficio"]

        canvas = tk.Canvas(pagina, bg=COR_FUNDO, highlightthickness=0)
        scrollbar = ttk.Scrollbar(pagina, orient="vertical", command=canvas.yview)
        cartao_externo = tk.Frame(canvas, bg=COR_FUNDO)

        cartao_externo.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        janela_canvas = canvas.create_window((0, 0), window=cartao_externo, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(janela_canvas, width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        cartao = tk.Frame(cartao_externo, bg=COR_CARTAO, padx=24, pady=20)
        cartao.pack(fill="x", padx=4, pady=4)

        tk.Label(cartao, text="Novo ofício", bg=COR_CARTAO, fg=COR_AZUL_ESCURO,
                 font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(0, 16))
        tk.Label(cartao,
                 text="O ofício segue o padrão oficial: carta corrida, sem quadros (diferente do memorando).",
                 bg=COR_CARTAO, fg=COR_TEXTO_FRACO, font=("Segoe UI", 9), justify="left").pack(anchor="w", pady=(0, 16))

        linha_numero = tk.Frame(cartao, bg=COR_CARTAO)
        linha_numero.pack(fill="x")
        self.campo_numero_oficio = CampoTexto(linha_numero, "Número do ofício", largura=15)
        self.campo_numero_oficio.pack(side="left", padx=(0, 20))
        self.campo_ano_oficio = CampoTexto(linha_numero, "Ano", str(datetime.date.today().year), largura=10)
        self.campo_ano_oficio.pack(side="left")

        self.campo_protocolo_oficio = CampoTexto(
            cartao, "Protocolo (ex: SIGED Nº 12345) - opcional, deixe em branco se não usar", largura=40)
        self.campo_protocolo_oficio.pack(fill="x")

        cidade = (self.config_escola.get("cidade") or "").strip()
        self._ultima_data_oficio_auto = data_por_extenso(cidade)
        self.campo_data_oficio = CampoTexto(
            cartao, "Local e data", self._ultima_data_oficio_auto, largura=45)
        self.campo_data_oficio.pack(fill="x")

        self.campo_para_oficio = CampoTexto(cartao, "Para (nome do destinatário)", largura=60)
        self.campo_para_oficio.pack(fill="x")

        self.campo_cargo_destinatario = CampoTexto(
            cartao, "Cargo/instituição do destinatário (opcional)", largura=60)
        self.campo_cargo_destinatario.pack(fill="x")

        self.campo_assunto_oficio = CampoTexto(cartao, "Assunto", largura=70)
        self.campo_assunto_oficio.pack(fill="x")

        self.campo_saudacao_oficio = CampoTexto(cartao, "Saudação", "Prezado(a) Senhor(a),", largura=40)
        self.campo_saudacao_oficio.pack(fill="x")

        tk.Label(cartao, text="Corpo do ofício (cada parágrafo em uma linha)",
                 bg=COR_CARTAO, fg=COR_TEXTO, font=FONTE_LABEL).pack(anchor="w")
        self.texto_corpo_oficio = tk.Text(cartao, height=10, font=FONTE_PADRAO, relief="solid", bd=1,
                                           highlightthickness=1, highlightbackground=COR_BORDA, wrap="word")
        self.texto_corpo_oficio.pack(fill="x", pady=(3, 10))
        self.texto_corpo_oficio.insert("1.0", "Vimos, por meio deste, ")

        self.campo_fecho_oficio = CampoTexto(cartao, "Fecho", "Atenciosamente,", largura=30)
        self.campo_fecho_oficio.pack(fill="x")

        self.campo_assinado_por_oficio = CampoTexto(
            cartao,
            "Assinado por (opcional - só aparece se a escola não tiver cadastrado uma\n"
            "imagem de assinatura em Dados da Escola)",
            largura=50)
        self.campo_assinado_por_oficio.pack(fill="x")

        self.campo_cargo_assinado_por_oficio = CampoTexto(
            cartao, "Cargo de quem assina (opcional)", largura=40)
        self.campo_cargo_assinado_por_oficio.pack(fill="x")

        btn_gerar = tk.Button(cartao, text="Gerar Ofício (.docx)", command=self._gerar_oficio,
                               bg=COR_AZUL, fg="white", font=("Segoe UI", 11, "bold"),
                               activebackground=COR_AZUL_ESCURO, activeforeground="white",
                               relief="flat", padx=20, pady=10, cursor="hand2")
        btn_gerar.pack(anchor="w", pady=(10, 0))

        self.label_status_oficio = tk.Label(cartao, text="", bg=COR_CARTAO, fg=COR_VERDE, font=FONTE_PADRAO)
        self.label_status_oficio.pack(anchor="w", pady=(8, 0))

    def _gerar_oficio(self):
        if not self.config_escola.get("nome_escola"):
            messagebox.showwarning(
                "Dados da escola pendentes",
                "Preencha e salve os Dados da Escola antes de gerar um ofício.")
            self._ir_para("escola")
            return

        numero = self.campo_numero_oficio.get()
        para = self.campo_para_oficio.get()
        assunto = self.campo_assunto_oficio.get()
        corpo = self.texto_corpo_oficio.get("1.0", "end").strip()

        if not numero or not para or not assunto or not corpo:
            messagebox.showwarning(
                "Campos obrigatórios", "Preencha ao menos Número, Para, Assunto e o corpo do ofício.")
            return

        dados_oficio = {
            "numero": numero,
            "ano": self.campo_ano_oficio.get(),
            "protocolo": self.campo_protocolo_oficio.get(),
            "data": self.campo_data_oficio.get(),
            "para": para,
            "cargo_destinatario": self.campo_cargo_destinatario.get(),
            "assunto": assunto,
            "saudacao": self.campo_saudacao_oficio.get(),
            "corpo": corpo,
            "fecho": self.campo_fecho_oficio.get(),
            "assinado_por": self.campo_assinado_por_oficio.get(),
            "cargo_assinado_por": self.campo_cargo_assinado_por_oficio.get(),
        }

        nome_sugerido = f"Oficio {numero}-{self.campo_ano_oficio.get()}.docx".replace("/", "-")
        caminho = filedialog.asksaveasfilename(
            title="Salvar ofício",
            initialfile=nome_sugerido,
            defaultextension=".docx",
            filetypes=[("Documento Word", "*.docx")])
        if not caminho:
            return

        try:
            documentos.gerar_oficio(self.config_escola, dados_oficio, caminho)
        except Exception as e:
            messagebox.showerror("Erro ao gerar ofício", str(e))
            return

        self.label_status_oficio.config(text=f"Ofício gerado: {os.path.basename(caminho)}")
        if messagebox.askyesno("Ofício gerado", "Ofício gerado com sucesso. Deseja abrir o arquivo agora?"):
            try:
                os.startfile(caminho)
            except Exception:
                pass

    # ---------------- Declaração Escolar ----------------
    def _montar_pagina_declaracao(self):
        pagina = self.paginas["declaracao"]

        canvas = tk.Canvas(pagina, bg=COR_FUNDO, highlightthickness=0)
        scrollbar = ttk.Scrollbar(pagina, orient="vertical", command=canvas.yview)
        cartao_externo = tk.Frame(canvas, bg=COR_FUNDO)

        cartao_externo.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        janela_canvas = canvas.create_window((0, 0), window=cartao_externo, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(janela_canvas, width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        cartao = tk.Frame(cartao_externo, bg=COR_CARTAO, padx=24, pady=20)
        cartao.pack(fill="x", padx=4, pady=4)

        tk.Label(cartao, text="Nova declaração escolar", bg=COR_CARTAO, fg=COR_AZUL_ESCURO,
                 font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(0, 16))

        self.campo_aluno = CampoTexto(cartao, "Nome completo do(a) aluno(a)", largura=60)
        self.campo_aluno.pack(fill="x")

        self.campo_tipo_codigo = CampoTipoCodigo(cartao, "Tipo de código do(a) aluno(a)")
        self.campo_tipo_codigo.pack(fill="x")

        self.campo_codigo_sigeam = CampoTexto(cartao, "Número/código - opcional", largura=30)
        self.campo_codigo_sigeam.pack(fill="x")

        linha1 = tk.Frame(cartao, bg=COR_CARTAO)
        linha1.pack(fill="x")
        self.campo_situacao_matricula = CampoOpcao(
            linha1, "Situação da matrícula",
            [("esta", "Está matriculado(a)"), ("foi", "Foi matriculado(a)")], largura=20)
        self.campo_situacao_matricula.pack(side="left", padx=(0, 20))
        self.campo_ano_letivo = CampoTexto(linha1, "Ano letivo", largura=12)
        self.campo_ano_letivo.pack(side="left")

        linha2 = tk.Frame(cartao, bg=COR_CARTAO)
        linha2.pack(fill="x")
        self.campo_situacao_curso = CampoOpcao(
            linha2, "Situação do curso", [("cursa", "Cursa"), ("cursou", "Cursou")], largura=15)
        self.campo_situacao_curso.pack(side="left", padx=(0, 20))
        self.campo_serie = CampoTexto(linha2, "Série", largura=18)
        self.campo_serie.pack(side="left", padx=(0, 20))
        self.campo_turno = CampoOpcao(
            linha2, "Turno",
            [("matutino", "Matutino"), ("vespertino", "Vespertino"),
             ("noturno", "Noturno"), ("intermediario", "Intermediário")], largura=15)
        self.campo_turno.pack(side="left")

        self.campo_turma = CampoTexto(cartao, "Turma", largura=30)
        self.campo_turma.pack(fill="x")

        tk.Label(cartao, text="Declarações para fins de", bg=COR_CARTAO, fg=COR_AZUL_ESCURO,
                 font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(10, 6))

        self.campo_finalidade = CampoOpcao(
            cartao, "Finalidade",
            [("trabalho", "Trabalho"), ("transferencia", "Transferência"), ("sinetram", "Sinetram"),
             ("bolsa_familia", "Bolsa Família / Frequência"), ("outros", "Outros")], largura=35)
        self.campo_finalidade.pack(fill="x")

        self.campo_finalidade_frequencia = CampoTexto(
            cartao, "Frequência (%) - só quando a finalidade é Bolsa Família", largura=15)
        self.campo_finalidade_frequencia.pack(fill="x")

        self.campo_finalidade_outros = CampoTexto(
            cartao, "Descrição - só quando a finalidade é Outros", largura=50)
        self.campo_finalidade_outros.pack(fill="x")

        tk.Label(cartao, text="Status do aluno", bg=COR_CARTAO, fg=COR_AZUL_ESCURO,
                 font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(10, 6))

        self.campo_status_aluno = CampoOpcao(
            cartao, "Status",
            [("promovido", "Promovido(a)"), ("retido", "Retido(a)"), ("desistente", "Desistente"),
             ("progressao_parcial", "Progressão Parcial"), ("cursando", "Cursando")], largura=25)
        self.campo_status_aluno.pack(fill="x")

        self.campo_status_desistente_data = CampoTexto(
            cartao, "Data da desistência (DD/MM/AAAA) - só quando o status é Desistente", largura=20)
        self.campo_status_desistente_data.pack(fill="x")

        self.campo_obs = CampoTexto(cartao, "OBS (opcional)", largura=70)
        self.campo_obs.pack(fill="x")

        cidade = (self.config_escola.get("cidade") or "").strip()
        self._ultima_data_declaracao_auto = data_por_extenso(cidade)
        self.campo_data_declaracao = CampoTexto(
            cartao, "Local e data", self._ultima_data_declaracao_auto, largura=45)
        self.campo_data_declaracao.pack(fill="x")

        btn_gerar = tk.Button(cartao, text="Gerar Declaração (.docx)", command=self._gerar_declaracao,
                               bg=COR_AZUL, fg="white", font=("Segoe UI", 11, "bold"),
                               activebackground=COR_AZUL_ESCURO, activeforeground="white",
                               relief="flat", padx=20, pady=10, cursor="hand2")
        btn_gerar.pack(anchor="w", pady=(10, 0))

        self.label_status_declaracao = tk.Label(cartao, text="", bg=COR_CARTAO, fg=COR_VERDE, font=FONTE_PADRAO)
        self.label_status_declaracao.pack(anchor="w", pady=(8, 0))

    def _gerar_declaracao(self):
        if not self.config_escola.get("nome_escola"):
            messagebox.showwarning(
                "Dados da escola pendentes",
                "Preencha e salve os Dados da Escola antes de gerar uma declaração.")
            self._ir_para("escola")
            return

        aluno = self.campo_aluno.get()
        if not aluno:
            messagebox.showwarning("Campo obrigatório", "Preencha o nome do(a) aluno(a).")
            return

        dados_decl = {
            "aluno": aluno,
            "codigo_sigeam": self.campo_codigo_sigeam.get(),
            "codigo_tipo": self.campo_tipo_codigo.get(),
            "situacao_matricula": self.campo_situacao_matricula.get(),
            "ano_letivo": self.campo_ano_letivo.get(),
            "situacao_curso": self.campo_situacao_curso.get(),
            "serie": self.campo_serie.get(),
            "turma": self.campo_turma.get(),
            "turno": self.campo_turno.get(),
            "finalidade": self.campo_finalidade.get(),
            "finalidade_frequencia": self.campo_finalidade_frequencia.get(),
            "finalidade_outros": self.campo_finalidade_outros.get(),
            "status_aluno": self.campo_status_aluno.get(),
            "status_desistente_data": self.campo_status_desistente_data.get(),
            "obs": self.campo_obs.get(),
            "data": self.campo_data_declaracao.get(),
        }

        nome_sugerido = f"Declaracao - {aluno}.docx"
        caminho = filedialog.asksaveasfilename(
            title="Salvar declaração",
            initialfile=nome_sugerido,
            defaultextension=".docx",
            filetypes=[("Documento Word", "*.docx")])
        if not caminho:
            return

        try:
            documentos.gerar_declaracao(self.config_escola, dados_decl, caminho)
        except Exception as e:
            messagebox.showerror("Erro ao gerar declaração", str(e))
            return

        self.label_status_declaracao.config(text=f"Declaração gerada: {os.path.basename(caminho)}")
        if messagebox.askyesno("Declaração gerada", "Declaração gerada com sucesso. Deseja abrir o arquivo agora?"):
            try:
                os.startfile(caminho)
            except Exception:
                pass

    # ---------------- Capa de Abertura de Livro ----------------
    def _montar_pagina_capa_livro(self):
        pagina = self.paginas["capa_livro"]

        canvas = tk.Canvas(pagina, bg=COR_FUNDO, highlightthickness=0)
        scrollbar = ttk.Scrollbar(pagina, orient="vertical", command=canvas.yview)
        cartao_externo = tk.Frame(canvas, bg=COR_FUNDO)

        cartao_externo.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        janela_canvas = canvas.create_window((0, 0), window=cartao_externo, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(janela_canvas, width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        cartao = tk.Frame(cartao_externo, bg=COR_CARTAO, padx=24, pady=20)
        cartao.pack(fill="x", padx=4, pady=4)

        tk.Label(cartao, text="Novo termo de abertura de livro", bg=COR_CARTAO, fg=COR_AZUL_ESCURO,
                 font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(0, 4))
        tk.Label(cartao,
                 text="Serve pra abrir qualquer tipo de livro de registro da escola (atas,\n"
                      "ponto, ocorrências etc.) - o texto é sempre o mesmo, só os dados abaixo mudam.",
                 bg=COR_CARTAO, fg=COR_TEXTO_FRACO, font=("Segoe UI", 9), justify="left").pack(anchor="w", pady=(0, 16))

        self.campo_tipo_livro = CampoTexto(
            cartao, "Tipo de livro (ex: Livro de Atas, Livro Ponto, Livro de Ocorrências)", largura=50)
        self.campo_tipo_livro.pack(fill="x")

        linha_livro = tk.Frame(cartao, bg=COR_CARTAO)
        linha_livro.pack(fill="x")
        self.campo_numero_livro = CampoTexto(linha_livro, "Número do livro", largura=12)
        self.campo_numero_livro.pack(side="left", padx=(0, 20))
        self.campo_qtd_folhas = CampoTexto(linha_livro, "Quantidade de folhas numeradas", largura=15)
        self.campo_qtd_folhas.pack(side="left")

        self.campo_finalidade_livro = CampoTexto(
            cartao, "Finalidade do livro (pra que ele vai servir)", largura=70)
        self.campo_finalidade_livro.pack(fill="x")

        cidade = (self.config_escola.get("cidade") or "").strip()
        self._ultima_data_capa_livro_auto = data_por_extenso(cidade)
        self.campo_data_capa_livro = CampoTexto(
            cartao, "Local e data", self._ultima_data_capa_livro_auto, largura=45)
        self.campo_data_capa_livro.pack(fill="x")

        self.campo_responsavel_abertura = CampoTexto(
            cartao, "Responsável pela abertura", self.config_escola.get("diretor_nome", ""), largura=50)
        self.campo_responsavel_abertura.pack(fill="x")

        self.campo_cargo_abertura = CampoTexto(
            cartao, "Cargo", self.config_escola.get("diretor_cargo", ""), largura=30)
        self.campo_cargo_abertura.pack(fill="x")

        btn_gerar = tk.Button(cartao, text="Gerar Termo de Abertura (.docx)", command=self._gerar_capa_livro,
                               bg=COR_AZUL, fg="white", font=("Segoe UI", 11, "bold"),
                               activebackground=COR_AZUL_ESCURO, activeforeground="white",
                               relief="flat", padx=20, pady=10, cursor="hand2")
        btn_gerar.pack(anchor="w", pady=(10, 0))

        self.label_status_capa_livro = tk.Label(cartao, text="", bg=COR_CARTAO, fg=COR_VERDE, font=FONTE_PADRAO)
        self.label_status_capa_livro.pack(anchor="w", pady=(8, 0))

    def _gerar_capa_livro(self):
        if not self.config_escola.get("nome_escola"):
            messagebox.showwarning(
                "Dados da escola pendentes",
                "Preencha e salve os Dados da Escola antes de gerar um termo de abertura.")
            self._ir_para("escola")
            return

        tipo_livro = self.campo_tipo_livro.get()
        finalidade = self.campo_finalidade_livro.get()
        if not tipo_livro or not finalidade:
            messagebox.showwarning(
                "Campos obrigatórios", "Preencha ao menos o Tipo de livro e a Finalidade.")
            return

        dados_livro = {
            "tipo_livro": tipo_livro,
            "numero": self.campo_numero_livro.get(),
            "qtd_folhas": self.campo_qtd_folhas.get(),
            "finalidade": finalidade,
            "data": self.campo_data_capa_livro.get(),
            "responsavel_abertura": self.campo_responsavel_abertura.get(),
            "cargo_abertura": self.campo_cargo_abertura.get(),
        }

        nome_sugerido = f"Termo de Abertura - {tipo_livro}.docx"
        caminho = filedialog.asksaveasfilename(
            title="Salvar termo de abertura",
            initialfile=nome_sugerido,
            defaultextension=".docx",
            filetypes=[("Documento Word", "*.docx")])
        if not caminho:
            return

        try:
            documentos.gerar_capa_livro(self.config_escola, dados_livro, caminho)
        except Exception as e:
            messagebox.showerror("Erro ao gerar termo de abertura", str(e))
            return

        self.label_status_capa_livro.config(text=f"Termo de abertura gerado: {os.path.basename(caminho)}")
        if messagebox.askyesno(
                "Termo de abertura gerado", "Termo de abertura gerado com sucesso. Deseja abrir o arquivo agora?"):
            try:
                os.startfile(caminho)
            except Exception:
                pass

    # ---------------- Ficha de Matrícula ----------------
    def _montar_pagina_ficha_matricula(self):
        pagina = self.paginas["ficha_matricula"]

        canvas = tk.Canvas(pagina, bg=COR_FUNDO, highlightthickness=0)
        scrollbar = ttk.Scrollbar(pagina, orient="vertical", command=canvas.yview)
        cartao_externo = tk.Frame(canvas, bg=COR_FUNDO)

        cartao_externo.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        janela_canvas = canvas.create_window((0, 0), window=cartao_externo, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(janela_canvas, width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        cartao = tk.Frame(cartao_externo, bg=COR_CARTAO, padx=24, pady=20)
        cartao.pack(fill="x", padx=4, pady=4)

        tk.Label(cartao, text="Nova ficha de matrícula", bg=COR_CARTAO, fg=COR_AZUL_ESCURO,
                 font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(0, 16))

        self.campo_ficha_tipo_ficha = CampoOpcao(
            cartao, "Tipo de ficha",
            [("Educação Infantil", "Educação Infantil"), ("Maternal", "Maternal"),
             ("Ensino Fundamental I", "Ensino Fundamental I"), ("Ensino Fundamental II", "Ensino Fundamental II")],
            largura=25)
        self.campo_ficha_tipo_ficha.pack(fill="x")

        self.campo_ficha_codigo_aluno = CampoTexto(cartao, "Código do aluno (do sistema da rede) - opcional",
                                                    largura=25)
        self.campo_ficha_codigo_aluno.pack(fill="x")
        self.campo_ficha_nis = CampoTexto(cartao, "N.I.S. (Número de Identificação Social) - opcional", largura=25)
        self.campo_ficha_nis.pack(fill="x")

        tk.Label(cartao, text="Dados pessoais da criança", bg=COR_CARTAO, fg=COR_AZUL_ESCURO,
                 font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(10, 6))

        self.campo_ficha_nome_social = CampoTexto(cartao, "Nome social - opcional", largura=50)
        self.campo_ficha_nome_social.pack(fill="x")
        self.campo_ficha_nome_crianca = CampoTexto(cartao, "Nome completo da criança - sem abreviaturas",
                                                     largura=60)
        self.campo_ficha_nome_crianca.pack(fill="x")

        linha1 = tk.Frame(cartao, bg=COR_CARTAO)
        linha1.pack(fill="x")
        self.campo_ficha_data_nascimento = CampoTexto(linha1, "Data de nascimento", largura=14)
        self.campo_ficha_data_nascimento.pack(side="left", padx=(0, 16))
        self.campo_ficha_sexo = CampoOpcao(
            linha1, "Sexo", [("masculino", "Masculino"), ("feminino", "Feminino")], largura=12)
        self.campo_ficha_sexo.pack(side="left", padx=(0, 16))
        self.campo_ficha_gemeo = CampoOpcao(linha1, "Gêmeo(a)", [("nao", "Não"), ("sim", "Sim")], largura=8)
        self.campo_ficha_gemeo.pack(side="left", padx=(0, 16))
        self.campo_ficha_tipo_sanguineo = CampoTexto(linha1, "Tipo sanguíneo - opcional", largura=10)
        self.campo_ficha_tipo_sanguineo.pack(side="left")

        linha2 = tk.Frame(cartao, bg=COR_CARTAO)
        linha2.pack(fill="x")
        self.campo_ficha_nacionalidade = CampoTexto(linha2, "Nacionalidade (só para estrangeiro) - opcional",
                                                      largura=30)
        self.campo_ficha_nacionalidade.pack(side="left", padx=(0, 16))
        self.campo_ficha_data_entrada_pais = CampoTexto(linha2, "Data de entrada no país - opcional", largura=16)
        self.campo_ficha_data_entrada_pais.pack(side="left")

        linha3 = tk.Frame(cartao, bg=COR_CARTAO)
        linha3.pack(fill="x")
        self.campo_ficha_naturalidade = CampoTexto(linha3, "Naturalidade/Município", largura=35)
        self.campo_ficha_naturalidade.pack(side="left", padx=(0, 16))
        self.campo_ficha_uf_naturalidade = CampoTexto(linha3, "UF", largura=6)
        self.campo_ficha_uf_naturalidade.pack(side="left")

        self.campo_ficha_nome_mae = CampoTexto(cartao, "Nome completo da mãe - sem abreviaturas", largura=60)
        self.campo_ficha_nome_mae.pack(fill="x")
        self.campo_ficha_nome_pai = CampoTexto(cartao, "Nome completo do pai - sem abreviaturas - opcional",
                                                largura=60)
        self.campo_ficha_nome_pai.pack(fill="x")

        self.campo_ficha_endereco = CampoTexto(cartao, "Endereço residencial", largura=60)
        self.campo_ficha_endereco.pack(fill="x")

        linha4 = tk.Frame(cartao, bg=COR_CARTAO)
        linha4.pack(fill="x")
        self.campo_ficha_numero = CampoTexto(linha4, "Número", largura=10)
        self.campo_ficha_numero.pack(side="left", padx=(0, 16))
        self.campo_ficha_complemento = CampoTexto(linha4, "Complemento - opcional", largura=30)
        self.campo_ficha_complemento.pack(side="left", padx=(0, 16))
        self.campo_ficha_tipo_logradouro = CampoTexto(linha4, "Tipo de logradouro (ex: Rua, Avenida)", largura=18)
        self.campo_ficha_tipo_logradouro.pack(side="left")

        linha5 = tk.Frame(cartao, bg=COR_CARTAO)
        linha5.pack(fill="x")
        self.campo_ficha_bairro = CampoTexto(linha5, "Bairro", largura=35)
        self.campo_ficha_bairro.pack(side="left", padx=(0, 16))
        self.campo_ficha_cep = CampoTexto(linha5, "CEP", largura=12)
        self.campo_ficha_cep.pack(side="left")

        tk.Label(cartao, text="Certidão de nascimento", bg=COR_CARTAO, fg=COR_AZUL_ESCURO,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(8, 4))

        linha6 = tk.Frame(cartao, bg=COR_CARTAO)
        linha6.pack(fill="x")
        self.campo_ficha_numero_termo = CampoTexto(linha6, "Número do termo", largura=14)
        self.campo_ficha_numero_termo.pack(side="left", padx=(0, 16))
        self.campo_ficha_folha = CampoTexto(linha6, "Folha", largura=8)
        self.campo_ficha_folha.pack(side="left", padx=(0, 16))
        self.campo_ficha_livro = CampoTexto(linha6, "Livro", largura=10)
        self.campo_ficha_livro.pack(side="left", padx=(0, 16))
        self.campo_ficha_data_emissao_certidao = CampoTexto(linha6, "Data de emissão", largura=14)
        self.campo_ficha_data_emissao_certidao.pack(side="left", padx=(0, 16))
        self.campo_ficha_uf_cartorio = CampoTexto(linha6, "UF do cartório", largura=8)
        self.campo_ficha_uf_cartorio.pack(side="left")

        self.campo_ficha_nome_cartorio = CampoTexto(cartao, "Nome do cartório - órgão emissor", largura=60)
        self.campo_ficha_nome_cartorio.pack(fill="x")
        self.campo_ficha_matricula_registro_civil = CampoTexto(
            cartao, "Matrícula do Registro Civil (número com 32 dígitos) - opcional", largura=45)
        self.campo_ficha_matricula_registro_civil.pack(fill="x")

        tk.Label(cartao, text="Identidade (RG) - opcional", bg=COR_CARTAO, fg=COR_AZUL_ESCURO,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(8, 4))

        linha7 = tk.Frame(cartao, bg=COR_CARTAO)
        linha7.pack(fill="x")
        self.campo_ficha_numero_identidade = CampoTexto(linha7, "Número da identidade", largura=16)
        self.campo_ficha_numero_identidade.pack(side="left", padx=(0, 16))
        self.campo_ficha_complemento_identidade = CampoTexto(linha7, "Complemento", largura=16)
        self.campo_ficha_complemento_identidade.pack(side="left", padx=(0, 16))
        self.campo_ficha_data_expedicao_identidade = CampoTexto(linha7, "Data de expedição", largura=14)
        self.campo_ficha_data_expedicao_identidade.pack(side="left", padx=(0, 16))
        self.campo_ficha_uf_rg = CampoTexto(linha7, "UF", largura=6)
        self.campo_ficha_uf_rg.pack(side="left", padx=(0, 16))
        self.campo_ficha_orgao_emissor_identidade = CampoTexto(linha7, "Órgão emissor", largura=14)
        self.campo_ficha_orgao_emissor_identidade.pack(side="left")

        linha8 = tk.Frame(cartao, bg=COR_CARTAO)
        linha8.pack(fill="x")
        self.campo_ficha_cpf = CampoTexto(linha8, "Número do CPF - opcional", largura=18)
        self.campo_ficha_cpf.pack(side="left", padx=(0, 16))
        self.campo_ficha_cor_raca = CampoOpcao(
            linha8, "Cor/Raça",
            [("branca", "Branca"), ("preta", "Preta"), ("parda", "Parda"), ("amarela", "Amarela"),
             ("indigena", "Indígena"), ("nao_declarada", "Não declarada")], largura=16)
        self.campo_ficha_cor_raca.pack(side="left", padx=(0, 16))
        self.campo_ficha_telefone_ficha = CampoTexto(linha8, "Telefone de contato", largura=18)
        self.campo_ficha_telefone_ficha.pack(side="left")

        tk.Label(cartao, text="Dados escolares", bg=COR_CARTAO, fg=COR_AZUL_ESCURO,
                 font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(14, 6))

        linha9 = tk.Frame(cartao, bg=COR_CARTAO)
        linha9.pack(fill="x")
        self.campo_ficha_data_ingresso = CampoTexto(
            linha9, "Data de ingresso na Unidade de Ensino", datetime.date.today().strftime("%d/%m/%Y"), largura=16)
        self.campo_ficha_data_ingresso.pack(side="left", padx=(0, 16))
        self.campo_ficha_deficiencia = CampoOpcao(
            linha9, "Criança com deficiência", [("nao", "Não"), ("sim", "Sim")], largura=10)
        self.campo_ficha_deficiencia.pack(side="left", padx=(0, 16))
        self.campo_ficha_bolsa_familia = CampoOpcao(
            linha9, "Participa do Bolsa Família", [("nao", "Não"), ("sim", "Sim")], largura=10)
        self.campo_ficha_bolsa_familia.pack(side="left")

        linha10 = tk.Frame(cartao, bg=COR_CARTAO)
        linha10.pack(fill="x")
        self.campo_ficha_tipo_deficiencia = CampoTexto(linha10, "Tipo de deficiência - opcional", largura=30)
        self.campo_ficha_tipo_deficiencia.pack(side="left", padx=(0, 16))
        self.campo_ficha_necessidades_especiais = CampoOpcao(
            linha10, "Necessidades educacionais especiais", [("nao", "Não"), ("sim", "Sim")], largura=10)
        self.campo_ficha_necessidades_especiais.pack(side="left", padx=(0, 16))
        self.campo_ficha_apoio_pedagogico = CampoOpcao(
            linha10, "Apoio pedagógico especializado",
            [("", "Não se aplica"), ("na_escola", "Na própria escola"), ("outra_escola", "Outra escola/Centro")],
            largura=20)
        self.campo_ficha_apoio_pedagogico.pack(side="left")

        linha11 = tk.Frame(cartao, bg=COR_CARTAO)
        linha11.pack(fill="x")
        self.campo_ficha_transporte_escolar = CampoOpcao(
            linha11, "Utiliza transporte escolar público", [("nao", "Não"), ("sim", "Sim")], largura=10)
        self.campo_ficha_transporte_escolar.pack(side="left", padx=(0, 16))
        self.campo_ficha_tipo_transporte = CampoOpcao(
            linha11, "Tipo de transporte oferecido",
            [("", "Não se aplica"), ("fluvial", "Fluvial"), ("rodoviario", "Rodoviário")], largura=16)
        self.campo_ficha_tipo_transporte.pack(side="left", padx=(0, 16))
        self.campo_ficha_zona_residencia = CampoOpcao(
            linha11, "Zona de residência", [("urbana", "Urbana"), ("rural", "Rural")], largura=10)
        self.campo_ficha_zona_residencia.pack(side="left")

        self.campo_ficha_movimento = CampoOpcao(
            cartao, "Movimento e rendimento escolar",
            [("nenhum", "Nenhum (matrícula normal)"), ("abandono", "Afastado por abandono"),
             ("transferencia", "Afastado por transferência"), ("matricula_final", "Matrícula Final")],
            largura=30)
        self.campo_ficha_movimento.pack(fill="x")

        self.campo_ficha_etapa = CampoTexto(cartao, "Série/Turma (ex: Maternal II, 5º Ano)", largura=30)
        self.campo_ficha_etapa.pack(fill="x")

        btn_gerar = tk.Button(cartao, text="Gerar Ficha de Matrícula (.docx)", command=self._gerar_ficha_matricula,
                               bg=COR_AZUL, fg="white", font=("Segoe UI", 11, "bold"),
                               activebackground=COR_AZUL_ESCURO, activeforeground="white",
                               relief="flat", padx=20, pady=10, cursor="hand2")
        btn_gerar.pack(anchor="w", pady=(14, 0))

        self.label_status_ficha_matricula = tk.Label(cartao, text="", bg=COR_CARTAO, fg=COR_VERDE,
                                                       font=FONTE_PADRAO)
        self.label_status_ficha_matricula.pack(anchor="w", pady=(8, 0))

    def _gerar_ficha_matricula(self):
        if not self.config_escola.get("nome_escola"):
            messagebox.showwarning(
                "Dados da escola pendentes",
                "Preencha e salve os Dados da Escola antes de gerar uma ficha de matrícula.")
            self._ir_para("escola")
            return

        nome_crianca = self.campo_ficha_nome_crianca.get()
        nome_mae = self.campo_ficha_nome_mae.get()
        if not nome_crianca or not nome_mae:
            messagebox.showwarning(
                "Campos obrigatórios", "Preencha ao menos o nome da criança e o nome da mãe.")
            return

        dados_ficha = {
            "tipo_ficha": self.campo_ficha_tipo_ficha.get(),
            "codigo_aluno": self.campo_ficha_codigo_aluno.get(),
            "nis": self.campo_ficha_nis.get(),
            "nome_social": self.campo_ficha_nome_social.get(),
            "nome_crianca": nome_crianca,
            "data_nascimento": self.campo_ficha_data_nascimento.get(),
            "sexo": self.campo_ficha_sexo.get(),
            "gemeo": self.campo_ficha_gemeo.get() == "sim",
            "tipo_sanguineo": self.campo_ficha_tipo_sanguineo.get(),
            "nacionalidade": self.campo_ficha_nacionalidade.get(),
            "data_entrada_pais": self.campo_ficha_data_entrada_pais.get(),
            "naturalidade": self.campo_ficha_naturalidade.get(),
            "uf_naturalidade": self.campo_ficha_uf_naturalidade.get(),
            "nome_mae": nome_mae,
            "nome_pai": self.campo_ficha_nome_pai.get(),
            "endereco": self.campo_ficha_endereco.get(),
            "numero": self.campo_ficha_numero.get(),
            "complemento": self.campo_ficha_complemento.get(),
            "tipo_logradouro": self.campo_ficha_tipo_logradouro.get(),
            "bairro": self.campo_ficha_bairro.get(),
            "cep": self.campo_ficha_cep.get(),
            "numero_termo": self.campo_ficha_numero_termo.get(),
            "folha": self.campo_ficha_folha.get(),
            "livro": self.campo_ficha_livro.get(),
            "data_emissao_certidao": self.campo_ficha_data_emissao_certidao.get(),
            "uf_cartorio": self.campo_ficha_uf_cartorio.get(),
            "nome_cartorio": self.campo_ficha_nome_cartorio.get(),
            "matricula_registro_civil": self.campo_ficha_matricula_registro_civil.get(),
            "numero_identidade": self.campo_ficha_numero_identidade.get(),
            "complemento_identidade": self.campo_ficha_complemento_identidade.get(),
            "data_expedicao_identidade": self.campo_ficha_data_expedicao_identidade.get(),
            "uf_rg": self.campo_ficha_uf_rg.get(),
            "orgao_emissor_identidade": self.campo_ficha_orgao_emissor_identidade.get(),
            "cpf": self.campo_ficha_cpf.get(),
            "cor_raca": self.campo_ficha_cor_raca.get(),
            "telefone": self.campo_ficha_telefone_ficha.get(),
            "data_ingresso": self.campo_ficha_data_ingresso.get(),
            "deficiencia": self.campo_ficha_deficiencia.get(),
            "bolsa_familia": self.campo_ficha_bolsa_familia.get(),
            "tipo_deficiencia": self.campo_ficha_tipo_deficiencia.get(),
            "necessidades_especiais": self.campo_ficha_necessidades_especiais.get(),
            "apoio_pedagogico": self.campo_ficha_apoio_pedagogico.get(),
            "transporte_escolar": self.campo_ficha_transporte_escolar.get(),
            "tipo_transporte": self.campo_ficha_tipo_transporte.get(),
            "zona_residencia": self.campo_ficha_zona_residencia.get(),
            "movimento": self.campo_ficha_movimento.get(),
            "etapa": self.campo_ficha_etapa.get(),
        }

        nome_sugerido = f"Ficha de Matricula - {nome_crianca}.docx"
        caminho = filedialog.asksaveasfilename(
            title="Salvar ficha de matrícula",
            initialfile=nome_sugerido,
            defaultextension=".docx",
            filetypes=[("Documento Word", "*.docx")])
        if not caminho:
            return

        try:
            documentos.gerar_ficha_matricula(self.config_escola, dados_ficha, caminho)
        except Exception as e:
            messagebox.showerror("Erro ao gerar ficha de matrícula", str(e))
            return

        self.label_status_ficha_matricula.config(text=f"Ficha de matrícula gerada: {os.path.basename(caminho)}")
        if messagebox.askyesno(
                "Ficha de matrícula gerada",
                "Ficha de matrícula gerada com sucesso. Deseja abrir o arquivo agora?"):
            try:
                os.startfile(caminho)
            except Exception:
                pass

    # ---------------- Lista de Reunião de Pais e Alunos ----------------
    def _montar_pagina_lista_reuniao(self):
        pagina = self.paginas["lista_reuniao"]

        canvas = tk.Canvas(pagina, bg=COR_FUNDO, highlightthickness=0)
        scrollbar = ttk.Scrollbar(pagina, orient="vertical", command=canvas.yview)
        cartao_externo = tk.Frame(canvas, bg=COR_FUNDO)

        cartao_externo.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        janela_canvas = canvas.create_window((0, 0), window=cartao_externo, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(janela_canvas, width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        cartao = tk.Frame(cartao_externo, bg=COR_CARTAO, padx=24, pady=20)
        cartao.pack(fill="x", padx=4, pady=4)

        tk.Label(cartao, text="Nova lista de reunião", bg=COR_CARTAO, fg=COR_AZUL_ESCURO,
                 font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(0, 16))

        self.campo_reuniao_tipo = CampoTexto(
            cartao, "Tipo de reunião/lista (ex: Reunião de Pais e Alunos, Entrega de Livros...)",
            "Reunião de Pais e Alunos", largura=50)
        self.campo_reuniao_tipo.pack(fill="x")

        self.campo_reuniao_data = CampoTexto(
            cartao, "Data (texto livre, ex: 21/08/2025)", datetime.date.today().strftime("%d/%m/%Y"), largura=16)
        self.campo_reuniao_data.pack(fill="x")

        linha1 = tk.Frame(cartao, bg=COR_CARTAO)
        linha1.pack(fill="x")
        self.campo_reuniao_ano = CampoTexto(
            linha1, "Ano letivo", str(datetime.date.today().year), largura=10)
        self.campo_reuniao_ano.pack(side="left", padx=(0, 16))
        self.campo_reuniao_fase = CampoTexto(linha1, "Fase (ex: 1º Período)", largura=16)
        self.campo_reuniao_fase.pack(side="left", padx=(0, 16))
        self.campo_reuniao_turma = CampoTexto(linha1, "Turma", largura=14)
        self.campo_reuniao_turma.pack(side="left", padx=(0, 16))
        self.campo_reuniao_turno = CampoTexto(linha1, "Turno", largura=16)
        self.campo_reuniao_turno.pack(side="left")

        self.campo_reuniao_ensino_projeto = CampoTexto(
            cartao, "Ensino/Projeto (ex: Pré-Escola - 1º e 2º Períodos)", largura=50)
        self.campo_reuniao_ensino_projeto.pack(fill="x")

        tk.Label(cartao, text="Nomes dos alunos (um por linha)",
                 bg=COR_CARTAO, fg=COR_TEXTO, font=FONTE_LABEL).pack(anchor="w")
        self.texto_reuniao_nomes = tk.Text(cartao, height=14, font=FONTE_PADRAO, relief="solid", bd=1,
                                            highlightthickness=1, highlightbackground=COR_BORDA, wrap="word")
        self.texto_reuniao_nomes.pack(fill="x", pady=(3, 10))

        self.campo_reuniao_linhas_brancas = CampoTexto(
            cartao, "Linhas em branco na tabela (se deixar os nomes em branco)", "25", largura=10)
        self.campo_reuniao_linhas_brancas.pack(anchor="w", pady=(0, 10))

        tk.Label(cartao,
                 text="Preencha os dados de uma turma e clique em \"Gerar\". O programa\n"
                      "vai perguntar se você quer adicionar outra turma - se disser que\n"
                      "sim, troque a Turma e os Nomes dos alunos e clique em \"Gerar\" de\n"
                      "novo. Repita até adicionar todas; no final ele junta tudo num\n"
                      "único arquivo .docx, uma página por turma.",
                 bg=COR_CARTAO, fg=COR_TEXTO_FRACO, font=("Segoe UI", 9), justify="left").pack(anchor="w", pady=(0, 10))

        btn_gerar = tk.Button(cartao, text="Gerar Lista de Reunião (.docx)", command=self._gerar_lista_reuniao,
                               bg=COR_AZUL, fg="white", font=("Segoe UI", 11, "bold"),
                               activebackground=COR_AZUL_ESCURO, activeforeground="white",
                               relief="flat", padx=20, pady=10, cursor="hand2")
        btn_gerar.pack(anchor="w", pady=(10, 0))

        self.turmas_acumuladas_reuniao = []
        self.label_turmas_acumuladas_reuniao = tk.Label(
            cartao, text="", bg=COR_CARTAO, fg=COR_TEXTO_FRACO, font=("Segoe UI", 9), justify="left")
        self.label_turmas_acumuladas_reuniao.pack(anchor="w", pady=(8, 0))

        self.label_status_lista_reuniao = tk.Label(cartao, text="", bg=COR_CARTAO, fg=COR_VERDE, font=FONTE_PADRAO)
        self.label_status_lista_reuniao.pack(anchor="w", pady=(8, 0))

    def _atualizar_label_turmas_acumuladas_reuniao(self):
        n = len(self.turmas_acumuladas_reuniao)
        if n == 0:
            self.label_turmas_acumuladas_reuniao.config(text="")
        else:
            nomes_turmas = ", ".join(t["turma"] for t in self.turmas_acumuladas_reuniao)
            self.label_turmas_acumuladas_reuniao.config(
                text=f"{n} turma(s) já adicionada(s), aguardando finalizar: {nomes_turmas}")

    def _limpar_campos_turma_reuniao(self):
        self.campo_reuniao_turma.set("")
        self.texto_reuniao_nomes.delete("1.0", "end")

    def _gerar_lista_reuniao(self):
        if not self.config_escola.get("nome_escola"):
            messagebox.showwarning(
                "Dados da escola pendentes",
                "Preencha e salve os Dados da Escola antes de gerar a lista de reunião.")
            self._ir_para("escola")
            return

        turma = self.campo_reuniao_turma.get()
        if not turma:
            messagebox.showwarning("Campo obrigatório", "Preencha a Turma.")
            return

        try:
            linhas_brancas = int(self.campo_reuniao_linhas_brancas.get() or "25")
        except ValueError:
            linhas_brancas = 25

        dados_turma_atual = {
            "tipo_reuniao": self.campo_reuniao_tipo.get(),
            "data_reuniao": self.campo_reuniao_data.get(),
            "ano_letivo": self.campo_reuniao_ano.get(),
            "ensino_projeto": self.campo_reuniao_ensino_projeto.get(),
            "fase": self.campo_reuniao_fase.get(),
            "turno": self.campo_reuniao_turno.get(),
            "turma": turma,
            "nomes_alunos": self.texto_reuniao_nomes.get("1.0", "end").strip(),
            "linhas_em_branco": linhas_brancas,
        }
        self.turmas_acumuladas_reuniao.append(dados_turma_atual)
        self._atualizar_label_turmas_acumuladas_reuniao()

        quer_mais = messagebox.askyesno(
            "Adicionar outra turma?",
            f"Turma \"{turma}\" adicionada ({len(self.turmas_acumuladas_reuniao)} turma(s) até agora).\n\n"
            "Deseja preencher e adicionar outra turma antes de gerar o arquivo?")
        if quer_mais:
            self._limpar_campos_turma_reuniao()
            return

        turmas = self.turmas_acumuladas_reuniao
        if len(turmas) == 1:
            nome_sugerido = f"Lista de Reuniao - Turma {turmas[0]['turma']}.docx"
        else:
            nome_sugerido = "Lista de Reuniao - Varias Turmas.docx"

        caminho = filedialog.asksaveasfilename(
            title="Salvar lista de reunião",
            initialfile=nome_sugerido,
            defaultextension=".docx",
            filetypes=[("Documento Word", "*.docx")])
        if not caminho:
            # cancelou o salvamento - desfaz a adicao dessa turma pra nao
            # duplicar se a pessoa clicar em Gerar de novo
            self.turmas_acumuladas_reuniao.pop()
            self._atualizar_label_turmas_acumuladas_reuniao()
            return

        try:
            if len(turmas) == 1:
                documentos.gerar_lista_reuniao(self.config_escola, turmas[0], caminho)
            else:
                documentos.gerar_lista_reuniao_varias_turmas(self.config_escola, {}, turmas, caminho)
        except Exception as e:
            messagebox.showerror("Erro ao gerar lista de reunião", str(e))
            return

        self.turmas_acumuladas_reuniao = []
        self._atualizar_label_turmas_acumuladas_reuniao()
        self._limpar_campos_turma_reuniao()

        self.label_status_lista_reuniao.config(text=f"Lista de reunião gerada: {os.path.basename(caminho)}")
        if messagebox.askyesno(
                "Lista de reunião gerada",
                "Lista de reunião gerada com sucesso. Deseja abrir o arquivo agora?"):
            try:
                os.startfile(caminho)
            except Exception:
                pass

    # ---------------- Lista de Frequência Escolar ----------------
    def _montar_pagina_lista_frequencia(self):
        pagina = self.paginas["lista_frequencia"]

        canvas = tk.Canvas(pagina, bg=COR_FUNDO, highlightthickness=0)
        scrollbar = ttk.Scrollbar(pagina, orient="vertical", command=canvas.yview)
        cartao_externo = tk.Frame(canvas, bg=COR_FUNDO)

        cartao_externo.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        janela_canvas = canvas.create_window((0, 0), window=cartao_externo, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(janela_canvas, width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        cartao = tk.Frame(cartao_externo, bg=COR_CARTAO, padx=24, pady=20)
        cartao.pack(fill="x", padx=4, pady=4)

        tk.Label(cartao, text="Nova ficha de frequência escolar", bg=COR_CARTAO, fg=COR_AZUL_ESCURO,
                 font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(0, 16))

        linha1 = tk.Frame(cartao, bg=COR_CARTAO)
        linha1.pack(fill="x")
        self.campo_freq_mes = CampoTexto(linha1, "Mês", NOMES_MESES[datetime.date.today().month - 1], largura=16)
        self.campo_freq_mes.pack(side="left", padx=(0, 16))
        self.campo_freq_ano = CampoTexto(linha1, "Ano letivo", str(datetime.date.today().year), largura=10)
        self.campo_freq_ano.pack(side="left")

        linha2 = tk.Frame(cartao, bg=COR_CARTAO)
        linha2.pack(fill="x")
        self.campo_freq_turma = CampoTexto(linha2, "Turma", largura=14)
        self.campo_freq_turma.pack(side="left", padx=(0, 16))
        self.campo_freq_serie = CampoTexto(linha2, "Série (ex: Maternal II, 5º Ano)", largura=18)
        self.campo_freq_serie.pack(side="left", padx=(0, 16))
        self.campo_freq_turno = CampoTexto(linha2, "Turno", largura=16)
        self.campo_freq_turno.pack(side="left", padx=(0, 16))
        self.campo_freq_professor = CampoTexto(linha2, "Professor(a)", largura=30)
        self.campo_freq_professor.pack(side="left")

        self.campo_freq_dias = CampoTexto(
            cartao, "Dias letivos do mês (separados por vírgula - aceita faixas com \"-\", ex: 1-5,8-12,15-19,22-26)",
            largura=70)
        self.campo_freq_dias.pack(fill="x")

        tk.Label(cartao, text="Nomes dos alunos (um por linha)",
                 bg=COR_CARTAO, fg=COR_TEXTO, font=FONTE_LABEL).pack(anchor="w")
        self.texto_freq_nomes = tk.Text(cartao, height=14, font=FONTE_PADRAO, relief="solid", bd=1,
                                         highlightthickness=1, highlightbackground=COR_BORDA, wrap="word")
        self.texto_freq_nomes.pack(fill="x", pady=(3, 10))

        self.campo_freq_linhas_brancas = CampoTexto(
            cartao, "Linhas em branco na tabela (se deixar os nomes em branco)", "25", largura=10)
        self.campo_freq_linhas_brancas.pack(anchor="w", pady=(0, 10))

        tk.Label(cartao,
                 text="Sai em formato paisagem (página na horizontal), já que tem uma\n"
                      "coluna pra cada dia letivo do mês, mais Faltas e Faltas\n"
                      "Justificadas no final - tudo em branco, pra marcar \"P\"/\"F\" à\n"
                      "mão dia a dia e somar no fim do mês.",
                 bg=COR_CARTAO, fg=COR_TEXTO_FRACO, font=("Segoe UI", 9), justify="left").pack(anchor="w", pady=(0, 10))

        tk.Label(cartao,
                 text="Preencha os dados de uma turma e clique em \"Gerar\". O programa\n"
                      "vai perguntar se você quer adicionar outra turma - se disser que\n"
                      "sim, troque a Turma/Série/Turno/Professor(a)/Nomes dos alunos e\n"
                      "clique em \"Gerar\" de novo. Repita até adicionar todas; no final\n"
                      "ele junta tudo num único arquivo .docx, uma página por turma.",
                 bg=COR_CARTAO, fg=COR_TEXTO_FRACO, font=("Segoe UI", 9), justify="left").pack(anchor="w", pady=(0, 10))

        btn_gerar = tk.Button(cartao, text="Gerar Ficha de Frequência (.docx)", command=self._gerar_lista_frequencia,
                               bg=COR_AZUL, fg="white", font=("Segoe UI", 11, "bold"),
                               activebackground=COR_AZUL_ESCURO, activeforeground="white",
                               relief="flat", padx=20, pady=10, cursor="hand2")
        btn_gerar.pack(anchor="w", pady=(10, 0))

        self.turmas_acumuladas_frequencia = []
        self.label_turmas_acumuladas_frequencia = tk.Label(
            cartao, text="", bg=COR_CARTAO, fg=COR_TEXTO_FRACO, font=("Segoe UI", 9), justify="left")
        self.label_turmas_acumuladas_frequencia.pack(anchor="w", pady=(8, 0))

        self.label_status_lista_frequencia = tk.Label(cartao, text="", bg=COR_CARTAO, fg=COR_VERDE, font=FONTE_PADRAO)
        self.label_status_lista_frequencia.pack(anchor="w", pady=(8, 0))

    def _atualizar_label_turmas_acumuladas_frequencia(self):
        n = len(self.turmas_acumuladas_frequencia)
        if n == 0:
            self.label_turmas_acumuladas_frequencia.config(text="")
        else:
            nomes_turmas = ", ".join(t["turma"] for t in self.turmas_acumuladas_frequencia)
            self.label_turmas_acumuladas_frequencia.config(
                text=f"{n} turma(s) já adicionada(s), aguardando finalizar: {nomes_turmas}")

    def _limpar_campos_turma_frequencia(self):
        self.campo_freq_turma.set("")
        self.campo_freq_serie.set("")
        self.campo_freq_turno.set("")
        self.campo_freq_professor.set("")
        self.texto_freq_nomes.delete("1.0", "end")

    def _gerar_lista_frequencia(self):
        if not self.config_escola.get("nome_escola"):
            messagebox.showwarning(
                "Dados da escola pendentes",
                "Preencha e salve os Dados da Escola antes de gerar a ficha de frequência.")
            self._ir_para("escola")
            return

        turma = self.campo_freq_turma.get()
        dias_texto = self.campo_freq_dias.get()
        if not turma or not dias_texto:
            messagebox.showwarning(
                "Campos obrigatórios", "Preencha ao menos a Turma e os Dias letivos do mês.")
            return

        try:
            linhas_brancas = int(self.campo_freq_linhas_brancas.get() or "25")
        except ValueError:
            linhas_brancas = 25

        dados_turma_atual = {
            "mes": self.campo_freq_mes.get(),
            "ano_letivo": self.campo_freq_ano.get(),
            "turma": turma,
            "serie": self.campo_freq_serie.get(),
            "turno": self.campo_freq_turno.get(),
            "professor": self.campo_freq_professor.get(),
            "dias_letivos": dias_texto,
            "nomes_alunos": self.texto_freq_nomes.get("1.0", "end").strip(),
            "linhas_em_branco": linhas_brancas,
        }
        self.turmas_acumuladas_frequencia.append(dados_turma_atual)
        self._atualizar_label_turmas_acumuladas_frequencia()

        quer_mais = messagebox.askyesno(
            "Adicionar outra turma?",
            f"Turma \"{turma}\" adicionada ({len(self.turmas_acumuladas_frequencia)} turma(s) até agora).\n\n"
            "Deseja preencher e adicionar outra turma antes de gerar o arquivo?")
        if quer_mais:
            self._limpar_campos_turma_frequencia()
            return

        turmas = self.turmas_acumuladas_frequencia
        if len(turmas) == 1:
            nome_sugerido = f"Ficha de Frequencia - Turma {turmas[0]['turma']}.docx"
        else:
            nome_sugerido = "Ficha de Frequencia - Varias Turmas.docx"

        caminho = filedialog.asksaveasfilename(
            title="Salvar ficha de frequência",
            initialfile=nome_sugerido,
            defaultextension=".docx",
            filetypes=[("Documento Word", "*.docx")])
        if not caminho:
            # cancelou o salvamento - desfaz a adicao dessa turma pra nao
            # duplicar se a pessoa clicar em Gerar de novo
            self.turmas_acumuladas_frequencia.pop()
            self._atualizar_label_turmas_acumuladas_frequencia()
            return

        try:
            if len(turmas) == 1:
                documentos.gerar_lista_frequencia(self.config_escola, turmas[0], caminho)
            else:
                documentos.gerar_lista_frequencia_varias_turmas(self.config_escola, {}, turmas, caminho)
        except Exception as e:
            messagebox.showerror("Erro ao gerar ficha de frequência", str(e))
            return

        self.turmas_acumuladas_frequencia = []
        self._atualizar_label_turmas_acumuladas_frequencia()
        self._limpar_campos_turma_frequencia()

        self.label_status_lista_frequencia.config(text=f"Ficha de frequência gerada: {os.path.basename(caminho)}")
        if messagebox.askyesno(
                "Ficha de frequência gerada",
                "Ficha de frequência gerada com sucesso. Deseja abrir o arquivo agora?"):
            try:
                os.startfile(caminho)
            except Exception:
                pass

    # ---------------- Relatório do Bolsa Família ----------------
    def _montar_pagina_bolsa(self):
        pagina = self.paginas["bolsa"]

        canvas = tk.Canvas(pagina, bg=COR_FUNDO, highlightthickness=0)
        scrollbar = ttk.Scrollbar(pagina, orient="vertical", command=canvas.yview)
        cartao_externo = tk.Frame(canvas, bg=COR_FUNDO)

        cartao_externo.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        janela_canvas = canvas.create_window((0, 0), window=cartao_externo, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(janela_canvas, width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        cartao = tk.Frame(cartao_externo, bg=COR_CARTAO, padx=24, pady=20)
        cartao.pack(fill="x", padx=4, pady=4)

        tk.Label(cartao, text="Relatório do Bolsa Família", bg=COR_CARTAO, fg=COR_AZUL_ESCURO,
                 font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(0, 4))
        tk.Label(cartao,
                 text="Cruza as fichas de frequência (Gier ou por matéria) com a lista de\n"
                      "alunos do Bolsa Família e calcula a % de presença de cada um - mesma\n"
                      "lógica já usada no Relatório de Presença.",
                 bg=COR_CARTAO, fg=COR_TEXTO_FRACO, font=("Segoe UI", 9),
                 justify="left").pack(anchor="w", pady=(0, 16))

        tk.Label(cartao, text="Pastas dos meses (uma pasta por mês, com os PDFs de frequência)",
                 bg=COR_CARTAO, fg=COR_TEXTO, font=FONTE_LABEL).pack(anchor="w")

        moldura_lista = tk.Frame(cartao, bg=COR_CARTAO)
        moldura_lista.pack(fill="x", pady=(4, 0))
        rolagem_pastas = ttk.Scrollbar(moldura_lista)
        rolagem_pastas.pack(side="right", fill="y")
        self.lista_pastas_bolsa = tk.Listbox(
            moldura_lista, height=4, font=FONTE_PADRAO, bg="#FAFBFD", relief="solid", bd=1,
            highlightthickness=1, highlightbackground=COR_BORDA, activestyle="none",
            yscrollcommand=rolagem_pastas.set)
        self.lista_pastas_bolsa.pack(side="left", fill="x", expand=True)
        rolagem_pastas.config(command=self.lista_pastas_bolsa.yview)

        botoes_pastas = tk.Frame(cartao, bg=COR_CARTAO)
        botoes_pastas.pack(fill="x", pady=(6, 16))
        tk.Button(botoes_pastas, text="+ Adicionar pasta do mês...", command=self._adicionar_pasta_bolsa,
                  bg="#E3F2ED", fg=COR_AZUL_ESCURO, font=("Segoe UI", 10, "bold"),
                  activebackground="#CFEAE0", activeforeground=COR_AZUL_ESCURO,
                  relief="flat", padx=14, pady=6, cursor="hand2").pack(side="left")
        tk.Button(botoes_pastas, text="Remover selecionada", command=self._remover_pasta_bolsa,
                  bg=COR_FUNDO, fg="#B3261E", font=("Segoe UI", 10),
                  activebackground="#F5C4C4", activeforeground="#B3261E",
                  relief="flat", padx=14, pady=6, cursor="hand2").pack(side="left", padx=(8, 0))

        tk.Label(cartao, text="Lista de alunos do Bolsa Família (.docx ou formulário do MEC em .pdf)",
                 bg=COR_CARTAO, fg=COR_TEXTO, font=FONTE_LABEL).pack(anchor="w")
        linha_lista = tk.Frame(cartao, bg=COR_CARTAO)
        linha_lista.pack(fill="x", pady=(4, 16))
        self.var_lista_bolsa = tk.StringVar()
        tk.Entry(linha_lista, textvariable=self.var_lista_bolsa, font=FONTE_PADRAO,
                  relief="solid", bd=1, highlightthickness=1,
                  highlightbackground=COR_BORDA).pack(side="left", fill="x", expand=True)
        tk.Button(linha_lista, text="Selecionar arquivo...", command=self._escolher_lista_bolsa,
                  bg=COR_FUNDO, fg=COR_AZUL_ESCURO, font=("Segoe UI", 9, "bold"),
                  activebackground="#E3F2ED", activeforeground=COR_AZUL_ESCURO,
                  relief="flat", padx=10, cursor="hand2").pack(side="left", padx=(8, 0))

        linha_opcoes = tk.Frame(cartao, bg=COR_CARTAO)
        linha_opcoes.pack(fill="x", pady=(0, 4))
        self.campo_percentual_bolsa = CampoTexto(
            linha_opcoes, "Percentual mínimo de presença (%)", "60", largura=8)
        self.campo_percentual_bolsa.pack(side="left", padx=(0, 24))

        frame_formatos = tk.Frame(linha_opcoes, bg=COR_CARTAO)
        frame_formatos.pack(side="left")
        tk.Label(frame_formatos, text="Formatos pra salvar", bg=COR_CARTAO, fg=COR_TEXTO,
                 font=FONTE_LABEL).pack(anchor="w")
        linha_checks = tk.Frame(frame_formatos, bg=COR_CARTAO)
        linha_checks.pack(anchor="w", pady=(3, 0))
        self.var_formato_xlsx_bolsa = tk.BooleanVar(value=True)
        self.var_formato_docx_bolsa = tk.BooleanVar(value=True)
        self.var_formato_pdf_bolsa = tk.BooleanVar(value=True)
        tk.Checkbutton(linha_checks, text="Excel (.xlsx)", variable=self.var_formato_xlsx_bolsa,
                        bg=COR_CARTAO, font=FONTE_PADRAO).pack(side="left")
        tk.Checkbutton(linha_checks, text="Word (.docx)", variable=self.var_formato_docx_bolsa,
                        bg=COR_CARTAO, font=FONTE_PADRAO).pack(side="left", padx=(10, 0))
        tk.Checkbutton(linha_checks, text="PDF (.pdf)", variable=self.var_formato_pdf_bolsa,
                        bg=COR_CARTAO, font=FONTE_PADRAO).pack(side="left", padx=(10, 0))

        self.btn_gerar_bolsa = tk.Button(
            cartao, text="Gerar Relatório do Bolsa Família", command=self._gerar_relatorio_bolsa,
            bg=COR_AZUL, fg="white", font=("Segoe UI", 11, "bold"),
            activebackground=COR_AZUL_ESCURO, activeforeground="white",
            relief="flat", padx=20, pady=10, cursor="hand2")
        self.btn_gerar_bolsa.pack(anchor="w", pady=(16, 10))

        tk.Label(cartao, text="Andamento", bg=COR_CARTAO, fg=COR_TEXTO, font=FONTE_LABEL).pack(anchor="w")
        self.texto_log_bolsa = tk.Text(
            cartao, height=12, font=("Consolas", 9), relief="solid", bd=1,
            highlightthickness=1, highlightbackground=COR_BORDA, wrap="word",
            state="disabled", bg="#FAFBFD")
        self.texto_log_bolsa.pack(fill="x", pady=(4, 0))

    def _adicionar_pasta_bolsa(self):
        pasta = filedialog.askdirectory(title="Escolha a pasta do mês (com os PDFs de frequência)")
        if pasta:
            self.lista_pastas_bolsa.insert("end", pasta)

    def _remover_pasta_bolsa(self):
        for indice in reversed(self.lista_pastas_bolsa.curselection()):
            self.lista_pastas_bolsa.delete(indice)

    def _escolher_lista_bolsa(self):
        caminho = filedialog.askopenfilename(
            title="Escolha a lista de alunos do Bolsa Família",
            filetypes=[("Word ou PDF", "*.docx *.pdf"), ("Documento Word", "*.docx"), ("PDF", "*.pdf")])
        if caminho:
            self.var_lista_bolsa.set(caminho)

    def _log_bolsa(self, mensagem):
        self.texto_log_bolsa.config(state="normal")
        self.texto_log_bolsa.insert("end", str(mensagem) + "\n")
        self.texto_log_bolsa.see("end")
        self.texto_log_bolsa.config(state="disabled")

    def _gerar_relatorio_bolsa(self):
        pastas = list(self.lista_pastas_bolsa.get(0, "end"))
        arquivo_lista = self.var_lista_bolsa.get().strip()
        if not pastas:
            messagebox.showwarning("Campo obrigatório", "Adicione ao menos uma pasta de mês.")
            return
        if not arquivo_lista:
            messagebox.showwarning(
                "Campo obrigatório", "Escolha o arquivo da lista de alunos do Bolsa Família.")
            return
        try:
            percentual_minimo = float(self.campo_percentual_bolsa.get().replace(",", "."))
        except ValueError:
            messagebox.showwarning("Campo inválido", "Preencha um percentual mínimo válido (ex: 60).")
            return

        formatos = []
        if self.var_formato_xlsx_bolsa.get():
            formatos.append("xlsx")
        if self.var_formato_docx_bolsa.get():
            formatos.append("docx")
        if self.var_formato_pdf_bolsa.get():
            formatos.append("pdf")
        if not formatos:
            messagebox.showwarning("Campo obrigatório", "Marque ao menos um formato pra salvar.")
            return

        self.texto_log_bolsa.config(state="normal")
        self.texto_log_bolsa.delete("1.0", "end")
        self.texto_log_bolsa.config(state="disabled")
        self.btn_gerar_bolsa.config(state="disabled", text="Gerando...")

        def log_thread_safe(mensagem):
            self.after(0, self._log_bolsa, mensagem)

        def worker():
            try:
                resultados, meses, resumo = relatorio_bolsa_familia.analisar(
                    pastas, arquivo_lista, percentual_minimo, log=log_thread_safe)
            except Exception as e:
                self.after(0, self._erro_relatorio_bolsa, str(e))
                return
            self.after(0, self._pedir_salvar_relatorio_bolsa,
                       resultados, meses, formatos, percentual_minimo)

        threading.Thread(target=worker, daemon=True).start()

    def _erro_relatorio_bolsa(self, mensagem):
        self.btn_gerar_bolsa.config(state="normal", text="Gerar Relatório do Bolsa Família")
        messagebox.showerror("Erro ao gerar relatório", mensagem)

    def _pedir_salvar_relatorio_bolsa(self, resultados, meses, formatos, percentual_minimo):
        self.btn_gerar_bolsa.config(state="normal", text="Gerar Relatório do Bolsa Família")
        nome_sugerido = relatorio_bolsa_familia.sugerir_nome(meses)
        primeiro_formato = formatos[0]
        caminho = filedialog.asksaveasfilename(
            title="Salvar relatório como",
            initialfile=f"{nome_sugerido}.{primeiro_formato}",
            defaultextension=f".{primeiro_formato}",
            filetypes=[("Documento", f"*.{primeiro_formato}")])
        if not caminho:
            self._log_bolsa("\nSalvamento cancelado.")
            return

        caminho_base = os.path.splitext(caminho)[0]
        try:
            arquivos = relatorio_bolsa_familia.salvar(
                resultados, caminho_base, percentual_minimo, meses, formatos, log=self._log_bolsa)
        except Exception as e:
            messagebox.showerror("Erro ao salvar relatório", str(e))
            return

        self._log_bolsa("\nConcluído.")
        if messagebox.askyesno(
                "Relatório gerado",
                "Relatório gerado com sucesso. Deseja abrir a pasta agora?"):
            try:
                os.startfile(os.path.dirname(arquivos[0]))
            except Exception:
                pass

    # ---------------- Notificação de Excesso de Faltas ----------------
    def _montar_pagina_justificativa_faltas(self):
        pagina = self.paginas["justificativa_faltas"]

        canvas = tk.Canvas(pagina, bg=COR_FUNDO, highlightthickness=0)
        scrollbar = ttk.Scrollbar(pagina, orient="vertical", command=canvas.yview)
        cartao_externo = tk.Frame(canvas, bg=COR_FUNDO)

        cartao_externo.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        janela_canvas = canvas.create_window((0, 0), window=cartao_externo, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(janela_canvas, width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        cartao = tk.Frame(cartao_externo, bg=COR_CARTAO, padx=24, pady=20)
        cartao.pack(fill="x", padx=4, pady=4)

        tk.Label(cartao, text="Nova justificativa de excesso de faltas", bg=COR_CARTAO, fg=COR_AZUL_ESCURO,
                 font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(0, 16))

        tk.Label(cartao,
                 text="Preenchido quando o(a) responsável comparece à escola para justificar\n"
                      "o excesso de faltas do(a) aluno(a).",
                 bg=COR_CARTAO, fg=COR_TEXTO_FRACO, font=("Segoe UI", 9),
                 justify="left").pack(anchor="w", pady=(0, 14))

        self.campo_justificativa_aluno = CampoTexto(cartao, "Nome do aluno", largura=50)
        self.campo_justificativa_aluno.pack(fill="x")

        linha1 = tk.Frame(cartao, bg=COR_CARTAO)
        linha1.pack(fill="x")
        self.campo_justificativa_turma = CampoTexto(linha1, "Turma", largura=12)
        self.campo_justificativa_turma.pack(side="left", padx=(0, 16))
        self.campo_justificativa_turno = CampoTexto(linha1, "Turno", largura=16)
        self.campo_justificativa_turno.pack(side="left", padx=(0, 16))
        self.campo_justificativa_responsavel = CampoTexto(linha1, "Nome do responsável", largura=30)
        self.campo_justificativa_responsavel.pack(side="left")

        linha2 = tk.Frame(cartao, bg=COR_CARTAO)
        linha2.pack(fill="x")
        self.campo_justificativa_etapa = CampoOpcao(
            linha2, "Etapa de ensino",
            [("infantil", "Educação Infantil (mínimo 60%)"),
             ("fundamental", "Ensino Fundamental em diante (mínimo 75%)")], largura=32)
        self.campo_justificativa_etapa.pack(side="left", padx=(0, 16))
        self.campo_justificativa_periodo = CampoTexto(linha2, "Período (ex: Agosto/2026)", largura=18)
        self.campo_justificativa_periodo.pack(side="left")

        linha3 = tk.Frame(cartao, bg=COR_CARTAO)
        linha3.pack(fill="x")
        self.campo_justificativa_dias = CampoTexto(linha3, "Dias letivos no período", largura=14)
        self.campo_justificativa_dias.pack(side="left", padx=(0, 16))
        self.campo_justificativa_faltas = CampoTexto(linha3, "Faltas no período", largura=14)
        self.campo_justificativa_faltas.pack(side="left")

        self.campo_justificativa_motivo = CampoOpcao(
            cartao, "Motivo apresentado pelo(a) responsável",
            [("doenca_aluno", "Doença do(a) aluno(a)"),
             ("doenca_familia", "Doença/problema de saúde na família"),
             ("mudanca_endereco", "Mudança de endereço/dificuldade de acesso"),
             ("transporte", "Dificuldade de transporte"),
             ("trabalho_renda", "Motivo de trabalho/renda familiar"),
             ("outro", "Outro")], largura=40)
        self.campo_justificativa_motivo.pack(fill="x")

        self.campo_justificativa_motivo_outro = CampoTexto(
            cartao, "Descrição do motivo - só quando o motivo é Outro", largura=50)
        self.campo_justificativa_motivo_outro.pack(fill="x")

        self.campo_justificativa_obs = CampoTexto(cartao, "Observações (opcional)", largura=70)
        self.campo_justificativa_obs.pack(fill="x")

        self.campo_data_justificativa = CampoTexto(
            cartao, "Local e data", data_por_extenso((self.config_escola.get("cidade") or "").strip()), largura=50)
        self.campo_data_justificativa.pack(fill="x")
        self._ultima_data_justificativa_auto = self.campo_data_justificativa.get()

        self.campo_justificativa_recebido_por = CampoTexto(
            cartao, "Recebido por (quem atendeu na escola)", self.config_escola.get("diretor_nome", ""),
            largura=50)
        self.campo_justificativa_recebido_por.pack(fill="x")
        self.campo_justificativa_cargo = CampoTexto(
            cartao, "Cargo", self.config_escola.get("diretor_cargo", ""), largura=30)
        self.campo_justificativa_cargo.pack(fill="x")

        btn_gerar = tk.Button(cartao, text="Gerar Justificativa (.docx)", command=self._gerar_justificativa_faltas,
                               bg=COR_AZUL, fg="white", font=("Segoe UI", 11, "bold"),
                               activebackground=COR_AZUL_ESCURO, activeforeground="white",
                               relief="flat", padx=20, pady=10, cursor="hand2")
        btn_gerar.pack(anchor="w", pady=(10, 0))

        self.label_status_justificativa = tk.Label(cartao, text="", bg=COR_CARTAO, fg=COR_VERDE, font=FONTE_PADRAO)
        self.label_status_justificativa.pack(anchor="w", pady=(8, 0))

    def _gerar_justificativa_faltas(self):
        if not self.config_escola.get("nome_escola"):
            messagebox.showwarning(
                "Dados da escola pendentes",
                "Preencha e salve os Dados da Escola antes de gerar a justificativa.")
            self._ir_para("escola")
            return

        aluno = self.campo_justificativa_aluno.get()
        if not aluno:
            messagebox.showwarning("Campo obrigatório", "Preencha o nome do aluno.")
            return

        dados_justificativa = {
            "nome_aluno": aluno,
            "turma": self.campo_justificativa_turma.get(),
            "turno": self.campo_justificativa_turno.get(),
            "nome_responsavel": self.campo_justificativa_responsavel.get(),
            "etapa_ensino": self.campo_justificativa_etapa.get(),
            "periodo": self.campo_justificativa_periodo.get(),
            "dias_letivos": self.campo_justificativa_dias.get(),
            "faltas": self.campo_justificativa_faltas.get(),
            "motivo": self.campo_justificativa_motivo.get(),
            "motivo_outro": self.campo_justificativa_motivo_outro.get(),
            "observacoes": self.campo_justificativa_obs.get(),
            "data": self.campo_data_justificativa.get(),
            "recebido_por": self.campo_justificativa_recebido_por.get(),
            "cargo_recebido_por": self.campo_justificativa_cargo.get(),
        }

        nome_sugerido = f"Justificativa de Faltas - {aluno}.docx"
        caminho = filedialog.asksaveasfilename(
            title="Salvar justificativa de excesso de faltas",
            initialfile=nome_sugerido,
            defaultextension=".docx",
            filetypes=[("Documento Word", "*.docx")])
        if not caminho:
            return

        try:
            documentos.gerar_justificativa_faltas(self.config_escola, dados_justificativa, caminho)
        except Exception as e:
            messagebox.showerror("Erro ao gerar justificativa", str(e))
            return

        self.label_status_justificativa.config(text=f"Justificativa gerada: {os.path.basename(caminho)}")
        if messagebox.askyesno(
                "Justificativa gerada",
                "Justificativa gerada com sucesso. Deseja abrir o arquivo agora?"):
            try:
                os.startfile(caminho)
            except Exception:
                pass

    # ---------------- Certificado de Conclusão ----------------
    def _montar_pagina_certificado(self):
        pagina = self.paginas["certificado"]

        canvas = tk.Canvas(pagina, bg=COR_FUNDO, highlightthickness=0)
        scrollbar = ttk.Scrollbar(pagina, orient="vertical", command=canvas.yview)
        cartao_externo = tk.Frame(canvas, bg=COR_FUNDO)

        cartao_externo.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        janela_canvas = canvas.create_window((0, 0), window=cartao_externo, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(janela_canvas, width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        cartao = tk.Frame(cartao_externo, bg=COR_CARTAO, padx=24, pady=20)
        cartao.pack(fill="x", padx=4, pady=4)

        tk.Label(cartao, text="Novo certificado de conclusão", bg=COR_CARTAO, fg=COR_AZUL_ESCURO,
                 font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(0, 16))

        self.campo_certificado_aluno = CampoTexto(cartao, "Nome do aluno", largura=50)
        self.campo_certificado_aluno.pack(fill="x")

        self.campo_certificado_etapa = CampoTexto(
            cartao, "Etapa concluída (ex: Educação Infantil, 5º Ano do Ensino Fundamental)", largura=50)
        self.campo_certificado_etapa.pack(fill="x")

        linha1 = tk.Frame(cartao, bg=COR_CARTAO)
        linha1.pack(fill="x")
        self.campo_certificado_ano = CampoTexto(
            linha1, "Ano letivo", str(datetime.date.today().year), largura=10)
        self.campo_certificado_ano.pack(side="left", padx=(0, 16))
        self.campo_certificado_turma = CampoTexto(linha1, "Turma (opcional)", largura=14)
        self.campo_certificado_turma.pack(side="left")

        self.campo_data_certificado = CampoTexto(
            cartao, "Local e data", data_por_extenso((self.config_escola.get("cidade") or "").strip()), largura=50)
        self.campo_data_certificado.pack(fill="x")
        self._ultima_data_certificado_auto = self.campo_data_certificado.get()

        self.campo_certificado_assinado_por = CampoTexto(
            cartao, "Assinado por", self.config_escola.get("diretor_nome", ""), largura=50)
        self.campo_certificado_assinado_por.pack(fill="x")
        self.campo_certificado_cargo = CampoTexto(
            cartao, "Cargo", self.config_escola.get("diretor_cargo", ""), largura=30)
        self.campo_certificado_cargo.pack(fill="x")

        tk.Label(cartao,
                 text="Sai com uma borda decorativa ao redor da página, pra dar a cara de\n"
                      "certificado/diploma.",
                 bg=COR_CARTAO, fg=COR_TEXTO_FRACO, font=("Segoe UI", 9),
                 justify="left").pack(anchor="w", pady=(10, 10))

        btn_gerar = tk.Button(cartao, text="Gerar Certificado (.docx)", command=self._gerar_certificado,
                               bg=COR_AZUL, fg="white", font=("Segoe UI", 11, "bold"),
                               activebackground=COR_AZUL_ESCURO, activeforeground="white",
                               relief="flat", padx=20, pady=10, cursor="hand2")
        btn_gerar.pack(anchor="w", pady=(10, 0))

        self.label_status_certificado = tk.Label(cartao, text="", bg=COR_CARTAO, fg=COR_VERDE, font=FONTE_PADRAO)
        self.label_status_certificado.pack(anchor="w", pady=(8, 0))

    def _gerar_certificado(self):
        if not self.config_escola.get("nome_escola"):
            messagebox.showwarning(
                "Dados da escola pendentes",
                "Preencha e salve os Dados da Escola antes de gerar o certificado.")
            self._ir_para("escola")
            return

        aluno = self.campo_certificado_aluno.get()
        if not aluno:
            messagebox.showwarning("Campo obrigatório", "Preencha o nome do aluno.")
            return

        dados_certificado = {
            "nome_aluno": aluno,
            "etapa_concluida": self.campo_certificado_etapa.get(),
            "ano_letivo": self.campo_certificado_ano.get(),
            "turma": self.campo_certificado_turma.get(),
            "data": self.campo_data_certificado.get(),
            "assinado_por": self.campo_certificado_assinado_por.get(),
            "cargo_assinado_por": self.campo_certificado_cargo.get(),
        }

        nome_sugerido = f"Certificado - {aluno}.docx"
        caminho = filedialog.asksaveasfilename(
            title="Salvar certificado",
            initialfile=nome_sugerido,
            defaultextension=".docx",
            filetypes=[("Documento Word", "*.docx")])
        if not caminho:
            return

        try:
            documentos.gerar_certificado(self.config_escola, dados_certificado, caminho)
        except Exception as e:
            messagebox.showerror("Erro ao gerar certificado", str(e))
            return

        self.label_status_certificado.config(text=f"Certificado gerado: {os.path.basename(caminho)}")
        if messagebox.askyesno(
                "Certificado gerado",
                "Certificado gerado com sucesso. Deseja abrir o arquivo agora?"):
            try:
                os.startfile(caminho)
            except Exception:
                pass

    # ---------------- Histórico Escolar ----------------
    def _montar_pagina_historico_escolar(self):
        pagina = self.paginas["historico_escolar"]

        canvas = tk.Canvas(pagina, bg=COR_FUNDO, highlightthickness=0)
        scrollbar = ttk.Scrollbar(pagina, orient="vertical", command=canvas.yview)
        cartao_externo = tk.Frame(canvas, bg=COR_FUNDO)

        cartao_externo.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        janela_canvas = canvas.create_window((0, 0), window=cartao_externo, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(janela_canvas, width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        cartao = tk.Frame(cartao_externo, bg=COR_CARTAO, padx=24, pady=20)
        cartao.pack(fill="x", padx=4, pady=4)

        tk.Label(cartao, text="Novo histórico escolar", bg=COR_CARTAO, fg=COR_AZUL_ESCURO,
                 font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(0, 16))

        self.campo_historico_aluno = CampoTexto(cartao, "Nome do aluno", largura=50)
        self.campo_historico_aluno.pack(fill="x")

        linha1 = tk.Frame(cartao, bg=COR_CARTAO)
        linha1.pack(fill="x")
        self.campo_historico_nascimento = CampoTexto(linha1, "Data de nascimento", largura=16)
        self.campo_historico_nascimento.pack(side="left", padx=(0, 16))
        self.campo_historico_naturalidade = CampoTexto(linha1, "Naturalidade", largura=20)
        self.campo_historico_naturalidade.pack(side="left", padx=(0, 16))
        self.campo_historico_nacionalidade = CampoTexto(linha1, "Nacionalidade", "Brasileira", largura=16)
        self.campo_historico_nacionalidade.pack(side="left")

        linha2 = tk.Frame(cartao, bg=COR_CARTAO)
        linha2.pack(fill="x")
        self.campo_historico_mae = CampoTexto(linha2, "Nome da mãe", largura=30)
        self.campo_historico_mae.pack(side="left", padx=(0, 16))
        self.campo_historico_pai = CampoTexto(linha2, "Nome do pai", largura=30)
        self.campo_historico_pai.pack(side="left")

        separador = tk.Frame(cartao, bg=COR_BORDA, height=1)
        separador.pack(fill="x", pady=(12, 12))

        tk.Label(cartao, text="Registro de escolaridade (um ano letivo por linha)",
                 bg=COR_CARTAO, fg=COR_AZUL_ESCURO, font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 8))

        cabecalho_tabela = tk.Frame(cartao, bg=COR_CARTAO)
        cabecalho_tabela.pack(fill="x")
        for texto, largura in (("Ano Letivo", 8), ("Etapa/Série", 18), ("Turma", 8),
                                ("Carga Horária", 12), ("Resultado Final", 16)):
            tk.Label(cabecalho_tabela, text=texto, bg=COR_CARTAO, fg=COR_TEXTO, font=FONTE_LABEL,
                     width=largura, anchor="w").pack(side="left", padx=(0, 4))

        self.linhas_historico_escolar = []
        frame_linhas = tk.Frame(cartao, bg=COR_CARTAO)
        frame_linhas.pack(fill="x", pady=(3, 0))
        for _ in range(8):
            linha = tk.Frame(frame_linhas, bg=COR_CARTAO)
            linha.pack(fill="x", pady=2)
            campos = {}
            for chave, largura in (("ano_letivo", 8), ("etapa", 18), ("turma", 8),
                                    ("carga_horaria", 12), ("resultado", 16)):
                var = tk.StringVar()
                tk.Entry(linha, textvariable=var, font=FONTE_PADRAO, width=largura,
                          relief="solid", bd=1, highlightthickness=1,
                          highlightbackground=COR_BORDA).pack(side="left", padx=(0, 4))
                campos[chave] = var
            self.linhas_historico_escolar.append(campos)

        self.campo_data_historico = CampoTexto(
            cartao, "Local e data", data_por_extenso((self.config_escola.get("cidade") or "").strip()), largura=50)
        self.campo_data_historico.pack(fill="x", pady=(12, 0))
        self._ultima_data_historico_auto = self.campo_data_historico.get()

        self.campo_historico_assinado_por = CampoTexto(
            cartao, "Assinado por", self.config_escola.get("secretario_nome", ""), largura=50)
        self.campo_historico_assinado_por.pack(fill="x")
        self.campo_historico_cargo = CampoTexto(cartao, "Cargo", "Secretário(a) Escolar", largura=30)
        self.campo_historico_cargo.pack(fill="x")

        btn_gerar = tk.Button(cartao, text="Gerar Histórico Escolar (.docx)", command=self._gerar_historico_escolar,
                               bg=COR_AZUL, fg="white", font=("Segoe UI", 11, "bold"),
                               activebackground=COR_AZUL_ESCURO, activeforeground="white",
                               relief="flat", padx=20, pady=10, cursor="hand2")
        btn_gerar.pack(anchor="w", pady=(16, 0))

        self.label_status_historico = tk.Label(cartao, text="", bg=COR_CARTAO, fg=COR_VERDE, font=FONTE_PADRAO)
        self.label_status_historico.pack(anchor="w", pady=(8, 0))

    def _gerar_historico_escolar(self):
        if not self.config_escola.get("nome_escola"):
            messagebox.showwarning(
                "Dados da escola pendentes",
                "Preencha e salve os Dados da Escola antes de gerar o histórico.")
            self._ir_para("escola")
            return

        aluno = self.campo_historico_aluno.get()
        if not aluno:
            messagebox.showwarning("Campo obrigatório", "Preencha o nome do aluno.")
            return

        registros = []
        for campos in self.linhas_historico_escolar:
            ano = campos["ano_letivo"].get().strip()
            if not ano:
                continue
            registros.append({chave: var.get().strip() for chave, var in campos.items()})

        dados_historico = {
            "nome_aluno": aluno,
            "data_nascimento": self.campo_historico_nascimento.get(),
            "naturalidade": self.campo_historico_naturalidade.get(),
            "nacionalidade": self.campo_historico_nacionalidade.get(),
            "nome_mae": self.campo_historico_mae.get(),
            "nome_pai": self.campo_historico_pai.get(),
            "registros": registros,
            "data": self.campo_data_historico.get(),
            "assinado_por": self.campo_historico_assinado_por.get(),
            "cargo_assinado_por": self.campo_historico_cargo.get(),
        }

        nome_sugerido = f"Historico Escolar - {aluno}.docx"
        caminho = filedialog.asksaveasfilename(
            title="Salvar histórico escolar",
            initialfile=nome_sugerido,
            defaultextension=".docx",
            filetypes=[("Documento Word", "*.docx")])
        if not caminho:
            return

        try:
            documentos.gerar_historico_escolar(self.config_escola, dados_historico, caminho)
        except Exception as e:
            messagebox.showerror("Erro ao gerar histórico escolar", str(e))
            return

        self.label_status_historico.config(text=f"Histórico escolar gerado: {os.path.basename(caminho)}")
        if messagebox.askyesno(
                "Histórico escolar gerado",
                "Histórico escolar gerado com sucesso. Deseja abrir o arquivo agora?"):
            try:
                os.startfile(caminho)
            except Exception:
                pass


if __name__ == "__main__":
    app = App()
    app.mainloop()
