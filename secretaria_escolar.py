"""App de Secretaria Escolar - gera memorandos, ofícios e outros documentos."""

import json
import os
import shutil
import sys
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import documentos

PASTA_BASE = os.path.dirname(os.path.abspath(__file__))
PASTA_DADOS = os.path.join(PASTA_BASE, "dados_escola")
PASTA_IMAGENS = os.path.join(PASTA_DADOS, "imagens")
ARQUIVO_CONFIG = os.path.join(PASTA_DADOS, "config.json")

os.makedirs(PASTA_IMAGENS, exist_ok=True)

COR_FUNDO = "#EEF1F7"
COR_CARTAO = "#FFFFFF"
COR_AZUL = "#4472C4"
COR_AZUL_ESCURO = "#2F528F"
COR_AZUL_CLARO = "#5B8AD9"
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
                         bg="#E8ECF4", fg=COR_AZUL_ESCURO, font=("Segoe UI", 9),
                         relief="flat", padx=10, pady=3, cursor="hand2")
        btn.pack(side="left", padx=(0, 6))
        btn_remover = tk.Button(linha, text="Remover", command=self._remover,
                                 bg="#FBE5E5", fg="#B3261E", font=("Segoe UI", 9),
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
    ("memorando", "📄", "Memorando", True),
    ("oficio", "📋", "Ofício", True),
    ("declaracao", "📃", "Declaração Escolar", True),
    ("bolsa", "🎓", "Relatório do Bolsa Família", False),
    ("capa_livro", "📚", "Capa de Abertura de Livro", True),
    ("ficha_matricula", "📝", "Ficha de Matrícula", True),
    ("lista_reuniao", "👥", "Lista de Reunião de Pais e Alunos", False),
    ("lista_frequencia", "🗓️", "Lista de Frequência Escolar", False),
]


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Secretaria Escolar")
        self.geometry("880x720")
        self.configure(bg=COR_FUNDO)

        self.config_escola = carregar_config()

        self._montar_topo()
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
                 fg="#C7D5EE", font=("Segoe UI", 9)).pack(side="left")
        tk.Button(faixa, text="⚙ Dados da escola", command=lambda: self._ir_para("escola"),
                  bg=COR_AZUL_ESCURO, fg="#C7D5EE", font=("Segoe UI", 9), relief="flat",
                  activebackground=COR_AZUL, activeforeground="white",
                  cursor="hand2", bd=0).pack(side="right", padx=20)

    def _montar_paginas(self):
        self.container = tk.Frame(self, bg=COR_FUNDO)
        self.container.pack(fill="both", expand=True, padx=14, pady=14)
        self.container.rowconfigure(0, weight=1)
        self.container.columnconfigure(0, weight=1)

        self.paginas = {}
        for nome in ("home", "escola", "memorando", "oficio", "declaracao", "capa_livro", "ficha_matricula"):
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

    def _ir_para(self, nome, primeira_vez=False):
        if nome == "home":
            self._atualizar_saudacao_home()
        if nome == "escola":
            if primeira_vez:
                self._banner_boas_vindas.pack(fill="x", pady=(0, 10), before=self._linha_voltar_escola)
            else:
                self._banner_boas_vindas.pack_forget()
        if nome == "oficio":
            self._atualizar_data_oficio()
        if nome == "declaracao":
            self._atualizar_data_declaracao()
        if nome == "capa_livro":
            self._atualizar_data_capa_livro()
        self.paginas[nome].tkraise()

    def _atualizar_data_oficio(self):
        # So reescreve se o campo ainda estiver com o ultimo valor que o
        # proprio app preencheu sozinho - se o usuario editou a mao, nao mexe.
        if self.campo_data_oficio.get() == getattr(self, "_ultima_data_oficio_auto", None):
            cidade = (self.config_escola.get("cidade") or "").strip()
            novo_valor = data_por_extenso(cidade)
            self.campo_data_oficio.set(novo_valor)
            self._ultima_data_oficio_auto = novo_valor

    def _atualizar_data_declaracao(self):
        if self.campo_data_declaracao.get() == getattr(self, "_ultima_data_declaracao_auto", None):
            cidade = (self.config_escola.get("cidade") or "").strip()
            novo_valor = data_por_extenso(cidade)
            self.campo_data_declaracao.set(novo_valor)
            self._ultima_data_declaracao_auto = novo_valor

    def _atualizar_data_capa_livro(self):
        if self.campo_data_capa_livro.get() == getattr(self, "_ultima_data_capa_livro_auto", None):
            cidade = (self.config_escola.get("cidade") or "").strip()
            novo_valor = data_por_extenso(cidade)
            self.campo_data_capa_livro.set(novo_valor)
            self._ultima_data_capa_livro_auto = novo_valor

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
        for col in range(2):
            grade.columnconfigure(col, weight=1)

        for i, (chave, emoji, titulo, disponivel) in enumerate(DOCUMENTOS_DISPONIVEIS):
            linha, col = divmod(i, 2)
            self._criar_cartao_documento(grade, chave, emoji, titulo, disponivel).grid(
                row=linha, column=col, sticky="nsew", padx=8, pady=8)

    def _criar_cartao_documento(self, pai, chave, emoji, titulo, disponivel):
        cartao = tk.Frame(pai, bg=COR_CARTAO, padx=18, pady=16, cursor="hand2" if disponivel else "arrow")
        tk.Label(cartao, text=emoji, bg=COR_CARTAO, font=("Segoe UI", 26)).pack(anchor="w")
        tk.Label(cartao, text=titulo, bg=COR_CARTAO,
                 fg=COR_TEXTO if disponivel else COR_TEXTO_FRACO,
                 font=("Segoe UI", 11, "bold"), wraplength=280, justify="left").pack(anchor="w", pady=(6, 0))
        tk.Label(cartao, text="Disponível" if disponivel else "Em breve", bg=COR_CARTAO,
                 fg=COR_VERDE if disponivel else COR_TEXTO_FRACO, font=("Segoe UI", 9)).pack(anchor="w", pady=(4, 0))

        if disponivel:
            def abrir(_e=None):
                self._ir_para(chave)
            for widget in (cartao, *cartao.winfo_children()):
                widget.bind("<Button-1>", abrir)
        else:
            def avisar(_e=None):
                messagebox.showinfo(titulo, f"{titulo} ainda não está disponível — em breve!")
            for widget in (cartao, *cartao.winfo_children()):
                widget.bind("<Button-1>", avisar)
        return cartao

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

        self._linha_voltar_escola = tk.Frame(pagina, bg=COR_FUNDO)
        self._linha_voltar_escola.pack(fill="x", pady=(0, 4))
        tk.Button(self._linha_voltar_escola, text="← Voltar ao início", command=lambda: self._ir_para("home"),
                  bg=COR_FUNDO, fg=COR_AZUL_ESCURO, font=("Segoe UI", 9, "bold"),
                  relief="flat", cursor="hand2", bd=0).pack(anchor="w")

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

        tk.Button(pagina, text="← Voltar ao início", command=lambda: self._ir_para("home"),
                  bg=COR_FUNDO, fg=COR_AZUL_ESCURO, font=("Segoe UI", 9, "bold"),
                  relief="flat", cursor="hand2", bd=0).pack(anchor="w", pady=(0, 4))

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

        tk.Button(pagina, text="← Voltar ao início", command=lambda: self._ir_para("home"),
                  bg=COR_FUNDO, fg=COR_AZUL_ESCURO, font=("Segoe UI", 9, "bold"),
                  relief="flat", cursor="hand2", bd=0).pack(anchor="w", pady=(0, 4))

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

        tk.Button(pagina, text="← Voltar ao início", command=lambda: self._ir_para("home"),
                  bg=COR_FUNDO, fg=COR_AZUL_ESCURO, font=("Segoe UI", 9, "bold"),
                  relief="flat", cursor="hand2", bd=0).pack(anchor="w", pady=(0, 4))

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

        tk.Button(pagina, text="← Voltar ao início", command=lambda: self._ir_para("home"),
                  bg=COR_FUNDO, fg=COR_AZUL_ESCURO, font=("Segoe UI", 9, "bold"),
                  relief="flat", cursor="hand2", bd=0).pack(anchor="w", pady=(0, 4))

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

        tk.Button(pagina, text="← Voltar ao início", command=lambda: self._ir_para("home"),
                  bg=COR_FUNDO, fg=COR_AZUL_ESCURO, font=("Segoe UI", 9, "bold"),
                  relief="flat", cursor="hand2", bd=0).pack(anchor="w", pady=(0, 4))

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


if __name__ == "__main__":
    app = App()
    app.mainloop()
