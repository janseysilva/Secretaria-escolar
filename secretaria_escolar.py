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
        tk.Label(self.frame, text=rotulo, bg=COR_CARTAO, fg=COR_TEXTO, font=FONTE_LABEL).pack(anchor="w")
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


class CampoImagem:
    """Label + botao de escolher arquivo + nome do arquivo escolhido."""

    def __init__(self, pai, rotulo, valor_inicial=None):
        self.frame = tk.Frame(pai, bg=COR_CARTAO)
        tk.Label(self.frame, text=rotulo, bg=COR_CARTAO, fg=COR_TEXTO, font=FONTE_LABEL).pack(anchor="w")
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
    ("oficio", "📋", "Ofício", False),
    ("declaracao", "📃", "Declaração Escolar", False),
    ("bolsa", "🎓", "Relatório do Bolsa Família", False),
    ("capa_livro", "📚", "Capa de Abertura de Livro", False),
    ("ficha_matricula", "📝", "Ficha de Matrícula", False),
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
        for nome in ("home", "escola", "memorando"):
            frame = tk.Frame(self.container, bg=COR_FUNDO)
            frame.grid(row=0, column=0, sticky="nsew")
            self.paginas[nome] = frame

        self._montar_pagina_home()
        self._montar_pagina_escola()
        self._montar_pagina_memorando()

    def _ir_para(self, nome, primeira_vez=False):
        if nome == "home":
            self._atualizar_saudacao_home()
        if nome == "escola":
            if primeira_vez:
                self._banner_boas_vindas.pack(fill="x", pady=(0, 10), before=self._linha_voltar_escola)
            else:
                self._banner_boas_vindas.pack_forget()
        self.paginas[nome].tkraise()

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


if __name__ == "__main__":
    app = App()
    app.mainloop()
