"""Geração de documentos oficiais de secretaria escolar (memorando, ofício, etc.)."""

import docx
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.table import _Cell

FONTE = "Arial"
TAM_NORMAL = 12
TAM_TITULO = 16

LARGURAS_COLUNAS_CM = [3.8, 7.4, 3.8, 3.5]  # soma = 18.5cm = largura util da pagina A4 (margens estreitas)
ALTURA_MINIMA_CABECALHO_CM = 1.7
ALTURA_MINIMA_DEPARA_CM = 1.3
ALTURA_MINIMA_ASSUNTO_CM = 1.8
ALTURA_MINIMA_CORPO_CM = 11.0  # garante que a pagina fique cheia mesmo com pouco texto
                                # (deixa espaco pro rodape crescer quando tem imagem de assinatura)
CHARS_POR_LINHA_CORPO = 85  # estimativa de quantos caracteres cabem numa linha do corpo
ALTURA_LINHA_CM = 0.53  # altura aproximada de uma linha de texto a 12pt


def _estimar_linhas(texto, chars_por_linha=CHARS_POR_LINHA_CORPO):
    """Estimativa grosseira de quantas linhas um paragrafo vai ocupar quando
    quebrado automaticamente pelo Word - usada so pra calcular quanto espaco
    em branco falta antes do fecho, nao precisa ser exata."""
    texto = texto.strip()
    if not texto:
        return 0
    return max(1, -(-len(texto) // chars_por_linha))


def _definir_margens_celulas(tabela, cima_cm=0.25, baixo_cm=0.25, esquerda_cm=0.2, direita_cm=0.2):
    """Da mais respiro interno as celulas (por padrao o python-docx nao coloca quase nada)."""
    tblPr = tabela._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for nome, valor in (("top", cima_cm), ("bottom", baixo_cm), ("left", esquerda_cm), ("right", direita_cm)):
        el = OxmlElement(f"w:{nome}")
        el.set(qn("w:w"), str(int(Cm(valor).twips)))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)


def _altura_minima(linha_tabela, altura_cm):
    linha_tabela.height = Cm(altura_cm)
    linha_tabela.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST


def _config_secao(document):
    secao = document.sections[0]
    secao.page_width = Cm(21)
    secao.page_height = Cm(29.7)
    secao.left_margin = Cm(1.5)
    secao.right_margin = Cm(1.0)
    secao.top_margin = Cm(1.8)
    secao.bottom_margin = Cm(1.0)


def _config_secao_paisagem(document):
    """Pagina A4 na horizontal (paisagem) - usada em documentos com muitas
    colunas, como a Ficha de Frequencia (uma coluna por dia letivo)."""
    secao = document.sections[0]
    secao.page_width = Cm(29.7)
    secao.page_height = Cm(21)
    secao.left_margin = Cm(1.5)
    secao.right_margin = Cm(1.0)
    secao.top_margin = Cm(0.7)
    secao.bottom_margin = Cm(0.6)


def _config_secao_certificado(document):
    """Pagina A4 na horizontal (paisagem), com margens mais generosas do
    que _config_secao_paisagem - o Certificado sempre sai em paisagem,
    formato tradicional de diploma."""
    secao = document.sections[0]
    secao.page_width = Cm(29.7)
    secao.page_height = Cm(21)
    secao.left_margin = Cm(2.5)
    secao.right_margin = Cm(2.5)
    secao.top_margin = Cm(2.0)
    secao.bottom_margin = Cm(2.0)


def _remover_bordas_tabela(tabela):
    tbl = tabela._tbl
    tblPr = tbl.tblPr
    bordas = OxmlElement("w:tblBorders")
    for nome in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{nome}")
        el.set(qn("w:val"), "nil")
        bordas.append(el)
    tblPr.append(bordas)


def _definir_largura_colunas(tabela, larguras_cm):
    tabela.autofit = False
    tabela.allow_autofit = False
    for linha in tabela.rows:
        for celula, largura in zip(linha.cells, larguras_cm):
            celula.width = Cm(largura)
    for idx, largura in enumerate(larguras_cm):
        tabela.columns[idx].width = Cm(largura)


def _set_fonte_padrao(document):
    estilo = document.styles["Normal"]
    estilo.font.name = FONTE
    estilo.font.size = Pt(TAM_NORMAL)
    estilo.paragraph_format.space_before = Pt(0)
    estilo.paragraph_format.space_after = Pt(0)


def _paragrafo(celula_ou_doc, texto="", negrito=False, tamanho=TAM_NORMAL,
               alinhamento=None, indice=0):
    # So reaproveita o primeiro paragrafo em branco quando o destino e uma
    # CELULA de tabela recem-criada (que sempre vem com 1 paragrafo vazio) -
    # nunca no documento inteiro, senao "document.paragraphs[0]" seria o
    # primeiro paragrafo em branco de TODO o documento (nao o mais recente),
    # fazendo o texto aparecer fora de ordem, antes de conteudo ja escrito.
    if isinstance(celula_ou_doc, _Cell):
        if indice == 0 and celula_ou_doc.paragraphs and not celula_ou_doc.paragraphs[0].runs:
            p = celula_ou_doc.paragraphs[0]
        else:
            p = celula_ou_doc.add_paragraph()
    else:
        p = celula_ou_doc.add_paragraph()
    if alinhamento is not None:
        p.alignment = alinhamento
    if texto:
        run = p.add_run(texto)
        run.font.name = FONTE
        run.font.size = Pt(tamanho)
        run.bold = negrito
    return p


def _impedir_quebra_de_linha(linha_tabela):
    """Evita que uma linha de tabela seja partida entre duas páginas."""
    trPr = linha_tabela._tr.get_or_add_trPr()
    cant_split = OxmlElement("w:cantSplit")
    trPr.append(cant_split)


def _imagem_centralizada(destino, caminho_imagem, largura_cm):
    if not caminho_imagem:
        return None
    p = destino.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(caminho_imagem, width=Cm(largura_cm))
    return p


def _celula_vazia(celula):
    """Centraliza verticalmente uma celula recem-criada (ja vem sem
    conteudo, nao precisa limpar nada - so ajustar o alinhamento).

    Nota: NAO fazer `paragrafo.text = ""` aqui - no python-docx 1.2.0 isso
    cria um run vazio de verdade, o que faz _paragrafo() achar que a celula
    ja tem conteudo e pular pra um paragrafo novo, deixando uma linha em
    branco indesejada antes do primeiro texto real."""
    celula.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    return celula


def _linhas_cabecalho(dados_escola):
    """Linhas de texto do cabecalho que aparece abaixo do brasao/logo - nome
    da escola e secretaria/orgao superior, iguais no Memorando, no Oficio e
    na Declaracao."""
    linhas = []
    nome_escola = (dados_escola.get("nome_escola") or "").strip()
    if nome_escola:
        linhas.append(nome_escola)
    secretaria = (dados_escola.get("secretaria") or "").strip()
    if secretaria:
        linhas.append(secretaria)
    return linhas


def _linhas_contato(dados_escola):
    """Linhas de contato da escola - endereco numa linha, telefone e email
    juntos na linha de baixo - pra completar o cabecalho do Memorando, do
    Oficio e do Termo de Abertura (a Declaracao ja mostra esses dados numa
    tabela propria, entao nao usa isso aqui, evita duplicar)."""
    linhas = []
    endereco = (dados_escola.get("endereco") or "").strip()
    if endereco:
        linhas.append(endereco)
    telefone = (dados_escola.get("telefone") or "").strip()
    email = (dados_escola.get("email") or "").strip()
    contato = []
    if telefone:
        contato.append(f"Telefone: {telefone}")
    if email:
        contato.append(f"E-mail: {email}")
    if contato:
        linhas.append("   ".join(contato))
    return linhas


def _cabecalho_logo_e_secretaria(document, dados_escola, largura_total_cm=16.0):
    """Logo a esquerda + nome da escola/secretaria a direita (maiusculo),
    igual ao timbre oficial da SEMED - usado no Oficio, no Termo de
    Abertura e na Declaracao. Retorna quantas linhas de texto foram
    escritas, pra entrar no calculo de quanto empurrar o rodape pra baixo.
    largura_total_cm: largura do bloco todo (logo + texto) - o padrao 16cm
    e o usado nos documentos em retrato; documentos em paisagem (mais
    largos) passam a largura util da pagina pra manter o timbre
    proporcional em vez de ficar torto pra esquerda."""
    logo_path = dados_escola.get("logo_path")
    linhas_cab = _linhas_cabecalho(dados_escola) or ["Secretaria Municipal de Educação"]
    largura_logo = largura_total_cm * 2.5 / 16.0
    largura_sec = largura_total_cm - largura_logo
    if logo_path:
        cab = document.add_table(rows=1, cols=2)
        _remover_bordas_tabela(cab)
        _definir_largura_colunas(cab, [largura_logo, largura_sec])
        cel_logo, cel_sec = cab.rows[0].cells
        _imagem_centralizada(cel_logo, logo_path, 2.0)
        cel_sec.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        for i, linha in enumerate(linhas_cab):
            _paragrafo(cel_sec, linha.upper(), negrito=(i == 0), tamanho=13 if i == 0 else TAM_NORMAL,
                       alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    else:
        for i, linha in enumerate(linhas_cab):
            _paragrafo(document, linha.upper(), negrito=(i == 0), tamanho=13 if i == 0 else TAM_NORMAL,
                       alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    return len(linhas_cab)


def _tabela_dados_escola(document, dados_escola):
    """Tabela com bordas mostrando os dados de identificacao da escola
    (CMEI/Escola Municipal, Endereco, Telefone, Email) - um por linha,
    igual ao modelo oficial da SEMED. Usada no Oficio, no Termo de Abertura
    e na Declaracao. Retorna quantas linhas a tabela ocupou."""
    linhas_info = [
        ("CMEI/Escola Municipal", dados_escola.get("nome_escola", "")),
        ("Endereço", dados_escola.get("endereco", "")),
        ("Telefone", dados_escola.get("telefone", "")),
        ("Email", dados_escola.get("email", "")),
    ]
    tabela_info = document.add_table(rows=len(linhas_info), cols=1)
    tabela_info.style = "Table Grid"
    for linha_tabela, (rotulo, valor) in zip(tabela_info.rows, linhas_info):
        celula = _celula_vazia(linha_tabela.cells[0])
        celula.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = celula.paragraphs[0]
        r1 = p.add_run(f"{rotulo}: ")
        r1.bold = True
        r1.font.name = FONTE
        r1.font.size = Pt(TAM_NORMAL)
        r2 = p.add_run(valor)
        r2.font.name = FONTE
        r2.font.size = Pt(TAM_NORMAL)
    return len(linhas_info)


def gerar_memorando(dados_escola, dados_memo, caminho_saida):
    """Gera um Memorando (.docx) a partir dos dados da escola e do memorando.

    Todo o documento fica em uma unica tabela com bordas visiveis, separando
    cada bloco (cabecalho, DE/PARA, ASSUNTO, corpo, rodape) em quadros, igual
    ao modelo de memorando usado nas secretarias escolares. Recebido por e a
    segunda Data ficam em branco no documento, para preencher a mao. O resto
    ja vem preenchido pelo que o usuario digitou/cadastrou.

    dados_escola: dict com nome_escola, secretaria, diretor_nome,
        diretor_cargo, diretor_portaria, logo_path, assinatura_path.
    dados_memo: dict com numero, ano, para, assunto, saudacao, corpo (texto
        com quebras de linha), fecho, assinado_por (nome de quem assina -
        so aparece na caixa ASSINATURA quando a escola nao tem uma imagem de
        assinatura cadastrada), protocolo (texto livre, ex: "SIGED Nº 12345"
        - se vazio, a linha de protocolo nem aparece no cabecalho), data.
    """
    document = docx.Document()
    _config_secao(document)
    _set_fonte_padrao(document)

    protocolo = (dados_memo.get("protocolo") or "").strip()
    usa_protocolo = bool(protocolo)
    linhas_cabecalho = 2 if usa_protocolo else 1
    total_linhas = linhas_cabecalho + 1 + 1 + 1 + 1  # cabecalho + DE/PARA + ASSUNTO + corpo + rodape

    tabela = document.add_table(rows=total_linhas, cols=4)
    tabela.style = "Table Grid"
    _definir_largura_colunas(tabela, LARGURAS_COLUNAS_CM)
    _definir_margens_celulas(tabela)

    for linha in tabela.rows[:linhas_cabecalho]:
        _altura_minima(linha, ALTURA_MINIMA_CABECALHO_CM)

    # --- Cabecalho: logo + titulo (+ protocolo, se a escola usar) ---
    if usa_protocolo:
        cel_logo = _celula_vazia(tabela.cell(0, 0).merge(tabela.cell(1, 0)))
        cel_titulo = _celula_vazia(tabela.cell(0, 1).merge(tabela.cell(0, 3)))
        cel_protocolo = _celula_vazia(tabela.cell(1, 1).merge(tabela.cell(1, 3)))
    else:
        cel_logo = _celula_vazia(tabela.cell(0, 0))
        cel_titulo = _celula_vazia(tabela.cell(0, 1).merge(tabela.cell(0, 3)))

    if dados_escola.get("logo_path"):
        _imagem_centralizada(cel_logo, dados_escola["logo_path"], 2.8)

    for i, linha in enumerate(_linhas_cabecalho(dados_escola)):
        _paragrafo(cel_titulo, linha, negrito=(i == 0), tamanho=13 if i == 0 else TAM_NORMAL)
    for linha in _linhas_contato(dados_escola):
        _paragrafo(cel_titulo, linha, tamanho=TAM_NORMAL - 1)

    secretaria = (dados_escola.get("secretaria") or "").strip()
    numero = (dados_memo.get("numero") or "").strip()
    ano = (dados_memo.get("ano") or "").strip()
    titulo_memo = f"MEMORANDO Nº {numero}/{ano}" if numero else "MEMORANDO Nº"
    if secretaria:
        titulo_memo += f" - {secretaria}"
    _paragrafo(cel_titulo, titulo_memo, negrito=True, tamanho=TAM_TITULO)

    if usa_protocolo:
        _paragrafo(cel_protocolo, protocolo, negrito=True)

    proxima_linha = linhas_cabecalho

    # --- DE / PARA ---
    _altura_minima(tabela.rows[proxima_linha], ALTURA_MINIMA_DEPARA_CM)
    cel_de = _celula_vazia(tabela.cell(proxima_linha, 0).merge(tabela.cell(proxima_linha, 1)))
    cel_para = _celula_vazia(tabela.cell(proxima_linha, 2).merge(tabela.cell(proxima_linha, 3)))
    _paragrafo(cel_de, f"DE: {dados_escola.get('nome_escola', '')}", negrito=True)
    para = (dados_memo.get("para") or "").strip()
    _paragrafo(cel_para, f"PARA: {para}" if para else "PARA:", negrito=True)
    proxima_linha += 1

    # --- ASSUNTO ---
    _altura_minima(tabela.rows[proxima_linha], ALTURA_MINIMA_ASSUNTO_CM)
    cel_assunto = _celula_vazia(tabela.cell(proxima_linha, 0).merge(tabela.cell(proxima_linha, 3)))
    cel_assunto.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    p_assunto = cel_assunto.paragraphs[0]
    r1 = p_assunto.add_run("ASSUNTO: ")
    r1.bold = True
    r1.font.name = FONTE
    r1.font.size = Pt(TAM_NORMAL)
    assunto = (dados_memo.get("assunto") or "").strip()
    if assunto:
        r2 = p_assunto.add_run(assunto.upper())
        r2.font.name = FONTE
        r2.font.size = Pt(TAM_NORMAL)
    proxima_linha += 1

    # --- Corpo do memorando ---
    _altura_minima(tabela.rows[proxima_linha], ALTURA_MINIMA_CORPO_CM)
    cel_corpo = _celula_vazia(tabela.cell(proxima_linha, 0).merge(tabela.cell(proxima_linha, 3)))
    cel_corpo.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    p_saud = cel_corpo.paragraphs[0]
    p_saud.paragraph_format.first_line_indent = Cm(1.25)
    run_saud = p_saud.add_run(dados_memo.get("saudacao") or "Prezado(a) Senhor(a),")
    run_saud.font.name = FONTE
    run_saud.font.size = Pt(TAM_NORMAL)
    linhas_usadas = 1

    cel_corpo.add_paragraph()
    linhas_usadas += 1

    corpo = dados_memo.get("corpo") or ""
    paragrafos_corpo = [linha.strip() for linha in corpo.split("\n") if linha.strip()]
    for linha in paragrafos_corpo:
        p = cel_corpo.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Cm(1.25)
        run = p.add_run(linha)
        run.font.name = FONTE
        run.font.size = Pt(TAM_NORMAL)
        linhas_usadas += _estimar_linhas(linha)

    cel_corpo.add_paragraph()
    linhas_usadas += 1

    # Empurra o fecho pra mais perto do rodape (como no modelo original),
    # preenchendo com linhas em branco o espaco que sobrar dentro da altura
    # minima do corpo - sem isso, textos curtos deixam o fecho "colado" logo
    # depois do ultimo paragrafo, bem no topo da caixa.
    altura_util_corpo_cm = ALTURA_MINIMA_CORPO_CM - 0.5  # desconta margens da celula
    linhas_disponiveis = altura_util_corpo_cm / ALTURA_LINHA_CM
    linhas_preenchimento = int(linhas_disponiveis - linhas_usadas - 1)
    linhas_preenchimento = max(0, min(linhas_preenchimento, 20))
    for _ in range(linhas_preenchimento):
        cel_corpo.add_paragraph()

    p_fecho = cel_corpo.add_paragraph()
    p_fecho.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_fecho = p_fecho.add_run(dados_memo.get("fecho") or "Atenciosamente,")
    run_fecho.font.name = FONTE
    run_fecho.font.size = Pt(TAM_NORMAL)
    proxima_linha += 1

    # --- Rodape: DATA | ASSINATURA | RECEBIDO POR | DATA ---
    linha_rodape = tabela.rows[proxima_linha]
    _impedir_quebra_de_linha(linha_rodape)
    cel_data, cel_assinatura, cel_recebido, cel_data2 = linha_rodape.cells
    for c in (cel_data, cel_assinatura, cel_recebido, cel_data2):
        _celula_vazia(c)

    # RECEBIDO POR e a segunda DATA ficam em branco de proposito: sao
    # preenchidos a mao na hora de entregar/receber o memorando.
    _paragrafo(cel_data, "DATA", negrito=True, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    data_memo = (dados_memo.get("data") or "").strip()
    if data_memo:
        _paragrafo(cel_data, data_memo, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

    _paragrafo(cel_assinatura, "ASSINATURA", negrito=True, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    if dados_escola.get("assinatura_path"):
        _imagem_centralizada(cel_assinatura, dados_escola["assinatura_path"], 3.0)
    else:
        # Sem imagem cadastrada: usa o nome digitado como alternativa.
        assinado_por = (dados_memo.get("assinado_por") or "").strip()
        if assinado_por:
            cel_assinatura.add_paragraph()
            _paragrafo(cel_assinatura, assinado_por, negrito=True, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

    _paragrafo(cel_recebido, "RECEBIDO POR", negrito=True, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    _paragrafo(cel_data2, "DATA", negrito=True, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

    for c in (cel_data, cel_assinatura, cel_recebido, cel_data2):
        c.add_paragraph()
        c.add_paragraph()
        c.add_paragraph()

    document.save(caminho_saida)
    return caminho_saida


def _config_secao_oficio(document):
    secao = document.sections[0]
    secao.page_width = Cm(21)
    secao.page_height = Cm(29.7)
    secao.left_margin = Cm(3.0)
    secao.right_margin = Cm(2.0)
    secao.top_margin = Cm(2.0)
    secao.bottom_margin = Cm(2.0)


ALTURA_UTIL_PAGINA_OFICIO_CM = 29.7 - 2.0 - 2.0  # altura da pagina menos margens topo/rodape


def gerar_oficio(dados_escola, dados_oficio, caminho_saida):
    """Gera um Ofício (.docx) no "padrão ofício" - uma carta corrida, sem
    quadros/caixas (diferente do Memorando, que usa uma tabela com bordas).
    Cabeçalho (logo + nome da escola), título, local e data por extenso,
    destinatário, assunto, corpo e fecho/assinatura, igual ao modelo oficial
    usado pela administração pública.

    dados_escola: dict com nome_escola, secretaria, logo_path,
        assinatura_path.
    dados_oficio: dict com numero, ano, protocolo (texto livre, opcional),
        data (ja formatada por extenso, ex: "14 de setembro de 2026"), para
        (nome do destinatario), cargo_destinatario (opcional), assunto,
        saudacao, corpo (texto com quebras de linha), fecho, assinado_por
        (nome de quem assina - so aparece se a escola nao tem imagem de
        assinatura cadastrada), cargo_assinado_por (opcional).
    """
    document = docx.Document()
    _config_secao_oficio(document)
    _set_fonte_padrao(document)
    linhas_usadas = 0

    # --- Cabecalho: logo + secretaria, e tabela com os dados da escola ---
    linhas_usadas += _cabecalho_logo_e_secretaria(document, dados_escola) + 2
    document.add_paragraph()
    linhas_usadas += _tabela_dados_escola(document, dados_escola) + 1
    secretaria = (dados_escola.get("secretaria") or "").strip()

    document.add_paragraph()
    linhas_usadas += 1

    # --- Titulo + protocolo ---
    numero = (dados_oficio.get("numero") or "").strip()
    ano = (dados_oficio.get("ano") or "").strip()
    titulo = f"OFÍCIO Nº {numero}/{ano}" if numero else "OFÍCIO Nº"
    if secretaria:
        titulo += f" - {secretaria}"
    _paragrafo(document, titulo, negrito=True, tamanho=TAM_TITULO)
    linhas_usadas += 1

    protocolo = (dados_oficio.get("protocolo") or "").strip()
    if protocolo:
        _paragrafo(document, protocolo, negrito=True)
        linhas_usadas += 1

    document.add_paragraph()
    linhas_usadas += 1

    # --- Destinatario ---
    para = (dados_oficio.get("para") or "").strip()
    if para:
        _paragrafo(document, f"A Sua Senhoria o(a) Senhor(a)")
        _paragrafo(document, para, negrito=True)
        linhas_usadas += 2
        cargo_dest = (dados_oficio.get("cargo_destinatario") or "").strip()
        if cargo_dest:
            _paragrafo(document, cargo_dest)
            linhas_usadas += 1

    document.add_paragraph()
    linhas_usadas += 1

    # --- Assunto ---
    p_assunto = document.add_paragraph()
    r1 = p_assunto.add_run("Assunto: ")
    r1.bold = True
    r1.font.name = FONTE
    r1.font.size = Pt(TAM_NORMAL)
    assunto = (dados_oficio.get("assunto") or "").strip()
    if assunto:
        r2 = p_assunto.add_run(assunto)
        r2.font.name = FONTE
        r2.font.size = Pt(TAM_NORMAL)
    linhas_usadas += _estimar_linhas("Assunto: " + assunto)

    document.add_paragraph()
    document.add_paragraph()
    linhas_usadas += 2

    # --- Corpo ---
    p_saud = document.add_paragraph()
    p_saud.paragraph_format.first_line_indent = Cm(1.25)
    run_saud = p_saud.add_run(dados_oficio.get("saudacao") or "Prezado(a) Senhor(a),")
    run_saud.font.name = FONTE
    run_saud.font.size = Pt(TAM_NORMAL)
    linhas_usadas += 1

    document.add_paragraph()
    linhas_usadas += 1

    corpo = dados_oficio.get("corpo") or ""
    for linha in corpo.split("\n"):
        linha = linha.strip()
        if not linha:
            continue
        p = document.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Cm(1.25)
        run = p.add_run(linha)
        run.font.name = FONTE
        run.font.size = Pt(TAM_NORMAL)
        linhas_usadas += _estimar_linhas(linha)

    # Empurra o fecho/assinatura pra mais perto do fim da pagina, preenchendo
    # com linhas em branco o espaco que sobrar - textos curtos empurram
    # bastante, textos longos empurram pouco ou nada.
    linhas_reservadas_assinatura = 15 if dados_escola.get("assinatura_path") else 11
    linhas_disponiveis = ALTURA_UTIL_PAGINA_OFICIO_CM / ALTURA_LINHA_CM
    linhas_preenchimento = int(linhas_disponiveis - linhas_usadas - linhas_reservadas_assinatura)
    linhas_preenchimento = max(2, min(linhas_preenchimento, 30))
    for _ in range(linhas_preenchimento):
        document.add_paragraph()

    # --- Fecho e assinatura ---
    # Todos os paragrafos daqui pra frente (fecho, assinatura, data) ficam
    # marcados "manter com o proximo", pra sempre ficarem juntos - ou tudo
    # cabe na pagina atual, ou tudo pula junto pra proxima (nunca fica so a
    # data sozinha numa pagina em branco).
    paragrafos_bloco_final = []

    p_fecho = document.add_paragraph()
    p_fecho.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_fecho = p_fecho.add_run(dados_oficio.get("fecho") or "Atenciosamente,")
    run_fecho.font.name = FONTE
    run_fecho.font.size = Pt(TAM_NORMAL)
    paragrafos_bloco_final.append(p_fecho)

    paragrafos_bloco_final.append(document.add_paragraph())

    if dados_escola.get("assinatura_path"):
        paragrafos_bloco_final.append(_imagem_centralizada(document, dados_escola["assinatura_path"], 3.0))
    else:
        assinado_por = (dados_oficio.get("assinado_por") or "").strip()
        if assinado_por:
            paragrafos_bloco_final.append(
                _paragrafo(document, assinado_por, negrito=True, alinhamento=WD_ALIGN_PARAGRAPH.CENTER))
            cargo_assina = (dados_oficio.get("cargo_assinado_por") or "").strip()
            if cargo_assina:
                paragrafos_bloco_final.append(
                    _paragrafo(document, cargo_assina, alinhamento=WD_ALIGN_PARAGRAPH.CENTER))

    paragrafos_bloco_final.append(document.add_paragraph())
    paragrafos_bloco_final.append(document.add_paragraph())

    # --- Local e data, no final de tudo, alinhado a direita ---
    data = (dados_oficio.get("data") or "").strip()
    if data:
        paragrafos_bloco_final.append(_paragrafo(document, f"{data}.", alinhamento=WD_ALIGN_PARAGRAPH.RIGHT))

    for p in paragrafos_bloco_final[:-1]:
        if p is not None:
            p.paragraph_format.keep_with_next = True

    document.save(caminho_saida)
    return caminho_saida


def _caixa(marcado):
    """Retorna a caixinha de opcao marcada ou vazia, ex: (X) / ( )."""
    return "(X)" if marcado else "( )"


def _linha_com_borda_inferior(destino):
    """Cria uma 'linha' (pra assinatura/preenchimento a mao) usando uma
    borda inferior no paragrafo - mais confiavel que espacos sublinhados,
    que podem ser cortados na conversao pra PDF."""
    if destino.paragraphs and not destino.paragraphs[0].runs:
        p = destino.paragraphs[0]
    else:
        p = destino.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    borda = OxmlElement("w:bottom")
    borda.set(qn("w:val"), "single")
    borda.set(qn("w:sz"), "6")
    borda.set(qn("w:space"), "1")
    borda.set(qn("w:color"), "000000")
    pBdr.append(borda)
    pPr.append(pBdr)
    return p


def gerar_declaracao(dados_escola, dados_decl, caminho_saida):
    """Gera uma Declaração Escolar (.docx), no formato de formulário usado
    pela SEMED Manaus (tabela com os dados da escola + texto com opções
    de caixinha marcadas automaticamente conforme o que foi escolhido no
    app, ao contrário de deixar tudo em branco pra marcar a mao).

    dados_escola: dict com nome_escola, secretaria, logo_path, endereco,
        telefone, email.
    dados_decl: dict com aluno, codigo_sigeam, codigo_tipo ("SIGEAM"/
        "Matrícula"/rotulo customizado), situacao_matricula
        ("esta" ou "foi"), ano_letivo, situacao_curso ("cursa" ou "cursou"),
        serie (texto livre), turma, turno ("matutino"/"vespertino"/
        "noturno"/"intermediario"),
        finalidade ("trabalho"/"transferencia"/"sinetram"/"bolsa_familia"/
        "outros"), finalidade_frequencia (numero, so p/ bolsa_familia),
        finalidade_outros (texto, so p/ outros), status_aluno
        ("promovido"/"retido"/"desistente"/"progressao_parcial"/"cursando"),
        status_desistente_data (so p/ desistente), obs (opcional), data
        (ja formatada por extenso).
    """
    document = docx.Document()
    _config_secao_oficio(document)
    _set_fonte_padrao(document)

    # --- Cabecalho: logo + nome da escola/secretaria, lado a lado ---
    logo_path = dados_escola.get("logo_path")
    linhas_cab = _linhas_cabecalho(dados_escola) or ["Secretaria Municipal de Educação"]
    if logo_path:
        cab = document.add_table(rows=1, cols=2)
        _remover_bordas_tabela(cab)
        _definir_largura_colunas(cab, [2.5, 13.5])
        cel_logo, cel_sec = cab.rows[0].cells
        _imagem_centralizada(cel_logo, logo_path, 2.0)
        cel_sec.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        for i, linha in enumerate(linhas_cab):
            _paragrafo(cel_sec, linha.upper(), negrito=(i == 0), tamanho=13 if i == 0 else TAM_NORMAL,
                       alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    else:
        for i, linha in enumerate(linhas_cab):
            _paragrafo(document, linha.upper(), negrito=(i == 0), tamanho=13 if i == 0 else TAM_NORMAL)

    document.add_paragraph()

    # --- Tabela com os dados da escola ---
    linhas_info = [
        ("CMEI/Escola Municipal", dados_escola.get("nome_escola", "")),
        ("Endereço", dados_escola.get("endereco", "")),
        ("Telefone", dados_escola.get("telefone", "")),
        ("Email", dados_escola.get("email", "")),
    ]
    tabela_info = document.add_table(rows=len(linhas_info), cols=1)
    tabela_info.style = "Table Grid"
    for linha_tabela, (rotulo, valor) in zip(tabela_info.rows, linhas_info):
        celula = _celula_vazia(linha_tabela.cells[0])
        celula.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = celula.paragraphs[0]
        r1 = p.add_run(f"{rotulo}: ")
        r1.bold = True
        r1.font.name = FONTE
        r1.font.size = Pt(TAM_NORMAL)
        r2 = p.add_run(valor)
        r2.font.name = FONTE
        r2.font.size = Pt(TAM_NORMAL)

    document.add_paragraph()

    # --- Titulo ---
    _paragrafo(document, "DECLARAÇÃO", negrito=True, tamanho=TAM_TITULO, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

    document.add_paragraph()

    # --- Corpo: texto corrido com as opcoes marcadas ---
    aluno = (dados_decl.get("aluno") or "").strip()
    codigo_sigeam = (dados_decl.get("codigo_sigeam") or "").strip()
    codigo_tipo = (dados_decl.get("codigo_tipo") or "SIGEAM").strip()
    ano_letivo = (dados_decl.get("ano_letivo") or "").strip()
    turma = (dados_decl.get("turma") or "").strip()
    situacao_matricula = dados_decl.get("situacao_matricula") or "esta"
    situacao_curso = dados_decl.get("situacao_curso") or "cursa"
    serie = (dados_decl.get("serie") or "").strip()
    turno = dados_decl.get("turno") or ""

    p_corpo = document.add_paragraph()
    p_corpo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    texto_corpo = (
        f"Declaramos para os devidos fins que o(a) aluno(a), {aluno or '_' * 45}, "
        f"sob o código ({codigo_tipo}) {codigo_sigeam or '_' * 15}, "
        f"{_caixa(situacao_matricula == 'esta')} está {_caixa(situacao_matricula == 'foi')} foi "
        f"matriculado(a) neste estabelecimento de ensino no ano letivo de {ano_letivo or '20____'}, onde "
        f"{_caixa(situacao_curso == 'cursa')} cursa {_caixa(situacao_curso == 'cursou')} cursou a série "
        f"{serie or '_' * 20}, na turma "
        f"{turma or '______'}, turno {_caixa(turno == 'matutino')} matutino {_caixa(turno == 'vespertino')} "
        f"vespertino {_caixa(turno == 'noturno')} noturno {_caixa(turno == 'intermediario')} intermediário, "
        f"conforme especificações abaixo:"
    )
    run_corpo = p_corpo.add_run(texto_corpo)
    run_corpo.font.name = FONTE
    run_corpo.font.size = Pt(TAM_NORMAL)

    document.add_paragraph()

    # --- Duas colunas: finalidade da declaracao + status do aluno ---
    finalidade = dados_decl.get("finalidade") or ""
    freq = (dados_decl.get("finalidade_frequencia") or "").strip()
    outros_texto = (dados_decl.get("finalidade_outros") or "").strip()
    status = dados_decl.get("status_aluno") or ""
    status_data = (dados_decl.get("status_desistente_data") or "").strip()

    tabela_opcoes = document.add_table(rows=6, cols=2)
    _remover_bordas_tabela(tabela_opcoes)
    _definir_largura_colunas(tabela_opcoes, [8.0, 8.0])

    _paragrafo(tabela_opcoes.cell(0, 0), "Declarações para fins de:", negrito=True)
    _paragrafo(tabela_opcoes.cell(1, 0), f"{_caixa(finalidade == 'trabalho')} Trabalho")
    _paragrafo(tabela_opcoes.cell(2, 0), f"{_caixa(finalidade == 'transferencia')} Transferência")
    _paragrafo(tabela_opcoes.cell(3, 0), f"{_caixa(finalidade == 'sinetram')} Sinetram")
    _paragrafo(tabela_opcoes.cell(4, 0),
               f"{_caixa(finalidade == 'bolsa_familia')} Bolsa Família / Frequência {freq or '_____'}")
    _paragrafo(tabela_opcoes.cell(5, 0), f"{_caixa(finalidade == 'outros')} Outros {outros_texto or '_' * 20}")

    _paragrafo(tabela_opcoes.cell(0, 1), "Status do Aluno:", negrito=True)
    _paragrafo(tabela_opcoes.cell(1, 1), f"{_caixa(status == 'promovido')} Promovido (a)")
    _paragrafo(tabela_opcoes.cell(2, 1), f"{_caixa(status == 'retido')} Retido (a)")
    _paragrafo(tabela_opcoes.cell(3, 1),
               f"{_caixa(status == 'desistente')} Desistente a partir de {status_data or '__/__/__'}")
    _paragrafo(tabela_opcoes.cell(4, 1), f"{_caixa(status == 'progressao_parcial')} Progressão Parcial")
    _paragrafo(tabela_opcoes.cell(5, 1), f"{_caixa(status == 'cursando')} Cursando")

    # --- OBS ---
    obs = (dados_decl.get("obs") or "").strip()
    p_obs = document.add_paragraph()
    r1 = p_obs.add_run("OBS: ")
    r1.bold = True
    r1.font.name = FONTE
    r1.font.size = Pt(TAM_NORMAL)
    if obs:
        r2 = p_obs.add_run(obs)
        r2.font.name = FONTE
        r2.font.size = Pt(TAM_NORMAL)
    else:
        _linha_com_borda_inferior(document)

    document.add_paragraph()

    # --- Nota legal fixa ---
    p_nota = document.add_paragraph()
    run_nota = p_nota.add_run(
        "Conforme o parágrafo único do art. 141 do Regimento Geral das Escolas da Rede Municipal, "
        "esta declaração tem validade de 30 dias a contar da data de sua expedição.")
    run_nota.font.name = FONTE
    run_nota.font.size = Pt(TAM_NORMAL - 1)
    run_nota.italic = True

    document.add_paragraph()

    # --- Data ---
    data = (dados_decl.get("data") or "").strip()
    if data:
        p_data = document.add_paragraph()
        run_data = p_data.add_run(f"{data}.")
        run_data.font.name = FONTE
        run_data.font.size = Pt(TAM_NORMAL)
        run_data.bold = True
        run_data.italic = True

    document.add_paragraph()
    document.add_paragraph()

    # --- Assinaturas: Secretario(a) e Diretor(a), lado a lado ---
    tabela_assinatura = document.add_table(rows=2, cols=2)
    _remover_bordas_tabela(tabela_assinatura)
    _definir_largura_colunas(tabela_assinatura, [8.0, 8.0])

    _linha_com_borda_inferior(tabela_assinatura.cell(0, 0))
    _linha_com_borda_inferior(tabela_assinatura.cell(0, 1))

    p_sec = _paragrafo(tabela_assinatura.cell(1, 0), "Secretário (a)", alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    p_sec.runs[0].italic = True
    p_dir = _paragrafo(tabela_assinatura.cell(1, 1), "Diretor (a)", alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    p_dir.runs[0].italic = True

    document.save(caminho_saida)
    return caminho_saida


def gerar_capa_livro(dados_escola, dados_livro, caminho_saida):
    """Gera um Termo de Abertura de Livro (.docx) - documento generico que
    serve pra abrir qualquer tipo de livro de registro da escola (atas,
    ponto, ocorrencias, etc.): o texto legal e sempre o mesmo, só os dados
    do livro em si mudam conforme o que a pessoa preencher.

    dados_escola: dict com nome_escola, secretaria, cidade, logo_path,
        endereco, telefone, email, assinatura_path.
    dados_livro: dict com tipo_livro (texto livre, ex: "Livro de Atas"),
        numero (numero do livro), finalidade (pra que ele serve), qtd_folhas
        (quantas folhas numeradas tem), data (local + data ja formatados por
        extenso, ex: "Manaus, 15 de setembro de 2026"), responsavel_abertura
        (nome de quem assina a abertura), cargo_abertura (cargo dessa
        pessoa) - ambos digitados na tela deste documento, nao vem do
        cadastro da escola (so pre-preenchidos com o diretor cadastrado).
    """
    document = docx.Document()
    _config_secao_oficio(document)
    _set_fonte_padrao(document)
    linhas_usadas = 0

    # --- Cabecalho: logo + secretaria, e tabela com os dados da escola ---
    linhas_usadas += _cabecalho_logo_e_secretaria(document, dados_escola) + 2
    document.add_paragraph()
    linhas_usadas += _tabela_dados_escola(document, dados_escola) + 1

    document.add_paragraph()
    linhas_usadas += 1

    # --- Titulo ---
    _paragrafo(document, "TERMO DE ABERTURA", negrito=True, tamanho=TAM_TITULO,
               alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    linhas_usadas += 1

    document.add_paragraph()
    document.add_paragraph()
    linhas_usadas += 2

    # --- Corpo: texto legal generico, com os dados do livro preenchidos ---
    tipo_livro = (dados_livro.get("tipo_livro") or "").strip()
    numero = (dados_livro.get("numero") or "").strip()
    finalidade = (dados_livro.get("finalidade") or "").strip()
    qtd_folhas = (dados_livro.get("qtd_folhas") or "").strip()
    data = (dados_livro.get("data") or "").strip()
    responsavel_abertura = (dados_livro.get("responsavel_abertura") or "").strip()
    cargo_abertura = (dados_livro.get("cargo_abertura") or "").strip()

    trecho_livro = f", denominado \"{tipo_livro}\"" if tipo_livro else ""
    texto_corpo = (
        "Na qualidade de responsável por esta unidade escolar, declaro aberto o "
        f"presente livro nº {numero or '_' * 5}{trecho_livro}, destinado a "
        f"{finalidade or '_' * 40}, contendo {qtd_folhas or '_' * 5} folhas numeradas "
        "e rubricadas, que servirão para o fim aqui declarado. E, para que surta os "
        "efeitos legais, firmo o presente Termo de Abertura."
    )
    p_corpo = document.add_paragraph()
    p_corpo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_corpo.paragraph_format.first_line_indent = Cm(1.25)
    run_corpo = p_corpo.add_run(texto_corpo)
    run_corpo.font.name = FONTE
    run_corpo.font.size = Pt(TAM_NORMAL)
    linhas_usadas += _estimar_linhas(texto_corpo)

    # Empurra o bloco final (Responsavel/Cargo/Local e data) pra mais perto
    # do fim da pagina, preenchendo com linhas em branco o espaco que sobrar
    # - texto curto empurra bastante, texto longo empurra pouco ou nada.
    linhas_reservadas_final = 15 if dados_escola.get("assinatura_path") else 11
    linhas_disponiveis = ALTURA_UTIL_PAGINA_OFICIO_CM / ALTURA_LINHA_CM
    linhas_preenchimento = int(linhas_disponiveis - linhas_usadas - linhas_reservadas_final)
    linhas_preenchimento = max(2, min(linhas_preenchimento, 30))
    for _ in range(linhas_preenchimento):
        document.add_paragraph()

    # --- Assinatura: imagem cadastrada (se houver) + Responsavel/Cargo/Data ---
    # Todos os paragrafos daqui pra frente ficam marcados "manter com o
    # proximo", pra sempre ficarem juntos - ou tudo cabe na pagina atual,
    # ou tudo pula junto pra proxima.
    paragrafos_bloco_final = []

    if dados_escola.get("assinatura_path"):
        paragrafos_bloco_final.append(_imagem_centralizada(document, dados_escola["assinatura_path"], 3.0))

    p_resp = document.add_paragraph()
    p_resp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p_resp.add_run("Responsável pela abertura: ")
    r1.bold = True
    r1.font.name = FONTE
    r1.font.size = Pt(TAM_NORMAL)
    r2 = p_resp.add_run(responsavel_abertura or "_" * 40)
    r2.font.name = FONTE
    r2.font.size = Pt(TAM_NORMAL)
    paragrafos_bloco_final.append(p_resp)

    p_cargo = document.add_paragraph()
    p_cargo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r3 = p_cargo.add_run("Cargo: ")
    r3.bold = True
    r3.font.name = FONTE
    r3.font.size = Pt(TAM_NORMAL)
    r4 = p_cargo.add_run(cargo_abertura or "_" * 25)
    r4.font.name = FONTE
    r4.font.size = Pt(TAM_NORMAL)
    paragrafos_bloco_final.append(p_cargo)

    paragrafos_bloco_final.append(document.add_paragraph())
    paragrafos_bloco_final.append(document.add_paragraph())

    # --- Local e data, alinhado a direita, por ultimo ---
    if data:
        paragrafos_bloco_final.append(_paragrafo(document, f"{data}.", alinhamento=WD_ALIGN_PARAGRAPH.RIGHT))

    for p in paragrafos_bloco_final[:-1]:
        if p is not None:
            p.paragraph_format.keep_with_next = True

    document.save(caminho_saida)
    return caminho_saida


LARGURA_UTIL_FICHA_CM = 18.5  # mesma largura util do Memorando (margens estreitas)


def _celula_campo(celula, rotulo, valor):
    """Celula de formulario com o rotulo pequeno em cima (itálico) e o
    valor maior embaixo (negrito) - imita o padrao visual de formulario
    oficial (rótulo/caixa de preenchimento)."""
    celula.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    p1 = celula.paragraphs[0]
    r1 = p1.add_run(rotulo)
    r1.italic = True
    r1.font.name = FONTE
    r1.font.size = Pt(TAM_NORMAL - 3)
    p2 = celula.add_paragraph()
    r2 = p2.add_run(valor or "")
    r2.bold = True
    r2.font.name = FONTE
    r2.font.size = Pt(TAM_NORMAL - 1)


def _celula_opcoes(celula, rotulo, opcoes):
    """Celula de formulario com um rotulo em cima e uma lista de opcoes
    (X)/( ) lado a lado embaixo - pra campos de escolha unica/multipla
    (Sim/Não, Urbana/Rural etc), marcadas automaticamente conforme o que
    foi escolhido no app."""
    celula.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    p1 = celula.paragraphs[0]
    r1 = p1.add_run(rotulo)
    r1.italic = True
    r1.font.name = FONTE
    r1.font.size = Pt(TAM_NORMAL - 3)
    p2 = celula.add_paragraph()
    texto_opcoes = "   ".join(f"{_caixa(marcado)} {texto}" for texto, marcado in opcoes)
    r2 = p2.add_run(texto_opcoes)
    r2.font.name = FONTE
    r2.font.size = Pt(TAM_NORMAL - 1)


def _linha_campos(document, campos, larguras=None):
    """Uma linha do formulario com varios campos lado a lado (lista de
    tuplas (rotulo, valor) OU (rotulo, [(texto, marcado), ...]) pra campos
    de opcao) - cada um numa celula com borda."""
    n = len(campos)
    larguras = larguras or [LARGURA_UTIL_FICHA_CM / n] * n
    tabela = document.add_table(rows=1, cols=n)
    tabela.style = "Table Grid"
    _definir_largura_colunas(tabela, larguras)
    _definir_margens_celulas(tabela)
    for celula, (rotulo, valor) in zip(tabela.rows[0].cells, campos):
        if isinstance(valor, list):
            _celula_opcoes(celula, rotulo, valor)
        else:
            _celula_campo(celula, rotulo, valor)
    return tabela


def _titulo_secao_ficha(document, texto, largura_cm=LARGURA_UTIL_FICHA_CM):
    """Barra de titulo de secao do formulario (fundo cinza claro)."""
    tabela = document.add_table(rows=1, cols=1)
    tabela.style = "Table Grid"
    _definir_largura_colunas(tabela, [largura_cm])
    celula = tabela.rows[0].cells[0]
    tcPr = celula._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), "E7E6E6")
    tcPr.append(shd)
    _paragrafo(celula, texto, negrito=True, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)


def gerar_ficha_matricula(dados_escola, dados_ficha, caminho_saida):
    """Gera uma Ficha de Matrícula (.docx), adaptada do modelo oficial usado
    pela SEMED Manaus pra servir qualquer escola do Brasil e qualquer etapa
    de ensino (cabecalho generico igual aos outros documentos, em vez do
    timbre fixo da prefeitura de Manaus; título e etapa/turma digitados
    livremente, não presos só à Educação Infantil).

    dados_escola: dict com nome_escola, secretaria, logo_path, endereco,
        telefone, email.
    dados_ficha: dict com tipo_ficha (texto livre pro título, ex: "Educação
        Infantil"/"Ensino Fundamental I" - default "Educação Infantil" se
        vazio), nome_social, nome_crianca, data_nascimento, sexo
        ("masculino"/"feminino"), gemeo (bool), tipo_sanguineo,
        nacionalidade, data_entrada_pais, naturalidade, uf_naturalidade,
        nome_mae, nome_pai, endereco, numero, complemento, tipo_logradouro,
        bairro, cep, numero_termo, folha, livro, data_emissao_certidao,
        uf_cartorio, nome_cartorio, numero_identidade,
        complemento_identidade, data_expedicao_identidade, uf_rg,
        orgao_emissor_identidade, cpf, cor_raca ("branca"/"preta"/"parda"/
        "amarela"/"indigena"/"nao_declarada"), telefone,
        matricula_registro_civil, codigo_aluno, nis, data_ingresso,
        deficiencia ("sim"/"nao"), bolsa_familia ("sim"/"nao"),
        tipo_deficiencia, necessidades_especiais ("sim"/"nao"),
        apoio_pedagogico ("na_escola"/"outra_escola"/""), transporte_escolar
        ("sim"/"nao"), tipo_transporte ("fluvial"/"rodoviario"/""),
        zona_residencia ("urbana"/"rural"), movimento ("nenhum"/"abandono"/
        "transferencia"/"matricula_final"), etapa (texto livre, ex:
        "Maternal II", "5º Ano").
    """
    document = docx.Document()
    _config_secao(document)
    _set_fonte_padrao(document)

    # --- Cabecalho: logo + secretaria, e tabela com os dados da escola ---
    _cabecalho_logo_e_secretaria(document, dados_escola)
    document.add_paragraph()
    _tabela_dados_escola(document, dados_escola)
    document.add_paragraph()

    tipo_ficha = (dados_ficha.get("tipo_ficha") or "Educação Infantil").strip()
    _paragrafo(document, f"FICHA DE MATRÍCULA – {tipo_ficha.upper()}", negrito=True, tamanho=TAM_TITULO,
               alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    document.add_paragraph()

    nome_escola = (dados_escola.get("nome_escola") or "").strip()
    _linha_campos(document, [
        ("Unidade de Ensino", nome_escola),
        ("Código Aluno", dados_ficha.get("codigo_aluno", "")),
        ("N.I.S.", dados_ficha.get("nis", "")),
    ], larguras=[9.5, 4.5, 4.5])

    document.add_paragraph()

    # --- DADOS PESSOAIS DA CRIANÇA ---
    _titulo_secao_ficha(document, "DADOS PESSOAIS DA CRIANÇA")

    nome_social = (dados_ficha.get("nome_social") or "").strip()
    if nome_social:
        _linha_campos(document, [("Nome Social", nome_social)])

    _linha_campos(document, [("Nome completo da criança – sem abreviaturas", dados_ficha.get("nome_crianca", ""))])

    sexo = dados_ficha.get("sexo") or ""
    _linha_campos(document, [
        ("Data de nascimento", dados_ficha.get("data_nascimento", "")),
        ("Sexo", [("Masculino", sexo == "masculino"), ("Feminino", sexo == "feminino")]),
        ("Gêmeo", [("Sim", bool(dados_ficha.get("gemeo")))]),
        ("Tipo Sanguíneo", dados_ficha.get("tipo_sanguineo", "")),
    ], larguras=[4.5, 5.0, 3.0, 6.0])

    _linha_campos(document, [
        ("Nacionalidade (para estrangeiro)", dados_ficha.get("nacionalidade", "")),
        ("Data de entrada no país", dados_ficha.get("data_entrada_pais", "")),
    ], larguras=[11.0, 7.5])

    _linha_campos(document, [
        ("Naturalidade/Município", dados_ficha.get("naturalidade", "")),
        ("UF", dados_ficha.get("uf_naturalidade", "")),
    ], larguras=[16.0, 2.5])

    _linha_campos(document, [("Nome completo da mãe – sem abreviaturas", dados_ficha.get("nome_mae", ""))])
    _linha_campos(document, [("Nome completo do pai – sem abreviaturas", dados_ficha.get("nome_pai", ""))])
    _linha_campos(document, [("Endereço residencial", dados_ficha.get("endereco", ""))])

    _linha_campos(document, [
        ("Número", dados_ficha.get("numero", "")),
        ("Complemento", dados_ficha.get("complemento", "")),
        ("Tipo logradouro", dados_ficha.get("tipo_logradouro", "")),
    ], larguras=[4.0, 8.5, 6.0])

    _linha_campos(document, [
        ("Bairro", dados_ficha.get("bairro", "")),
        ("CEP", dados_ficha.get("cep", "")),
    ], larguras=[12.0, 6.5])

    _linha_campos(document, [
        ("Certidão Civil", [("Nascimento", True)]),
        ("Número do Termo", dados_ficha.get("numero_termo", "")),
        ("Folha", dados_ficha.get("folha", "")),
        ("Livro", dados_ficha.get("livro", "")),
        ("Data de emissão", dados_ficha.get("data_emissao_certidao", "")),
        ("UF/Cartório", dados_ficha.get("uf_cartorio", "")),
    ], larguras=[3.5, 3.5, 2.5, 2.5, 3.5, 3.0])

    _linha_campos(document, [
        ("Nome do Cartório – Órgão emissor", dados_ficha.get("nome_cartorio", "")),
        ("Número da Identidade", dados_ficha.get("numero_identidade", "")),
    ], larguras=[11.0, 7.5])

    _linha_campos(document, [
        ("Complemento de Identidade", dados_ficha.get("complemento_identidade", "")),
        ("Data de Expedição", dados_ficha.get("data_expedicao_identidade", "")),
        ("UF Rg", dados_ficha.get("uf_rg", "")),
        ("Órgão emissor", dados_ficha.get("orgao_emissor_identidade", "")),
    ], larguras=[5.5, 4.5, 2.5, 6.0])

    cor_raca = dados_ficha.get("cor_raca") or ""
    _linha_campos(document, [
        ("Número do CPF", dados_ficha.get("cpf", "")),
        ("Cor/Raça", [("Branca", cor_raca == "branca"), ("Preta", cor_raca == "preta"),
                      ("Parda", cor_raca == "parda"), ("Amarela", cor_raca == "amarela"),
                      ("Indígena", cor_raca == "indigena"), ("Não declarada", cor_raca == "nao_declarada")]),
        ("Telefone", dados_ficha.get("telefone", "")),
    ], larguras=[3.5, 10.0, 5.0])

    _linha_campos(document, [("Matrícula do Registro Civil", dados_ficha.get("matricula_registro_civil", ""))])

    document.add_paragraph()

    # --- DADOS ESCOLARES ---
    _titulo_secao_ficha(document, "DADOS ESCOLARES")

    deficiencia = dados_ficha.get("deficiencia") or "nao"
    bolsa_familia = dados_ficha.get("bolsa_familia") or "nao"
    _linha_campos(document, [
        ("Data de ingresso na Unidade de Ensino", dados_ficha.get("data_ingresso", "")),
        ("Criança com Deficiência", [("Sim", deficiencia == "sim"), ("Não", deficiencia == "nao")]),
        ("Participa do Programa Bolsa Família",
         [("Sim", bolsa_familia == "sim"), ("Não", bolsa_familia == "nao")]),
    ], larguras=[7.0, 5.5, 6.0])

    tipo_deficiencia = (dados_ficha.get("tipo_deficiencia") or "").strip()
    necessidades = dados_ficha.get("necessidades_especiais") or "nao"
    apoio = dados_ficha.get("apoio_pedagogico") or ""
    _linha_campos(document, [
        ("Tipo de Deficiência", tipo_deficiencia),
        ("Necessidades Educacionais Especiais", [("Sim", necessidades == "sim"), ("Não", necessidades == "nao")]),
        ("Apoio Pedagógico Especializado",
         [("Na própria escola", apoio == "na_escola"), ("Outra escola/Centro", apoio == "outra_escola")]),
    ], larguras=[5.0, 6.5, 7.0])

    transporte = dados_ficha.get("transporte_escolar") or "nao"
    tipo_transp = dados_ficha.get("tipo_transporte") or ""
    zona = dados_ficha.get("zona_residencia") or ""
    _linha_campos(document, [
        ("Utiliza Transporte Escolar Público Municipal?",
         [("Sim", transporte == "sim"), ("Não", transporte == "nao")]),
        ("Tipo de Transporte Público Oferecido",
         [("Fluvial", tipo_transp == "fluvial"), ("Rodoviário", tipo_transp == "rodoviario")]),
        ("Zona de Residência", [("Urbana", zona == "urbana"), ("Rural", zona == "rural")]),
    ], larguras=[6.0, 6.5, 6.0])

    movimento = dados_ficha.get("movimento") or "nenhum"
    _linha_campos(document, [
        ("Movimento e Rendimento Escolar", [
            ("Afastado por abandono", movimento == "abandono"),
            ("Afastado por transferência", movimento == "transferencia"),
            ("Matrícula Final", movimento == "matricula_final"),
        ]),
    ])

    _linha_campos(document, [("Série/Turma", dados_ficha.get("etapa", ""))])

    document.save(caminho_saida)
    return caminho_saida


def _secao_lista_reuniao(document, dados_escola, dados_lista):
    """Desenha uma secao completa (cabecalho + titulo + identificacao da
    turma + tabela de alunos) no documento - usada tanto pra uma lista de
    turma unica quanto, repetida com quebra de pagina entre cada chamada,
    pra gerar varias turmas dentro do MESMO arquivo (ver
    gerar_lista_reuniao_varias_turmas)."""
    # --- Cabecalho: logo + secretaria, e tabela com os dados da escola ---
    _cabecalho_logo_e_secretaria(document, dados_escola)
    document.add_paragraph()
    _tabela_dados_escola(document, dados_escola)
    document.add_paragraph()

    # --- Barra de titulo: data + tipo de reuniao/lista ---
    data_reuniao = (dados_lista.get("data_reuniao") or "").strip()
    tipo_reuniao = (dados_lista.get("tipo_reuniao") or "").strip() or "Reunião de Pais e Alunos"
    titulo = f"{data_reuniao} – {tipo_reuniao.upper()}" if data_reuniao else tipo_reuniao.upper()
    _titulo_secao_ficha(document, titulo)

    # --- Identificacao da turma ---
    _linha_campos(document, [
        ("Ano", dados_lista.get("ano_letivo", "")),
        ("Ensino/Projeto", dados_lista.get("ensino_projeto", "")),
        ("Fase", dados_lista.get("fase", "")),
        ("Turma", dados_lista.get("turma", "")),
        ("Turno", dados_lista.get("turno", "")),
    ], larguras=[2.0, 6.5, 4.0, 3.0, 3.0])

    # --- Tabela de alunos: Nº | Nome do Aluno | Data | Assinatura ---
    # Fonte menor e margens de celula mais estreitas que o padrao (so aqui,
    # pra caber uma turma cheia numa pagina so - o resto do documento usa
    # o tamanho normal).
    TAM_TABELA_ALUNOS = 10
    nomes = [linha.strip() for linha in (dados_lista.get("nomes_alunos") or "").split("\n") if linha.strip()]
    if nomes:
        linhas_tabela = len(nomes) + 3  # nomes + margem extra pra acrescentar aluno depois
    else:
        # sem nomes (ex: essa turma ainda nao teve a lista preenchida) -
        # tabela sai em branco, com a quantidade de linhas pedida.
        try:
            linhas_tabela = int(dados_lista.get("linhas_em_branco") or 3)
        except (TypeError, ValueError):
            linhas_tabela = 3
    total_linhas = linhas_tabela + 1  # +1 pro cabecalho da tabela

    tabela = document.add_table(rows=total_linhas, cols=4)
    tabela.style = "Table Grid"
    larguras_tabela = [1.3, 9.5, 2.2, 5.5]  # nome mais largo - nome completo com 2 sobrenomes nao quebra linha
    _definir_largura_colunas(tabela, larguras_tabela)
    _definir_margens_celulas(tabela, cima_cm=0.08, baixo_cm=0.08, esquerda_cm=0.15, direita_cm=0.15)

    for celula, texto in zip(tabela.rows[0].cells, ["Nº", "Nome do Aluno", "Data", "Assinatura do Responsável"]):
        celula.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _paragrafo(celula, texto, negrito=True, tamanho=TAM_TABELA_ALUNOS, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

    for i in range(linhas_tabela):
        linha_tabela = tabela.rows[i + 1]
        cel_num, cel_nome, cel_data, cel_assinatura = linha_tabela.cells
        for c in (cel_num, cel_nome, cel_data, cel_assinatura):
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _paragrafo(cel_num, f"{i + 1:03d}", tamanho=TAM_TABELA_ALUNOS, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
        if i < len(nomes):
            _paragrafo(cel_nome, nomes[i], tamanho=TAM_TABELA_ALUNOS)


def gerar_lista_reuniao(dados_escola, dados_lista, caminho_saida):
    """Gera uma Lista de Reunião de Pais e Alunos (.docx), adaptada de um
    modelo real de lista de assinatura usado pela SEMED Manaus (a mesma
    estrutura serve pra outras listas tipo "entrega de livros" - aqui fixada
    pra reunião de pais). Cabecalho generico igual aos outros documentos
    (nao usa os campos especificos de Manaus do modelo original, como
    INEP/Ato de Criacao, pra continuar servindo qualquer escola do Brasil).

    dados_escola: dict com nome_escola, secretaria, logo_path, endereco,
        telefone, email.
    dados_lista: dict com tipo_reuniao (texto livre, ex: "Reunião de Pais e
        Alunos" ou "Entrega de Livros"), data_reuniao (texto livre, ex:
        "21/08/2025"), ano_letivo, ensino_projeto (texto livre, ex:
        "Pré-Escola - 1º e 2º Períodos"), fase (texto livre, ex: "1º
        Período"), turma (texto livre, ex: "G"), turno (texto livre, ex:
        "Vespertino"), nomes_alunos (texto com um nome de aluno por linha -
        se vazio, a tabela sai em branco com linhas_em_branco linhas pra
        preencher a mao), linhas_em_branco (numero de linhas da tabela
        quando nomes_alunos estiver vazio, padrao 3).
    """
    document = docx.Document()
    _config_secao(document)
    _set_fonte_padrao(document)
    _secao_lista_reuniao(document, dados_escola, dados_lista)
    document.save(caminho_saida)
    return caminho_saida


def gerar_lista_reuniao_varias_turmas(dados_escola, dados_comuns, turmas, caminho_saida):
    """Gera UM UNICO .docx com uma pagina por turma (cabecalho + titulo +
    tabela repetidos, com quebra de pagina entre elas) - pra quando a mesma
    reuniao/lista precisa sair pra varias turmas de uma vez, sem gerar um
    arquivo separado pra cada uma.

    dados_comuns: mesmos campos de dados_lista em gerar_lista_reuniao, MENOS
        turma/nomes_alunos/linhas_em_branco (que vem de cada item de turmas).
    turmas: lista de dicts, cada um com turma (texto), nomes_alunos (texto,
        pode ser vazio) e opcionalmente linhas_em_branco.
    """
    document = docx.Document()
    _config_secao(document)
    _set_fonte_padrao(document)

    for indice, dados_turma in enumerate(turmas):
        if indice > 0:
            document.add_page_break()
        dados_lista = dict(dados_comuns)
        dados_lista.update(dados_turma)
        _secao_lista_reuniao(document, dados_escola, dados_lista)

    document.save(caminho_saida)
    return caminho_saida


LARGURA_UTIL_FREQUENCIA_CM = 29.7 - 1.5 - 1.0  # pagina em paisagem, margens de _config_secao_paisagem


def _parse_dias_letivos(texto):
    """Converte "1-5,8-12,15" em [1,2,3,4,5,8,9,10,11,12,15] - aceita
    numeros soltos e faixas com "-", separados por virgula. Pedacos
    invalidos sao ignorados silenciosamente (o formulario ja valida que o
    campo nao ficou vazio antes de chamar isso)."""
    dias = []
    for pedaco in (texto or "").split(","):
        pedaco = pedaco.strip()
        if not pedaco:
            continue
        if "-" in pedaco:
            try:
                inicio, fim = pedaco.split("-", 1)
                dias.extend(range(int(inicio.strip()), int(fim.strip()) + 1))
            except ValueError:
                continue
        else:
            try:
                dias.append(int(pedaco))
            except ValueError:
                continue
    return dias


def _secao_lista_frequencia(document, dados_escola, dados_frequencia):
    """Desenha uma secao completa (cabecalho + titulo + identificacao +
    grade de frequencia) no documento - usada tanto pra uma ficha de turma
    unica quanto, repetida com quebra de pagina entre cada chamada, pra
    gerar varias turmas dentro do MESMO arquivo (ver
    gerar_lista_frequencia_varias_turmas).

    dados_frequencia: dict com mes (texto livre, ex: "Setembro"),
        ano_letivo, turma, serie, turno, professor, dias_letivos (texto
        livre, ex: "1-5,8-12,15-19,22-26" - dias do mes que tem aula, sem
        contar fins de semana/feriados, digitados pela secretaria),
        nomes_alunos (texto com um nome de aluno por linha - se vazio, a
        tabela sai em branco com linhas_em_branco linhas), linhas_em_branco
        (numero de linhas quando nomes_alunos estiver vazio, padrao 3).
    """
    # Cabecalho compacto (so logo + nome da escola/secretaria, sem a tabela
    # de endereco/telefone/email) - esse documento e paisagem e mais curto
    # na vertical, e o espaco importa mais aqui pra caber a turma inteira
    # numa pagina so; a identificacao de Turma/Turno/Professor(a) abaixo ja
    # basta pro uso do dia a dia.
    _cabecalho_logo_e_secretaria(document, dados_escola, largura_total_cm=LARGURA_UTIL_FREQUENCIA_CM)

    mes = (dados_frequencia.get("mes") or "").strip()
    ano = (dados_frequencia.get("ano_letivo") or "").strip()
    mes_ano = "/".join(parte for parte in (mes.upper(), ano) if parte)
    titulo = f"{mes_ano} – FICHA DE FREQUÊNCIA ESCOLAR" if mes_ano else "FICHA DE FREQUÊNCIA ESCOLAR"
    _titulo_secao_ficha(document, titulo, largura_cm=LARGURA_UTIL_FREQUENCIA_CM)

    dias = _parse_dias_letivos(dados_frequencia.get("dias_letivos"))

    _linha_campos(document, [
        ("Turma", dados_frequencia.get("turma", "")),
        ("Série", dados_frequencia.get("serie", "")),
        ("Turno", dados_frequencia.get("turno", "")),
        ("Professor(a)", dados_frequencia.get("professor", "")),
        ("Dias letivos dados", str(len(dias)) if dias else ""),
    ], larguras=[4.5, 5.5, 4.0, 9.0, LARGURA_UTIL_FREQUENCIA_CM - 23.0])

    # --- Tabela de frequencia: Nº | Nome do Aluno | 1 coluna por dia (com
    # cabecalho agrupador "PRESENÇA" por cima) | Faltas | F. Just. ---
    TAM_TABELA_FREQ = 8
    nomes = [linha.strip() for linha in (dados_frequencia.get("nomes_alunos") or "").split("\n") if linha.strip()]

    largura_num, largura_nome, largura_faltas, largura_fjust = 1.0, 6.5, 1.3, 1.6
    largura_fixas = largura_num + largura_nome + largura_faltas + largura_fjust
    n_dias = max(len(dias), 1)
    largura_por_dia = max((LARGURA_UTIL_FREQUENCIA_CM - largura_fixas) / n_dias, 0.45)

    if nomes:
        linhas_tabela = len(nomes) + 3
    else:
        try:
            linhas_tabela = int(dados_frequencia.get("linhas_em_branco") or 3)
        except (TypeError, ValueError):
            linhas_tabela = 3
    linhas_cabecalho = 2 if dias else 1
    total_linhas = linhas_tabela + linhas_cabecalho

    n_cols = 2 + len(dias) + 2
    tabela = document.add_table(rows=total_linhas, cols=n_cols)
    tabela.style = "Table Grid"
    larguras_tabela = [largura_num, largura_nome] + [largura_por_dia] * len(dias) + [largura_faltas, largura_fjust]
    _definir_largura_colunas(tabela, larguras_tabela)
    _definir_margens_celulas(tabela, cima_cm=0.04, baixo_cm=0.04, esquerda_cm=0.08, direita_cm=0.08)

    if dias:
        # cabecalho em 2 linhas: Nº/Nome/Faltas/F.Just mesclados na vertical
        # (ocupam as 2 linhas), e "PRESENÇA" mesclado na horizontal por
        # cima dos dias - deixa claro que sao VARIAS colunas de presenca,
        # uma por dia letivo, agrupadas visualmente.
        col_faltas = 2 + len(dias)
        col_fjust = col_faltas + 1

        cel_num = tabela.cell(0, 0).merge(tabela.cell(1, 0))
        cel_nome = tabela.cell(0, 1).merge(tabela.cell(1, 1))
        cel_faltas = tabela.cell(0, col_faltas).merge(tabela.cell(1, col_faltas))
        cel_fjust = tabela.cell(0, col_fjust).merge(tabela.cell(1, col_fjust))
        if len(dias) > 1:
            cel_presenca = tabela.cell(0, 2).merge(tabela.cell(0, col_faltas - 1))
        else:
            cel_presenca = tabela.cell(0, 2)

        for celula, texto in ((cel_num, "Nº"), (cel_nome, "Nome do Aluno"),
                               (cel_faltas, "Faltas"), (cel_fjust, "F. Just.")):
            celula.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            _paragrafo(celula, texto, negrito=True, tamanho=TAM_TABELA_FREQ, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

        cel_presenca.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _paragrafo(cel_presenca, "PRESENÇA (dias letivos)", negrito=True,
                   tamanho=TAM_TABELA_FREQ, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

        for idx, dia in enumerate(dias):
            celula_dia = tabela.rows[1].cells[2 + idx]
            celula_dia.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            _paragrafo(celula_dia, str(dia), negrito=True, tamanho=TAM_TABELA_FREQ,
                       alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    else:
        for celula, texto in zip(tabela.rows[0].cells, ["Nº", "Nome do Aluno", "Faltas", "F. Just."]):
            celula.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            _paragrafo(celula, texto, negrito=True, tamanho=TAM_TABELA_FREQ, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

    for i in range(linhas_tabela):
        linha_tabela = tabela.rows[i + linhas_cabecalho]
        celulas = linha_tabela.cells
        for c in celulas:
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _paragrafo(celulas[0], f"{i + 1:02d}", tamanho=TAM_TABELA_FREQ, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
        if i < len(nomes):
            _paragrafo(celulas[1], nomes[i], tamanho=TAM_TABELA_FREQ)
        # as colunas de dias, Faltas e F. Just. ficam em branco, pra marcar a mao

    p_nota = _paragrafo(document, "Marcar \"P\" para presença e \"F\" para falta em cada dia letivo. "
                                   "Totalizar Faltas e Faltas Justificadas ao final do mês.",
                         tamanho=9)
    p_nota.paragraph_format.space_before = Pt(6)


def gerar_lista_frequencia(dados_escola, dados_frequencia, caminho_saida):
    """Gera uma Ficha de Frequência Escolar (.docx) - grade mensal com um
    aluno por linha e uma coluna por dia letivo do mes, pra marcar
    presenca/falta a mao dia a dia (mesmo uso que alimenta o calculo de
    frequencia do Bolsa Familia no Relatorio de Presenca). Sai em pagina
    paisagem (horizontal), ja que o numero de colunas de dias normalmente
    nao caberia numa pagina em retrato.

    dados_escola: dict com nome_escola, secretaria, logo_path, endereco,
        telefone, email.
    dados_frequencia: ver _secao_lista_frequencia.
    """
    document = docx.Document()
    _config_secao_paisagem(document)
    _set_fonte_padrao(document)
    _secao_lista_frequencia(document, dados_escola, dados_frequencia)
    document.save(caminho_saida)
    return caminho_saida


def gerar_lista_frequencia_varias_turmas(dados_escola, dados_comuns, turmas, caminho_saida):
    """Gera UM UNICO .docx com uma pagina por turma (cabecalho + titulo +
    grade repetidos, com quebra de pagina entre elas) - pra quando a mesma
    ficha de frequencia precisa sair pra varias turmas de uma vez, sem
    gerar um arquivo separado pra cada uma.

    dados_comuns: mesmos campos de dados_frequencia em gerar_lista_frequencia,
        que valem pra todas as turmas (normalmente so mes/ano_letivo, ja que
        turma/serie/turno/professor/nomes_alunos mudam por turma). Pode ser
        {} quando cada item de turmas ja vem com todos os campos completos.
    turmas: lista de dicts, cada um com os campos especificos daquela turma.
    """
    document = docx.Document()
    _config_secao_paisagem(document)
    _set_fonte_padrao(document)

    for indice, dados_turma in enumerate(turmas):
        if indice > 0:
            document.add_page_break()
        dados_frequencia = dict(dados_comuns)
        dados_frequencia.update(dados_turma)
        _secao_lista_frequencia(document, dados_escola, dados_frequencia)

    document.save(caminho_saida)
    return caminho_saida


def gerar_justificativa_faltas(dados_escola, dados_justificativa, caminho_saida):
    """Gera uma Justificativa de Excesso de Faltas (.docx) - documento
    preenchido/assinado QUANDO o responsável comparece à escola pra
    explicar o excesso de faltas do aluno (diferente de uma notificação
    que a escola manda pedindo pra ele comparecer - aqui ele já está lá,
    dando a justificativa). O percentual mínimo exigido muda conforme a
    etapa: 60% na Educação Infantil, 75% do 1º ano do Ensino Fundamental
    em diante (conforme a LDB, Lei nº 9.394/96, art. 24, inciso VI).

    dados_escola: dict com nome_escola, secretaria, logo_path, endereco,
        telefone, email.
    dados_justificativa: dict com nome_aluno, turma, etapa_ensino
        ("infantil" -> mínimo 60%, qualquer outro valor -> mínimo 75%),
        turno, nome_responsavel, periodo (texto livre, ex: "Agosto/2026"),
        dias_letivos (texto/num), faltas (texto/num), motivo ("doenca_aluno"/
        "doenca_familia"/"mudanca_endereco"/"transporte"/"trabalho_renda"/
        "outro"), motivo_outro (texto - só usado quando motivo == "outro"),
        observacoes (texto livre, opcional), data (já formatada por
        extenso), recebido_por (nome de quem atendeu na escola),
        cargo_recebido_por.
    """
    document = docx.Document()
    _config_secao_oficio(document)
    _set_fonte_padrao(document)

    # --- Cabecalho: logo + secretaria, e tabela com os dados da escola ---
    _cabecalho_logo_e_secretaria(document, dados_escola)
    document.add_paragraph()
    _tabela_dados_escola(document, dados_escola)
    document.add_paragraph()

    _paragrafo(document, "JUSTIFICATIVA DE EXCESSO DE FALTAS", negrito=True, tamanho=TAM_TITULO,
               alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    document.add_paragraph()

    # --- Identificacao do aluno ---
    _linha_campos(document, [
        ("Nome do aluno", dados_justificativa.get("nome_aluno", "")),
        ("Turma", dados_justificativa.get("turma", "")),
        ("Turno", dados_justificativa.get("turno", "")),
        ("Responsável", dados_justificativa.get("nome_responsavel", "")),
    ], larguras=[6.0, 2.5, 2.5, 5.0])

    etapa = dados_justificativa.get("etapa_ensino") or "infantil"
    percentual_minimo = 60 if etapa == "infantil" else 75
    _linha_campos(document, [
        ("Etapa de ensino", "Educação Infantil" if etapa == "infantil" else "Ensino Fundamental em diante"),
        ("Período de referência", dados_justificativa.get("periodo", "")),
    ], larguras=[8.0, 8.0])

    # --- Dados de frequencia (percentual apurado calculado automaticamente) ---
    dias_texto = (dados_justificativa.get("dias_letivos") or "").strip()
    faltas_texto = (dados_justificativa.get("faltas") or "").strip()
    percentual_texto = ""
    try:
        dias_num = int(dias_texto)
        faltas_num = int(faltas_texto)
        if dias_num > 0:
            percentual_texto = f"{round((dias_num - faltas_num) / dias_num * 100, 1):g}%"
    except (TypeError, ValueError):
        pass

    _linha_campos(document, [
        ("Dias letivos no período", dias_texto),
        ("Faltas no período", faltas_texto),
        ("% apurado", percentual_texto),
        ("% mínimo exigido", f"{percentual_minimo}%"),
    ], larguras=[4.0, 4.0, 4.0, 4.0])

    document.add_paragraph()

    # --- Corpo: contextualizacao + compromisso, tudo num paragrafo so ---
    p_corpo = document.add_paragraph()
    p_corpo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_corpo.paragraph_format.first_line_indent = Cm(1.25)
    run_corpo = p_corpo.add_run(
        "O(a) responsável acima identificado(a) compareceu a esta unidade escolar para "
        "justificar o excesso de faltas do(a) aluno(a) acima identificado(a), que "
        f"apresentou frequência de {percentual_texto or '____'} no período informado, abaixo "
        f"do mínimo de {percentual_minimo}% exigido pela legislação vigente (LDB - Lei nº "
        "9.394/96, art. 24, inciso VI), declarando que as informações prestadas são "
        "verdadeiras e comprometendo-se a garantir a frequência regular do(a) aluno(a) nos "
        "dias letivos subsequentes, ciente de que a reincidência poderá acarretar o "
        "encaminhamento do caso ao Conselho Tutelar (ECA - Lei nº 8.069/90) e comprometer a "
        "manutenção de benefícios sociais vinculados à frequência, como o Bolsa Família."
    )
    run_corpo.font.name = FONTE
    run_corpo.font.size = Pt(TAM_NORMAL)

    document.add_paragraph()

    motivo = dados_justificativa.get("motivo") or ""
    motivo_outro = (dados_justificativa.get("motivo_outro") or "").strip()
    opcoes_motivo = [
        ("Doença do(a) aluno(a)", motivo == "doenca_aluno"),
        ("Doença/problema de saúde na família", motivo == "doenca_familia"),
        ("Mudança de endereço/dificuldade de acesso", motivo == "mudanca_endereco"),
        ("Dificuldade de transporte", motivo == "transporte"),
        ("Motivo de trabalho/renda familiar", motivo == "trabalho_renda"),
        (f"Outro: {motivo_outro}" if motivo == "outro" and motivo_outro else "Outro", motivo == "outro"),
    ]

    _linha_campos(document, [
        ("Motivo apresentado pelo(a) responsável", opcoes_motivo),
    ], larguras=[16.0])

    document.add_paragraph()

    # --- OBS ---
    obs = (dados_justificativa.get("observacoes") or "").strip()
    p_obs = document.add_paragraph()
    r1 = p_obs.add_run("Observações: ")
    r1.bold = True
    r1.font.name = FONTE
    r1.font.size = Pt(TAM_NORMAL)
    if obs:
        r2 = p_obs.add_run(obs)
        r2.font.name = FONTE
        r2.font.size = Pt(TAM_NORMAL)
    else:
        _linha_com_borda_inferior(document)

    document.add_paragraph()

    data = (dados_justificativa.get("data") or "").strip()
    if data:
        _paragrafo(document, f"{data}.", alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

    document.add_paragraph()

    # --- Assinaturas: Responsavel e quem atendeu na escola, lado a lado ---
    tabela_assinatura = document.add_table(rows=2, cols=2)
    _remover_bordas_tabela(tabela_assinatura)
    _definir_largura_colunas(tabela_assinatura, [8.0, 8.0])

    _linha_com_borda_inferior(tabela_assinatura.cell(0, 0))
    _linha_com_borda_inferior(tabela_assinatura.cell(0, 1))

    p_resp = _paragrafo(tabela_assinatura.cell(1, 0), "Assinatura do responsável",
                         alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    p_resp.runs[0].italic = True

    recebido_por = (dados_justificativa.get("recebido_por") or "").strip()
    cargo_recebido = (dados_justificativa.get("cargo_recebido_por") or "").strip()
    rotulo_escola = recebido_por or "Assinatura de quem atendeu"
    if recebido_por and cargo_recebido:
        rotulo_escola = f"{recebido_por} - {cargo_recebido}"
    p_esc = _paragrafo(tabela_assinatura.cell(1, 1), rotulo_escola, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    p_esc.runs[0].italic = True

    document.save(caminho_saida)
    return caminho_saida


def _borda_pagina(document, cor_hex="2F528F", espessura_pt=18, estilo="double"):
    """Desenha uma borda decorativa ao redor da pagina inteira - usado no
    Certificado, pra dar a cara de "diploma" em vez de carta comum."""
    secao = document.sections[0]
    sectPr = secao._sectPr
    bordas = OxmlElement("w:pgBorders")
    bordas.set(qn("w:offsetFrom"), "page")
    for nome in ("top", "left", "bottom", "right"):
        borda = OxmlElement(f"w:{nome}")
        borda.set(qn("w:val"), estilo)
        borda.set(qn("w:sz"), str(espessura_pt))
        borda.set(qn("w:space"), "24")
        borda.set(qn("w:color"), cor_hex)
        bordas.append(borda)
    sectPr.append(bordas)


LARGURA_UTIL_CERTIFICADO_CM = 29.7 - 2.5 - 2.5  # pagina em paisagem, margens de _config_secao_certificado


def gerar_certificado(dados_escola, dados_certificado, caminho_saida):
    """Gera um Certificado de Conclusão (.docx) - documento genérico pra
    qualquer etapa/série concluída (Educação Infantil, Ensino Fundamental
    etc.), com uma borda decorativa na página pra dar a cara de diploma.
    Assinaturas de Secretário(a) e Diretor(a) lado a lado, mais uma linha
    de assinatura do(a) próprio(a) aluno(a) como concludente. Uma 2ª
    página opcional registra a emissão no livro da escola (Registro/
    Livro/Folha) - mesmo tipo de controle do Termo de Abertura de Livro.

    dados_escola: dict com nome_escola, secretaria, logo_path, endereco,
        telefone, email.
    dados_certificado: dict com nome_aluno, etapa_concluida (texto livre,
        ex: "Educação Infantil", "5º Ano do Ensino Fundamental"), ano_letivo,
        turma (opcional), data (já formatada por extenso), secretario_nome,
        diretor_nome, diretor_cargo, registro_numero, livro_numero, folha,
        registrado_por, amparo_legal (todos os últimos 5 opcionais - só
        pra 2ª página de registro no livro).
    """
    document = docx.Document()
    _config_secao_certificado(document)
    _set_fonte_padrao(document)
    _borda_pagina(document)

    _cabecalho_logo_e_secretaria(document, dados_escola, largura_total_cm=LARGURA_UTIL_CERTIFICADO_CM)
    document.add_paragraph()
    _tabela_dados_escola(document, dados_escola)

    document.add_paragraph()
    document.add_paragraph()

    _paragrafo(document, "CERTIFICADO", negrito=True, tamanho=TAM_TITULO + 6,
               alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

    document.add_paragraph()
    document.add_paragraph()

    nome_aluno = (dados_certificado.get("nome_aluno") or "_" * 45).strip() or "_" * 45
    etapa = (dados_certificado.get("etapa_concluida") or "_" * 30).strip() or "_" * 30
    ano_letivo = (dados_certificado.get("ano_letivo") or "____").strip() or "____"
    turma = (dados_certificado.get("turma") or "").strip()
    trecho_turma = f", na turma {turma}" if turma else ""

    texto_corpo = (
        f"Certificamos que {nome_aluno} concluiu com aproveitamento "
        f"{etapa}, no ano letivo de {ano_letivo}{trecho_turma}, nesta unidade de ensino, "
        "fazendo jus ao presente certificado."
    )
    p_corpo = document.add_paragraph()
    p_corpo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_corpo = p_corpo.add_run(texto_corpo)
    run_corpo.font.name = FONTE
    run_corpo.font.size = Pt(TAM_NORMAL + 1)

    document.add_paragraph()

    data = (dados_certificado.get("data") or "").strip()
    if data:
        _paragrafo(document, f"{data}.", alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

    document.add_paragraph()
    document.add_paragraph()

    # --- Assinaturas: Secretario(a) e Diretor(a) lado a lado ---
    tabela_assinatura = document.add_table(rows=2, cols=2)
    _remover_bordas_tabela(tabela_assinatura)
    _definir_largura_colunas(tabela_assinatura, [LARGURA_UTIL_CERTIFICADO_CM / 2] * 2)

    _linha_com_borda_inferior(tabela_assinatura.cell(0, 0))
    _linha_com_borda_inferior(tabela_assinatura.cell(0, 1))

    secretario_nome = (dados_certificado.get("secretario_nome") or "").strip()
    p_sec = _paragrafo(tabela_assinatura.cell(1, 0), secretario_nome or "Secretário(a) Escolar",
                        alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    p_sec.runs[0].italic = not secretario_nome
    _paragrafo(tabela_assinatura.cell(1, 0), "SECRETÁRIO(A)", tamanho=TAM_NORMAL - 1,
               alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

    diretor_nome = (dados_certificado.get("diretor_nome") or "").strip()
    p_dir = _paragrafo(tabela_assinatura.cell(1, 1), diretor_nome or "Diretor(a)",
                        alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    p_dir.runs[0].italic = not diretor_nome
    _paragrafo(tabela_assinatura.cell(1, 1), "DIRETOR(A)", tamanho=TAM_NORMAL - 1,
               alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

    document.add_paragraph()

    # --- Assinatura do(a) proprio(a) aluno(a), como concludente ---
    _linha_com_borda_inferior(document)
    _paragrafo(document, nome_aluno, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    _paragrafo(document, "CONCLUDENTE", tamanho=TAM_NORMAL - 1, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

    # --- 2a pagina: registro da emissao no livro da escola ---
    registro_numero = (dados_certificado.get("registro_numero") or "").strip()
    livro_numero = (dados_certificado.get("livro_numero") or "").strip()
    folha = (dados_certificado.get("folha") or "").strip()
    registrado_por = (dados_certificado.get("registrado_por") or "").strip()
    amparo_legal = (dados_certificado.get("amparo_legal") or "").strip()

    document.add_page_break()
    _cabecalho_logo_e_secretaria(document, dados_escola, largura_total_cm=LARGURA_UTIL_CERTIFICADO_CM)
    document.add_paragraph()
    _tabela_dados_escola(document, dados_escola)
    document.add_paragraph()

    tabela_registro = document.add_table(rows=1, cols=2)
    tabela_registro.style = "Table Grid"
    _definir_largura_colunas(tabela_registro, [LARGURA_UTIL_CERTIFICADO_CM / 2] * 2)
    _definir_margens_celulas(tabela_registro)

    cel_amparo, cel_registro = tabela_registro.rows[0].cells
    cel_amparo.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    _paragrafo(cel_amparo, "Amparo Legal:", negrito=True)
    _paragrafo(cel_amparo, amparo_legal)

    cel_registro.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    _paragrafo(cel_registro, "Registro Nº:", negrito=True)
    _paragrafo(cel_registro, registro_numero)
    _paragrafo(cel_registro, "Livro Nº:", negrito=True)
    _paragrafo(cel_registro, livro_numero)
    _paragrafo(cel_registro, "Folha:", negrito=True)
    _paragrafo(cel_registro, folha)
    _paragrafo(cel_registro, "Registrado por:", negrito=True)
    _paragrafo(cel_registro, registrado_por)

    document.save(caminho_saida)
    return caminho_saida


LARGURA_UTIL_HISTORICO_CM = 18.5  # mesma largura util do retrato (margens estreitas)
TAM_TABELA_HISTORICO = 9  # fonte pequena de proposito - precisa caber varios anos letivos

SITUACOES_HISTORICO = {"promovido": "PRO", "retido": "RET", "retido_frequencia": "RFR"}


def gerar_historico_escolar(dados_escola, dados_historico, caminho_saida):
    """Gera um Histórico Escolar (.docx) - identificação do aluno + uma
    tabela detalhada por disciplina em cada ano letivo (nota, carga
    horária, faltas), igual ao modelo oficial usado pela SEMED Manaus.

    dados_escola: dict com nome_escola, secretaria, logo_path, endereco,
        telefone, email, cidade.
    dados_historico: dict com nome_aluno, codigo_aluno, data_nascimento,
        registro_geral, municipio, naturalidade, nacionalidade,
        anos (lista de dicts, cada um com ano_letivo, ensino,
        fase, estabelecimento, situacao ("promovido"/"retido"/
        "retido_frequencia"), disciplinas - lista de dicts com nome,
        nota, carga_horaria, faltas - anos sem ano_letivo sao ignorados),
        amparo_legal, regras (ambos opcionais, texto livre), data (já
        formatada por extenso), diretor_nome, diretor_portaria,
        secretario_nome, secretario_portaria.
    """
    document = docx.Document()
    _config_secao(document)
    _set_fonte_padrao(document)

    _cabecalho_logo_e_secretaria(document, dados_escola)
    document.add_paragraph()
    _tabela_dados_escola(document, dados_escola)
    document.add_paragraph()

    _titulo_secao_ficha(document, "HISTÓRICO ESCOLAR", largura_cm=LARGURA_UTIL_HISTORICO_CM)

    # --- Identificacao do aluno ---
    _linha_campos(document, [
        ("Código do aluno", dados_historico.get("codigo_aluno", "")),
        ("Nome do aluno", dados_historico.get("nome_aluno", "")),
    ], larguras=[4.5, 14.0])
    _linha_campos(document, [
        ("Data de nascimento", dados_historico.get("data_nascimento", "")),
        ("Nº Registro Geral", dados_historico.get("registro_geral", "")),
        ("Município", dados_historico.get("municipio", "")),
    ], larguras=[5.5, 6.5, 6.5])
    _linha_campos(document, [
        ("Naturalidade", dados_historico.get("naturalidade", "")),
        ("Nacionalidade", dados_historico.get("nacionalidade", "")),
    ], larguras=[9.25, 9.25])

    document.add_paragraph()
    _titulo_secao_ficha(document, "REGISTRO DE ESCOLARIDADE", largura_cm=LARGURA_UTIL_HISTORICO_CM)

    # --- Tabela: uma linha por disciplina, ANO/ENSINO/FASE/SIT./ESTAB. so
    # aparecem na primeira linha de cada ano (nao repete nas seguintes,
    # igual ao modelo real) ---
    anos = [a for a in (dados_historico.get("anos") or []) if (a.get("ano_letivo") or "").strip()]

    linhas_totais = sum(max(len(a.get("disciplinas") or []), 1) for a in anos)
    tabela = document.add_table(rows=linhas_totais + 1, cols=9)
    tabela.style = "Table Grid"
    larguras_tabela = [1.6, 2.6, 1.5, 3.4, 1.3, 1.2, 1.7, 1.3, 3.9]
    _definir_largura_colunas(tabela, larguras_tabela)
    # Margens bem apertadas e fonte reduzida - com muitos anos letivos
    # lançados (ex: 9 anos x varias disciplinas), cada centimetro por
    # linha conta pra nao estourar em paginas demais.
    _definir_margens_celulas(tabela, cima_cm=0.03, baixo_cm=0.03, esquerda_cm=0.08, direita_cm=0.08)

    cabecalhos = ["Ano", "Ensino", "Fase", "Disciplina", "Nota", "C.H.", "Faltas", "Sit.", "Estab."]
    for celula, texto in zip(tabela.rows[0].cells, cabecalhos):
        celula.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _paragrafo(celula, texto, negrito=True, tamanho=TAM_TABELA_HISTORICO, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

    linha_atual = 1
    for ano in anos:
        disciplinas = ano.get("disciplinas") or [{}]
        situacao = SITUACOES_HISTORICO.get(ano.get("situacao"), "")
        for i, disciplina in enumerate(disciplinas):
            linha_tabela = tabela.rows[linha_atual]
            valores = [
                ano.get("ano_letivo", "") if i == 0 else "",
                ano.get("ensino", "") if i == 0 else "",
                ano.get("fase", "") if i == 0 else "",
                disciplina.get("nome", ""),
                disciplina.get("nota", ""),
                disciplina.get("carga_horaria", ""),
                disciplina.get("faltas", ""),
                situacao if i == 0 else "",
                ano.get("estabelecimento", "") if i == 0 else "",
            ]
            for celula, valor in zip(linha_tabela.cells, valores):
                celula.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                _paragrafo(celula, valor, tamanho=TAM_TABELA_HISTORICO, alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
            linha_atual += 1

    _paragrafo(document, "Legenda:  PRO – Promovido    RET – Retido    RFR – Retido por Frequência",
               tamanho=TAM_NORMAL - 2)

    # --- Amparo legal / regras (opcional, texto livre por escola) ---
    amparo_legal = (dados_historico.get("amparo_legal") or "").strip()
    regras = (dados_historico.get("regras") or "").strip()
    if amparo_legal or regras:
        document.add_paragraph()
        tabela_legal = document.add_table(rows=1, cols=2)
        tabela_legal.style = "Table Grid"
        _definir_largura_colunas(tabela_legal, [7.0, 11.5])
        _definir_margens_celulas(tabela_legal)
        cel_amparo, cel_regras = tabela_legal.rows[0].cells
        cel_amparo.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        _paragrafo(cel_amparo, "Amparo Legal", negrito=True)
        _paragrafo(cel_amparo, amparo_legal)
        cel_regras.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        _paragrafo(cel_regras, "Regras", negrito=True)
        _paragrafo(cel_regras, regras)

    document.add_paragraph()
    document.add_paragraph()

    data = (dados_historico.get("data") or "").strip()
    if data:
        _paragrafo(document, f"{data}.", alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

    document.add_paragraph()

    # --- Assinaturas: Diretor(a) e Secretario(a) lado a lado, cada uma
    # com a portaria de nomeacao (opcional) acima do nome ---
    tabela_assinatura = document.add_table(rows=3, cols=2)
    _remover_bordas_tabela(tabela_assinatura)
    _definir_largura_colunas(tabela_assinatura, [9.25, 9.25])

    diretor_portaria = (dados_historico.get("diretor_portaria") or "").strip()
    diretor_nome = (dados_historico.get("diretor_nome") or "").strip()
    secretario_portaria = (dados_historico.get("secretario_portaria") or "").strip()
    secretario_nome = (dados_historico.get("secretario_nome") or "").strip()

    _paragrafo(tabela_assinatura.cell(0, 0), diretor_portaria, tamanho=TAM_NORMAL - 1,
               alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    _paragrafo(tabela_assinatura.cell(0, 1), secretario_portaria, tamanho=TAM_NORMAL - 1,
               alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

    _linha_com_borda_inferior(tabela_assinatura.cell(1, 0))
    _linha_com_borda_inferior(tabela_assinatura.cell(1, 1))

    p_dir = _paragrafo(tabela_assinatura.cell(2, 0), diretor_nome or "Diretor(a)",
                        alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    p_dir.runs[0].italic = not diretor_nome
    _paragrafo(tabela_assinatura.cell(2, 0), "DIRETOR(A)", tamanho=TAM_NORMAL - 1,
               alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

    p_sec = _paragrafo(tabela_assinatura.cell(2, 1), secretario_nome or "Secretário(a) Escolar",
                        alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    p_sec.runs[0].italic = not secretario_nome
    _paragrafo(tabela_assinatura.cell(2, 1), "SECRETÁRIO(A)", tamanho=TAM_NORMAL - 1,
               alinhamento=WD_ALIGN_PARAGRAPH.CENTER)

    document.save(caminho_saida)
    return caminho_saida
