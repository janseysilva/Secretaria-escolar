"""Geração de documentos oficiais de secretaria escolar (memorando, ofício, etc.)."""

import docx
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

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
    if hasattr(celula_ou_doc, "paragraphs") and hasattr(celula_ou_doc, "add_paragraph"):
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
    """Limpa o paragrafo padrao vazio que toda celula nova ja vem com."""
    if celula.paragraphs and not celula.paragraphs[0].runs:
        celula.paragraphs[0].text = ""
    celula.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    return celula


def gerar_memorando(dados_escola, dados_memo, caminho_saida):
    """Gera um Memorando (.docx) a partir dos dados da escola e do memorando.

    Todo o documento fica em uma unica tabela com bordas visiveis, separando
    cada bloco (cabecalho, DE/PARA, ASSUNTO, corpo, rodape) em quadros, igual
    ao modelo de memorando usado nas secretarias escolares. Recebido por e a
    segunda Data ficam em branco no documento, para preencher a mao. O resto
    ja vem preenchido pelo que o usuario digitou/cadastrou.

    dados_escola: dict com nome_escola, secretaria, diretor_nome, diretor_cargo,
        diretor_portaria, logo_path, assinatura_path.
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
