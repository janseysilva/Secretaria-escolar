"""Controle do teste gratis de 7 dias e da licenca do app de Secretaria
Escolar. O app abre normalmente mesmo com o teste vencido - so a
geracao de documentos fica bloqueada ate a licenca ser ativada.

O pagamento em si (PIX/QR code) ainda nao esta integrado - por enquanto
"licenca_ativa" so pode ser ligada manualmente editando o arquivo
licenca.json, o mesmo tipo de "modo de espera" usado nos anuncios/compra
do Espertinhos antes das contas de pagamento reais existirem.
"""

import datetime
import json
import os

DIAS_TESTE_GRATIS = 7

PASTA_BASE = os.path.dirname(os.path.abspath(__file__))
PASTA_DADOS = os.path.join(PASTA_BASE, "dados_escola")
ARQUIVO_LICENCA = os.path.join(PASTA_DADOS, "licenca.json")


def _carregar():
    if os.path.exists(ARQUIVO_LICENCA):
        with open(ARQUIVO_LICENCA, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def _salvar(dados):
    os.makedirs(PASTA_DADOS, exist_ok=True)
    with open(ARQUIVO_LICENCA, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)


def garantir_licenca_iniciada():
    """Cria o arquivo de licenca na primeira vez que o app roda,
    marcando hoje como o inicio do teste gratis. Chamado uma vez, ao
    abrir o app."""
    dados = _carregar()
    if "data_primeiro_uso" not in dados:
        dados["data_primeiro_uso"] = datetime.date.today().isoformat()
        dados.setdefault("licenca_ativa", False)
        _salvar(dados)


def dias_restantes_teste():
    """Quantos dias ainda restam do teste gratis (pode ser negativo se
    ja venceu ha algum tempo)."""
    dados = _carregar()
    data_texto = dados.get("data_primeiro_uso")
    if not data_texto:
        return DIAS_TESTE_GRATIS
    inicio = datetime.date.fromisoformat(data_texto)
    dias_passados = (datetime.date.today() - inicio).days
    return DIAS_TESTE_GRATIS - dias_passados


def licenca_ativa():
    return bool(_carregar().get("licenca_ativa"))


def licenca_valida():
    """True se o programa pode gerar documentos - licenca paga ativa,
    ou ainda dentro do periodo de teste gratis."""
    return licenca_ativa() or dias_restantes_teste() > 0
