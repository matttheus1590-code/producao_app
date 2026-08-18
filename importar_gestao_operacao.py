"""Import único (Fase 13) da planilha "Gestão de Fluxo Produtivo" para dentro
dos pedidos que já existem no banco.

Roda por cima de um banco JÁ POVOADO (diferente de seed.py, que só roda em
banco vazio) — por isso fica isolado neste módulo, chamado por
`_importar_gestao_operacao(app)` em app.py, protegido por uma checagem de
"já rodou antes" (ver comentário lá).

Estratégia de casamento por `Nº PEDIDO TOTVS` (== Pedido.pedido_venda):
  1. Igual exato.
  2. Aproximado: um valor "contido" no outro (cobre casos como planilha
     tendo "253" e o banco já tendo "253 (229)").
  3. Sem nenhum match -> cria um Pedido novo só com os dados comerciais
     (sem item de produção — aparece com "—" na Listagem Geral, igual
     qualquer pedido sem item hoje).

Quando VÁRIOS registros Pedido do banco compartilham o mesmo pedido_venda
(o caso legado do import original: 1146 registros para 385 pedidos reais),
os campos novos (go_*) são gravados em TODOS eles — assim a listagem/filtro
funciona igual não importa qual registro você está vendo.

Importante: nunca sobrescreve campos que já existiam antes desta fase
(cliente, vendedor, pedido_venda, data_inclusao_pedido, prioridade, frete,
país, estado, cidade) em pedidos que já casaram — só define esses campos
quando o pedido é novo (sem nenhum match). Os campos go_* são sempre
gravados/atualizados, pois são 100% novos.

Limitação conhecida: ~4 das 420 linhas da planilha têm "Nº PEDIDO TOTVS"
inválido/placeholder ("-", "N/A") — sem um número real pra casar, viram
sempre um Pedido novo. Em uso normal isso roda só UMA vez (protegido pelo
guard em app.py), então não duplica; só duplicaria se alguém apagasse
manualmente os campos go_* do banco pra forçar o import rodar de novo — um
cenário que não acontece no fluxo normal do site.
"""

import re
import unicodedata
from datetime import date, datetime

import openpyxl

from extensions import db
from models import Pedido, Transportadora

SHEET_NAME = "GESTAO OPERACAO"

COL = {
    "cliente": 1,
    "vendedor": 2,
    "tipo_pedido": 3,
    "contrato": 4,
    "pedido_compra_cliente": 5,
    "proposta": 6,
    "pedido_venda": 7,
    "data_inclusao_pedido": 8,
    "data_solicitada_entrega": 9,
    "modalidade_frete": 11,
    "pais": 12,
    "estado": 13,
    "cidade": 14,
    "status_pedido_info": 15,
    "valor_total_pedido": 16,
    "criticidade": 17,
    "previsao_liberacao_pcp": 20,
    "data_efetiva_liberacao_pcp": 21,
    "data_solicitada_cliente_retira": 22,
    "custo_producao_real": 26,
    "termino_semanal_pcp": 27,
    "data_emissao_nf": 28,
    "valor_nf_emitida": 29,
    "numero_nf": 30,
    "status_logistica": 31,
    "data_pedido_expedido": 32,
    "transportadora": 34,
    "custo_frete_previsto": 35,
    "custo_frete_final": 36,
    "custo_frete_sobre_nota": 37,
    "data_prevista_entrega": 38,
    "data_real_entrega": 39,
    "otd": 41,
    "data_solicitada_cliente_final": 42,
    "data_entregue_cliente": 43,
    "obs": 47,
    "status_final_alinhamento": 48,
}

FIRST_ROW = 2
LAST_ROW = 421  # linha 421 é a última com CLIENTE preenchido (420 pedidos)

_VALORES_INVALIDOS_PEDIDO_VENDA = {"-", "N/A", "NA", "", "NONE"}


def _cell(ws, row, campo):
    return ws.cell(row=row, column=COL[campo]).value


_MARCADORES_VAZIO = {"-", "--", "N/A", "NA", "N/D", "ND"}


def _parse_texto(v, max_len=None):
    if v is None:
        return None
    s = str(v).strip()
    if not s or s.upper() in _MARCADORES_VAZIO:
        return None
    if max_len:
        s = s[:max_len]
    return s


def _parse_data(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    if not s:
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _parse_numero(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("R$", "").replace(".", "").replace(",", ".").strip()
    try:
        return float(s)
    except ValueError:
        return None


def _parse_otd(v):
    s = _parse_texto(v)
    if not s:
        return None
    s = s.strip().upper()
    if s in ("SIM", "S", "YES", "TRUE"):
        return "SIM"
    if s in ("NÃO", "NAO", "N", "NO", "FALSE"):
        return "NÃO"
    return s[:10]


# ---------------------------------------------------------------------------
# Normalização de transportadoras: agrupa grafias diferentes da mesma empresa
# (variação de espaço, pontuação, truncamento) por prefixo comum.
# ---------------------------------------------------------------------------
def _chave_normalizada(s):
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    s = s.strip().upper()
    s = re.sub(r"[.\-,/]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def construir_mapa_transportadoras(nomes_brutos):
    """Recebe uma lista de nomes brutos (como aparecem na planilha) e devolve
    um dict {nome_bruto_normalizado_de_exibicao: nome_canonico}."""
    display_por_chave = {}
    for n in nomes_brutos:
        disp = re.sub(r"\s+", " ", str(n).strip().upper())
        chave = _chave_normalizada(disp)
        display_por_chave.setdefault(chave, []).append(disp)

    chaves = sorted(display_por_chave.keys(), key=lambda k: -len(display_por_chave[k]))
    usadas = set()
    canonico_por_chave = {}
    for k in chaves:
        if k in usadas:
            continue
        grupo = [k]
        usadas.add(k)
        for outra in chaves:
            if outra in usadas:
                continue
            menor, maior = (k, outra) if len(k) <= len(outra) else (outra, k)
            if len(menor) >= 8 and maior.startswith(menor):
                grupo.append(outra)
                usadas.add(outra)
        todos_disp = [d for gk in grupo for d in display_por_chave[gk]]
        canonico = max(todos_disp, key=len)
        for gk in grupo:
            canonico_por_chave[gk] = canonico

    mapa = {}
    for n in nomes_brutos:
        disp = re.sub(r"\s+", " ", str(n).strip().upper())
        chave = _chave_normalizada(disp)
        mapa[disp] = canonico_por_chave[chave]
    return mapa


def _obter_ou_criar_transportadora(nome_canonico, cache):
    if nome_canonico in cache:
        return cache[nome_canonico]
    t = Transportadora.query.filter_by(nome=nome_canonico).first()
    if not t:
        t = Transportadora(nome=nome_canonico, ativo=True)
        db.session.add(t)
        db.session.flush()
    cache[nome_canonico] = t
    return t


_TOKEN_MIN_LEN = 2  # ignora tokens numéricos muito curtos ("1", "2") — geram falso positivo


def _tokens_numericos(s):
    """Extrai os "números de pedido" de dentro de um valor composto, ex.:
    "253 (229)" -> {"253", "229"}; "184/235" -> {"184", "235"}. Usa igualdade
    exata de cada token (não substring "contém"), pra não confundir "146" com
    "14" (que seria um falso positivo perigoso)."""
    return {t for t in re.findall(r"\d+", s) if len(t) >= _TOKEN_MIN_LEN}


def _construir_indice_tokens(exatos):
    """A partir do índice {pedido_venda_exato: [pedidos]} já existente, monta
    um índice adicional {token_numerico: [pedidos]} pra casamento aproximado."""
    indice = {}
    for pv, pedidos in exatos.items():
        for tok in _tokens_numericos(pv):
            indice.setdefault(tok, []).extend(pedidos)
    return indice


def _match_pedidos(valor, exatos, indice_tokens):
    """Devolve (lista_de_pedidos, tipo) onde tipo é 'exato', 'aproximado' ou None.

    Casamento aproximado exige que algum "número de pedido" dentro do valor da
    planilha seja EXATAMENTE igual a algum número de pedido já indexado a
    partir dos pedido_venda do banco — nunca por "contém como substring"
    (isso é o que evita casar "146" com "14")."""
    valor = (valor or "").strip()
    if not valor or valor.upper() in _VALORES_INVALIDOS_PEDIDO_VENDA:
        return [], None
    if valor in exatos:
        return exatos[valor], "exato"

    encontrados = []
    for tok in _tokens_numericos(valor):
        if tok in indice_tokens:
            encontrados.extend(indice_tokens[tok])
    if encontrados:
        vistos = set()
        unicos = []
        for p in encontrados:
            if p.id not in vistos:
                vistos.add(p.id)
                unicos.append(p)
        return unicos, "aproximado"
    return [], None


def importar_gestao_operacao(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb[SHEET_NAME]

    # -- monta índice de pedido_venda existentes no banco --
    exatos = {}
    for p in Pedido.query.all():
        pv = (p.pedido_venda or "").strip()
        if pv:
            exatos.setdefault(pv, []).append(p)
    indice_tokens = _construir_indice_tokens(exatos)

    # -- coleta nomes brutos de transportadora e monta o mapa canônico --
    nomes_transportadora = []
    for row in range(FIRST_ROW, LAST_ROW + 1):
        v = _cell(ws, row, "transportadora")
        if v:
            nomes_transportadora.append(v)
    mapa_transportadora = construir_mapa_transportadoras(nomes_transportadora)
    cache_transportadora = {}

    stats = {"total": 0, "exato": 0, "aproximado": 0, "novos": 0, "sem_match": 0}
    linhas_sem_match = []
    linhas_aproximadas = []

    for row in range(FIRST_ROW, LAST_ROW + 1):
        cliente = _parse_texto(_cell(ws, row, "cliente"))
        if not cliente:
            continue
        stats["total"] += 1

        pedido_venda_raw = _parse_texto(_cell(ws, row, "pedido_venda")) or ""
        pedidos, tipo_match = _match_pedidos(pedido_venda_raw, exatos, indice_tokens)

        if tipo_match == "exato":
            stats["exato"] += 1
        elif tipo_match == "aproximado":
            stats["aproximado"] += 1
            linhas_aproximadas.append((row, pedido_venda_raw, [p.pedido_venda for p in pedidos]))

        transportadora_obj = None
        transp_raw = _cell(ws, row, "transportadora")
        if transp_raw:
            disp = re.sub(r"\s+", " ", str(transp_raw).strip().upper())
            canonico = mapa_transportadora.get(disp, disp)
            transportadora_obj = _obter_ou_criar_transportadora(canonico, cache_transportadora)

        campos_go = dict(
            go_tipo_pedido=_parse_texto(_cell(ws, row, "tipo_pedido"), 60),
            go_contrato=_parse_texto(_cell(ws, row, "contrato"), 60),
            go_pedido_compra_cliente=_parse_texto(_cell(ws, row, "pedido_compra_cliente"), 60),
            go_proposta=_parse_texto(_cell(ws, row, "proposta"), 60),
            go_data_solicitada_entrega=_parse_data(_cell(ws, row, "data_solicitada_entrega")),
            go_status_pedido_info=_parse_texto(_cell(ws, row, "status_pedido_info"), 120),
            go_valor_pedido_operacao=_parse_numero(_cell(ws, row, "valor_total_pedido")),
            go_previsao_liberacao_pcp=_parse_data(_cell(ws, row, "previsao_liberacao_pcp")),
            go_data_efetiva_liberacao_pcp=_parse_data(_cell(ws, row, "data_efetiva_liberacao_pcp")),
            go_data_solicitada_cliente_retira=_parse_data(_cell(ws, row, "data_solicitada_cliente_retira")),
            go_custo_producao_real=_parse_numero(_cell(ws, row, "custo_producao_real")),
            go_termino_semanal_pcp=_parse_texto(_cell(ws, row, "termino_semanal_pcp"), 40),
            go_data_emissao_nf=_parse_data(_cell(ws, row, "data_emissao_nf")),
            go_valor_nf_emitida=_parse_numero(_cell(ws, row, "valor_nf_emitida")),
            go_numero_nf=_parse_texto(_cell(ws, row, "numero_nf"), 30),
            go_status_logistica=_parse_texto(_cell(ws, row, "status_logistica"), 60),
            go_data_pedido_expedido=_parse_data(_cell(ws, row, "data_pedido_expedido")),
            go_transportadora_id=transportadora_obj.id if transportadora_obj else None,
            go_custo_frete_previsto=_parse_numero(_cell(ws, row, "custo_frete_previsto")),
            go_custo_frete_final=_parse_numero(_cell(ws, row, "custo_frete_final")),
            go_custo_frete_sobre_nota=_parse_numero(_cell(ws, row, "custo_frete_sobre_nota")),
            go_data_prevista_entrega=_parse_data(_cell(ws, row, "data_prevista_entrega")),
            go_data_real_entrega=_parse_data(_cell(ws, row, "data_real_entrega")),
            go_otd_realizado=_parse_otd(_cell(ws, row, "otd")),
            go_data_solicitada_cliente_final=_parse_data(_cell(ws, row, "data_solicitada_cliente_final")),
            go_data_entregue_cliente=_parse_data(_cell(ws, row, "data_entregue_cliente")),
            go_obs_operacao=_parse_texto(_cell(ws, row, "obs")),
            go_status_final_alinhamento=_parse_texto(_cell(ws, row, "status_final_alinhamento"), 60),
        )

        if pedidos:
            # já existe(m) — só grava os campos novos (go_*), nunca mexe nos
            # campos que já existiam antes desta fase.
            for p in pedidos:
                for campo, valor in campos_go.items():
                    setattr(p, campo, valor)
        else:
            stats["sem_match"] += 1
            linhas_sem_match.append((row, pedido_venda_raw, cliente))
            stats["novos"] += 1

            criticidade = _parse_texto(_cell(ws, row, "criticidade"))
            prioridade = None
            if criticidade:
                cu = criticidade.strip().upper()
                if cu in ("MEDIA", "MÉDIA"):
                    prioridade = "MÉDIA"
                elif cu in ("BAIXA", "ALTA"):
                    prioridade = cu

            pedido_venda_valido = (
                pedido_venda_raw if pedido_venda_raw.upper() not in _VALORES_INVALIDOS_PEDIDO_VENDA else None
            )
            novo = Pedido(
                cliente=cliente,
                vendedor=_parse_texto(_cell(ws, row, "vendedor")),
                pedido_venda=pedido_venda_valido,
                data_inclusao_pedido=_parse_data(_cell(ws, row, "data_inclusao_pedido")),
                pais=_parse_texto(_cell(ws, row, "pais")) or "Brasil",
                estado=_parse_texto(_cell(ws, row, "estado")),
                cidade=_parse_texto(_cell(ws, row, "cidade")),
                frete=_parse_texto(_cell(ws, row, "modalidade_frete")),
                prioridade=prioridade or "MÉDIA",
                **campos_go,
            )
            db.session.add(novo)
            # registra nos índices pra eventuais linhas seguintes com o mesmo pedido_venda
            if pedido_venda_raw:
                exatos.setdefault(pedido_venda_raw, []).append(novo)
                for tok in _tokens_numericos(pedido_venda_raw):
                    indice_tokens.setdefault(tok, []).append(novo)

    db.session.commit()

    stats["linhas_sem_match"] = linhas_sem_match
    stats["linhas_aproximadas"] = linhas_aproximadas
    stats["transportadoras_canonicas"] = sorted(set(mapa_transportadora.values()))
    return stats
