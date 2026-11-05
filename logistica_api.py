"""Integração com o site da LOGÍSTICA (API v1, somente leitura).

A Logística é a fonte oficial de tudo que é expedição e transporte; este site só
LÊ e exibe. Contrato: `briefing_api_logistica.md` (GET /api/v1/expedicoes,
GET /api/v1/pedidos/{pedido_venda}/expedicao, GET /api/v1/saude).

Este módulo é puro (não conhece Flask nem banco): cliente HTTP + conversão do
JSON da Logística para os campos `go_*` de PedidoOperacao. Quem grava no banco é
o app.py (ver `_logistica_aplicar_payload`).

Configuração (variáveis de ambiente, nunca no código):
  LOGISTICA_API_URL     endereço base do site da Logística, ex. https://logistica.exemplo.com.br
  LOGISTICA_API_TOKEN   chave secreta (Bearer)
  LOGISTICA_SYNC_MINUTOS  (opcional) intervalo da sincronização automática; padrão 5
Sem URL e chave, a integração fica DESLIGADA e o site funciona exatamente como antes.
"""

import json
import os
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, datetime
from urllib.parse import urlparse

TIMEOUT_SEGUNDOS = 10
LIMITE_PAGINA = 200
MAX_PAGINAS = 100  # trava de segurança contra paginação infinita


class LogisticaErro(Exception):
    """Falha ao falar com a Logística (rede, 401, 5xx, JSON inválido...)."""


class LogisticaNaoEncontrado(LogisticaErro):
    """404 — o pedido não existe na Logística."""


def url_base():
    """Endereço base da Logística. Aceita com ou sem o sufixo /api/v1 no final
    (quem cadastra costuma colar o endereço completo da API)."""
    base = (os.environ.get("LOGISTICA_API_URL") or "").strip().rstrip("/")
    if base.endswith("/api/v1"):
        base = base[: -len("/api/v1")]
    return base


def _token():
    return (os.environ.get("LOGISTICA_API_TOKEN") or "").strip()


def configurado():
    return bool(url_base() and _token())


def intervalo_sync_minutos():
    try:
        return max(1, int(os.environ.get("LOGISTICA_SYNC_MINUTOS", "5")))
    except ValueError:
        return 5


def _validar_url(base):
    """Só HTTPS — exceto máquina local (testes)."""
    p = urlparse(base)
    if p.scheme == "https":
        return
    if p.scheme == "http" and (p.hostname or "") in ("localhost", "127.0.0.1", "::1"):
        return
    raise LogisticaErro("LOGISTICA_API_URL precisa ser https:// (http só é aceito em localhost).")


def _get(caminho, params=None, timeout=TIMEOUT_SEGUNDOS):
    base = url_base()
    if not base or not _token():
        raise LogisticaErro("Integração com a Logística não configurada.")
    _validar_url(base)
    url = base + caminho
    if params:
        limpos = {k: v for k, v in params.items() if v is not None}
        if limpos:
            url += "?" + urllib.parse.urlencode(limpos)
    req = urllib.request.Request(url, headers={
        "Authorization": f"Bearer {_token()}",  # a chave nunca é registrada em log
        "Accept": "application/json",
        "User-Agent": "gestao-producao-operacao/1",
    })
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            corpo = resp.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        if e.code == 404:
            raise LogisticaNaoEncontrado("Pedido não encontrado na Logística.") from None
        if e.code in (401, 403):
            raise LogisticaErro("A Logística recusou a chave de acesso (HTTP %s)." % e.code) from None
        raise LogisticaErro("A Logística respondeu com erro HTTP %s." % e.code) from None
    except (urllib.error.URLError, TimeoutError, OSError) as e:
        motivo = getattr(e, "reason", e)
        raise LogisticaErro("Não foi possível falar com a Logística (%s)." % motivo) from None
    try:
        return json.loads(corpo)
    except ValueError:
        raise LogisticaErro("Resposta da Logística não é um JSON válido.") from None


def buscar_saude():
    return _get("/api/v1/saude")


def buscar_mudancas(atualizado_desde=None):
    """Percorre todas as páginas de /api/v1/expedicoes. Devolve lista de pedidos
    (dicts do contrato), em ordem crescente de `atualizado_em`."""
    pedidos = []
    pagina = 1
    while pagina <= MAX_PAGINAS:
        dados = _get("/api/v1/expedicoes", {
            "atualizado_desde": atualizado_desde, "limite": LIMITE_PAGINA, "pagina": pagina,
        })
        lote = dados.get("pedidos") if isinstance(dados, dict) else None
        if not isinstance(lote, list):
            raise LogisticaErro("Resposta de /expedicoes sem a lista 'pedidos'.")
        pedidos.extend(p for p in lote if isinstance(p, dict))
        if not dados.get("tem_mais"):
            break
        pagina += 1
    return pedidos


def buscar_pedido(pedido_venda, timeout=TIMEOUT_SEGUNDOS):
    texto = urllib.parse.quote(str(pedido_venda).strip(), safe="")
    dados = _get(f"/api/v1/pedidos/{texto}/expedicao", timeout=timeout)
    # formato combinado: {versao, gerado_em, pedido}; aceita também o objeto
    # direto ou o envelope {"pedidos": [ ... ]}
    if isinstance(dados, dict) and isinstance(dados.get("pedido"), dict):
        return dados["pedido"]
    if isinstance(dados, dict) and isinstance(dados.get("pedidos"), list):
        return dados["pedidos"][0] if dados["pedidos"] else None
    return dados if isinstance(dados, dict) else None


# ----------------------------------------------------------------------------
# Conversão do contrato -> campos do site
# ----------------------------------------------------------------------------

def _data(v):
    if not v:
        return None
    if isinstance(v, date) and not isinstance(v, datetime):
        return v
    try:
        return datetime.strptime(str(v)[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def _num(v):
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _texto(v):
    if v is None:
        return None
    t = str(v).strip()
    return t or None


def parse_instante(v):
    """ISO 8601 UTC -> datetime ingênuo em UTC (padrão do banco); None se inválido."""
    if not v:
        return None
    try:
        t = str(v).strip().replace("Z", "+00:00")
        d = datetime.fromisoformat(t)
        if d.tzinfo is not None:
            from datetime import timezone
            d = d.astimezone(timezone.utc).replace(tzinfo=None)
        return d
    except ValueError:
        return None


# campos de PedidoOperacao que a Logística passa a controlar (ficam travados na tela)
CAMPOS_CONTROLADOS = (
    "go_status_logistica",
    "go_data_prevista_coleta",
    "go_data_pedido_expedido",
    "go_data_prevista_entrega",
    "go_data_real_entrega",
    "go_data_entregue_cliente",
    "go_transportadora_nome",  # virtual: vira go_transportadora_id no app.py
    "go_custo_frete_previsto",
    "go_custo_frete_final",
    "go_custo_frete_sobre_nota",
    "go_numero_nf",
    "go_data_emissao_nf",
    "go_valor_nf_emitida",
)

# nomes reais de coluna travados no formulário (go_transportadora_nome -> go_transportadora_id)
CAMPOS_TRAVADOS_FORM = (
    "go_status_logistica",
    "go_data_prevista_coleta",
    "go_data_pedido_expedido",
    "go_data_prevista_entrega",
    "go_data_real_entrega",
    "go_data_entregue_cliente",
    "go_transportadora_id",
    "go_custo_frete_previsto",
    "go_custo_frete_final",
    "go_custo_frete_sobre_nota",
    "go_numero_nf",
    "go_data_emissao_nf",
    "go_valor_nf_emitida",
)


def normalizar_pedido(texto):
    """Mesmo critério do site: sem espaços e sem zeros à esquerda ("000872" = "872").
    A parte depois da barra é mantida ("931/1")."""
    t = str(texto or "").strip()
    if not t:
        return ""
    return t.lstrip("0") or "0"


def base_do_pedido(texto):
    """"931/1" -> "931" (número base, que é a chave na Logística)."""
    return normalizar_pedido(str(texto or "").split("/")[0])


def _remessas(payload):
    return [r for r in (payload.get("remessas") or []) if isinstance(r, dict)]


def _remessa_da_parte(payload, parte):
    alvo = normalizar_pedido(parte)
    for r in _remessas(payload):
        if normalizar_pedido(r.get("id")) == alvo:
            return r
    return None


def mapear(payload, parte=None):
    """JSON de UM pedido da Logística -> {campo_go: valor | None} para a PARTE
    pedida (ex.: "931", "931/1"); None se a Logística ainda não tem dado dessa parte.

    A Logística usa o nº base ("931") e lista cada parte em remessas[] (ids
    "931", "931/1", "931/2"...). Pedido de uma remessa só (ou parte sem
    divisão) usa o resumo do pedido; com várias partes, cada linha do site pega
    os dados da sua própria remessa.

    Várias NFs: número = todos separados por vírgula; valor = soma; data de
    emissão = a mais recente. Campo desconhecido na Logística vem None."""
    remessas = _remessas(payload)
    frete = payload.get("frete") if isinstance(payload.get("frete"), dict) else {}
    nfs = [n for n in (payload.get("notas_fiscais") or []) if isinstance(n, dict)]
    status = _texto(payload.get("status_logistica"))
    transp = _texto(payload.get("transportadora"))
    coleta = _data(payload.get("data_prevista_coleta"))
    exped = _data(payload.get("data_expedicao"))
    prev_ent = _data(payload.get("data_prevista_entrega"))
    entrega = _data(payload.get("data_entrega_cliente"))
    chegada = _data(payload.get("data_chegada_destino"))
    usar_frete = True

    parte_norm = normalizar_pedido(parte) if parte else ""
    tem_sufixo = "/" in parte_norm
    rem = _remessa_da_parte(payload, parte_norm) if parte_norm else None
    if rem is not None and len(remessas) > 1:
        # linha do site = uma parte específica: dados dessa remessa
        exped = _data(rem.get("data_expedicao"))
        prev_ent = _data(rem.get("data_prevista_entrega"))
        entrega = _data(rem.get("data_entrega_cliente"))
        chegada = entrega
        transp = _texto(rem.get("transportadora"))
        coleta = None
        numeros_rem = {str(x).strip() for x in (rem.get("notas_fiscais") or [])}
        nfs = [n for n in nfs if str(n.get("numero")).strip() in numeros_rem]
        usar_frete = not tem_sufixo  # frete é do pedido: fica na linha base
        if entrega:
            status = "ENTREGUE"
        elif exped:
            status = "EM TRANSPORTE"
    elif parte_norm and tem_sufixo and rem is None:
        return None  # parte que a Logística ainda não conhece
    elif parte_norm and len(remessas) > 1 and rem is None and not tem_sufixo:
        pass  # linha base sem remessa própria: usa o resumo do pedido

    numeros = [_texto(n.get("numero")) for n in nfs]
    numeros = [n for n in numeros if n]
    valores = [_num(n.get("valor")) for n in nfs]
    valores = [v for v in valores if v is not None]
    emissoes = [_data(n.get("data_emissao")) for n in nfs]
    emissoes = [d for d in emissoes if d]
    return {
        "go_status_logistica": status,
        "go_data_prevista_coleta": coleta,
        "go_data_pedido_expedido": exped,
        "go_data_prevista_entrega": prev_ent,
        "go_data_real_entrega": chegada,
        "go_data_entregue_cliente": entrega,
        "go_transportadora_nome": transp,
        "go_custo_frete_previsto": _num(frete.get("previsto")) if usar_frete else None,
        "go_custo_frete_final": _num(frete.get("final")) if usar_frete else None,
        "go_custo_frete_sobre_nota": _num(frete.get("percentual_sobre_nota")) if usar_frete else None,
        "go_numero_nf": ", ".join(numeros)[:30] if numeros else None,
        "go_data_emissao_nf": max(emissoes) if emissoes else None,
        "go_valor_nf_emitida": round(sum(valores), 2) if valores else None,
    }


def resumo_remessas(payload):
    """Para exibir "2 de 3 remessas entregues" etc.
    Devolve dict(total, entregues, expedidas, lista=[...]) ou None sem remessas."""
    rem = [r for r in (payload.get("remessas") or []) if isinstance(r, dict)]
    if not rem:
        return None
    lista = []
    for r in rem:
        lista.append({
            "id": _texto(r.get("id")),
            "data_expedicao": _data(r.get("data_expedicao")),
            "data_prevista_entrega": _data(r.get("data_prevista_entrega")),
            "data_entrega_cliente": _data(r.get("data_entrega_cliente")),
            "transportadora": _texto(r.get("transportadora")),
            "rastreio": _texto(r.get("rastreio")),
            "notas_fiscais": [str(x) for x in (r.get("notas_fiscais") or [])],
        })
    return {
        "total": len(lista),
        "expedidas": sum(1 for r in lista if r["data_expedicao"]),
        "entregues": sum(1 for r in lista if r["data_entrega_cliente"]),
        "lista": lista,
    }
