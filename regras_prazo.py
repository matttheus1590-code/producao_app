"""Regras ÚNICAS de prazo: "em aberto", "atrasado no plano" e "vencido ao cliente".

Por que existe: a mesma palavra "atrasado" aparecia calculada de jeitos
diferentes em várias telas. Aqui ficam as definições oficiais; as telas e os
relatórios chamam estas funções em vez de reescrever a conta.

Definições (itens = ItemPedido, pedidos = Pedido):

* EM ABERTO (item): status_producao != "FINALIZADO".
* ATRASADO NO PLANO (item): em aberto E liberação prevista (PCP) já vencida
  (liberacao_prevista < hoje). É o desvio em relação ao que o próprio PCP
  planejou. Sem previsão -> NÃO conta como atrasado (conta como "sem previsão").
* VENCIDO AO CLIENTE (pedido/item): em aberto E data prometida ao cliente
  (Pedido.data_cliente) já vencida. É o atraso que o cliente sente.
* PEDIDO ATRASADO NO PLANO: tem ao menos um item atrasado no plano.

"Hoje" é sempre um parâmetro (default date.today()) para permitir teste.
Este módulo não importa app nem models: as funções SQL recebem a classe.
"""
from datetime import date

STATUS_FINAL = "FINALIZADO"


def _hoje(hoje=None):
    return hoje or date.today()


# ----------------------------- Python (objetos) -----------------------------
def item_aberto(item):
    return item.status_producao != STATUS_FINAL


def item_atrasado_plano(item, hoje=None):
    prev = item.liberacao_prevista
    return bool(item_aberto(item) and prev and prev < _hoje(hoje))


def dias_atraso_plano(item, hoje=None):
    """Dias de atraso no plano (0 se não está atrasado)."""
    if item_atrasado_plano(item, hoje):
        return (_hoje(hoje) - item.liberacao_prevista).days
    return 0


def item_vencido_cliente(item, hoje=None):
    pedido = getattr(item, "pedido", None)
    cli = pedido.data_cliente if pedido is not None else None
    return bool(item_aberto(item) and cli and cli < _hoje(hoje))


def pedido_atrasado_plano(pedido, hoje=None):
    return any(item_atrasado_plano(i, hoje) for i in pedido.itens)


def pedido_vencido_cliente(pedido, hoje=None):
    cli = pedido.data_cliente
    if not cli or cli >= _hoje(hoje):
        return False
    return any(item_aberto(i) for i in pedido.itens)


# ------------------------------ SQL (consultas) ------------------------------
def sql_item_aberto(ItemPedido):
    return ItemPedido.status_producao != STATUS_FINAL


def sql_item_atrasado_plano(ItemPedido, hoje=None):
    from sqlalchemy import and_
    return and_(
        ItemPedido.status_producao != STATUS_FINAL,
        ItemPedido.liberacao_prevista.isnot(None),
        ItemPedido.liberacao_prevista < _hoje(hoje),
    )
