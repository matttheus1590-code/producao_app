from datetime import date, timedelta
from types import SimpleNamespace as NS

import regras_prazo as R

HOJE = date(2026, 6, 15)
ONTEM = HOJE - timedelta(days=1)
AMANHA = HOJE + timedelta(days=1)


def item(status="ANDAMENTO", prev=None, cli=None):
    return NS(status_producao=status, liberacao_prevista=prev, pedido=NS(data_cliente=cli, itens=[]))


def test_finalizado_nunca_esta_atrasado():
    assert not R.item_atrasado_plano(item("FINALIZADO", ONTEM), HOJE)
    assert R.dias_atraso_plano(item("FINALIZADO", ONTEM), HOJE) == 0


def test_sem_previsao_nao_conta_como_atrasado():
    assert not R.item_atrasado_plano(item("PENDENTE", None), HOJE)


def test_previsao_vencida_em_aberto_e_atrasado():
    i = item("PENDENTE", HOJE - timedelta(days=3))
    assert R.item_atrasado_plano(i, HOJE)
    assert R.dias_atraso_plano(i, HOJE) == 3


def test_previsao_hoje_ou_futura_nao_e_atrasado():
    assert not R.item_atrasado_plano(item("ANDAMENTO", HOJE), HOJE)
    assert not R.item_atrasado_plano(item("ANDAMENTO", AMANHA), HOJE)


def test_em_tratativa_conta_como_aberto():
    assert R.item_aberto(item("EM TRATATIVA"))
    assert R.item_atrasado_plano(item("EM TRATATIVA", ONTEM), HOJE)


def test_vencido_ao_cliente_e_diferente_de_atrasado_no_plano():
    i = item("ANDAMENTO", prev=AMANHA, cli=ONTEM)
    assert R.item_vencido_cliente(i, HOJE)
    assert not R.item_atrasado_plano(i, HOJE)
    j = item("ANDAMENTO", prev=ONTEM, cli=AMANHA)
    assert R.item_atrasado_plano(j, HOJE)
    assert not R.item_vencido_cliente(j, HOJE)


def test_pedido():
    ok = item("FINALIZADO", ONTEM)
    atras = item("ANDAMENTO", ONTEM)
    ped = NS(data_cliente=ONTEM, itens=[ok, atras])
    assert R.pedido_atrasado_plano(ped, HOJE)
    assert R.pedido_vencido_cliente(ped, HOJE)
    tudo_final = NS(data_cliente=ONTEM, itens=[ok])
    assert not R.pedido_atrasado_plano(tudo_final, HOJE)
    assert not R.pedido_vencido_cliente(tudo_final, HOJE)
