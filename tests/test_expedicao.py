"""Liberação efetiva NÃO é expedição (Bruno, 27/11/2026).

Pedido FOB liberado para o cliente coletar ganhava "data de expedição" e aparecia
no aviso "Material expedido" do Painel sem NF e sem transportadora. A data só vale
como expedição quando há NF emitida (ou registro manual de embarque/baixa)."""
from datetime import date, datetime

import app as A
from conftest import entrar
from models import PedidoOperacao, db


def _po(pv, cliente, **kw):
    po = PedidoOperacao(pedido_venda=pv, cliente=cliente, frete="FOB", **kw)
    db.session.add(po)
    return po


def test_sem_nf_e_so_liberado_para_coleta(app):
    hoje = date(2026, 11, 27)
    with app.app_context():
        _po("T-EXP-1", "SO LIBERADO", go_data_pedido_expedido=hoje, go_data_efetiva_liberacao_pcp=hoje)
        _po("T-EXP-2", "COM NF", go_data_pedido_expedido=hoje, go_data_efetiva_liberacao_pcp=hoje, go_numero_nf="8127")
        _po("T-EXP-3", "EMBARQUE MANUAL", go_data_pedido_expedido=hoje, go_embarque_manual_em=datetime(2026, 11, 27, 9, 0))
        db.session.commit()
        r = A._expedidos_recentes_painel(hoje)
        expedidos = {i["pedido_venda"] for i in r["hoje"]}
        assert "T-EXP-2" in expedidos and "T-EXP-3" in expedidos
        assert "T-EXP-1" not in expedidos
        assert "T-EXP-1" in r["liberados_sem_expedicao"]
        assert len(r["hoje"]) == 2


def test_painel_mostra_aviso_de_liberado_sem_expedicao(app):
    with app.app_context():
        hoje = A._hoje_brt()
        _po("T-EXP-9", "CLIENTE SO LIBERADO", go_data_pedido_expedido=hoje)
        db.session.commit()
    c = app.test_client()
    entrar(c, "t_admin")
    html = c.get("/painel").get_data(as_text=True)
    assert "liberado(s) para coleta, ainda sem NF/expedição" in html and "T-EXP-9" in html
