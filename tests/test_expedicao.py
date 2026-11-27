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


def test_fob_liberado_sem_nf_continua_parado_na_expedicao(app, monkeypatch):
    """O FOB 'liberado p/ coleta' (Logística manda data de expedição, mas sem NF) tem que aparecer
    no Kanban de parados em vez de sumir; com NF ele é expedição de verdade e sai do quadro."""
    import json
    from models import LogisticaSync
    hoje = date(2026, 11, 27)
    monkeypatch.setattr(A.logistica_api, "configurado", lambda: True)
    payload = json.dumps({"data_expedicao": "2026-11-27", "data_entrega_cliente": "2026-11-27"})
    with app.app_context():
        _po("T-EXP-20", "FOB SEM NF", go_data_efetiva_liberacao_pcp=hoje)
        _po("T-EXP-21", "FOB COM NF", go_data_efetiva_liberacao_pcp=hoje, go_numero_nf="8127")
        for pv in ("T-EXP-20", "T-EXP-21"):
            db.session.add(LogisticaSync(pedido_venda_norm=pv, pedido_venda_logistica=pv, payload_json=payload))
        db.session.commit()
        with app.test_request_context("/gestao-operacao/kanban-expedicao"):
            r = A._kanban_parados_expedicao({}, hoje=hoje)
        parados = {c["pedido_venda"] for col in r["colunas"] for c in col["cards"]}
        assert "T-EXP-20" in parados      # liberado p/ coleta, sem NF: segue parado
        assert "T-EXP-21" not in parados  # NF emitida + data da Logística: expedido de verdade
