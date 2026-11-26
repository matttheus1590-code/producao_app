from datetime import date

import pytest

from conftest import entrar, token_do_html
from models import ItemPedido, Pedido, db


@pytest.fixture()
def pedido_id(app):
    with app.app_context():
        p = Pedido(cliente="CLIENTE TESTE BAIXA", pedido_venda="T-9001", data_cliente=date(2026, 12, 20))
        p.itens.append(ItemPedido(descricao_produto="ITEM A", quantidade=1, estacao="PU"))
        p.itens.append(ItemPedido(descricao_produto="ITEM B", quantidade=2, estacao="PU"))
        db.session.add(p)
        db.session.commit()
        return p.id


def _token(c):
    return token_do_html(c.get("/").get_data(as_text=True))


def test_baixa_grava_data_em_todos_os_itens(app, pedido_id):
    c = app.test_client()
    entrar(c, "t_pcp")
    r = c.post(f"/pedidos/{pedido_id}/baixa", data={"liberacao_real": "2026-11-25", "csrf_token": _token(c), "voltar": "/?visao=painel"})
    assert r.status_code == 302 and r.headers["Location"].endswith("/?visao=painel")
    with app.app_context():
        itens = db.session.get(Pedido, pedido_id).itens
        assert {i.liberacao_real for i in itens} == {date(2026, 11, 25)}


def test_baixa_sem_data_e_recusada(app, pedido_id):
    c = app.test_client()
    entrar(c, "t_pcp")
    c.post(f"/pedidos/{pedido_id}/baixa", data={"liberacao_real": "", "csrf_token": _token(c)})
    with app.app_context():
        assert all(i.liberacao_real is None for i in db.session.get(Pedido, pedido_id).itens)


def test_baixa_nao_redireciona_para_fora(app, pedido_id):
    c = app.test_client()
    entrar(c, "t_pcp")
    r = c.post(f"/pedidos/{pedido_id}/baixa", data={"liberacao_real": "2026-11-25", "csrf_token": _token(c), "voltar": "https://mal.example/x"})
    assert "mal.example" not in r.headers["Location"]


@pytest.mark.parametrize("usuario", ["t_gestao", "t_comercial"])
def test_perfis_sem_permissao_nao_dao_baixa(app, pedido_id, usuario):
    c = app.test_client()
    entrar(c, usuario)
    c.post(f"/pedidos/{pedido_id}/baixa", data={"liberacao_real": "2026-11-25", "csrf_token": _token(c) or ""})
    with app.app_context():
        assert all(i.liberacao_real is None for i in db.session.get(Pedido, pedido_id).itens)
