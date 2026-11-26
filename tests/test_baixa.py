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


# ------------------------- autorização de faturamento -------------------------
import autorizacao_faturamento as AF

DADOS_AUT = {
    "liberacao_real": "2026-11-25", "data_emissao": "2026-11-26", "cliente_doc": "DELP", "observacoes": "Doca 2",
    "vol_qtd[]": ["1", "2"], "vol_tipo[]": ["VOL.", "CAIXA"], "vol_c[]": ["500", "460"], "vol_l[]": ["360", "460"],
    "vol_a[]": ["360", "460"], "vol_pb[]": ["3", "8,5"],
    "hora_emissao": "15:30", "coleta_modo": "imediata",
}


def test_formatacao_das_linhas():
    assert AF.fmt_num(3.0) == "3" and AF.fmt_num(8.5) == "8,5" and AF.fmt_num(None) == ""
    assert AF.num_br("1.250,75") == 1250.75 and AF.num_br("3,5") == 3.5 and AF.num_br("abc") is None and AF.num_br("-2") is None
    assert AF.texto_linha({"quantidade": 1, "tipo": "VOL.", "comprimento": 500.0, "largura": 360.0, "altura": 360.0}) == "1X VOL. 500 X 360 X 360"
    assert AF.data_extenso(date(2026, 11, 26)) == "26 de novembro de 2026"


def test_salva_autorizacao_e_gera_documentos(app, pedido_id):
    c = app.test_client()
    entrar(c, "t_pcp")
    r = c.post(f"/pedidos/{pedido_id}/baixa", data=dict(DADOS_AUT, csrf_token=_token(c)), headers={"X-Requested-With": "fetch"})
    assert r.status_code == 200 and r.get_json()["ok"]
    j = c.get(f"/pedidos/{pedido_id}/autorizacao.json").get_json()
    assert j["cliente_doc"] == "DELP" and len(j["volumes"]) == 2 and j["volumes"][1]["peso_bruto"] == 8.5
    pdf = c.get(f"/pedidos/{pedido_id}/autorizacao-faturamento.pdf")
    assert pdf.status_code == 200 and pdf.data.startswith(b"%PDF") and "attachment" in pdf.headers["Content-Disposition"]
    html = c.get(f"/pedidos/{pedido_id}/autorizacao-faturamento").get_data(as_text=True)
    assert "1X VOL. 500 X 360 X 360" in html and "2X CAIXA 460 X 460 X 460" in html and "T-9001" in html


def test_reimprimir_atualiza_a_mesma_autorizacao(app, pedido_id):
    c = app.test_client()
    entrar(c, "t_pcp")
    for obs in ("primeira", "segunda"):
        c.post(f"/pedidos/{pedido_id}/baixa", data=dict(DADOS_AUT, observacoes=obs, csrf_token=_token(c)), headers={"X-Requested-With": "fetch"})
    with app.app_context():
        from models import AutorizacaoFaturamento
        auts = AutorizacaoFaturamento.query.filter_by(pedido_id=pedido_id).all()
        assert len(auts) == 1 and auts[0].observacoes == "segunda" and len(auts[0].volumes) == 2


def test_documento_sem_autorizacao_nao_abre(app, pedido_id):
    c = app.test_client()
    entrar(c, "t_pcp")
    assert c.get(f"/pedidos/{pedido_id}/autorizacao-faturamento.pdf").status_code == 302


def test_gestao_nao_acessa_a_autorizacao(app, pedido_id):
    c = app.test_client()
    entrar(c, "t_gestao")
    assert c.get(f"/pedidos/{pedido_id}/autorizacao.json").status_code in (302, 403)


def test_hora_e_coleta_agendada(app, pedido_id):
    assert AF.ler_hora("9:05") == "09:05" and AF.ler_hora("1430") == "14:30" and AF.ler_hora("25:00") is None and AF.ler_hora("") is None
    c = app.test_client()
    entrar(c, "t_pcp")
    c.post(f"/pedidos/{pedido_id}/baixa", data=dict(DADOS_AUT, coleta_modo="agendada", coleta_data="2026-11-27", coleta_hora="14:00", csrf_token=_token(c)), headers={"X-Requested-With": "fetch"})
    j = c.get(f"/pedidos/{pedido_id}/autorizacao.json").get_json()
    assert j["coleta_modo"] == "agendada" and j["coleta_data"] == "2026-11-27" and j["coleta_hora"] == "14:00" and j["hora_emissao"] == "15:30"
    html = c.get(f"/pedidos/{pedido_id}/autorizacao-faturamento?sem=1").get_data(as_text=True)
    assert "COLETA LIBERADA A PARTIR DE" in html and "27/11/2026 ÀS 14:00" in html and "26 de novembro de 2026 · 15:30" in html
    assert "PESO LIQ" not in html.upper()


def test_coleta_imediata_e_sem_agendamento_sem_data(app, pedido_id):
    c = app.test_client()
    entrar(c, "t_pcp")
    # "a partir de" sem data cai para imediata (não inventa agendamento)
    c.post(f"/pedidos/{pedido_id}/baixa", data=dict(DADOS_AUT, coleta_modo="agendada", coleta_data="", csrf_token=_token(c)), headers={"X-Requested-With": "fetch"})
    html = c.get(f"/pedidos/{pedido_id}/autorizacao-faturamento?sem=1").get_data(as_text=True)
    assert "COLETA LIBERADA IMEDIATA" in html
