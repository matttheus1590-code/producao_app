import pytest

from conftest import entrar, token_do_html

ESCRITA_QUALIDADE = ["/qualidade/novo", "/qualidade/rdim/novo"]
ESCRITA_PD = ["/pd/novo"]


def _logado(app, usuario):
    c = app.test_client()
    assert entrar(c, usuario).status_code == 302
    return c


@pytest.mark.parametrize("usuario,pode", [("t_admin", True), ("t_pcp", True), ("t_lider", True),
                                          ("t_gestao", False), ("t_comercial", False), ("t_pd", False)])
def test_quem_pode_abrir_cadastro_de_qualidade(app, usuario, pode):
    c = _logado(app, usuario)
    codigos = {c.get(u).status_code for u in ESCRITA_QUALIDADE}
    if pode:
        assert codigos <= {200}
    else:
        assert 200 not in codigos


@pytest.mark.parametrize("usuario,pode", [("t_admin", True), ("t_pcp", True), ("t_lider", True), ("t_pd", True),
                                          ("t_gestao", False), ("t_comercial", False)])
def test_quem_pode_abrir_cadastro_de_pd(app, usuario, pode):
    c = _logado(app, usuario)
    codigos = {c.get(u).status_code for u in ESCRITA_PD}
    if pode:
        assert codigos <= {200}
    else:
        assert 200 not in codigos


def test_gestao_continua_vendo_qualidade_e_pd(app):
    c = _logado(app, "t_gestao")
    assert c.get("/qualidade").status_code == 200
    assert c.get("/pd").status_code == 200
