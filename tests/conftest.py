"""Configuração dos testes automáticos.

Rodar:  pip install pytest   e depois   python -m pytest -q
Os testes usam um banco SQLite temporário, nunca o banco de produção.
"""
import os
import sys
import tempfile

_tmp = tempfile.mkdtemp(prefix="producao_testes_")
os.environ["DATABASE_URL"] = "sqlite:///" + os.path.join(_tmp, "teste.db")
os.environ.pop("RENDER", None)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest  # noqa: E402

import app as app_mod  # noqa: E402
import seguranca  # noqa: E402
from models import Usuario, db  # noqa: E402


@pytest.fixture(scope="session")
def app():
    a = app_mod.app
    a.config["TESTING"] = True
    with a.app_context():
        db.create_all()
        for nome, role in (("t_admin", "ADMIN"), ("t_pcp", "PCP"), ("t_lider", "LIDER"),
                           ("t_gestao", "GESTAO"), ("t_pd", "PD"), ("t_comercial", "COMERCIAL")):
            if not Usuario.query.filter_by(username=nome).first():
                u = Usuario(username=nome, nome=nome, role=role, ativo=True)
                u.set_senha("SenhaTeste2026x")
                db.session.add(u)
        db.session.commit()
        seguranca.recalcular_senhas_iniciais(Usuario)
    return a


@pytest.fixture()
def cliente(app):
    return app.test_client()


def token_do_html(html):
    import re
    m = re.search(r'name="csrf_token" value="([^"]+)"', html)
    return m.group(1) if m else None


def entrar(c, usuario, senha="SenhaTeste2026x"):
    t = token_do_html(c.get("/login").get_data(as_text=True))
    return c.post("/login", data={"username": usuario, "senha": senha, "csrf_token": t})
