import re

import pytest

import seguranca
from conftest import entrar, token_do_html


def test_login_traz_token_e_cabecalhos(cliente):
    r = cliente.get("/login")
    h = r.get_data(as_text=True)
    assert 'name="csrf-token"' in h and token_do_html(h)
    assert r.headers["X-Frame-Options"] == "SAMEORIGIN"
    assert r.headers["X-Content-Type-Options"] == "nosniff"


def test_post_sem_token_e_recusado(cliente):
    r = cliente.post("/login", data={"username": "t_admin", "senha": "SenhaTeste2026x"})
    assert r.status_code == 400


def test_post_de_outro_site_e_recusado(cliente):
    r = cliente.post("/login", data={"username": "t_admin", "senha": "x"}, headers={"Origin": "https://mal.example"})
    assert r.status_code == 400


def test_login_valido_funciona(cliente):
    r = entrar(cliente, "t_admin")
    assert r.status_code == 302
    assert cliente.get("/painel").status_code == 200


def test_limite_de_tentativas_de_login(app):
    c = app.test_client()
    t = token_do_html(c.get("/login").get_data(as_text=True))
    codigos = [c.post("/login", data={"username": "vitima_teste", "senha": "x%d" % i, "csrf_token": t}).status_code
               for i in range(12)]
    assert codigos[0] == 200 and 429 in codigos
    # outro usuário continua entrando
    r = c.post("/login", data={"username": "t_pcp", "senha": "SenhaTeste2026x", "csrf_token": t})
    assert r.status_code == 302


def test_senha_fraca_e_recusada():
    from types import SimpleNamespace
    u = SimpleNamespace(username="admin")
    v = seguranca._validar_nova_senha
    assert v("admin123", "admin123", u)          # senha inicial conhecida
    assert v("curta1", "curta1", u)              # curta
    assert v("somenteletrasaqui", "somenteletrasaqui", u)
    assert v("SenhaBoa2026x", "outra", u)        # confirmação diferente
    assert v("SenhaBoa2026x", "SenhaBoa2026x", u) is None


def test_formularios_post_das_paginas_trazem_token(app):
    c = app.test_client()
    entrar(c, "t_admin")
    faltando = []
    for p in ("/painel", "/listagem", "/estacoes", "/qualidade", "/pd", "/pedidos/novo", "/usuarios"):
        r = c.get(p)
        if r.status_code != 200:
            continue
        for m in re.finditer(r"<form\b([^>]*)>(.{0,400})", r.get_data(as_text=True), re.S | re.I):
            if re.search(r'method\s*=\s*["\']?post', m.group(1), re.I) and "csrf_token" not in m.group(2):
                faltando.append(p)
    assert not faltando
