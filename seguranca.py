"""Proteções de segurança do sistema (auditoria de 26/11/2026).

Tudo aqui é independente de bibliotecas pagas ou novas dependências:

1. Limite de tentativas de login (por usuário e por IP), em memória.
2. Proteção CSRF: token na sessão, validado em toda requisição que altera
   dados (POST/PUT/PATCH/DELETE). O token é colocado nos formulários e nas
   chamadas fetch/XHR automaticamente (ver _injetar_csrf), sem precisar mexer
   em cada template.
3. Troca obrigatória das senhas iniciais que estão escritas no código
   (admin, Gustavo, Fabiano e Daniel) — enquanto a senha não for trocada, o
   usuário só consegue abrir a tela de troca de senha.
4. Cabeçalhos de segurança e cookies de sessão mais restritos.

Como ligar: `init_seguranca(app)` em create_app(), depois de db.init_app.
"""

import hmac
import re
import secrets
import threading
import time

from flask import (
    abort,
    current_app,
    flash,
    redirect,
    render_template,
    request,
    session,
    url_for,
)
from flask_login import current_user, login_required, logout_user
from werkzeug.security import check_password_hash, generate_password_hash

# ---------------------------------------------------------------------------
# 1. Limite de tentativas de login
# ---------------------------------------------------------------------------
JANELA_SEG = 15 * 60          # janela em que as falhas contam
MAX_FALHAS_USUARIO = 8        # falhas por usuário dentro da janela
MAX_FALHAS_IP = 30            # falhas por IP dentro da janela
BLOQUEIO_SEG = 10 * 60        # tempo de bloqueio ao estourar o limite

_trava = threading.Lock()
_falhas = {}                  # chave -> lista de timestamps de falha
_bloqueios = {}               # chave -> timestamp em que o bloqueio termina


def ip_do_cliente():
    """IP do cliente. Atrás do proxy do Render vem em X-Forwarded-For."""
    xff = request.headers.get("X-Forwarded-For", "")
    if xff:
        return xff.split(",")[0].strip()
    return request.remote_addr or "?"


def _chaves(username):
    return ("u:" + (username or "").strip().lower(), "i:" + ip_do_cliente())


def login_bloqueado(username):
    """Retorna os segundos restantes de bloqueio (0 se pode tentar)."""
    agora = time.time()
    restante = 0
    with _trava:
        for chave in _chaves(username):
            fim = _bloqueios.get(chave, 0)
            if fim > agora:
                restante = max(restante, int(fim - agora) + 1)
            elif chave in _bloqueios:
                del _bloqueios[chave]
    return restante


def registrar_falha_login(username):
    agora = time.time()
    chave_u, chave_i = _chaves(username)
    with _trava:
        for chave, limite in ((chave_u, MAX_FALHAS_USUARIO), (chave_i, MAX_FALHAS_IP)):
            lista = [t for t in _falhas.get(chave, []) if agora - t < JANELA_SEG]
            lista.append(agora)
            _falhas[chave] = lista
            if len(lista) >= limite:
                _bloqueios[chave] = agora + BLOQUEIO_SEG
                _falhas[chave] = []
        # limpeza simples pra o dicionário não crescer sem fim
        if len(_falhas) > 5000:
            for chave in [c for c, v in _falhas.items() if not v or agora - v[-1] > JANELA_SEG]:
                _falhas.pop(chave, None)
    current_app.logger.warning(
        "Falha de login: usuario=%r ip=%s", (username or "")[:80], ip_do_cliente()
    )


def limpar_falhas_login(username):
    with _trava:
        for chave in _chaves(username):
            _falhas.pop(chave, None)
            _bloqueios.pop(chave, None)


# ---------------------------------------------------------------------------
# 2. CSRF
# ---------------------------------------------------------------------------
_METODOS_SEGUROS = {"GET", "HEAD", "OPTIONS", "TRACE"}
# Endpoints chamados por automações externas sem sessão. Hoje não há nenhum
# POST assim (a rota /api/relatorio-diario é GET com token próprio).
ENDPOINTS_SEM_CSRF = set()


def token_csrf():
    token = session.get("_csrf")
    if not token:
        token = secrets.token_urlsafe(32)
        session["_csrf"] = token
    return token


def _validar_csrf():
    if request.method in _METODOS_SEGUROS:
        return
    if request.endpoint in ENDPOINTS_SEM_CSRF:
        return
    esperado = session.get("_csrf")
    enviado = request.form.get("csrf_token") or request.headers.get("X-CSRF-Token") or ""
    if not esperado or not hmac.compare_digest(str(esperado), str(enviado)):
        current_app.logger.warning(
            "CSRF recusado: %s %s ip=%s", request.method, request.path, ip_do_cliente()
        )
        abort(400, description="csrf")


_RX_FORM = re.compile(r"<form\b([^>]*)>", re.IGNORECASE)
_RX_METODO_POST = re.compile(r"""\bmethod\s*=\s*["']?post\b""", re.IGNORECASE)
_RX_ACAO_EXTERNA = re.compile(r"""\baction\s*=\s*["']?https?://""", re.IGNORECASE)
_RX_HEAD = re.compile(r"<head\b[^>]*>", re.IGNORECASE)

_SCRIPT_CSRF = (
    "<script>(function(){var m=document.querySelector('meta[name=csrf-token]');"
    "var t=m&&m.getAttribute('content');if(!t)return;"
    "var of=window.fetch;if(of){window.fetch=function(i,o){o=o||{};"
    "var me=String(o.method||(i&&i.method)||'GET').toUpperCase();"
    "if(me!=='GET'&&me!=='HEAD'){var h=new Headers(o.headers||(i&&i.headers)||{});"
    "h.set('X-CSRF-Token',t);o.headers=h;}return of.call(this,i,o);};}"
    "var xo=XMLHttpRequest.prototype.open,xs=XMLHttpRequest.prototype.send;"
    "XMLHttpRequest.prototype.open=function(me){this._csrfMe=String(me).toUpperCase();"
    "return xo.apply(this,arguments);};"
    "XMLHttpRequest.prototype.send=function(){if(this._csrfMe&&this._csrfMe!=='GET'"
    "&&this._csrfMe!=='HEAD'){try{this.setRequestHeader('X-CSRF-Token',t);}catch(e){}}"
    "return xs.apply(this,arguments);};"
    "document.addEventListener('submit',function(e){var f=e.target;"
    "if(f&&f.method&&f.method.toLowerCase()==='post'&&!f.querySelector('input[name=csrf_token]')"
    "&&!/^https?:/i.test(f.getAttribute('action')||'')){var i=document.createElement('input');"
    "i.type='hidden';i.name='csrf_token';i.value=t;f.appendChild(i);}},true);})();</script>"
)


def _injetar_csrf(resposta):
    """Coloca o token em toda página HTML: meta + script (fetch/XHR/forms
    criados depois) e um campo oculto em cada <form method=post>."""
    if resposta.direct_passthrough or resposta.is_streamed or resposta.mimetype != "text/html":
        return resposta
    try:
        html = resposta.get_data(as_text=True)
    except Exception:  # noqa: BLE001 - resposta em streaming etc.
        return resposta
    token = token_csrf()
    campo = f'<input type="hidden" name="csrf_token" value="{token}">'

    def _form(m):
        attrs = m.group(1)
        if _RX_METODO_POST.search(attrs) and not _RX_ACAO_EXTERNA.search(attrs):
            return m.group(0) + campo
        return m.group(0)

    html = _RX_FORM.sub(_form, html)
    meta = f'<meta name="csrf-token" content="{token}">' + _SCRIPT_CSRF
    if _RX_HEAD.search(html):
        html = _RX_HEAD.sub(lambda m: m.group(0) + meta, html, count=1)
    resposta.set_data(html)
    return resposta


# ---------------------------------------------------------------------------
# 3. Senhas iniciais que estão escritas no código
# ---------------------------------------------------------------------------
SENHAS_INICIAIS = {
    "admin": "admin123",
    "gustavo.fugita": "PD@Fugita2026!",
    "fabiano.pcp": "PCP@Fabiano2026!",
    "daniel.pcp": "PCP@Daniel2026!",
}
_SENHAS_CONHECIDAS = set(SENHAS_INICIAIS.values())
_com_senha_inicial = set()   # usernames que ainda usam a senha escrita no código
MIN_SENHA = 10


def recalcular_senhas_iniciais(Usuario):
    """Descobre quais usuários ainda usam a senha inicial (só 4 comparações)."""
    _com_senha_inicial.clear()
    for username, senha in SENHAS_INICIAIS.items():
        u = Usuario.query.filter_by(username=username).first()
        if u is not None and check_password_hash(u.senha_hash, senha):
            _com_senha_inicial.add(username)
    if _com_senha_inicial:
        current_app.logger.warning(
            "Usuários ainda com a senha inicial do código (troca será exigida): %s",
            ", ".join(sorted(_com_senha_inicial)),
        )


def precisa_trocar_senha(usuario):
    return bool(usuario and getattr(usuario, "username", None) in _com_senha_inicial)


def _validar_nova_senha(nova, confirmacao, usuario):
    if nova != confirmacao:
        return "A confirmação não é igual à nova senha."
    if len(nova) < MIN_SENHA:
        return f"A nova senha precisa ter pelo menos {MIN_SENHA} caracteres."
    if nova in _SENHAS_CONHECIDAS:
        return "Escolha uma senha diferente das senhas iniciais do sistema."
    if nova.strip().lower() == (usuario.username or "").lower():
        return "A senha não pode ser igual ao nome de usuário."
    if nova.isdigit() or nova.isalpha():
        return "Use letras e números (ou símbolos) na senha."
    return None


# ---------------------------------------------------------------------------
# 4. Ligação com o app
# ---------------------------------------------------------------------------
def init_seguranca(app, Usuario, db):
    import os

    # Cookies de sessão mais restritos. O Secure só é ligado no Render
    # (https); em teste local sem https o login deixaria de funcionar.
    app.config.setdefault("SESSION_COOKIE_HTTPONLY", True)
    app.config.setdefault("SESSION_COOKIE_SAMESITE", "Lax")
    app.config.setdefault("REMEMBER_COOKIE_HTTPONLY", True)
    app.config.setdefault("REMEMBER_COOKIE_SAMESITE", "Lax")
    if os.environ.get("RENDER"):
        app.config.setdefault("SESSION_COOKIE_SECURE", True)
        app.config.setdefault("REMEMBER_COOKIE_SECURE", True)

    app.jinja_env.globals["csrf_token"] = token_csrf

    @app.before_request
    def _seguranca_antes():
        if request.endpoint in (None, "static"):
            return None
        _validar_csrf()
        if (
            current_user.is_authenticated
            and precisa_trocar_senha(current_user)
            and request.endpoint not in ("trocar_senha", "logout", "health")
        ):
            return redirect(url_for("trocar_senha"))
        return None

    @app.after_request
    def _seguranca_depois(resposta):
        resposta.headers.setdefault("X-Content-Type-Options", "nosniff")
        resposta.headers.setdefault("X-Frame-Options", "SAMEORIGIN")
        resposta.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
        if request.is_secure or request.headers.get("X-Forwarded-Proto", "") == "https":
            resposta.headers.setdefault("Strict-Transport-Security", "max-age=31536000")
        return _injetar_csrf(resposta)

    @app.errorhandler(400)
    def _erro_400(e):
        if getattr(e, "description", "") == "csrf":
            return (
                "<!doctype html><meta charset='utf-8'><title>Sessão expirada</title>"
                "<body style='font-family:sans-serif;max-width:32rem;margin:4rem auto'>"
                "<h3>Sua página ficou desatualizada</h3>"
                "<p>Por segurança, o envio foi recusado. Volte, recarregue a página "
                "(F5) e tente de novo. Se continuar, entre no sistema outra vez.</p>"
                "<p><a href='/'>Voltar ao início</a></p></body>",
                400,
            )
        return e

    @app.route("/trocar-senha", methods=["GET", "POST"], endpoint="trocar_senha")
    @login_required
    def trocar_senha():
        obrigatoria = precisa_trocar_senha(current_user)
        if request.method == "POST":
            atual = request.form.get("senha_atual", "")
            nova = request.form.get("nova_senha", "")
            conf = request.form.get("confirmacao", "")
            if not current_user.checar_senha(atual):
                flash("A senha atual está incorreta.", "danger")
            else:
                erro = _validar_nova_senha(nova, conf, current_user)
                if erro:
                    flash(erro, "danger")
                else:
                    current_user.set_senha(nova)
                    db.session.commit()
                    _com_senha_inicial.discard(current_user.username)
                    current_app.logger.info("Senha trocada: usuario=%s", current_user.username)
                    flash("Senha trocada com sucesso.", "success")
                    return redirect("/")
        return render_template("trocar_senha.html", obrigatoria=obrigatoria)

    with app.app_context():
        try:
            recalcular_senhas_iniciais(Usuario)
        except Exception as exc:  # noqa: BLE001 - tabela ainda não existe etc.
            current_app.logger.warning("Checagem de senhas iniciais adiada: %s", exc)
