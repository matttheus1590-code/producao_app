"""Controle de acesso por papel (role).

Papéis existentes:
  ADMIN  - acesso total, inclusive cadastro de usuários.
  PCP    - cadastra pedidos, programação e produção de qualquer estação.
  LIDER  - só edita a produção da própria estação (Usuario.setor).
  GESTAO - acompanha tudo (painel, KPIs, gargalos, faturamento...) mas não
           edita pedidos nem produção.

Uso:
  @requer_role("ADMIN", "PCP")
  def minha_rota(): ...

  if pode_editar_estacao(current_user, item.estacao):
      ...
"""

from functools import wraps

from flask import abort
from flask_login import current_user, login_required

ROLES = ("ADMIN", "PCP", "LIDER", "GESTAO")

ROLES_LABELS = {
    "ADMIN": "Administrador",
    "PCP": "PCP",
    "LIDER": "Líder de setor",
    "GESTAO": "Gestão",
}


def requer_role(*roles_permitidos):
    """Decorator: exige login (como @login_required) e, além disso, que o
    usuário tenha um dos papéis informados. Quem não tiver o papel recebe 403."""

    def decorator(view):
        @wraps(view)
        @login_required
        def wrapped(*args, **kwargs):
            if current_user.role not in roles_permitidos:
                abort(403)
            return view(*args, **kwargs)

        return wrapped

    return decorator


def pode_editar_estacao(usuario, estacao):
    """ADMIN/PCP podem editar a produção de qualquer estação; LIDER só a sua."""
    if not usuario or not usuario.is_authenticated:
        return False
    if usuario.role in ("ADMIN", "PCP"):
        return True
    if usuario.role == "LIDER":
        return bool(estacao) and usuario.setor == estacao
    return False
