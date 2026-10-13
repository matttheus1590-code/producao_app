"""Controle de acesso por papel (role).

Papéis existentes:
  ADMIN  - acesso total, inclusive cadastro de usuários.
  PCP    - cadastra pedidos, programação e produção de qualquer estação.
  LIDER  - só edita a produção da própria estação (Usuario.setor).
  GESTAO - acompanha tudo (painel, KPIs, gargalos, faturamento...) mas não
           edita pedidos nem produção.
  PD     - líder de P&D (pedido do Bruno, 03/09/2026, caso do Gustavo
           Fugita): acesso completo (visualização + edição) só na área de
           P&D, e só visualização em Gestão Produção > Estações. Diferente
           dos outros papéis acima — que sempre puderam VER qualquer tela do
           sistema, só a EDIÇÃO é que era restrita por papel — PD é o
           primeiro papel restrito também por ÁREA: qualquer rota fora do
           allowlist ENDPOINTS_PERMITIDOS_PD abaixo devolve 403 (ver
           `pode_acessar_endpoint`, chamada no `before_request` de app.py).

  COMERCIAL - equipe comercial (pedido do Bruno, 09/10/2026): SÓ leitura e só
           a tela "Consulta Pedido/Planejamento PCP" (consulta de pedido +
           programação do PCP por semana). Mesmo mecanismo do PD: qualquer
           rota fora do allowlist ENDPOINTS_PERMITIDOS_COMERCIAL devolve 403.

  COMERCIAL_GERAL - comercial com visão ampla (pedido do Bruno, 13/10/2026,
           caso da Andrieli): vê o sistema todo como o papel GESTAO (só
           leitura), EXCETO P&D, Qualidade (RNC/RDIM) e Gestão de Custos —
           bloqueados por prefixo de rota (PREFIXOS_BLOQUEADOS_COMERCIAL_GERAL
           abaixo), tanto no menu quanto direto pela URL (403).

Uso:
  @requer_role("ADMIN", "PCP")
  def minha_rota(): ...

  if pode_editar_estacao(current_user, item.estacao):
      ...
"""

from functools import wraps

from flask import abort
from flask_login import current_user, login_required

ROLES = ("ADMIN", "PCP", "LIDER", "GESTAO", "PD", "COMERCIAL", "COMERCIAL_GERAL")

ROLES_LABELS = {
    "ADMIN": "Administrador",
    "PCP": "PCP",
    "LIDER": "Líder de setor",
    "GESTAO": "Gestão",
    "PD": "Líder de P&D",
    "COMERCIAL": "Comercial",
    "COMERCIAL_GERAL": "Comercial (acesso geral)",
}

# Endpoints (nomes de rota Flask, não caminhos de URL) que o papel PD pode
# acessar — P&D inteiro (view + edição, sem diferença aqui: cada rota já
# controla o que faz) e só a VISUALIZAÇÃO de Gestão Produção > Estações
# (estacao_kanban_mover, que move item de coluna, fica de fora de propósito:
# além deste allowlist bloquear com 403, pode_editar_estacao já devolve
# False pra qualquer papel que não seja ADMIN/PCP/LIDER, então o template
# nem mostra o botão de avançar).
ENDPOINTS_PERMITIDOS_PD = {
    "pd_dashboard", "pd_lista", "pd_novo", "pd_editar", "pd_teste_novo", "pd_evento_novo",
    "pd_kanban", "pd_mover_etapa", "pd_cronograma", "pd_testes_lista", "pd_custos", "pd_conhecimento",
    "estacoes_lista", "estacao_kanban",
}

# Endpoints que o papel COMERCIAL pode acessar — só Consulta Pedido/Planejamento
# PCP (as 2 abas, o autocomplete/detalhe do pedido e o PDF do planejamento
# mensal que sai da aba de Planejamento).
ENDPOINTS_PERMITIDOS_COMERCIAL = {
    "consulta_pedido", "consulta_pedido_busca_sugestoes", "consulta_pedido_busca_detalhe",
    "consulta_planejamento_pcp", "relatorio_listagem_geral_semanal_pdf",
}

# Áreas bloqueadas pro COMERCIAL_GERAL: o nome do endpoint (rota Flask) começa
# com um destes prefixos — P&D (pd_), Qualidade/RNC (qualidade_), Inspeções/
# RDIM (rdim_) e Gestão de Custos (custos_).
PREFIXOS_BLOQUEADOS_COMERCIAL_GERAL = ("pd_", "qualidade_", "rdim_", "custos_")

# Rotas sempre liberadas, mesmo pra um papel restrito por área — sem isso o
# PD nem conseguiria fazer login/logout, e os arquivos estáticos (CSS/JS)
# parariam de carregar.
ENDPOINTS_SEMPRE_LIVRES = {"login", "logout", "static"}


def pode_acessar_endpoint(usuario, endpoint):
    """True se `usuario` pode acessar esta rota. Hoje só o papel PD tem
    restrição de VISUALIZAÇÃO por área — todo o resto do sistema sempre
    deixou qualquer usuário autenticado ver qualquer tela (só editar é que
    já era restrito por papel, via `requer_role`/`pode_editar_estacao`)."""
    if not usuario or not usuario.is_authenticated:
        return True  # quem não está logado nem chega aqui — @login_required cuida disso antes
    if usuario.role not in ("PD", "COMERCIAL", "COMERCIAL_GERAL"):
        return True
    if endpoint in ENDPOINTS_SEMPRE_LIVRES:
        return True
    if usuario.role == "COMERCIAL_GERAL":
        return not endpoint.startswith(PREFIXOS_BLOQUEADOS_COMERCIAL_GERAL)
    if usuario.role == "COMERCIAL":
        return endpoint in ENDPOINTS_PERMITIDOS_COMERCIAL
    return endpoint in ENDPOINTS_PERMITIDOS_PD


def ve_pd_qualidade_custos(usuario):
    """False pros papéis que NÃO enxergam P&D, Qualidade e Gestão de Custos
    (COMERCIAL e COMERCIAL_GERAL) — usado nos templates pra esconder menus,
    colunas e links dessas áreas (senão virariam links que dão 403)."""
    return not (usuario and usuario.is_authenticated and usuario.role in ("COMERCIAL", "COMERCIAL_GERAL"))


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
