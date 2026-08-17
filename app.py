import os
from datetime import date, datetime, timedelta

from flask import Flask, flash, jsonify, redirect, render_template, request, url_for
from flask_login import current_user, login_required, login_user, logout_user
from sqlalchemy import and_, func, inspect, or_, text
from sqlalchemy.orm import selectinload

from extensions import db, login_manager
from models import (
    ESTACOES,
    FRETE_OPCOES,
    PRAZO_ALERTA_DIAS,
    PRIORIDADE_CORES,
    PRIORIDADE_OPCOES,
    REGIAO_POR_UF,
    REGIOES_OPCOES,
    SEMAFORO_CORES,
    SEMAFORO_LABELS,
    STATUS_CHAO_CORES,
    STATUS_CHAO_LABELS,
    STATUS_CHAO_OPCOES,
    STATUS_CORES,
    STATUS_OPCOES,
    Estacao,
    HistoricoAlteracao,
    ItemPedido,
    Pedido,
    Usuario,
)
from permissoes import ROLES, ROLES_LABELS, pode_editar_estacao, requer_role

MESES_PT = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PAGE_SIZE = 25

# Campos "importantes" que geram uma linha no histórico de alterações quando
# mudam de valor (não registramos tudo, só o que realmente importa acompanhar).
CAMPOS_HISTORICO_PEDIDO = ["cliente", "prioridade", "vendedor", "frete"]
CAMPOS_HISTORICO_ITEM = [
    "estacao",
    "status_producao",
    "inicio_producao",
    "inicio_inspecao",
    "termino_inspecao",
    "liberacao_faturamento",
    "liberacao_prevista",
    "rnc",
]


def _resolve_database_uri():
    """Usa DATABASE_URL (Postgres do Render) quando existir; senão, SQLite local."""
    url = os.environ.get("DATABASE_URL")
    if url:
        # Render/Heroku às vezes fornecem "postgres://", mas o SQLAlchemy 2.x exige "postgresql://"
        if url.startswith("postgres://"):
            url = url.replace("postgres://", "postgresql://", 1)
        return url
    return "sqlite:///" + os.path.join(BASE_DIR, "instance", "pedidos.db")


def create_app():
    app = Flask(__name__)
    app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "troque-esta-chave-em-producao")
    app.config["SQLALCHEMY_DATABASE_URI"] = _resolve_database_uri()
    app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

    db.init_app(app)
    login_manager.init_app(app)

    @login_manager.user_loader
    def load_user(user_id):
        return db.session.get(Usuario, int(user_id))

    with app.app_context():
        os.makedirs(os.path.join(BASE_DIR, "instance"), exist_ok=True)
        db.create_all()
        _migrar_itens_pedido(app)
        _migrar_producao_para_itens(app)
        _migrar_usuarios_role(app)
        _migrar_estacoes_tabela(app)
        _seed_inicial(app)

    @app.context_processor
    def inject_globals():
        estacoes_ativas = [e.nome for e in Estacao.query.filter_by(ativo=True).order_by(Estacao.ordem_exibicao).all()]
        return dict(
            ESTACOES=estacoes_ativas or ESTACOES,
            STATUS_OPCOES=STATUS_OPCOES,
            PRIORIDADE_OPCOES=PRIORIDADE_OPCOES,
            FRETE_OPCOES=FRETE_OPCOES,
            STATUS_CORES=STATUS_CORES,
            PRIORIDADE_CORES=PRIORIDADE_CORES,
            ROLES=ROLES,
            ROLES_LABELS=ROLES_LABELS,
            REGIOES_OPCOES=REGIOES_OPCOES,
            SEMAFORO_CORES=SEMAFORO_CORES,
            SEMAFORO_LABELS=SEMAFORO_LABELS,
            PRAZO_ALERTA_DIAS=PRAZO_ALERTA_DIAS,
            STATUS_CHAO_OPCOES=STATUS_CHAO_OPCOES,
            STATUS_CHAO_LABELS=STATUS_CHAO_LABELS,
            STATUS_CHAO_CORES=STATUS_CHAO_CORES,
        )

    register_routes(app)
    return app


def _migrar_itens_pedido(app):
    """Migra bancos criados antes de pedidos aceitarem vários produtos.

    Antes, cada Pedido tinha um único produto (colunas descricao_produto,
    quantidade e custo_unitario direto na tabela "pedidos"). Agora esses
    dados moram na tabela "itens_pedido" (um pedido -> vários itens). Esta
    função roda sozinha a cada start do site e só faz alguma coisa se
    detectar o formato antigo — não apaga nenhum pedido, só reorganiza os
    dados de produto para o novo formato.
    """
    inspector = inspect(db.engine)
    if "pedidos" not in inspector.get_table_names():
        return  # banco novo — db.create_all() já cuidou de tudo

    colunas_pedidos = {c["name"] for c in inspector.get_columns("pedidos")}
    formato_antigo = {"descricao_produto", "quantidade", "custo_unitario"}.issubset(colunas_pedidos)
    if not formato_antigo:
        return  # já está no formato novo

    with db.engine.begin() as conn:
        conn.execute(
            text(
                "INSERT INTO itens_pedido (pedido_id, descricao_produto, quantidade, custo_unitario) "
                "SELECT id, descricao_produto, quantidade, custo_unitario FROM pedidos"
            )
        )
        conn.execute(text("ALTER TABLE pedidos DROP COLUMN descricao_produto"))
        conn.execute(text("ALTER TABLE pedidos DROP COLUMN quantidade"))
        conn.execute(text("ALTER TABLE pedidos DROP COLUMN custo_unitario"))
    app.logger.info("Migração automática: produtos dos pedidos movidos para itens_pedido com sucesso.")


_COLUNAS_PRODUCAO = [
    "estacao",
    "status_producao",
    "status_manual",
    "inicio_producao",
    "inicio_inspecao",
    "termino_inspecao",
    "liberacao_faturamento",
    "liberacao_prevista",
    "rnc",
]

_TIPOS_COLUNAS_PRODUCAO = {
    "estacao": "VARCHAR(40)",
    "status_producao": "VARCHAR(20)",
    "status_manual": "BOOLEAN",
    "inicio_producao": "DATE",
    "inicio_inspecao": "DATE",
    "termino_inspecao": "DATE",
    "liberacao_faturamento": "DATE",
    "liberacao_prevista": "DATE",
    "rnc": "VARCHAR(120)",
}


def _migrar_producao_para_itens(app):
    """Migra bancos criados antes da produção (estação, status, datas, RNC) virar
    algo por ITEM em vez de por pedido inteiro.

    Como cada pedido tinha exatamente 1 item na época em que essas colunas ainda
    ficavam em "pedidos", a migração é uma cópia direta: cada item recebe os
    dados de produção do pedido que o originou. Não apaga nenhum dado.
    """
    inspector = inspect(db.engine)
    if "pedidos" not in inspector.get_table_names():
        return

    colunas_pedidos = {c["name"] for c in inspector.get_columns("pedidos")}
    formato_antigo = set(_COLUNAS_PRODUCAO).issubset(colunas_pedidos)
    if not formato_antigo:
        return  # já está no formato novo

    # db.create_all() não altera tabelas já existentes — "itens_pedido" pode já
    # existir (criada pela migração anterior) sem essas colunas novas ainda.
    colunas_itens = {c["name"] for c in inspector.get_columns("itens_pedido")}

    with db.engine.begin() as conn:
        for coluna, tipo in _TIPOS_COLUNAS_PRODUCAO.items():
            if coluna not in colunas_itens:
                conn.execute(text(f"ALTER TABLE itens_pedido ADD COLUMN {coluna} {tipo}"))
        for coluna in _COLUNAS_PRODUCAO:
            conn.execute(
                text(
                    f"UPDATE itens_pedido SET {coluna} = ("
                    f"SELECT p.{coluna} FROM pedidos p WHERE p.id = itens_pedido.pedido_id"
                    f")"
                )
            )
        for coluna in _COLUNAS_PRODUCAO:
            conn.execute(text(f"ALTER TABLE pedidos DROP COLUMN {coluna}"))
    app.logger.info("Migração automática: dados de produção movidos dos pedidos para os itens com sucesso.")


def _migrar_usuarios_role(app):
    """Adiciona os campos de papel de acesso (role/setor/ativo) em usuários
    criados antes do Dashboard Gerencial de PCP existir.

    Roda sozinha a cada início do site e só faz alguma coisa se detectar que
    essas colunas ainda não existem — não apaga nem altera nenhum usuário além
    de garantir que o "admin" original continue com acesso total (ADMIN) e que
    todo o resto continue podendo fazer o que já fazia hoje (PCP).
    """
    inspector = inspect(db.engine)
    if "usuarios" not in inspector.get_table_names():
        return  # banco novo — db.create_all() já cuidou de tudo

    colunas = {c["name"] for c in inspector.get_columns("usuarios")}
    faltando = [c for c in ("role", "setor", "ativo") if c not in colunas]
    if not faltando:
        return  # já está no formato novo

    with db.engine.begin() as conn:
        if "role" not in colunas:
            conn.execute(text("ALTER TABLE usuarios ADD COLUMN role VARCHAR(20)"))
        if "setor" not in colunas:
            conn.execute(text("ALTER TABLE usuarios ADD COLUMN setor VARCHAR(40)"))
        if "ativo" not in colunas:
            conn.execute(text("ALTER TABLE usuarios ADD COLUMN ativo BOOLEAN"))

        conn.execute(text("UPDATE usuarios SET role = 'PCP' WHERE role IS NULL"))
        conn.execute(text("UPDATE usuarios SET role = 'ADMIN' WHERE username = 'admin'"))
        conn.execute(text("UPDATE usuarios SET ativo = TRUE WHERE ativo IS NULL"))
    app.logger.info("Migração automática: usuários existentes receberam papel de acesso (role).")


def _migrar_estacoes_tabela(app):
    """Cadastra as estações como registros de verdade na tabela "estacoes", com
    os mesmos nomes e a mesma ordem da lista fixa ESTACOES (models.py).

    Como itens continuam guardando a estação pelo nome (texto), nenhum pedido
    existente precisa mudar — só passa a existir um registro correspondente
    pra cada nome, que os cadastros/tela de Estações usam."""
    if Estacao.query.count() > 0:
        return  # já foi semeado (ou o usuário já está gerenciando pelo cadastro)
    for i, nome in enumerate(ESTACOES):
        db.session.add(Estacao(nome=nome, ordem_exibicao=i, ativo=True))
    db.session.commit()
    app.logger.info("Migração automática: %d estações cadastradas como tabela.", len(ESTACOES))


def _seed_inicial(app):
    """Cria o usuário admin padrão e importa a planilha na primeira execução."""
    if Usuario.query.count() == 0:
        admin = Usuario(nome="Administrador", username="admin", role="ADMIN")
        admin.set_senha("admin123")
        db.session.add(admin)
        db.session.commit()
        app.logger.info("Usuário admin criado (login: admin / senha: admin123 — troque depois!)")

    if Pedido.query.count() == 0:
        xlsx_path = os.path.join(BASE_DIR, "data", "controle_producao_base.xlsx")
        if os.path.exists(xlsx_path):
            from seed import importar_planilha

            total = importar_planilha(xlsx_path)
            app.logger.info(f"{total} pedidos importados da planilha original.")


def _parse_data_form(valor):
    if not valor:
        return None
    try:
        return datetime.strptime(valor, "%Y-%m-%d").date()
    except ValueError:
        return None


def _parse_float_form(valor, default=0.0):
    if valor in (None, ""):
        return default
    try:
        return float(str(valor).replace(",", "."))
    except ValueError:
        return default


def _registrar_alteracoes(entidade_tipo, entidade_id, pedido_id, antes, depois, campos):
    """Compara os valores de "antes"/"depois" de uma lista de campos e grava uma
    linha no histórico para cada campo que realmente mudou de valor. Gravado na
    mesma sessão do banco que a alteração em si — só vira permanente quando o
    resto da rota der commit()."""
    for campo in campos:
        v_antes = antes.get(campo)
        v_depois = depois.get(campo)
        if v_antes == v_depois:
            continue
        db.session.add(
            HistoricoAlteracao(
                entidade_tipo=entidade_tipo,
                entidade_id=entidade_id,
                pedido_id=pedido_id,
                usuario_id=current_user.id if current_user.is_authenticated else None,
                usuario_nome=current_user.nome if current_user.is_authenticated else None,
                campo=campo,
                valor_anterior=None if v_antes is None else str(v_antes),
                valor_novo=None if v_depois is None else str(v_depois),
            )
        )


def _predicado_status(status):
    """Traduz o status "calculado" do pedido (Pedido.status_producao) para uma
    condição de SQL equivalente, para poder filtrar/contar direto no banco em
    vez de carregar todo mundo em Python (como a tela inicial fazia antes).

    Precisa reproduzir exatamente as mesmas regras da property Pedido.status_producao:
      - algum item "EM TRATATIVA"   -> EM TRATATIVA (tem prioridade sobre o resto)
      - todos os itens "FINALIZADO" -> FINALIZADO
      - todos os itens "PENDENTE" (ou pedido sem item) -> PENDENTE
      - qualquer outra mistura      -> ANDAMENTO
    """
    tem_itens = Pedido.itens.any()

    if status == "EM TRATATIVA":
        return Pedido.itens.any(ItemPedido.status_producao == "EM TRATATIVA")

    if status == "FINALIZADO":
        return and_(tem_itens, ~Pedido.itens.any(ItemPedido.status_producao != "FINALIZADO"))

    if status == "PENDENTE":
        return or_(~tem_itens, and_(tem_itens, ~Pedido.itens.any(ItemPedido.status_producao != "PENDENTE")))

    if status == "ANDAMENTO":
        return and_(
            tem_itens,
            ~Pedido.itens.any(ItemPedido.status_producao == "EM TRATATIVA"),
            Pedido.itens.any(ItemPedido.status_producao != "FINALIZADO"),
            Pedido.itens.any(ItemPedido.status_producao != "PENDENTE"),
        )

    return None


def _predicado_atrasado():
    """Pedido "atrasado": tem pelo menos um item com liberação prevista já
    vencida e que ainda não foi finalizado — mesma regra do semáforo vermelho."""
    return Pedido.itens.any(
        and_(ItemPedido.liberacao_prevista < date.today(), ItemPedido.status_producao != "FINALIZADO")
    )


def _calcular_resumo():
    """Contagens gerais (cards do topo) calculadas direto no banco — antes esta
    função carregava TODOS os pedidos em Python a cada acesso à tela inicial."""
    total = Pedido.query.count()
    pendente = Pedido.query.filter(_predicado_status("PENDENTE")).count()
    em_tratativa = Pedido.query.filter(_predicado_status("EM TRATATIVA")).count()
    andamento = Pedido.query.filter(_predicado_status("ANDAMENTO")).count()
    finalizado = Pedido.query.filter(_predicado_status("FINALIZADO")).count()
    valor_total = db.session.query(func.sum(ItemPedido.quantidade * ItemPedido.custo_unitario)).scalar() or 0.0
    return {
        "total": total,
        "pendente": pendente,
        "em_tratativa": em_tratativa,
        "andamento": andamento,
        "finalizado": finalizado,
        "valor_total": round(valor_total, 2),
    }


def _predicado_vencendo():
    """Pedido "vencendo": tem item com liberação prevista nos próximos
    PRAZO_ALERTA_DIAS dias (e ainda não atrasado nenhum item) — mesma regra do
    semáforo amarelo."""
    limite = date.today() + timedelta(days=PRAZO_ALERTA_DIAS)
    tem_item_vencendo = Pedido.itens.any(
        and_(
            ItemPedido.liberacao_prevista.isnot(None),
            ItemPedido.liberacao_prevista >= date.today(),
            ItemPedido.liberacao_prevista <= limite,
            ItemPedido.status_producao != "FINALIZADO",
        )
    )
    return and_(tem_item_vencendo, ~_predicado_atrasado())


def _faturamento_mes(ano, mes):
    """Soma o valor (quantidade × custo) dos itens com liberação PREVISTA e com
    liberação REALIZADA (faturamento) dentro de um mês — usado no previsto × realizado."""
    inicio = date(ano, mes, 1)
    fim = date(ano + 1, 1, 1) if mes == 12 else date(ano, mes + 1, 1)
    valor_expr = ItemPedido.quantidade * ItemPedido.custo_unitario

    previsto = (
        db.session.query(func.sum(valor_expr))
        .filter(ItemPedido.liberacao_prevista >= inicio, ItemPedido.liberacao_prevista < fim)
        .scalar()
        or 0.0
    )
    realizado = (
        db.session.query(func.sum(valor_expr))
        .filter(ItemPedido.liberacao_faturamento >= inicio, ItemPedido.liberacao_faturamento < fim)
        .scalar()
        or 0.0
    )
    return round(previsto, 2), round(realizado, 2)


def _faturamento_tendencia(meses=6):
    """Previsto × realizado dos últimos N meses (incluindo o atual), para o gráfico de tendência."""
    hoje = date.today()
    pontos = []
    ano, mes = hoje.year, hoje.month
    for _ in range(meses):
        pontos.append((ano, mes))
        mes -= 1
        if mes == 0:
            mes, ano = 12, ano - 1
    pontos.reverse()

    resultado = []
    for ano, mes in pontos:
        previsto, realizado = _faturamento_mes(ano, mes)
        resultado.append({"mes": f"{MESES_PT[mes - 1]}/{ano}", "previsto": previsto, "realizado": realizado})
    return resultado


def _lead_time_medio_dias():
    """Média de dias entre início de produção e término de inspeção/embalagem,
    entre os itens que já têm as duas datas preenchidas.

    Cálculo feito em Python (não em SQL) de propósito: diferença de datas tem
    sintaxe diferente entre SQLite e Postgres, e a quantidade de itens aqui é
    pequena o bastante pra isso não pesar."""
    itens = ItemPedido.query.filter(
        ItemPedido.inicio_producao.isnot(None), ItemPedido.termino_inspecao.isnot(None)
    ).all()
    if not itens:
        return None
    dias = [(i.termino_inspecao - i.inicio_producao).days for i in itens]
    return round(sum(dias) / len(dias), 1)


def _otd_percentual():
    """OTD (On-Time Delivery): % dos itens finalizados cujo término de inspeção
    aconteceu até a data de liberação prevista. Primeira proposta de cálculo —
    ajustável se o critério de "no prazo" precisar ser outro."""
    itens = ItemPedido.query.filter(
        ItemPedido.status_producao == "FINALIZADO",
        ItemPedido.liberacao_prevista.isnot(None),
        ItemPedido.termino_inspecao.isnot(None),
    ).all()
    if not itens:
        return None
    no_prazo = sum(1 for i in itens if i.termino_inspecao <= i.liberacao_prevista)
    return round(100 * no_prazo / len(itens), 1)


def _backlog_por_estacao():
    """Quantidade de itens não finalizados por estação — usado no gráfico do
    Painel e é a base do ranking de Gargalos (fase 8)."""
    linhas = (
        db.session.query(ItemPedido.estacao, func.count(ItemPedido.id))
        .filter(ItemPedido.status_producao != "FINALIZADO", ItemPedido.estacao.isnot(None))
        .group_by(ItemPedido.estacao)
        .order_by(func.count(ItemPedido.id).desc())
        .all()
    )
    return [{"estacao": e or "—", "quantidade": q} for e, q in linhas]


def _construir_timeline(pedido):
    """Monta uma lista de eventos (data + descrição) a partir das datas já
    preenchidas no pedido e em cada item, para exibir como linha do tempo."""
    eventos = []
    if pedido.data_cliente:
        eventos.append({"data": pedido.data_cliente, "titulo": "Data do cliente", "item": None})
    if pedido.data_inclusao_pedido:
        eventos.append({"data": pedido.data_inclusao_pedido, "titulo": "Pedido incluído no sistema", "item": None})

    for item in pedido.itens:
        rotulo = item.descricao_produto
        if item.inicio_producao:
            eventos.append({"data": item.inicio_producao, "titulo": f"Início de produção — {rotulo}", "item": item})
        if item.inicio_inspecao:
            eventos.append({"data": item.inicio_inspecao, "titulo": f"Início de inspeção/embalagem — {rotulo}", "item": item})
        if item.termino_inspecao:
            eventos.append({"data": item.termino_inspecao, "titulo": f"Término de inspeção/embalagem — {rotulo}", "item": item})
        if item.liberacao_faturamento:
            eventos.append({"data": item.liberacao_faturamento, "titulo": f"Liberado para faturamento — {rotulo}", "item": item})

    eventos.sort(key=lambda e: e["data"])
    return eventos


def register_routes(app):
    @app.route("/login", methods=["GET", "POST"])
    def login():
        if current_user.is_authenticated:
            return redirect(url_for("dashboard"))

        if request.method == "POST":
            username = request.form.get("username", "").strip()
            senha = request.form.get("senha", "")
            usuario = Usuario.query.filter_by(username=username).first()
            if usuario and usuario.checar_senha(senha):
                login_user(usuario)
                destino = request.args.get("next") or url_for("dashboard")
                return redirect(destino)
            flash("Usuário ou senha inválidos.", "danger")

        return render_template("login.html")

    @app.route("/logout")
    @login_required
    def logout():
        logout_user()
        return redirect(url_for("login"))

    @app.route("/painel")
    @login_required
    def painel():
        resumo = _calcular_resumo()
        atrasados = Pedido.query.filter(_predicado_atrasado()).count()
        vencendo = Pedido.query.filter(_predicado_vencendo()).count()
        backlog = resumo["total"] - resumo["finalizado"]

        hoje = date.today()
        previsto_mes, realizado_mes = _faturamento_mes(hoje.year, hoje.month)

        pedidos_atrasados = (
            Pedido.query.options(selectinload(Pedido.itens))
            .filter(_predicado_atrasado())
            .order_by(Pedido.data_inclusao_pedido.desc().nullslast())
            .limit(10)
            .all()
        )

        return render_template(
            "painel.html",
            resumo=resumo,
            atrasados=atrasados,
            vencendo=vencendo,
            backlog=backlog,
            previsto_mes=previsto_mes,
            realizado_mes=realizado_mes,
            lead_time_medio=_lead_time_medio_dias(),
            otd=_otd_percentual(),
            tendencia_faturamento=_faturamento_tendencia(),
            backlog_estacao=_backlog_por_estacao(),
            pedidos_atrasados=pedidos_atrasados,
        )

    @app.route("/")
    @login_required
    def dashboard():
        query = Pedido.query.options(selectinload(Pedido.itens))

        cliente = request.args.get("cliente", "").strip()
        status = request.args.get("status", "").strip()
        estacao = request.args.get("estacao", "").strip()
        vendedor = request.args.get("vendedor", "").strip()
        busca = request.args.get("busca", "").strip()
        produto = request.args.get("produto", "").strip()
        regiao = request.args.get("regiao", "").strip()
        data_inicio = request.args.get("data_inicio", "").strip()
        data_fim = request.args.get("data_fim", "").strip()
        atrasados = request.args.get("atrasados", "").strip()
        page = request.args.get("page", 1, type=int)

        if cliente:
            query = query.filter(Pedido.cliente.ilike(f"%{cliente}%"))
        if estacao:
            query = query.filter(Pedido.itens.any(ItemPedido.estacao == estacao))
        if vendedor:
            query = query.filter(Pedido.vendedor.ilike(f"%{vendedor}%"))
        if busca:
            like = f"%{busca}%"
            query = query.filter(
                or_(
                    Pedido.pedido_venda.ilike(like),
                    Pedido.cliente.ilike(like),
                    Pedido.itens.any(ItemPedido.descricao_produto.ilike(like)),
                )
            )
        if produto:
            query = query.filter(Pedido.itens.any(ItemPedido.descricao_produto.ilike(f"%{produto}%")))
        if regiao:
            ufs_da_regiao = [uf for uf, r in REGIAO_POR_UF.items() if r == regiao]
            if ufs_da_regiao:
                query = query.filter(Pedido.estado.in_(ufs_da_regiao))
        if data_inicio:
            data_inicio_parsed = _parse_data_form(data_inicio)
            if data_inicio_parsed:
                query = query.filter(Pedido.data_inclusao_pedido >= data_inicio_parsed)
        if data_fim:
            data_fim_parsed = _parse_data_form(data_fim)
            if data_fim_parsed:
                query = query.filter(Pedido.data_inclusao_pedido <= data_fim_parsed)
        if atrasados:
            query = query.filter(_predicado_atrasado())
        if status:
            predicado = _predicado_status(status)
            if predicado is not None:
                query = query.filter(predicado)

        query = query.order_by(Pedido.data_inclusao_pedido.desc().nullslast(), Pedido.id.desc())

        total_filtrado = query.count()
        pedidos = query.offset((page - 1) * PAGE_SIZE).limit(PAGE_SIZE).all()
        total_paginas = max(1, (total_filtrado + PAGE_SIZE - 1) // PAGE_SIZE)

        resumo = _calcular_resumo()

        return render_template(
            "dashboard.html",
            pedidos=pedidos,
            resumo=resumo,
            page=page,
            total_paginas=total_paginas,
            total_filtrado=total_filtrado,
            filtros=dict(
                cliente=cliente,
                status=status,
                estacao=estacao,
                vendedor=vendedor,
                busca=busca,
                produto=produto,
                regiao=regiao,
                data_inicio=data_inicio,
                data_fim=data_fim,
                atrasados=atrasados,
            ),
        )

    @app.route("/pedidos/novo", methods=["GET", "POST"])
    @requer_role("ADMIN", "PCP")
    def novo_pedido():
        if request.method == "POST":
            f = request.form
            descricoes = f.getlist("item_descricao[]")
            quantidades = f.getlist("item_quantidade[]")
            custos = f.getlist("item_custo[]")

            itens = []
            itens_recarregados = []
            for desc, qtd, custo in zip(descricoes, quantidades, custos):
                itens_recarregados.append({"descricao": desc, "quantidade": qtd, "custo": custo})
                desc = desc.strip()
                if not desc:
                    continue
                itens.append(
                    ItemPedido(
                        descricao_produto=desc,
                        quantidade=_parse_float_form(qtd),
                        custo_unitario=_parse_float_form(custo),
                    )
                )

            pedido = Pedido(
                data_cliente=_parse_data_form(f.get("data_cliente")),
                data_inclusao_pedido=_parse_data_form(f.get("data_inclusao_pedido")),
                cliente=f.get("cliente", "").strip(),
                cnpj=f.get("cnpj", "").strip() or None,
                cidade=f.get("cidade", "").strip() or None,
                estado=f.get("estado", "").strip() or None,
                pais=f.get("pais", "").strip() or "Brasil",
                frete=f.get("frete") or None,
                vendedor=f.get("vendedor", "").strip() or None,
                pedido_venda=f.get("pedido_venda", "").strip() or None,
                prioridade=f.get("prioridade") or "MÉDIA",
            )
            pedido.itens = itens

            if not pedido.cliente or not itens:
                flash("Cliente e ao menos um produto (com descrição) são obrigatórios.", "danger")
                return render_template("novo_pedido.html", form=f, itens=itens_recarregados)

            for item in itens:
                item.atualizar_status_automatico()
            db.session.add(pedido)
            db.session.commit()
            flash(f"Pedido de {pedido.cliente} incluído com sucesso ({len(itens)} item(ns)).", "success")
            return redirect(url_for("editar_pedido", pedido_id=pedido.id))

        return render_template("novo_pedido.html", form={}, itens=[])

    @app.route("/pedidos/<int:pedido_id>/editar", methods=["GET", "POST"])
    @requer_role("ADMIN", "PCP", "LIDER")
    def editar_pedido(pedido_id):
        pedido = db.session.get(Pedido, pedido_id)
        if pedido is None:
            flash("Pedido não encontrado.", "danger")
            return redirect(url_for("dashboard"))

        if request.method == "POST":
            f = request.form

            pedido_antes = {c: getattr(pedido, c) for c in CAMPOS_HISTORICO_PEDIDO}

            pedido.data_cliente = _parse_data_form(f.get("data_cliente"))
            pedido.data_inclusao_pedido = _parse_data_form(f.get("data_inclusao_pedido"))
            pedido.cliente = f.get("cliente", "").strip() or pedido.cliente
            pedido.cnpj = f.get("cnpj", "").strip() or None
            pedido.cidade = f.get("cidade", "").strip() or None
            pedido.estado = f.get("estado", "").strip() or None
            pedido.pais = f.get("pais", "").strip() or "Brasil"
            pedido.frete = f.get("frete") or None
            pedido.vendedor = f.get("vendedor", "").strip() or None
            pedido.pedido_venda = f.get("pedido_venda", "").strip() or None

            pedido.prioridade = f.get("prioridade") or pedido.prioridade
            pedido.obs = f.get("obs", "").strip() or None

            pedido_depois = {c: getattr(pedido, c) for c in CAMPOS_HISTORICO_PEDIDO}
            _registrar_alteracoes("pedido", pedido.id, pedido.id, pedido_antes, pedido_depois, CAMPOS_HISTORICO_PEDIDO)

            # ---- itens do pedido (vários produtos, cada um com sua própria produção) ----
            item_ids = f.getlist("item_id[]")
            descricoes = f.getlist("item_descricao[]")
            quantidades = f.getlist("item_quantidade[]")
            custos = f.getlist("item_custo[]")
            estacoes = f.getlist("item_estacao[]")
            status_itens = f.getlist("item_status[]")
            rncs = f.getlist("item_rnc[]")
            inicios_producao = f.getlist("item_inicio_producao[]")
            inicios_inspecao = f.getlist("item_inicio_inspecao[]")
            terminos_inspecao = f.getlist("item_termino_inspecao[]")
            liberacoes_faturamento = f.getlist("item_liberacao_faturamento[]")
            liberacoes_prevista = f.getlist("item_liberacao_prevista[]")

            itens_originais = {item.id: item for item in pedido.itens}
            # snapshot ANTES de qualquer alteração, só dos itens que já existiam
            # (itens novos não têm "antes" pra comparar — são uma inclusão, não uma mudança)
            itens_antes = {iid: {c: getattr(item, c) for c in CAMPOS_HISTORICO_ITEM} for iid, item in itens_originais.items()}
            ids_mantidos = set()

            linhas = zip(
                item_ids, descricoes, quantidades, custos, estacoes, status_itens, rncs,
                inicios_producao, inicios_inspecao, terminos_inspecao,
                liberacoes_faturamento, liberacoes_prevista,
            )
            for (item_id, desc, qtd, custo, estacao_item, status_item, rnc,
                 ini_prod, ini_insp, term_insp, lib_fat, lib_prev) in linhas:
                desc = desc.strip()

                if item_id:
                    iid = int(item_id)
                    item = itens_originais.get(iid)
                    if item is None:
                        continue
                    if not desc:
                        continue  # descrição apagada -> item será removido abaixo
                else:
                    if not desc:
                        continue
                    item = ItemPedido()
                    pedido.itens.append(item)

                item.descricao_produto = desc
                item.quantidade = _parse_float_form(qtd)
                item.custo_unitario = _parse_float_form(custo)
                item.estacao = estacao_item or None
                item.rnc = rnc.strip() or None
                item.inicio_producao = _parse_data_form(ini_prod)
                item.inicio_inspecao = _parse_data_form(ini_insp)
                item.termino_inspecao = _parse_data_form(term_insp)
                item.liberacao_faturamento = _parse_data_form(lib_fat)
                item.liberacao_prevista = _parse_data_form(lib_prev)

                if status_item == "EM TRATATIVA":
                    item.status_manual = True
                    item.status_producao = "EM TRATATIVA"
                else:
                    item.status_manual = False
                item.atualizar_status_automatico()

                if item_id:
                    ids_mantidos.add(int(item_id))
                    item_depois = {c: getattr(item, c) for c in CAMPOS_HISTORICO_ITEM}
                    _registrar_alteracoes(
                        "item_pedido", int(item_id), pedido.id, itens_antes[int(item_id)], item_depois, CAMPOS_HISTORICO_ITEM
                    )

            for iid, item in itens_originais.items():
                if iid not in ids_mantidos:
                    pedido.itens.remove(item)

            if not pedido.itens:
                flash("O pedido precisa ter ao menos um produto.", "danger")
                return render_template("editar_pedido.html", pedido=pedido)

            db.session.commit()
            flash("Pedido atualizado com sucesso.", "success")
            return redirect(url_for("editar_pedido", pedido_id=pedido.id))

        return render_template("editar_pedido.html", pedido=pedido)

    @app.route("/pedidos/<int:pedido_id>")
    @login_required
    def detalhe_pedido(pedido_id):
        pedido = db.session.get(Pedido, pedido_id)
        if pedido is None:
            flash("Pedido não encontrado.", "danger")
            return redirect(url_for("dashboard"))
        historico = (
            HistoricoAlteracao.query.filter_by(pedido_id=pedido.id)
            .order_by(HistoricoAlteracao.criado_em.desc())
            .all()
        )
        timeline = _construir_timeline(pedido)
        return render_template("detalhe_pedido.html", pedido=pedido, historico=historico, timeline=timeline)

    @app.route("/pedidos/<int:pedido_id>/excluir", methods=["POST"])
    @requer_role("ADMIN", "PCP")
    def excluir_pedido(pedido_id):
        pedido = db.session.get(Pedido, pedido_id)
        if pedido is not None:
            db.session.delete(pedido)
            db.session.commit()
            flash("Pedido excluído.", "info")
        return redirect(url_for("dashboard"))

    @app.route("/api/pedidos/<int:pedido_id>/resumo")
    @login_required
    def api_resumo_pedido(pedido_id):
        pedido = db.session.get(Pedido, pedido_id)
        if pedido is None:
            return jsonify({"erro": "não encontrado"}), 404
        return jsonify(
            {
                "valor_total": pedido.valor_total,
                "lt_comercial_dias": pedido.lt_comercial_dias,
                "prazo_total_dias": pedido.prazo_total_dias,
                "status_producao": pedido.status_producao,
                "itens": [
                    {
                        "descricao_produto": item.descricao_produto,
                        "estacao": item.estacao,
                        "status_producao": item.status_producao,
                        "tempo_espera_dias": item.tempo_espera_dias,
                        "lt_producao_dias": item.lt_producao_dias,
                        "prazo_total_dias": item.prazo_total_dias,
                        "valor_total": item.valor_total,
                    }
                    for item in pedido.itens
                ],
            }
        )

    # ------------------------------------------------------------------
    # Usuários (papéis de acesso) — só ADMIN cadastra/edita usuários
    # ------------------------------------------------------------------
    @app.route("/usuarios")
    @requer_role("ADMIN")
    def usuarios_lista():
        usuarios = Usuario.query.order_by(Usuario.nome).all()
        return render_template("usuarios_lista.html", usuarios=usuarios)

    @app.route("/usuarios/novo", methods=["GET", "POST"])
    @requer_role("ADMIN")
    def usuarios_novo():
        if request.method == "POST":
            f = request.form
            nome = f.get("nome", "").strip()
            username = f.get("username", "").strip()
            senha = f.get("senha", "")
            role = f.get("role") or "PCP"
            setor = f.get("setor", "").strip() or None

            if not nome or not username or not senha:
                flash("Nome, usuário e senha são obrigatórios.", "danger")
                return render_template("usuarios_form.html", usuario=None, form=f)

            if Usuario.query.filter_by(username=username).first():
                flash("Já existe um usuário com esse nome de login.", "danger")
                return render_template("usuarios_form.html", usuario=None, form=f)

            if role not in ROLES:
                role = "PCP"

            novo = Usuario(nome=nome, username=username, role=role, setor=setor if role == "LIDER" else None)
            novo.set_senha(senha)
            db.session.add(novo)
            db.session.commit()
            flash(f"Usuário {nome} criado com sucesso.", "success")
            return redirect(url_for("usuarios_lista"))

        return render_template("usuarios_form.html", usuario=None, form={})

    @app.route("/usuarios/<int:usuario_id>/editar", methods=["GET", "POST"])
    @requer_role("ADMIN")
    def usuarios_editar(usuario_id):
        usuario = db.session.get(Usuario, usuario_id)
        if usuario is None:
            flash("Usuário não encontrado.", "danger")
            return redirect(url_for("usuarios_lista"))

        if request.method == "POST":
            f = request.form
            nome = f.get("nome", "").strip()
            username = f.get("username", "").strip()
            role = f.get("role") or usuario.role
            setor = f.get("setor", "").strip() or None
            senha = f.get("senha", "")

            if not nome or not username:
                flash("Nome e usuário são obrigatórios.", "danger")
                return render_template("usuarios_form.html", usuario=usuario, form=f)

            outro = Usuario.query.filter(Usuario.username == username, Usuario.id != usuario.id).first()
            if outro:
                flash("Já existe outro usuário com esse nome de login.", "danger")
                return render_template("usuarios_form.html", usuario=usuario, form=f)

            if role not in ROLES:
                role = usuario.role

            usuario.nome = nome
            usuario.username = username
            usuario.role = role
            usuario.setor = setor if role == "LIDER" else None
            usuario.ativo = bool(f.get("ativo"))
            if senha:
                usuario.set_senha(senha)

            # o usuário admin original nunca perde acesso total nem fica desativado,
            # mesmo que alguém manipule o formulário
            if usuario.username == "admin":
                usuario.role = "ADMIN"
                usuario.setor = None
                usuario.ativo = True

            db.session.commit()
            flash(f"Usuário {usuario.nome} atualizado com sucesso.", "success")
            return redirect(url_for("usuarios_lista"))

        return render_template("usuarios_form.html", usuario=usuario, form={})

    @app.route("/usuarios/<int:usuario_id>/desativar", methods=["POST"])
    @requer_role("ADMIN")
    def usuarios_desativar(usuario_id):
        usuario = db.session.get(Usuario, usuario_id)
        if usuario is None:
            flash("Usuário não encontrado.", "danger")
        elif usuario.username == "admin":
            flash("O usuário admin original não pode ser desativado.", "danger")
        else:
            usuario.ativo = not usuario.ativo
            db.session.commit()
            flash(
                f"Usuário {usuario.nome} {'reativado' if usuario.ativo else 'desativado'} com sucesso.",
                "info",
            )
        return redirect(url_for("usuarios_lista"))

    # ------------------------------------------------------------------
    # Estações — visão geral por setor + Kanban de produção
    # ------------------------------------------------------------------
    @app.route("/estacoes")
    @login_required
    def estacoes_lista():
        estacoes = Estacao.query.order_by(Estacao.ordem_exibicao).all()
        hoje = date.today()
        linhas = []
        for e in estacoes:
            fila = ItemPedido.query.filter(
                ItemPedido.estacao == e.nome, ItemPedido.status_producao != "FINALIZADO"
            ).count()
            criticos = ItemPedido.query.filter(
                ItemPedido.estacao == e.nome,
                ItemPedido.status_producao != "FINALIZADO",
                ItemPedido.liberacao_prevista.isnot(None),
                ItemPedido.liberacao_prevista < hoje,
            ).count()
            itens_lt = ItemPedido.query.filter(
                ItemPedido.estacao == e.nome,
                ItemPedido.inicio_producao.isnot(None),
                ItemPedido.termino_inspecao.isnot(None),
            ).all()
            lt_medio = (
                round(sum((i.termino_inspecao - i.inicio_producao).days for i in itens_lt) / len(itens_lt), 1)
                if itens_lt
                else None
            )
            linhas.append({"estacao": e, "fila": fila, "criticos": criticos, "lt_medio": lt_medio})
        return render_template("estacoes_lista.html", linhas=linhas)

    @app.route("/estacoes/<nome>")
    @login_required
    def estacao_kanban(nome):
        estacao = Estacao.query.filter_by(nome=nome).first()
        if estacao is None:
            flash("Estação não encontrada.", "danger")
            return redirect(url_for("estacoes_lista"))

        itens = ItemPedido.query.options(selectinload(ItemPedido.pedido)).filter(ItemPedido.estacao == nome).all()
        itens.sort(key=lambda i: (i.pedido.data_inclusao_pedido or date.min, i.pedido_id))

        colunas = {chave: [] for chave in STATUS_CHAO_OPCOES}
        for item in itens:
            colunas[item.status_chao].append(item)

        return render_template(
            "estacoes_kanban.html",
            estacao=estacao,
            colunas=colunas,
            pode_editar=pode_editar_estacao(current_user, nome),
        )

    @app.route("/estacoes/<nome>/kanban/mover", methods=["POST"])
    @login_required
    def estacao_kanban_mover(nome):
        item_id = request.form.get("item_id", type=int)
        item = db.session.get(ItemPedido, item_id) if item_id else None
        if item is None:
            flash("Item não encontrado.", "danger")
            return redirect(url_for("estacao_kanban", nome=nome))

        if not pode_editar_estacao(current_user, item.estacao or nome):
            flash("Você não tem permissão para mover itens desta estação.", "danger")
            return redirect(url_for("estacao_kanban", nome=nome))

        antes = {c: getattr(item, c) for c in CAMPOS_HISTORICO_ITEM}
        hoje = date.today()
        etapa_atual = item.status_chao

        if etapa_atual in ("NAO_INICIADO", "PROGRAMADO"):
            item.inicio_producao = hoje
        elif etapa_atual == "EM_PRODUCAO":
            item.inicio_inspecao = hoje
        elif etapa_atual == "INSPECAO":
            item.termino_inspecao = hoje
        elif etapa_atual == "EMBALAGEM":
            item.liberacao_faturamento = hoje
        else:
            flash("Este item já está finalizado.", "info")
            return redirect(url_for("estacao_kanban", nome=nome))

        item.status_manual = False
        item.atualizar_status_automatico()

        depois = {c: getattr(item, c) for c in CAMPOS_HISTORICO_ITEM}
        _registrar_alteracoes("item_pedido", item.id, item.pedido_id, antes, depois, CAMPOS_HISTORICO_ITEM)
        db.session.commit()

        flash(
            f"\"{item.descricao_produto}\" avançou para {STATUS_CHAO_LABELS.get(item.status_chao, item.status_chao)}.",
            "success",
        )
        return redirect(url_for("estacao_kanban", nome=nome))

    # ------------------------------------------------------------------
    # Cadastros > Estações — CRUD simples (ADMIN/PCP)
    # ------------------------------------------------------------------
    @app.route("/cadastros/estacoes")
    @requer_role("ADMIN", "PCP")
    def cadastros_estacoes():
        estacoes = Estacao.query.order_by(Estacao.ordem_exibicao).all()
        return render_template("cadastros_estacoes.html", estacoes=estacoes)

    @app.route("/cadastros/estacoes/novo", methods=["GET", "POST"])
    @requer_role("ADMIN", "PCP")
    def cadastros_estacoes_novo():
        if request.method == "POST":
            f = request.form
            nome = f.get("nome", "").strip().upper()
            if not nome:
                flash("Informe o nome da estação.", "danger")
                return render_template("cadastros_estacoes_form.html", estacao=None, form=f)
            if Estacao.query.filter_by(nome=nome).first():
                flash("Já existe uma estação com esse nome.", "danger")
                return render_template("cadastros_estacoes_form.html", estacao=None, form=f)

            maior_ordem = db.session.query(func.max(Estacao.ordem_exibicao)).scalar() or 0
            meta = f.get("meta_lead_time_dias", "").strip()
            nova = Estacao(
                nome=nome,
                ordem_exibicao=maior_ordem + 1,
                meta_lead_time_dias=int(meta) if meta.isdigit() else None,
                ativo=True,
            )
            db.session.add(nova)
            db.session.commit()
            flash(f"Estação {nome} cadastrada com sucesso.", "success")
            return redirect(url_for("cadastros_estacoes"))

        return render_template("cadastros_estacoes_form.html", estacao=None, form={})

    @app.route("/cadastros/estacoes/<int:estacao_id>/editar", methods=["GET", "POST"])
    @requer_role("ADMIN", "PCP")
    def cadastros_estacoes_editar(estacao_id):
        estacao = db.session.get(Estacao, estacao_id)
        if estacao is None:
            flash("Estação não encontrada.", "danger")
            return redirect(url_for("cadastros_estacoes"))

        if request.method == "POST":
            f = request.form
            nome = f.get("nome", "").strip().upper()
            if not nome:
                flash("Informe o nome da estação.", "danger")
                return render_template("cadastros_estacoes_form.html", estacao=estacao, form=f)
            outra = Estacao.query.filter(Estacao.nome == nome, Estacao.id != estacao.id).first()
            if outra:
                flash("Já existe outra estação com esse nome.", "danger")
                return render_template("cadastros_estacoes_form.html", estacao=estacao, form=f)

            estacao.nome = nome
            meta = f.get("meta_lead_time_dias", "").strip()
            estacao.meta_lead_time_dias = int(meta) if meta.isdigit() else None
            estacao.ativo = bool(f.get("ativo"))
            db.session.commit()
            flash(f"Estação {estacao.nome} atualizada com sucesso.", "success")
            return redirect(url_for("cadastros_estacoes"))

        return render_template("cadastros_estacoes_form.html", estacao=estacao, form={})


app = create_app()

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
