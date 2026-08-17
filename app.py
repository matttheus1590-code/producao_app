import os
from datetime import datetime

from flask import Flask, flash, jsonify, redirect, render_template, request, url_for
from flask_login import current_user, login_required, login_user, logout_user
from sqlalchemy import inspect, text
from sqlalchemy.orm import selectinload

from extensions import db, login_manager
from models import (
    ESTACOES,
    FRETE_OPCOES,
    PRIORIDADE_CORES,
    PRIORIDADE_OPCOES,
    STATUS_CORES,
    STATUS_OPCOES,
    ItemPedido,
    Pedido,
    Usuario,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PAGE_SIZE = 25


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
        _seed_inicial(app)

    @app.context_processor
    def inject_globals():
        return dict(
            ESTACOES=ESTACOES,
            STATUS_OPCOES=STATUS_OPCOES,
            PRIORIDADE_OPCOES=PRIORIDADE_OPCOES,
            FRETE_OPCOES=FRETE_OPCOES,
            STATUS_CORES=STATUS_CORES,
            PRIORIDADE_CORES=PRIORIDADE_CORES,
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


def _seed_inicial(app):
    """Cria o usuário admin padrão e importa a planilha na primeira execução."""
    if Usuario.query.count() == 0:
        admin = Usuario(nome="Administrador", username="admin")
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

    @app.route("/")
    @login_required
    def dashboard():
        query = Pedido.query

        cliente = request.args.get("cliente", "").strip()
        status = request.args.get("status", "").strip()
        estacao = request.args.get("estacao", "").strip()
        vendedor = request.args.get("vendedor", "").strip()
        busca = request.args.get("busca", "").strip()
        page = request.args.get("page", 1, type=int)

        if cliente:
            query = query.filter(Pedido.cliente.ilike(f"%{cliente}%"))
        if status:
            query = query.filter(Pedido.status_producao == status)
        if estacao:
            query = query.filter(Pedido.estacao == estacao)
        if vendedor:
            query = query.filter(Pedido.vendedor.ilike(f"%{vendedor}%"))
        if busca:
            like = f"%{busca}%"
            query = query.filter(
                db.or_(
                    Pedido.pedido_venda.ilike(like),
                    Pedido.cliente.ilike(like),
                    Pedido.itens.any(ItemPedido.descricao_produto.ilike(like)),
                )
            )

        total_filtrado = query.count()
        pedidos = (
            query.options(selectinload(Pedido.itens))
            .order_by(Pedido.data_inclusao_pedido.desc().nullslast(), Pedido.id.desc())
            .offset((page - 1) * PAGE_SIZE)
            .limit(PAGE_SIZE)
            .all()
        )
        total_paginas = max(1, (total_filtrado + PAGE_SIZE - 1) // PAGE_SIZE)

        resumo = {
            "total": Pedido.query.count(),
            "pendente": Pedido.query.filter_by(status_producao="PENDENTE").count(),
            "em_tratativa": Pedido.query.filter_by(status_producao="EM TRATATIVA").count(),
            "andamento": Pedido.query.filter_by(status_producao="ANDAMENTO").count(),
            "finalizado": Pedido.query.filter_by(status_producao="FINALIZADO").count(),
            "valor_total": sum(p.valor_total for p in Pedido.query.options(selectinload(Pedido.itens)).all()),
        }

        return render_template(
            "dashboard.html",
            pedidos=pedidos,
            resumo=resumo,
            page=page,
            total_paginas=total_paginas,
            total_filtrado=total_filtrado,
            filtros=dict(cliente=cliente, status=status, estacao=estacao, vendedor=vendedor, busca=busca),
        )

    @app.route("/pedidos/novo", methods=["GET", "POST"])
    @login_required
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

            pedido.atualizar_status_automatico()
            db.session.add(pedido)
            db.session.commit()
            flash(f"Pedido de {pedido.cliente} incluído com sucesso ({len(itens)} item(ns)).", "success")
            return redirect(url_for("editar_pedido", pedido_id=pedido.id))

        return render_template("novo_pedido.html", form={}, itens=[])

    @app.route("/pedidos/<int:pedido_id>/editar", methods=["GET", "POST"])
    @login_required
    def editar_pedido(pedido_id):
        pedido = db.session.get(Pedido, pedido_id)
        if pedido is None:
            flash("Pedido não encontrado.", "danger")
            return redirect(url_for("dashboard"))

        if request.method == "POST":
            f = request.form

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

            # ---- itens do pedido (vários produtos) ----
            item_ids = f.getlist("item_id[]")
            descricoes = f.getlist("item_descricao[]")
            quantidades = f.getlist("item_quantidade[]")
            custos = f.getlist("item_custo[]")

            itens_originais = {item.id: item for item in pedido.itens}
            ids_mantidos = set()

            for item_id, desc, qtd, custo in zip(item_ids, descricoes, quantidades, custos):
                desc = desc.strip()
                if item_id:
                    iid = int(item_id)
                    item = itens_originais.get(iid)
                    if item is None:
                        continue
                    if not desc:
                        continue  # descrição apagada -> item será removido abaixo
                    item.descricao_produto = desc
                    item.quantidade = _parse_float_form(qtd)
                    item.custo_unitario = _parse_float_form(custo)
                    ids_mantidos.add(iid)
                else:
                    if not desc:
                        continue
                    pedido.itens.append(
                        ItemPedido(
                            descricao_produto=desc,
                            quantidade=_parse_float_form(qtd),
                            custo_unitario=_parse_float_form(custo),
                        )
                    )

            for iid, item in itens_originais.items():
                if iid not in ids_mantidos:
                    pedido.itens.remove(item)

            if not pedido.itens:
                flash("O pedido precisa ter ao menos um produto.", "danger")
                return render_template("editar_pedido.html", pedido=pedido)

            # ---- campos parametrizados ----
            pedido.prioridade = f.get("prioridade") or pedido.prioridade
            pedido.estacao = f.get("estacao") or None
            pedido.inicio_producao = _parse_data_form(f.get("inicio_producao"))
            pedido.inicio_inspecao = _parse_data_form(f.get("inicio_inspecao"))
            pedido.termino_inspecao = _parse_data_form(f.get("termino_inspecao"))
            pedido.liberacao_faturamento = _parse_data_form(f.get("liberacao_faturamento"))
            pedido.liberacao_prevista = _parse_data_form(f.get("liberacao_prevista"))
            pedido.rnc = f.get("rnc", "").strip() or None
            pedido.obs = f.get("obs", "").strip() or None

            status_form = f.get("status_producao")
            if status_form == "EM TRATATIVA":
                pedido.status_manual = True
                pedido.status_producao = "EM TRATATIVA"
            else:
                pedido.status_manual = False

            pedido.atualizar_status_automatico()
            db.session.commit()
            flash("Pedido atualizado com sucesso.", "success")
            return redirect(url_for("editar_pedido", pedido_id=pedido.id))

        return render_template("editar_pedido.html", pedido=pedido)

    @app.route("/pedidos/<int:pedido_id>/excluir", methods=["POST"])
    @login_required
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
                "tempo_espera_dias": pedido.tempo_espera_dias,
                "lt_producao_dias": pedido.lt_producao_dias,
                "prazo_total_dias": pedido.prazo_total_dias,
                "status_producao": pedido.status_producao,
            }
        )


app = create_app()

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
