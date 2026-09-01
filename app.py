import csv
import io
import os
import re
from calendar import monthrange
from datetime import date, datetime, timedelta

from flask import Flask, Response, flash, jsonify, redirect, render_template, request, url_for
from flask_login import current_user, login_required, login_user, logout_user
from openpyxl import Workbook
from openpyxl.styles import Font
from sqlalchemy import and_, extract, false, func, inspect, or_, text
from sqlalchemy.orm import selectinload

from extensions import db, login_manager
from models import (
    ESTACOES,
    FRETE_OPCOES,
    GO_OTD_META_PERCENTUAL,
    PRAZO_ALERTA_DIAS,
    PRIORIDADE_CORES,
    PRIORIDADE_OPCOES,
    REGIAO_POR_UF,
    REGIOES_OPCOES,
    RNC_DISPOSICAO_OPCOES,
    RNC_EFICACIA_CORES,
    RNC_EFICACIA_OPCOES,
    RNC_EMITENTE_OPCOES,
    RNC_FERRAMENTA_ANALISE_OPCOES,
    RNC_LOCAL_SETOR_OPCOES,
    RNC_ORIGEM_OPCOES,
    RNC_SETOR_OPCOES,
    RNC_SEVERIDADE_CORES,
    RNC_SEVERIDADE_OPCOES,
    RNC_SIM_NAO_OPCOES,
    RNC_STATUS_ACAO_OPCOES,
    RNC_STATUS_GERAL_ABERTOS,
    RNC_STATUS_GERAL_CORES,
    RNC_STATUS_GERAL_OPCOES,
    RNC_TIPO_NC_OPCOES,
    SEMAFORO_CORES,
    SEMAFORO_LABELS,
    STATUS_CHAO_CORES,
    STATUS_CHAO_LABELS,
    STATUS_CHAO_OPCOES,
    STATUS_CORES,
    STATUS_OPCOES,
    ControleSistema,
    Estacao,
    HistoricoAlteracao,
    ItemPedido,
    Pedido,
    PedidoOperacao,
    Programacao,
    RncQualidade,
    Transportadora,
    Usuario,
    gerar_semanas_pcp,
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
    "liberacao_real",
    "planejamento_semanal",
    "rnc",
    "numero_nota_fiscal",
    "valor_faturado",
    "transportadora_id",
    "data_envio",
]

# Campos de Gestão Operação (Fase 13) que geram histórico de alteração —
# igual em espírito a CAMPOS_HISTORICO_PEDIDO/ITEM, só que namespaced com
# "go_" (não precisa de tabela nova: usa o mesmo HistoricoAlteracao de sempre).
CAMPOS_HISTORICO_GESTAO_OPERACAO = [
    "go_tipo_pedido",
    "go_status_pedido_info",
    "go_previsao_liberacao_pcp",
    "go_data_efetiva_liberacao_pcp",
    "go_status_logistica",
    "go_data_pedido_expedido",
    "go_transportadora_id",
    "go_data_real_entrega",
    "go_otd_realizado",
    "go_data_entregue_cliente",
    "go_status_final_alinhamento",
]

# Pedido do Bruno (31/08/2026): a tela de edição de Gestão Operação deixa de
# mostrar TODOS os blocos (Comercial/PCP/Logística/Resultados) de uma vez —
# cada aba só edita os campos da própria área. GO_CAMPOS_POR_SECAO é a única
# fonte de verdade de "quais campos pertencem a cada aba", usada tanto pra
# decidir o que o template desenha quanto pra decidir o que a rota lê do
# formulário — assim um campo fora da seção atual nunca é tocado (nem lido,
# nem zerado) ao salvar.
GO_SECOES = ("comercial", "pcp", "logistica", "resultados")
GO_CAMPOS_POR_SECAO = {
    "comercial": [
        "go_tipo_pedido", "go_contrato", "go_pedido_compra_cliente", "go_proposta",
        "go_data_solicitada_entrega", "go_status_pedido_info", "go_valor_pedido_operacao",
    ],
    "pcp": [
        "go_previsao_liberacao_pcp", "go_data_efetiva_liberacao_pcp",
        "go_data_solicitada_cliente_retira", "go_custo_producao_real", "go_termino_semanal_pcp",
    ],
    "logistica": [
        "go_data_emissao_nf", "go_valor_nf_emitida", "go_numero_nf", "go_status_logistica",
        "go_data_pedido_expedido", "go_transportadora_id", "go_custo_frete_previsto",
        "go_custo_frete_final", "go_custo_frete_sobre_nota", "go_data_prevista_entrega",
        "go_data_real_entrega",
    ],
    "resultados": [
        "go_otd_realizado", "go_data_solicitada_cliente_final", "go_data_entregue_cliente",
        "go_obs_operacao", "go_status_final_alinhamento",
    ],
}
GO_SECAO_ENDPOINT = {
    # "comercial" aponta pra Listagem Geral (pedido do Bruno, 01/09/2026: a
    # aba/lista "Comercial" separada foi apagada porque a Listagem Geral já
    # mostra as mesmas informações — só o FORMULÁRIO de edição da seção
    # Comercial continua existindo, dentro de gestao_operacao_editar).
    "comercial": "gestao_operacao_listagem_geral",
    "pcp": "gestao_operacao_pcp",
    "logistica": "gestao_operacao_logistica",
    "resultados": "gestao_operacao_resultados",
}
GO_SECAO_LABEL = {
    "comercial": "Comercial", "pcp": "PCP", "logistica": "Logística / NF", "resultados": "Resultados / OTD",
}
_GO_CAMPOS_DATA = {
    "go_data_solicitada_entrega", "go_previsao_liberacao_pcp", "go_data_efetiva_liberacao_pcp",
    "go_data_solicitada_cliente_retira", "go_data_emissao_nf", "go_data_pedido_expedido",
    "go_data_prevista_entrega", "go_data_real_entrega", "go_data_solicitada_cliente_final",
    "go_data_entregue_cliente",
}
_GO_CAMPOS_FLOAT = {
    "go_valor_pedido_operacao", "go_custo_producao_real", "go_valor_nf_emitida",
    "go_custo_frete_previsto", "go_custo_frete_final", "go_custo_frete_sobre_nota",
}


def _parse_campo_go(campo, f):
    """Lê e converte UM campo go_* do formulário — usado pela edição
    seccionada de Gestão Operação, campo por campo, só para os campos da
    seção que está sendo salva."""
    if campo == "go_transportadora_id":
        valor = f.get("go_transportadora_id", "")
        return int(valor) if valor.strip().isdigit() else None
    if campo in _GO_CAMPOS_DATA:
        return _parse_data_form(f.get(campo))
    if campo in _GO_CAMPOS_FLOAT:
        valor = f.get(campo, "")
        return _parse_float_form(valor, default=None) if valor.strip() else None
    return f.get(campo, "").strip() or None


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
        _migrar_faturamento_itens(app)
        _migrar_logistica_itens(app)
        _migrar_planejamento_semanal_itens(app)
        _migrar_liberacao_real_itens(app)
        _migrar_atualizado_em_itens(app)
        _migrar_gestao_operacao_pedidos(app)
        _migrar_dados_go_para_pedidos_operacao(app)
        # Só depois de TODAS as colunas de pedidos/itens existirem de verdade
        # (senão o ORM tenta selecionar coluna que ainda não foi criada nesta
        # execução, em bancos antigos que ainda não passaram pelas migrações
        # acima) — consolida os pedidos fragmentados em vários registros.
        _consolidar_pedidos_duplicados(app)
        _seed_inicial(app)
        _importar_gestao_operacao(app)
        # Roda por último: depende de Pedido/ItemPedido já existirem com todas
        # as colunas migradas, e de _consolidar_pedidos_duplicados já ter
        # deixado (no máximo) 1 Pedido por pedido_venda.
        _sincronizar_planilha_producao_25_08_2026(app)
        # Idem para Gestão Operação: depende de _importar_gestao_operacao já
        # ter rodado (banco com PedidoOperacao populado) antes de sincronizar
        # por cima com a planilha mais nova.
        _sincronizar_gestao_operacao_28_08_2026(app)
        # Correção pontual pedida pelo Bruno depois de conferir os números da
        # semana 04/Ago manualmente — precisa rodar depois da sincronização
        # acima (que é quem trouxe o pedido 835 com a semana errada).
        _corrigir_semana_pcp_morken_835(app)
        _seed_rnc_qualidade(app)

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
            GO_OTD_META_PERCENTUAL=GO_OTD_META_PERCENTUAL,
            GO_SEMANAS_PCP=gerar_semanas_pcp(),
            RNC_EMITENTE_OPCOES=RNC_EMITENTE_OPCOES,
            RNC_SETOR_OPCOES=RNC_SETOR_OPCOES,
            RNC_ORIGEM_OPCOES=RNC_ORIGEM_OPCOES,
            RNC_LOCAL_SETOR_OPCOES=RNC_LOCAL_SETOR_OPCOES,
            RNC_TIPO_NC_OPCOES=RNC_TIPO_NC_OPCOES,
            RNC_SEVERIDADE_OPCOES=RNC_SEVERIDADE_OPCOES,
            RNC_FERRAMENTA_ANALISE_OPCOES=RNC_FERRAMENTA_ANALISE_OPCOES,
            RNC_DISPOSICAO_OPCOES=RNC_DISPOSICAO_OPCOES,
            RNC_STATUS_ACAO_OPCOES=RNC_STATUS_ACAO_OPCOES,
            RNC_EFICACIA_OPCOES=RNC_EFICACIA_OPCOES,
            RNC_SIM_NAO_OPCOES=RNC_SIM_NAO_OPCOES,
            RNC_STATUS_GERAL_OPCOES=RNC_STATUS_GERAL_OPCOES,
            RNC_SEVERIDADE_CORES=RNC_SEVERIDADE_CORES,
            RNC_STATUS_GERAL_CORES=RNC_STATUS_GERAL_CORES,
            RNC_EFICACIA_CORES=RNC_EFICACIA_CORES,
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


def _consolidar_pedidos_duplicados(app):
    """Corrige um problema histórico do import original: pedidos com mais de
    um produto foram importados como VÁRIOS registros `Pedido` separados (um
    por linha da planilha, cada um com 1 item só) em vez de 1 `Pedido` com
    vários `ItemPedido` dentro. Isso fazia "Venda total pedido" somar só 1
    produto por vez, e a tela de editar não mostrar todos os produtos de um
    pedido juntos (ex.: CATTALINI 728 = 4 registros `Pedido` separados, cada
    um com 1 item, em vez de 1 registro com 4 itens).

    Roda uma vez: agrupa os `Pedido` que compartilham o mesmo `pedido_venda`,
    escolhe o de menor id como "principal", move todos os itens — e o
    histórico de alterações — dos outros pra ele, e apaga os registros que
    sobraram vazios. Não perde nenhum item nem histórico, só reorganiza quem
    é o dono (pedido_id). Idempotente: numa segunda execução não encontra
    mais grupos com mais de 1 registro por pedido_venda, então não faz nada."""
    grupos = (
        db.session.query(Pedido.pedido_venda)
        .filter(Pedido.pedido_venda.isnot(None), Pedido.pedido_venda != "")
        .group_by(Pedido.pedido_venda)
        .having(func.count(Pedido.id) > 1)
        .all()
    )
    if not grupos:
        return

    total_pedidos_removidos = 0
    total_itens_movidos = 0

    for (pedido_venda,) in grupos:
        registros = Pedido.query.filter(Pedido.pedido_venda == pedido_venda).order_by(Pedido.id).all()
        if len(registros) < 2:
            continue  # já foi consolidado nesta mesma rodada (não deveria acontecer, mas por segurança)

        primario, duplicados = registros[0], registros[1:]

        datas_inclusao = [r.data_inclusao_pedido for r in registros if r.data_inclusao_pedido]
        if datas_inclusao:
            primario.data_inclusao_pedido = min(datas_inclusao)

        # Todos os outros campos escalares do pedido (comerciais + os go_*
        # legados de Gestão Operação, que ainda vivem fisicamente na tabela
        # mesmo sem serem mais lidos pelo app) — mantém o valor do principal
        # se já tiver algo, senão pega o primeiro valor não vazio encontrado
        # entre os duplicados. Genérico de propósito, pra não perder nenhum
        # dado esquecido numa lista manual de campos.
        campos_genericos = [
            c.name for c in Pedido.__table__.columns if c.name not in ("id", "pedido_venda", "data_inclusao_pedido")
        ]
        for campo in campos_genericos:
            if not getattr(primario, campo):
                for dup in duplicados:
                    valor = getattr(dup, campo)
                    if valor:
                        setattr(primario, campo, valor)
                        break

        for dup in duplicados:
            for item in list(dup.itens):
                item.pedido = primario  # via relationship, não só pedido_id — mantém o cascade consistente
                total_itens_movidos += 1
            HistoricoAlteracao.query.filter_by(pedido_id=dup.id).update(
                {"pedido_id": primario.id}, synchronize_session=False
            )
            db.session.flush()
            db.session.delete(dup)
            total_pedidos_removidos += 1

    db.session.commit()
    app.logger.info(
        "Migração automática: %d pedidos duplicados consolidados (%d registros removidos, %d itens reagrupados).",
        len(grupos),
        total_pedidos_removidos,
        total_itens_movidos,
    )


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


def _migrar_faturamento_itens(app):
    """Adiciona os campos de faturamento (número da nota fiscal e valor
    faturado) nos itens criados antes dessa fase existir.

    São campos 100% novos — a planilha original não tinha essa informação —
    então não há nada pra migrar/copiar: só ficam em branco nos itens antigos
    e passam a ser preenchidos dali pra frente."""
    inspector = inspect(db.engine)
    if "itens_pedido" not in inspector.get_table_names():
        return

    colunas = {c["name"] for c in inspector.get_columns("itens_pedido")}
    faltando = [c for c in ("numero_nota_fiscal", "valor_faturado") if c not in colunas]
    if not faltando:
        return

    with db.engine.begin() as conn:
        if "numero_nota_fiscal" not in colunas:
            conn.execute(text("ALTER TABLE itens_pedido ADD COLUMN numero_nota_fiscal VARCHAR(30)"))
        if "valor_faturado" not in colunas:
            conn.execute(text("ALTER TABLE itens_pedido ADD COLUMN valor_faturado FLOAT"))
    app.logger.info("Migração automática: campos de faturamento (NF e valor faturado) adicionados aos itens.")


def _migrar_logistica_itens(app):
    """Adiciona os campos de logística (transportadora e data de envio) nos
    itens criados antes dessa fase existir.

    Assim como faturamento, são campos 100% novos — a planilha original não
    tinha essa informação — então ficam em branco nos itens antigos e passam
    a ser preenchidos dali pra frente. "transportadora_id" não usa FK de
    verdade (ver comentário no models.py), então a coluna é só um INTEGER."""
    inspector = inspect(db.engine)
    if "itens_pedido" not in inspector.get_table_names():
        return

    colunas = {c["name"] for c in inspector.get_columns("itens_pedido")}
    faltando = [c for c in ("transportadora_id", "data_envio") if c not in colunas]
    if not faltando:
        return

    with db.engine.begin() as conn:
        if "transportadora_id" not in colunas:
            conn.execute(text("ALTER TABLE itens_pedido ADD COLUMN transportadora_id INTEGER"))
        if "data_envio" not in colunas:
            conn.execute(text("ALTER TABLE itens_pedido ADD COLUMN data_envio DATE"))
    app.logger.info("Migração automática: campos de logística (transportadora e data de envio) adicionados aos itens.")


def _migrar_planejamento_semanal_itens(app):
    """Adiciona o campo de planejamento semanal (preenchido manualmente pelo
    PCP, junto com a liberação prevista) nos itens criados antes dessa fase
    existir.

    É um campo 100% novo — não deriva de nenhum dado existente — então fica
    em branco nos itens antigos e passa a ser preenchido dali pra frente."""
    inspector = inspect(db.engine)
    if "itens_pedido" not in inspector.get_table_names():
        return

    colunas = {c["name"] for c in inspector.get_columns("itens_pedido")}
    if "planejamento_semanal" in colunas:
        return

    with db.engine.begin() as conn:
        conn.execute(text("ALTER TABLE itens_pedido ADD COLUMN planejamento_semanal VARCHAR(40)"))
    app.logger.info("Migração automática: campo de planejamento semanal adicionado aos itens.")


def _migrar_liberacao_real_itens(app):
    """Adiciona o campo de liberação real (preenchido manualmente, pra
    comparar com a liberação prevista) nos itens criados antes dessa fase
    existir.

    Campo 100% novo — não deriva de nenhum dado existente — fica em branco
    nos itens antigos e passa a ser preenchido dali pra frente."""
    inspector = inspect(db.engine)
    if "itens_pedido" not in inspector.get_table_names():
        return

    colunas = {c["name"] for c in inspector.get_columns("itens_pedido")}
    if "liberacao_real" in colunas:
        return

    with db.engine.begin() as conn:
        conn.execute(text("ALTER TABLE itens_pedido ADD COLUMN liberacao_real DATE"))
    app.logger.info("Migração automática: campo de liberação real adicionado aos itens.")


def _migrar_atualizado_em_itens(app):
    """Adiciona um carimbo de "última alteração" (atualizado_em) em cada item
    — o SQLAlchemy atualiza sozinho (onupdate) toda vez que o item é salvo,
    seja editando o pedido ou clicando em "Avançar" no Kanban das Estações.
    Usado pra ordenar o Kanban das Estações sempre com os itens mais
    novos/recém movimentados no topo de cada coluna (pedido do Bruno,
    25/08/2026).

    Como o campo não existia antes, faz um backfill único pros itens já
    existentes: usa a data mais avançada que o item já tem registrada
    (término de produção > início de produção > inclusão do pedido) como
    aproximação de "última atividade conhecida" — sem isso, todo item legado
    nasceria com o mesmo timestamp (o momento do deploy) e a ordenação
    ficaria arbitrária entre eles."""
    inspector = inspect(db.engine)
    if "itens_pedido" not in inspector.get_table_names():
        return

    colunas = {c["name"] for c in inspector.get_columns("itens_pedido")}
    if "atualizado_em" in colunas:
        return

    with db.engine.begin() as conn:
        # TIMESTAMP, não DATETIME: "DATETIME" é aceito pelo SQLite (que ignora
        # o nome do tipo) mas não existe no Postgres de produção — quebrou o
        # primeiro deploy dessa migração (psycopg2.errors.UndefinedObject).
        conn.execute(text("ALTER TABLE itens_pedido ADD COLUMN atualizado_em TIMESTAMP"))

    itens = ItemPedido.query.options(selectinload(ItemPedido.pedido)).all()
    agora = datetime.utcnow()
    for item in itens:
        melhor_data = (
            item.termino_inspecao
            or item.liberacao_faturamento
            or item.inicio_producao
            or (item.pedido.data_inclusao_pedido if item.pedido else None)
        )
        item.atualizado_em = datetime.combine(melhor_data, datetime.min.time()) if melhor_data else agora
    db.session.commit()
    app.logger.info("Migração automática: campo atualizado_em adicionado e preenchido em %d itens.", len(itens))


# Colunas novas da Fase 13 (Gestão Operação) e seu tipo SQL — todas opcionais,
# nenhuma substitui nada que já existe em Pedido.
_COLUNAS_GESTAO_OPERACAO = {
    # Comercial
    "go_tipo_pedido": "VARCHAR(60)",
    "go_contrato": "VARCHAR(60)",
    "go_pedido_compra_cliente": "VARCHAR(60)",
    "go_proposta": "VARCHAR(60)",
    "go_data_solicitada_entrega": "DATE",
    "go_status_pedido_info": "VARCHAR(120)",
    "go_valor_pedido_operacao": "FLOAT",
    # PCP
    "go_previsao_liberacao_pcp": "DATE",
    "go_data_efetiva_liberacao_pcp": "DATE",
    "go_data_solicitada_cliente_retira": "DATE",
    "go_custo_producao_real": "FLOAT",
    "go_termino_semanal_pcp": "VARCHAR(40)",
    # Logística / NF
    "go_data_emissao_nf": "DATE",
    "go_valor_nf_emitida": "FLOAT",
    "go_numero_nf": "VARCHAR(30)",
    "go_status_logistica": "VARCHAR(60)",
    "go_data_pedido_expedido": "DATE",
    "go_transportadora_id": "INTEGER",
    "go_custo_frete_previsto": "FLOAT",
    "go_custo_frete_final": "FLOAT",
    "go_custo_frete_sobre_nota": "FLOAT",
    "go_data_prevista_entrega": "DATE",
    "go_data_real_entrega": "DATE",
    # Resultados / OTD
    "go_otd_realizado": "VARCHAR(10)",
    "go_data_solicitada_cliente_final": "DATE",
    "go_data_entregue_cliente": "DATE",
    "go_obs_operacao": "TEXT",
    "go_status_final_alinhamento": "VARCHAR(60)",
}


def _migrar_gestao_operacao_pedidos(app):
    """Adiciona as colunas novas da Fase 13 (Gestão Operação: Comercial / PCP /
    Logística-NF / Resultados-OTD) em pedidos criados antes dessa fase existir.

    Mesmo padrão das migrações anteriores: só adiciona o que ainda não existe,
    roda sozinha a cada início do site, não apaga nem altera nenhum pedido."""
    inspector = inspect(db.engine)
    if "pedidos" not in inspector.get_table_names():
        return

    colunas = {c["name"] for c in inspector.get_columns("pedidos")}
    faltando = [c for c in _COLUNAS_GESTAO_OPERACAO if c not in colunas]
    if not faltando:
        return

    with db.engine.begin() as conn:
        for coluna in faltando:
            tipo = _COLUNAS_GESTAO_OPERACAO[coluna]
            conn.execute(text(f"ALTER TABLE pedidos ADD COLUMN {coluna} {tipo}"))
    app.logger.info("Migração automática: %d campos de Gestão Operação adicionados aos pedidos.", len(faltando))


def _migrar_dados_go_para_pedidos_operacao(app):
    """Backfill único: cria em `pedidos_operacao` (tabela nova, independente de
    `pedidos`) uma linha por pedido comercial, a partir dos dados de Gestão
    Operação que já existem nos `Pedido` legados (tabela de Gestão Produção).

    Motivo: até agora Gestão Operação vivia dentro da tabela `pedidos`, com o
    problema de que um mesmo pedido comercial podia corresponder a vários
    registros Pedido legados (o import histórico original criava 1 Pedido por
    LINHA de planilha, não por pedido comercial — por isso hoje existem 1146+
    Pedido mas só ~385 pedido_venda únicos). A partir de agora Gestão Operação
    e Gestão Produção são independentes: este backfill roda uma única vez
    (protegido por `PedidoOperacao.query.count() == 0`) e agrupa os Pedido
    legados por pedido_venda, usando o de menor id como representante — os
    campos go_* já estavam sincronizados entre eles (ver
    importar_gestao_operacao.py da Fase 13), então não há perda de dado.

    Os campos go_* continuam fisicamente na tabela `pedidos` depois disso
    (nada é apagado) — só o app para de lê-los/escrevê-los por ali."""
    if PedidoOperacao.query.count() > 0:
        return

    pedidos_com_dado_go = (
        Pedido.query.filter(
            or_(*(getattr(Pedido, campo).isnot(None) for campo in _COLUNAS_GESTAO_OPERACAO))
        )
        .order_by(Pedido.id)
        .all()
    )
    if not pedidos_com_dado_go:
        return

    representantes = {}
    ordem = []
    for p in pedidos_com_dado_go:
        chave = p.pedido_venda.strip() if p.pedido_venda else f"__pedido_{p.id}"
        if chave not in representantes:
            representantes[chave] = p  # primeiro da lista (menor id) = representante
            ordem.append(chave)

    for chave in ordem:
        rep = representantes[chave]
        db.session.add(
            PedidoOperacao(
                pedido_venda=rep.pedido_venda,
                cliente=rep.cliente,
                vendedor=rep.vendedor,
                data_inclusao_pedido=rep.data_inclusao_pedido,
                prioridade=rep.prioridade,
                frete=rep.frete,
                pais=rep.pais,
                estado=rep.estado,
                cidade=rep.cidade,
                **{campo: getattr(rep, campo) for campo in _COLUNAS_GESTAO_OPERACAO},
            )
        )

    db.session.commit()
    app.logger.info(
        "Backfill Gestão Operação -> tabela própria (pedidos_operacao): %d pedidos "
        "comerciais migrados (a partir de %d registros legados em `pedidos`).",
        len(ordem), len(pedidos_com_dado_go),
    )


# Marcador de "o Bruno zerou os dados de propósito, pra testar manualmente" —
# ver rota /admin/zerar-dados. Sem isso, tanto _seed_inicial (Pedido vazio)
# quanto _importar_gestao_operacao (PedidoOperacao sem go_tipo_pedido) veriam
# a tabela vazia depois do zerar e reimportariam a planilha antiga sozinhos no
# deploy seguinte — desfazendo a limpeza sem o Bruno pedir de novo.
_CHAVE_DADOS_ZERADOS_MANUALMENTE = "dados_pedidos_zerados_manualmente_28_08_2026"


def _seed_inicial(app):
    """Cria o usuário admin padrão e importa a planilha na primeira execução."""
    if Usuario.query.count() == 0:
        admin = Usuario(nome="Administrador", username="admin", role="ADMIN")
        admin.set_senha("admin123")
        db.session.add(admin)
        db.session.commit()
        app.logger.info("Usuário admin criado (login: admin / senha: admin123 — troque depois!)")

    zerado_manualmente = (
        ControleSistema.query.filter_by(chave=_CHAVE_DADOS_ZERADOS_MANUALMENTE).first() is not None
    )
    if Pedido.query.count() == 0 and not zerado_manualmente:
        xlsx_path = os.path.join(BASE_DIR, "data", "controle_producao_base.xlsx")
        if os.path.exists(xlsx_path):
            from seed import importar_planilha

            total = importar_planilha(xlsx_path)
            app.logger.info(f"{total} pedidos importados da planilha original.")


def _importar_gestao_operacao(app):
    """Importa (uma única vez) a planilha "Gestão de Fluxo Produtivo" pra dentro
    de `pedidos_operacao` (Gestão Operação — tabela própria, independente de
    `pedidos`/Gestão Produção), casando por pedido_venda.

    Roda DEPOIS de `_migrar_dados_go_para_pedidos_operacao` (o backfill a
    partir dos dados legados) — na prática isso quase sempre já deixa
    `go_tipo_pedido` preenchido em `pedidos_operacao`, então esta função só
    chega a importar de verdade da planilha se o backfill não tiver rodado
    (ex.: banco novo, sem nenhum pedido legado com dado de Gestão Operação)."""
    if PedidoOperacao.query.filter(PedidoOperacao.go_tipo_pedido.isnot(None)).first() is not None:
        return
    if ControleSistema.query.filter_by(chave=_CHAVE_DADOS_ZERADOS_MANUALMENTE).first() is not None:
        return

    xlsx_path = os.path.join(BASE_DIR, "data", "gestao_fluxo_2026.xlsx")
    if not os.path.exists(xlsx_path):
        return

    from importar_gestao_operacao import importar_gestao_operacao

    resultado = importar_gestao_operacao(xlsx_path)
    app.logger.info(
        "Importação Gestão Operação: %d linhas | %d exatas | %d aproximadas | "
        "%d pedidos novos | %d transportadoras.",
        resultado["total"],
        resultado["exato"],
        resultado["aproximado"],
        resultado["novos"],
        len(resultado["transportadoras_canonicas"]),
    )


_CHAVE_SINCRONIZACAO_25_08_2026 = "sincronizacao_planilha_producao_25_08_2026"


def _sincronizar_planilha_producao_25_08_2026(app):
    """Sincroniza Gestão Produção (Pedido/ItemPedido) com a aba "GERAL TESTE"
    da planilha enviada pelo Bruno em 25/08/2026 — atualiza pedidos/itens que
    já existem e cria os que estão na planilha mas ainda não existem no site.
    Protegido por `ControleSistema` (chave abaixo) porque, ao contrário das
    outras importações, roda por cima de dados que já existem — precisa
    rodar exatamente uma vez, mesmo com o banco de produção já povoado. Ver
    sincronizar_planilha_producao.py para as regras completas."""
    if ControleSistema.query.filter_by(chave=_CHAVE_SINCRONIZACAO_25_08_2026).first() is not None:
        return

    xlsx_path = os.path.join(BASE_DIR, "data", "sincronizacao_25_08_2026.xlsx")
    if not os.path.exists(xlsx_path):
        return

    from sincronizar_planilha_producao import sincronizar_planilha_producao

    stats = sincronizar_planilha_producao(xlsx_path)
    db.session.add(ControleSistema(chave=_CHAVE_SINCRONIZACAO_25_08_2026))
    db.session.commit()
    app.logger.info(
        "Sincronização planilha 25/08/2026: %d linhas | %d pedidos atualizados | "
        "%d pedidos criados | %d itens atualizados | %d itens criados | "
        "%d pedidos sem cliente ignorados.",
        stats["linhas_lidas"],
        stats["pedidos_atualizados"],
        stats["pedidos_criados"],
        stats["itens_atualizados"],
        stats["itens_criados"],
        len(stats["pedidos_sem_cliente_ignorados"]),
    )


_CHAVE_SINCRONIZACAO_GO_28_08_2026 = "sincronizacao_gestao_operacao_28_08_2026"


def _sincronizar_gestao_operacao_28_08_2026(app):
    """Sincroniza Gestão Operação (PedidoOperacao) com a planilha "28_08 Gestão
    de Fluxo Produtivo 2026" enviada pelo Bruno — atualiza pedidos que já
    existem e cria os que estão na planilha mas ainda não existem no site.
    Protegido por `ControleSistema` (mesmo padrão de
    `_sincronizar_planilha_producao_25_08_2026`) porque roda por cima de dados
    que já existem — precisa rodar exatamente uma vez, mesmo com o banco de
    produção já povoado. Ver sincronizar_gestao_operacao.py para as regras
    completas."""
    if ControleSistema.query.filter_by(chave=_CHAVE_SINCRONIZACAO_GO_28_08_2026).first() is not None:
        return

    xlsx_path = os.path.join(BASE_DIR, "data", "sincronizacao_gestao_operacao_28_08_2026.xlsx")
    if not os.path.exists(xlsx_path):
        return

    from sincronizar_gestao_operacao import sincronizar_gestao_operacao

    stats = sincronizar_gestao_operacao(xlsx_path)
    db.session.add(ControleSistema(chave=_CHAVE_SINCRONIZACAO_GO_28_08_2026))
    db.session.commit()
    app.logger.info(
        "Sincronização Gestão Operação 28/08/2026: %d linhas | %d pedidos atualizados | "
        "%d pedidos criados | %d campos atualizados | %d casamentos exatos | %d aproximados.",
        stats["linhas_lidas"],
        stats["pedidos_atualizados"],
        stats["pedidos_criados"],
        stats["campos_atualizados"],
        stats["exato"],
        stats["aproximado"],
    )


_CHAVE_CORRECAO_SEMANA_MORKEN_835 = "correcao_semana_pcp_morken_835_28_08_2026"


def _corrigir_semana_pcp_morken_835(app):
    """Correção pontual pedida pelo Bruno em 28/08/2026: o pedido 835 (Morken)
    veio da planilha com Término Semanal PCP "SEMANA 04 / AGO / 2026", mas a
    Previsão de Liberação PCP dele é 04/09/2026 (entrega solicitada 24/09) —
    ele não é um pedido de agosto. O Bruno confirmou que é pra mover pra
    setembro, então troco pra "SEMANA 01 / SET / 2026" (semana que contém o
    dia 4, mesmo critério de gerar_semanas_pcp). Com isso ele some do
    Faturamento por Semana de agosto e passa a aparecer em setembro.

    Guardado por ControleSistema — correção pontual, roda só uma vez, mesmo
    padrão das outras correções/sincronizações desta leva."""
    if ControleSistema.query.filter_by(chave=_CHAVE_CORRECAO_SEMANA_MORKEN_835).first() is not None:
        return

    pedido = PedidoOperacao.query.filter_by(pedido_venda="835").first()
    if pedido is not None and pedido.cliente and "MORKEN" in pedido.cliente.upper():
        pedido.go_termino_semanal_pcp = "SEMANA 01 / SET / 2026"
        app.logger.info(
            "Correção pontual: pedido 835 (Morken) movido de Semana 04/Ago para Semana 01/Set/2026."
        )
    else:
        # Não achou o pedido esperado (ou o cliente não bate) — não mexe em
        # nada pra não arriscar corrigir o pedido errado, só registra que já
        # tentou (pra não ficar reavaliando isso a cada boot).
        app.logger.warning(
            "Correção pontual pedido 835 (Morken): pedido não encontrado ou cliente não confere — nada foi alterado."
        )
    db.session.add(ControleSistema(chave=_CHAVE_CORRECAO_SEMANA_MORKEN_835))
    db.session.commit()


_CHAVE_SEED_RNC_QUALIDADE_31_08_2026 = "seed_rnc_qualidade_31_08_2026"


def _seed_rnc_qualidade(app):
    """Importa (uma única vez) a planilha "Controle RNC - Qualidade" que o
    Bruno enviou em 31/08/2026 pra dentro da tabela `rnc_qualidade`, nova
    (nasce vazia — `db.create_all()` já criou a tabela nesta mesma execução
    do boot, antes desta função rodar).

    Guardado por `ControleSistema` (mesmo padrão das outras importações/
    sincronizações pontuais) em vez de só checar "tabela vazia": assim, se o
    Bruno apagar algum RNC de teste depois, o próximo boot não reimporta os
    10 RNCs de agosto por cima — roda exatamente uma vez, para sempre."""
    if ControleSistema.query.filter_by(chave=_CHAVE_SEED_RNC_QUALIDADE_31_08_2026).first() is not None:
        return

    xlsx_path = os.path.join(BASE_DIR, "data", "rnc_qualidade_31_08_2026.xlsx")
    if not os.path.exists(xlsx_path):
        return

    from seed_rnc_qualidade import importar_rnc_qualidade

    total = importar_rnc_qualidade(xlsx_path)
    db.session.add(ControleSistema(chave=_CHAVE_SEED_RNC_QUALIDADE_31_08_2026))
    db.session.commit()
    app.logger.info("Importação Qualidade/RNC: %d RNCs importados da planilha 31/08/2026.", total)


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


def _resumo_otd():
    """Estatísticas de OTD (On-Time Delivery) da Gestão Operação — usa o campo
    go_otd_realizado (SIM/NÃO preenchido manualmente na planilha/tela), bem
    mais confiável que o _otd_percentual() antigo (que depende de datas quase
    nunca preenchidas no histórico). Só considera pedidos com OTD preenchido —
    quem ainda não tem essa informação fica de fora do percentual (não conta
    como "não cumpriu").

    Gestão Operação tem tabela própria (PedidoOperacao) — cada linha já é um
    pedido comercial, sem precisar agrupar nada em tempo de execução."""
    pedidos = PedidoOperacao.query.filter(PedidoOperacao.go_otd_realizado.isnot(None)).all()

    total = len(pedidos)
    no_prazo = sum(1 for p in pedidos if p.go_otd_realizado == "SIM")
    percentual = round((no_prazo / total) * 100, 1) if total else None

    def _quebra_por(atributo):
        contagem = {}
        for p in pedidos:
            chave = getattr(p, atributo)
            if not chave:
                continue
            c = contagem.setdefault(chave, {"total": 0, "no_prazo": 0})
            c["total"] += 1
            if p.go_otd_realizado == "SIM":
                c["no_prazo"] += 1
        top10 = sorted(contagem.items(), key=lambda kv: -kv[1]["total"])[:10]
        return [
            {
                "chave": chave,
                "total": v["total"],
                "no_prazo": v["no_prazo"],
                "percentual": round((v["no_prazo"] / v["total"]) * 100, 1) if v["total"] else None,
            }
            for chave, v in top10
        ]

    return {
        "total": total,
        "no_prazo": no_prazo,
        "percentual": percentual,
        "atinge_meta": (percentual is not None and percentual >= GO_OTD_META_PERCENTUAL),
        "por_vendedor": _quebra_por("vendedor"),
        "por_cliente": _quebra_por("cliente"),
    }


def _semana_label_curto(semana):
    """Encurta "SEMANA 03 / AGO / 2026" pra "S03 AGO/26" (rótulo do eixo do
    gráfico de projeção) — valores antigos que não seguem esse padrão (ex.:
    "DEZEMBRO/2025", digitado à mão antes desta convenção existir) aparecem
    como estão, sem tentar encurtar."""
    m = re.match(r"SEMANA (\d{2}) / (\w{3}) / (\d{4})", semana or "")
    if not m:
        return semana
    semana_num, mes, ano = m.groups()
    return f"S{semana_num} {mes}/{ano[2:]}"


def _projecao_pcp():
    """Projeção de carga por semana de PCP (Painel) — quantos PEDIDOS
    COMERCIAIS estão marcados pra cada "Término Semanal PCP" (Gestão
    Operação), separando o que já foi finalizado do que ainda está em aberto.
    Usa a mesma lista/ordem cronológica de gerar_semanas_pcp() (1 mês atrás
    até 6 meses à frente) — semanas sem nenhum pedido nas pontas são cortadas
    pra não poluir o gráfico."""
    semanas = gerar_semanas_pcp()

    pedidos = PedidoOperacao.query.filter(PedidoOperacao.go_termino_semanal_pcp.in_(semanas)).all()

    contagem = {}
    for p in pedidos:
        c = contagem.setdefault(p.go_termino_semanal_pcp, {"finalizado": 0, "em_aberto": 0})
        if p.status_producao == "FINALIZADO":
            c["finalizado"] += 1
        else:
            c["em_aberto"] += 1

    linhas = [
        {
            "semana": s,
            "semana_curta": _semana_label_curto(s),
            "finalizado": contagem.get(s, {}).get("finalizado", 0),
            "em_aberto": contagem.get(s, {}).get("em_aberto", 0),
        }
        for s in semanas
    ]
    while linhas and not linhas[0]["finalizado"] and not linhas[0]["em_aberto"]:
        linhas.pop(0)
    while linhas and not linhas[-1]["finalizado"] and not linhas[-1]["em_aberto"]:
        linhas.pop()
    return linhas


# Nomes de mês (abreviados e por extenso) que já apareceram no campo Término
# Semanal PCP — cobre tanto o formato novo ("SEMANA 04 / JUL / 2026") quanto
# valores antigos digitados à mão antes dessa convenção existir ("DEZEMBRO/2025").
_MESES_NUM_PCP = {
    "JAN": 1, "JANEIRO": 1,
    "FEV": 2, "FEVEREIRO": 2,
    "MAR": 3, "MARÇO": 3, "MARCO": 3,
    "ABR": 4, "ABRIL": 4,
    "MAI": 5, "MAIO": 5,
    "JUN": 6, "JUNHO": 6,
    "JUL": 7, "JULHO": 7,
    "AGO": 8, "AGOSTO": 8,
    "SET": 9, "SETEMBRO": 9,
    "OUT": 10, "OUTUBRO": 10,
    "NOV": 11, "NOVEMBRO": 11,
    "DEZ": 12, "DEZEMBRO": 12,
}


def _mes_ano_da_semana_pcp(semana):
    """Extrai (ano, mês) de dentro do valor de Término Semanal PCP, não
    importa se é "SEMANA 04 / JUL / 2026" (formato novo) ou "DEZEMBRO/2025"
    (formato antigo, digitado à mão) — devolve None se não conseguir entender."""
    if not semana:
        return None
    m = re.search(r"([A-ZÇÃÕ]+)\s*/\s*(\d{4})", semana.upper())
    if not m:
        return None
    nome_mes, ano = m.groups()
    mes = _MESES_NUM_PCP.get(nome_mes)
    if not mes:
        return None
    return (int(ano), mes)


def _chave_semana_pcp(semana):
    """Chave de ordenação cronológica pro rótulo de semana (ano, mês, nº da
    semana) — usada tanto pelo planejamento semanal da Listagem Geral quanto,
    se precisar no futuro, por qualquer outro campo no mesmo formato."""
    if not semana:
        return None
    mes_ano = _mes_ano_da_semana_pcp(semana)
    m = re.search(r"SEMANA\s*(\d+)", semana.upper())
    semana_num = int(m.group(1)) if m else 0
    if mes_ano is None:
        return (9999, 99, semana_num)
    ano, mes = mes_ano
    return (ano, mes, semana_num)


def _somar_meses(ano, mes, delta):
    total = (ano * 12 + (mes - 1)) + delta
    return (total // 12, total % 12 + 1)


# ---------------------------------------------------------------------------
# Calendário PCP da tela de Programação (pedido do Bruno, 31/08/2026) — semana
# de verdade (domingo a sábado), só pra esta tela. NÃO usa nem mexe no padrão
# "SEMANA NN / MÊS / ANO" (dia 1-7, 8-14...) usado em Gestão Operação,
# Faturamento por Semana e no Planejamento Semanal da Listagem Geral — esses
# três continuam exatamente como estão, comparando pelo texto já gravado nos
# pedidos importados da planilha. Misturar as duas convenções faria pedido
# sumir de semana sem ninguém perceber (testado: quase metade dos meses tem
# número de semanas diferente entre os dois padrões), por isso são
# propositalmente independentes.
def _semanas_calendario_pcp(ano, mes):
    """Devolve as semanas (domingo a sábado) cujo domingo de início cai no
    mês/ano dado, cada uma como {"numero", "inicio", "fim"}. Uma semana pode
    terminar no mês seguinte (ex.: semana 5 de agosto/2026 vai até 05/09) —
    isso é esperado, é só o card ficando "encostado" no mês seguinte."""
    primeiro_dia = date(ano, mes, 1)
    dias_ate_domingo = (6 - primeiro_dia.weekday()) % 7  # weekday(): segunda=0 ... domingo=6
    domingo = primeiro_dia + timedelta(days=dias_ate_domingo)
    semanas = []
    numero = 1
    while domingo.year == ano and domingo.month == mes:
        semanas.append({"numero": numero, "inicio": domingo, "fim": domingo + timedelta(days=6)})
        numero += 1
        domingo += timedelta(days=7)
    return semanas


def _parse_mes_ano_form(valor, default):
    """Lê um <input type=month> (formato "YYYY-MM") — devolve `default` se
    vier vazio ou num formato que não reconhece."""
    if not valor:
        return default
    try:
        ano_s, mes_s = valor.split("-")
        ano, mes = int(ano_s), int(mes_s)
        if 1 <= mes <= 12:
            return (ano, mes)
    except (ValueError, AttributeError):
        pass
    return default


def _projecao_pcp_mensal(mes_de, mes_ate):
    """Soma a projeção semanal de PCP por MÊS (soma de todas as semanas
    dentro do mês), separando o que já foi finalizado do que ainda está em
    aberto — em quantidade de PEDIDOS COMERCIAIS e em valor (R$, a partir de
    go_valor_pedido_operacao — já é o total do pedido, não soma nenhum item).
    `mes_de`/`mes_ate` são tuplas (ano, mês), intervalo fechado."""
    pedidos = PedidoOperacao.query.filter(PedidoOperacao.go_termino_semanal_pcp.isnot(None)).all()

    baldes = {}
    for p in pedidos:
        chave = _mes_ano_da_semana_pcp(p.go_termino_semanal_pcp)
        if chave is None or not (mes_de <= chave <= mes_ate):
            continue
        b = baldes.setdefault(chave, {"pedidos_fin": 0, "valor_fin": 0.0, "pedidos_aberto": 0, "valor_aberto": 0.0})
        valor = p.go_valor_pedido_operacao or 0.0
        if p.status_producao == "FINALIZADO":
            b["pedidos_fin"] += 1
            b["valor_fin"] += valor
        else:
            b["pedidos_aberto"] += 1
            b["valor_aberto"] += valor

    linhas = []
    ano, mes = mes_de
    while (ano, mes) <= mes_ate:
        b = baldes.get((ano, mes), {"pedidos_fin": 0, "valor_fin": 0.0, "pedidos_aberto": 0, "valor_aberto": 0.0})
        linhas.append(
            {
                "ano": ano,
                "mes": mes,
                "label": f"{MESES_PT[mes - 1]}/{ano}",
                "pedidos_finalizados": b["pedidos_fin"],
                "valor_finalizado": round(b["valor_fin"], 2),
                "pedidos_em_aberto": b["pedidos_aberto"],
                "valor_em_aberto": round(b["valor_aberto"], 2),
                "pedidos_total": b["pedidos_fin"] + b["pedidos_aberto"],
                "valor_total_mes": round(b["valor_fin"] + b["valor_aberto"], 2),
            }
        )
        ano, mes = _somar_meses(ano, mes, 1)
    return linhas


def _faturamento_semanal_pcp(ano, mes):
    """Faturamento de Agosto (ou qualquer mês) por semana de PCP — pedido do
    Bruno em 28/08/2026.

    Revisado em 28/08/2026 depois que o Bruno conferiu os números da semana
    04/Ago manualmente: a versão original desta função exigia Data Efetiva de
    Liberação PCP preenchida pra contar "Qtd/Valor liberado" (replicando a
    fórmula SUMPRODUCT/IFERROR que eu tinha lido da planilha) — mas o Bruno
    confirmou que quer TODO pedido cujo Término Semanal PCP cai naquela
    semana, tenha ele já sido liberado ou não (ver AskUserQuestion — "Somar
    todos os pedidos da semana, liberados ou não"). Ou seja, "liberado" aqui
    não significa "já com Data Efetiva de Liberação preenchida", e sim
    "planejado pro PCP encerrar naquela semana" — mesmo agrupamento por
    Término Semanal PCP (coluna AA), só que sem o filtro extra que eu tinha
    adicionado por conta própria.

      - Qtd/Valor liberado: todo pedido cujo Término Semanal PCP
        (go_termino_semanal_pcp) cai nessa semana — valor é a soma de
        go_valor_pedido_operacao (valor total do pedido).
      - Qtd/Valor faturado: pedidos cujo Término Semanal PCP cai nessa
        semana, somando go_valor_nf_emitida.

    Diferença deliberada da planilha do Bruno: lá, valores de "VALOR NF
    EMITIDA" digitados em formato brasileiro com vírgula decimal (texto, não
    número) são silenciosamente zerados pelo IFERROR(...*1, 0) da fórmula
    dele — aqui esses valores são interpretados corretamente como número
    (mesmo parser usado no resto do site, ver _parse_numero em
    importar_gestao_operacao.py), então o total pode ficar um pouco MAIOR
    que o da planilha nesses meses com célula de texto — reportado ao Bruno
    junto com a entrega."""
    semanas = gerar_semanas_pcp(meses_atras=0, meses_frente=0, hoje=date(ano, mes, 1))
    pedidos = PedidoOperacao.query.filter(PedidoOperacao.go_termino_semanal_pcp.in_(semanas)).all()

    baldes = {
        s: {"qtd_liberada": 0, "valor_liberado": 0.0, "qtd_faturada": 0, "valor_faturado": 0.0}
        for s in semanas
    }
    for p in pedidos:
        b = baldes[p.go_termino_semanal_pcp]
        b["qtd_liberada"] += 1
        b["valor_liberado"] += p.go_valor_pedido_operacao or 0.0
        if p.go_valor_nf_emitida:
            b["qtd_faturada"] += 1
            b["valor_faturado"] += p.go_valor_nf_emitida

    linhas = [
        {
            "semana": s,
            "semana_curta": _semana_label_curto(s),
            "qtd_liberada": baldes[s]["qtd_liberada"],
            "valor_liberado": round(baldes[s]["valor_liberado"], 2),
            "qtd_faturada": baldes[s]["qtd_faturada"],
            "valor_faturado": round(baldes[s]["valor_faturado"], 2),
        }
        for s in semanas
    ]
    totais = {
        "qtd_liberada": sum(l["qtd_liberada"] for l in linhas),
        "valor_liberado": round(sum(l["valor_liberado"] for l in linhas), 2),
        "qtd_faturada": sum(l["qtd_faturada"] for l in linhas),
        "valor_faturado": round(sum(l["valor_faturado"] for l in linhas), 2),
    }
    return {"linhas": linhas, "totais": totais}


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


def _faturamento_mes(ano, mes, filtros=None):
    """Soma o valor dos itens com liberação PREVISTA (usa quantidade × custo)
    e com liberação REALIZADA (usa o valor faturado de verdade quando
    preenchido, senão cai pro quantidade × custo) dentro de um mês — usado no
    previsto × realizado. `filtros` (opcional) é uma lista de condições extras
    (ex.: cliente/região/vendedor) aplicadas a ambas as somas via join com Pedido."""
    inicio = date(ano, mes, 1)
    fim = date(ano + 1, 1, 1) if mes == 12 else date(ano, mes + 1, 1)
    valor_calculado = ItemPedido.quantidade * ItemPedido.custo_unitario
    valor_realizado = func.coalesce(ItemPedido.valor_faturado, valor_calculado)
    filtros = filtros or []

    previsto = (
        db.session.query(func.sum(valor_calculado))
        .join(Pedido, ItemPedido.pedido_id == Pedido.id)
        .filter(ItemPedido.liberacao_prevista >= inicio, ItemPedido.liberacao_prevista < fim, *filtros)
        .scalar()
        or 0.0
    )
    realizado = (
        db.session.query(func.sum(valor_realizado))
        .join(Pedido, ItemPedido.pedido_id == Pedido.id)
        .filter(ItemPedido.liberacao_faturamento >= inicio, ItemPedido.liberacao_faturamento < fim, *filtros)
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


def _lead_time_por_estacao(desde=None):
    """Lead time médio (dias) agrupado por estação, entre os itens que já têm
    início de produção e término de inspeção preenchidos."""
    query = ItemPedido.query.filter(
        ItemPedido.inicio_producao.isnot(None), ItemPedido.termino_inspecao.isnot(None), ItemPedido.estacao.isnot(None)
    )
    if desde:
        query = query.filter(ItemPedido.termino_inspecao >= desde)
    agrupado = {}
    for item in query.all():
        agrupado.setdefault(item.estacao, []).append((item.termino_inspecao - item.inicio_producao).days)

    resultado = [
        {"estacao": estacao, "lt_medio": round(sum(dias) / len(dias), 1), "quantidade": len(dias)}
        for estacao, dias in agrupado.items()
    ]
    resultado.sort(key=lambda r: r["lt_medio"], reverse=True)
    return resultado


def _otd_por_vendedor(desde=None):
    """OTD (% no prazo) agrupado por vendedor, entre os itens finalizados que
    têm liberação prevista e término de inspeção preenchidos."""
    query = ItemPedido.query.options(selectinload(ItemPedido.pedido)).filter(
        ItemPedido.status_producao == "FINALIZADO",
        ItemPedido.liberacao_prevista.isnot(None),
        ItemPedido.termino_inspecao.isnot(None),
    )
    if desde:
        query = query.filter(ItemPedido.termino_inspecao >= desde)

    agrupado = {}
    for item in query.all():
        vendedor = (item.pedido.vendedor if item.pedido and item.pedido.vendedor else "Sem vendedor")
        grupo = agrupado.setdefault(vendedor, {"no_prazo": 0, "total": 0})
        grupo["total"] += 1
        if item.termino_inspecao <= item.liberacao_prevista:
            grupo["no_prazo"] += 1

    resultado = [
        {"vendedor": vendedor, "otd": round(100 * g["no_prazo"] / g["total"], 1), "total": g["total"]}
        for vendedor, g in agrupado.items()
    ]
    resultado.sort(key=lambda r: r["otd"])
    return resultado


def _tendencia_kpis(meses=6):
    """Finalizados / lead time médio / OTD por mês, dos últimos N meses — para
    os gráficos de tendência da tela de KPIs."""
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
        inicio = date(ano, mes, 1)
        fim = date(ano + 1, 1, 1) if mes == 12 else date(ano, mes + 1, 1)
        itens = ItemPedido.query.filter(
            ItemPedido.termino_inspecao.isnot(None),
            ItemPedido.termino_inspecao >= inicio,
            ItemPedido.termino_inspecao < fim,
        ).all()

        lt_medio, otd = None, None
        if itens:
            lts = [(i.termino_inspecao - i.inicio_producao).days for i in itens if i.inicio_producao]
            if lts:
                lt_medio = round(sum(lts) / len(lts), 1)
            com_prazo = [i for i in itens if i.liberacao_prevista]
            if com_prazo:
                no_prazo = sum(1 for i in com_prazo if i.termino_inspecao <= i.liberacao_prevista)
                otd = round(100 * no_prazo / len(com_prazo), 1)

        resultado.append({"mes": f"{MESES_PT[mes - 1]}/{ano}", "finalizados": len(itens), "lt_medio": lt_medio, "otd": otd})
    return resultado


def _gargalos_por_estacao():
    """Ranking de estações por "quanto está travado ali": fila, atraso, tempo
    de espera médio, lead time médio e valor parado (não finalizado).

    Consultado estação por estação (poucas estações, poucas linhas cada) em vez
    de carregar TODOS os pedidos em Python — o mesmo cuidado de performance já
    aplicado na tela inicial (fase 3) e na visão geral de Estações (fase 5)."""
    hoje = date.today()
    estacoes = Estacao.query.filter_by(ativo=True).order_by(Estacao.ordem_exibicao).all()
    resultado = []

    for e in estacoes:
        nome = e.nome
        abertos = ItemPedido.query.filter(ItemPedido.estacao == nome, ItemPedido.status_producao != "FINALIZADO")
        fila = abertos.count()
        atraso = abertos.filter(
            ItemPedido.liberacao_prevista.isnot(None), ItemPedido.liberacao_prevista < hoje
        ).count()
        valor_parado = (
            db.session.query(func.sum(ItemPedido.quantidade * ItemPedido.custo_unitario))
            .filter(ItemPedido.estacao == nome, ItemPedido.status_producao != "FINALIZADO")
            .scalar()
            or 0.0
        )

        itens_com_inicio = (
            ItemPedido.query.options(selectinload(ItemPedido.pedido))
            .filter(ItemPedido.estacao == nome, ItemPedido.inicio_producao.isnot(None))
            .all()
        )
        esperas = [
            (i.inicio_producao - i.pedido.data_inclusao_pedido).days
            for i in itens_com_inicio
            if i.pedido and i.pedido.data_inclusao_pedido
        ]
        tempo_espera_medio = round(sum(esperas) / len(esperas), 1) if esperas else None

        lts = [(i.termino_inspecao - i.inicio_producao).days for i in itens_com_inicio if i.termino_inspecao]
        lt_medio = round(sum(lts) / len(lts), 1) if lts else None

        resultado.append(
            {
                "estacao": nome,
                "fila": fila,
                "atraso": atraso,
                "tempo_espera_medio": tempo_espera_medio,
                "lt_medio": lt_medio,
                "valor_parado": round(valor_parado, 2),
            }
        )

    resultado.sort(key=lambda r: (r["fila"], r["atraso"]), reverse=True)
    return resultado


def _faturamento_detalhado(ano, mes, cliente=None, regiao=None, vendedor=None):
    """Previsto × realizado de um mês específico, com o mesmo valor já
    quebrado por cliente / região / vendedor — usado na tela de Faturamento."""
    inicio = date(ano, mes, 1)
    fim = date(ano + 1, 1, 1) if mes == 12 else date(ano, mes + 1, 1)

    base = ItemPedido.query.options(selectinload(ItemPedido.pedido)).join(Pedido, ItemPedido.pedido_id == Pedido.id)
    if cliente:
        base = base.filter(Pedido.cliente.ilike(f"%{cliente}%"))
    if vendedor:
        base = base.filter(Pedido.vendedor.ilike(f"%{vendedor}%"))
    if regiao:
        ufs_da_regiao = [uf for uf, r in REGIAO_POR_UF.items() if r == regiao]
        if ufs_da_regiao:
            base = base.filter(Pedido.estado.in_(ufs_da_regiao))

    itens_previstos = base.filter(ItemPedido.liberacao_prevista >= inicio, ItemPedido.liberacao_prevista < fim).all()
    itens_realizados = base.filter(
        ItemPedido.liberacao_faturamento >= inicio, ItemPedido.liberacao_faturamento < fim
    ).all()

    def _agrupar(itens, chave_fn):
        agrupado = {}
        for i in itens:
            chave = chave_fn(i) or "—"
            agrupado[chave] = agrupado.get(chave, 0.0) + i.valor_faturamento_realizado
        linhas = [{"chave": k, "valor": round(v, 2)} for k, v in agrupado.items()]
        linhas.sort(key=lambda l: l["valor"], reverse=True)
        return linhas

    return {
        "previsto_total": round(sum(i.valor_total for i in itens_previstos), 2),
        "realizado_total": round(sum(i.valor_faturamento_realizado for i in itens_realizados), 2),
        "itens_realizados": len(itens_realizados),
        "por_cliente": _agrupar(itens_realizados, lambda i: i.pedido.cliente if i.pedido else None),
        "por_regiao": _agrupar(itens_realizados, lambda i: REGIAO_POR_UF.get(i.pedido.estado) if i.pedido and i.pedido.estado else None),
        "por_vendedor": _agrupar(itens_realizados, lambda i: i.pedido.vendedor if i.pedido else None),
        # lista "crua" dos itens realizados (não usada na tela, só na exportação
        # de relatório — mantida separada da contagem "itens_realizados" acima
        # pra não mudar o que a tela de Faturamento já espera receber)
        "itens_realizados_lista": itens_realizados,
    }


def _faturamento_previsto_nao_realizado():
    """Itens cuja liberação PREVISTA já passou mas a liberação de FATURAMENTO
    ainda não aconteceu — ou seja, o faturamento que era esperado até agora
    ainda não se realizou (mesmo conceito de "previsto" usado na tela de
    Faturamento, só que olhando pro atraso em vez de somar por mês)."""
    hoje = date.today()
    return (
        ItemPedido.query.options(selectinload(ItemPedido.pedido))
        .join(Pedido, ItemPedido.pedido_id == Pedido.id)
        .filter(
            ItemPedido.liberacao_prevista.isnot(None),
            ItemPedido.liberacao_prevista < hoje,
            ItemPedido.liberacao_faturamento.is_(None),
        )
        .order_by(ItemPedido.liberacao_prevista)
        .all()
    )


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


def _filtrar_pedidos(args):
    """Aplica exatamente os mesmos filtros da tela de Listagem (rota "/") a
    partir de um dict tipo request.args, devolvendo a query já filtrada (sem
    paginação) e o dicionário de filtros usado.

    Extraído da rota "dashboard" pra ser compartilhado com a exportação de
    relatório (fase 12) — assim a exportação nunca corre o risco de aplicar
    uma regra de filtro diferente da que a tela usa."""
    query = Pedido.query.options(selectinload(Pedido.itens))

    cliente = args.get("cliente", "").strip()
    status = args.get("status", "").strip()
    estacao = args.get("estacao", "").strip()
    vendedor = args.get("vendedor", "").strip()
    busca = args.get("busca", "").strip()
    produto = args.get("produto", "").strip()
    regiao = args.get("regiao", "").strip()
    data_inicio = args.get("data_inicio", "").strip()
    data_fim = args.get("data_fim", "").strip()
    atrasados = args.get("atrasados", "").strip()
    planejamento_semanal = args.get("planejamento_semanal", "").strip()
    planejamento_mensal = args.get("planejamento_mensal", "").strip()
    mes_inclusao = args.get("mes_inclusao", "").strip()
    mes_entrega_cliente = args.get("mes_entrega_cliente", "").strip()

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
    if planejamento_semanal:
        query = query.filter(Pedido.itens.any(ItemPedido.planejamento_semanal == planejamento_semanal))
    if planejamento_mensal:
        mes_ano = _parse_mes_ano_form(planejamento_mensal, None)
        if mes_ano:
            semanas_do_mes = [
                s for (s,) in db.session.query(ItemPedido.planejamento_semanal)
                .filter(ItemPedido.planejamento_semanal.isnot(None))
                .distinct()
                if _mes_ano_da_semana_pcp(s) == mes_ano
            ]
            if semanas_do_mes:
                query = query.filter(Pedido.itens.any(ItemPedido.planejamento_semanal.in_(semanas_do_mes)))
            else:
                # Mês escolhido não tem nenhum planejamento semanal preenchido
                # ainda — não deve mostrar nada (em vez de ignorar o filtro).
                query = query.filter(false())
    if mes_inclusao:
        mes_ano = _parse_mes_ano_form(mes_inclusao, None)
        if mes_ano:
            ano, mes = mes_ano
            query = query.filter(
                extract("year", Pedido.data_inclusao_pedido) == ano,
                extract("month", Pedido.data_inclusao_pedido) == mes,
            )
        else:
            query = query.filter(false())
    if mes_entrega_cliente:
        mes_ano = _parse_mes_ano_form(mes_entrega_cliente, None)
        if mes_ano:
            ano, mes = mes_ano
            query = query.filter(
                extract("year", Pedido.data_cliente) == ano,
                extract("month", Pedido.data_cliente) == mes,
            )
        else:
            query = query.filter(false())

    query = query.order_by(Pedido.data_inclusao_pedido.desc().nullslast(), Pedido.id.desc())

    filtros = dict(
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
        planejamento_semanal=planejamento_semanal,
        planejamento_mensal=planejamento_mensal,
        mes_inclusao=mes_inclusao,
        mes_entrega_cliente=mes_entrega_cliente,
    )
    return query, filtros


class _LinhaListagemGeral:
    """Uma linha da Listagem Geral = 1 pedido + 1 item (produto) dele — o
    mesmo número de pedido pode aparecer em várias linhas, uma por produto
    distinto, igual à planilha de referência. Só une os dois objetos num
    lugar só pra o template não precisar fazer `linha.pedido.x` /
    `linha.item.y` o tempo todo."""

    def __init__(self, pedido, item):
        self.pedido = pedido
        self.item = item

    # ---- identidade do pedido ----
    @property
    def pedido_id(self):
        return self.pedido.id

    @property
    def pedido_venda(self):
        return self.pedido.pedido_venda

    @property
    def cliente(self):
        return self.pedido.cliente

    @property
    def vendedor(self):
        return self.pedido.vendedor

    @property
    def data_inclusao_pedido(self):
        return self.pedido.data_inclusao_pedido

    @property
    def data_cliente(self):
        """Data solicitada pelo cliente (prazo comercial) — já existia no
        pedido (campo "Data do cliente" na tela de editar), só não aparecia
        na Listagem Geral."""
        return self.pedido.data_cliente

    @property
    def prioridade(self):
        return self.pedido.prioridade

    @property
    def frete(self):
        return self.pedido.frete

    @property
    def pais(self):
        return self.pedido.pais

    @property
    def estado(self):
        return self.pedido.estado

    @property
    def cidade(self):
        return self.pedido.cidade

    # ---- dados do item (produto) ----
    @property
    def item_id(self):
        return self.item.id

    @property
    def descricao_produto(self):
        return self.item.descricao_produto

    @property
    def quantidade(self):
        return self.item.quantidade

    @property
    def venda_unidade(self):
        """Mesmo dado de custo que já existe (custo_unitario) — só exibido
        sob o rótulo "Venda" pedido pelo Bruno."""
        return self.item.custo_unitario

    @property
    def venda_total(self):
        return self.item.valor_total

    @property
    def estacao(self):
        return self.item.estacao

    @property
    def status_producao(self):
        return self.item.status_producao

    @property
    def liberacao_prevista(self):
        return self.item.liberacao_prevista

    @property
    def liberacao_real(self):
        return self.item.liberacao_real

    @property
    def planejamento_semanal(self):
        return self.item.planejamento_semanal

    @property
    def venda_total_pedido(self):
        """Soma o custo de TODOS os itens do pedido (não só deste produto) —
        reaproveita Pedido.valor_total, que já faz exatamente essa soma."""
        return self.pedido.valor_total

    @property
    def semaforo(self):
        return self.item.semaforo


def _linhas_listagem_geral(pedidos, args):
    """Achata a lista de Pedido (com itens já carregados) em 1 linha por
    ItemPedido — a granularidade que a Listagem Geral usa agora (1 linha por
    produto, igual ao print de referência).

    `_filtrar_pedidos` já decidiu quais PEDIDOS entram (algum item bate o
    filtro); aqui, pros filtros que são naturalmente por ITEM — estação,
    produto, planejamento semanal/mensal —, mostra só os produtos que batem,
    não o pedido inteiro. Sem isso, filtrar por "semana X" mostraria também
    os outros produtos do mesmo pedido que caem em semanas diferentes."""
    estacao = args.get("estacao", "").strip()
    produto = args.get("produto", "").strip().upper()
    planejamento_semanal = args.get("planejamento_semanal", "").strip()
    planejamento_mensal = args.get("planejamento_mensal", "").strip()
    mes_ano = _parse_mes_ano_form(planejamento_mensal, None) if planejamento_mensal else None

    linhas = []
    for pedido in pedidos:
        for item in pedido.itens:
            if estacao and item.estacao != estacao:
                continue
            if produto and produto not in (item.descricao_produto or "").upper():
                continue
            if planejamento_semanal and item.planejamento_semanal != planejamento_semanal:
                continue
            if mes_ano and _mes_ano_da_semana_pcp(item.planejamento_semanal) != mes_ano:
                continue
            linhas.append(_LinhaListagemGeral(pedido, item))
    return linhas


def _ordenar_com_nulos_no_fim(linhas, chave, reverse):
    """Ordena por `chave(linha)`, deixando quem não tem valor (None) sempre
    no fim, não importa a direção — comportamento mais previsível pro
    usuário do que deixar o Python inverter os vazios junto com o resto."""
    com_valor = [l for l in linhas if chave(l) is not None]
    sem_valor = [l for l in linhas if chave(l) is None]
    com_valor.sort(key=chave, reverse=reverse)
    return com_valor + sem_valor


SORT_KEYS_LISTAGEM_GERAL = {
    "pedido_venda": lambda l: (l.pedido_venda or "").upper() or None,
    "cliente": lambda l: (l.cliente or "").upper() or None,
    "vendedor": lambda l: (l.vendedor or "").upper() or None,
    "data_inclusao": lambda l: l.data_inclusao_pedido,
    "data_cliente": lambda l: l.data_cliente,
    "prioridade": lambda l: PRIORIDADE_OPCOES.index(l.prioridade) if l.prioridade in PRIORIDADE_OPCOES else None,
    "produto": lambda l: (l.descricao_produto or "").upper() or None,
    "quantidade": lambda l: l.quantidade,
    "venda_unidade": lambda l: l.venda_unidade,
    "venda_total": lambda l: l.venda_total,
    "venda_total_pedido": lambda l: l.venda_total_pedido,
    "estacao": lambda l: (l.estacao or "").upper() or None,
    "status": lambda l: STATUS_OPCOES.index(l.status_producao) if l.status_producao in STATUS_OPCOES else None,
    "liberacao_prevista": lambda l: l.liberacao_prevista,
    "liberacao_real": lambda l: l.liberacao_real,
    "planejamento_semanal": lambda l: _chave_semana_pcp(l.planejamento_semanal),
    "frete": lambda l: (l.frete or "").upper() or None,
    "pais": lambda l: (l.pais or "").upper() or None,
    "estado": lambda l: (l.estado or "").upper() or None,
    "cidade": lambda l: (l.cidade or "").upper() or None,
}

SORT_PADRAO = "data_inclusao"
DIR_PADRAO = "desc"


def _filtrar_pedidos_operacao(args):
    """Filtros das 4 sub-abas de Gestão Operação (Comercial/PCP/Logística/
    Resultados). Independente de _filtrar_pedidos (Gestão Produção) — opera só
    em PedidoOperacao, sem nenhum join com Pedido/ItemPedido/estação."""
    query = PedidoOperacao.query

    cliente = args.get("cliente", "").strip()
    vendedor = args.get("vendedor", "").strip()
    busca = args.get("busca", "").strip()

    if cliente:
        query = query.filter(PedidoOperacao.cliente.ilike(f"%{cliente}%"))
    if vendedor:
        query = query.filter(PedidoOperacao.vendedor.ilike(f"%{vendedor}%"))
    if busca:
        like = f"%{busca}%"
        query = query.filter(
            or_(
                PedidoOperacao.pedido_venda.ilike(like),
                PedidoOperacao.cliente.ilike(like),
            )
        )

    query = query.order_by(PedidoOperacao.data_inclusao_pedido.desc().nullslast(), PedidoOperacao.id.desc())

    filtros = dict(cliente=cliente, vendedor=vendedor, busca=busca)
    return query, filtros


def _linhas_gestao_operacao(args):
    """Usado pelas 4 sub-abas de Gestão Operação: pagina o resultado de
    _filtrar_pedidos_operacao. Cada linha já é 1 pedido comercial (tabela
    própria PedidoOperacao) — sem duplicidade legada, sem precisar agrupar
    nada em Python."""
    page = args.get("page", 1, type=int)
    query, filtros = _filtrar_pedidos_operacao(args)
    total_filtrado = query.count()
    total_paginas = max(1, (total_filtrado + PAGE_SIZE - 1) // PAGE_SIZE)
    pagina = query.offset((page - 1) * PAGE_SIZE).limit(PAGE_SIZE).all()
    return pagina, page, total_paginas, total_filtrado, filtros


def _itens_producao_por_pedido_venda(pedidos_venda):
    """Pra tela "Listagem Geral" de Gestão Operação (pedido do Bruno: passar o
    mouse num pedido mostra os itens/quantidades já preenchidos no PCP, em
    Gestão Produção). PedidoOperacao e Pedido/ItemPedido são tabelas
    INDEPENDENTES (sem FK) — o único jeito de ligar um ao outro é o texto do
    "nº pedido de venda" que aparece nos dois. Casa só por igualdade exata
    (já normalizado/tirado espaço), NUNCA por aproximação — é só pra exibir
    numa dica visual, e uma correspondência errada mostraria os itens do
    pedido errado. Uma única query com IN (nunca N+1) — recebe a lista de
    pedido_venda já normalizada da página atual."""
    valores = sorted({v.strip() for v in pedidos_venda if v and v.strip()})
    if not valores:
        return {}
    pedidos = (
        Pedido.query.options(selectinload(Pedido.itens))
        # func.trim() nos dois lados: o texto salvo em Pedido.pedido_venda às
        # vezes tem espaço a mais (import antigo de planilha) — sem isso, a
        # igualdade exata falharia por causa só do espaço, escondendo itens
        # que na prática são do mesmo pedido.
        .filter(func.trim(Pedido.pedido_venda).in_(valores))
        .all()
    )
    mapa = {}
    for pedido in pedidos:
        chave = (pedido.pedido_venda or "").strip()
        if not chave:
            continue
        mapa.setdefault(chave, []).extend(pedido.itens)
    return mapa


# ----------------------------------------------------------------------
# Qualidade — RNC (Relatório de Não Conformidade). Área nova, independente
# de Gestão Produção/Operação (RncQualidade não tem FK com nada). Mesmo
# padrão de filtro/paginação de _filtrar_pedidos_operacao/_linhas_gestao_operacao.
# ----------------------------------------------------------------------
CAMPOS_HISTORICO_RNC = [
    "descricao_nc", "tipo_nc", "severidade", "causa_raiz", "disposicao_produto",
    "acao_corretiva_descricao", "responsavel_acao_corretiva", "prazo_acao_corretiva",
    "status_acao_corretiva", "eficacia_acao", "status_geral", "data_fechamento",
]


def _filtrar_rnc_qualidade(args):
    """`args` é sempre `request.args` (MultiDict) — os 5 filtros de "múltiplas
    opções" (status geral, severidade, origem, tipo de NC, setor) usam
    `getlist`, porque um <select multiple> manda um par nome=valor repetido
    pra cada opção marcada (pedido do Bruno: poder marcar mais de uma opção
    no mesmo filtro, ex.: Espumagem + PU, ou Dureza + Dimensional)."""
    query = RncQualidade.query

    busca = args.get("busca", "").strip()
    status_geral = [v for v in args.getlist("status_geral") if v]
    severidade = [v for v in args.getlist("severidade") if v]
    origem = [v for v in args.getlist("origem") if v]
    tipo_nc = [v for v in args.getlist("tipo_nc") if v]
    setor = [v for v in args.getlist("setor") if v]
    mes_emissao = args.get("mes_emissao", "").strip()  # "AAAA-MM", do <input type="month">
    apenas_abertas = args.get("apenas_abertas", "").strip()

    if busca:
        like = f"%{busca}%"
        query = query.filter(
            or_(
                RncQualidade.numero_rnc.ilike(like),
                RncQualidade.cliente_projeto.ilike(like),
                RncQualidade.produto_equipamento.ilike(like),
                RncQualidade.numero_pedido_contrato.ilike(like),
            )
        )
    if status_geral:
        query = query.filter(RncQualidade.status_geral.in_(status_geral))
    if severidade:
        query = query.filter(RncQualidade.severidade.in_(severidade))
    if origem:
        query = query.filter(RncQualidade.origem.in_(origem))
    if tipo_nc:
        query = query.filter(RncQualidade.tipo_nc.in_(tipo_nc))
    if setor:
        query = query.filter(RncQualidade.setor.in_(setor))
    if mes_emissao:
        try:
            ano_m, mes_m = (int(p) for p in mes_emissao.split("-"))
            inicio = date(ano_m, mes_m, 1)
            fim = date(ano_m, mes_m, monthrange(ano_m, mes_m)[1])
            query = query.filter(RncQualidade.data_emissao.between(inicio, fim))
        except (ValueError, TypeError):
            mes_emissao = ""  # valor incompreensível — ignora o filtro em vez de quebrar a busca
    if apenas_abertas == "1":
        query = query.filter(RncQualidade.status_geral.in_(RNC_STATUS_GERAL_ABERTOS))

    query = query.order_by(RncQualidade.data_emissao.desc().nullslast(), RncQualidade.id.desc())

    filtros = dict(
        busca=busca, status_geral=status_geral, severidade=severidade,
        origem=origem, tipo_nc=tipo_nc, setor=setor, mes_emissao=mes_emissao,
        apenas_abertas=apenas_abertas,
    )
    return query, filtros


def _rnc_opcoes_filtro(campo, opcoes_curadas):
    """Opções de um filtro de RNC (multi-seleção) — combina os valores REAIS
    já cadastrados nesse campo (que podem não bater com a lista de sugestão,
    já que todo campo de "lista" do RNC é texto livre — ver RNC_*_OPCOES em
    models.py) com a lista de sugestão, pra sempre dar pra filtrar por
    qualquer valor que já apareça em algum RNC, mesmo com grafia diferente
    da lista padrão (ex.: "Dureza" na planilha do Bruno vs. "Dureza /
    Material" na lista de sugestão)."""
    coluna = getattr(RncQualidade, campo)
    existentes = [v for (v,) in db.session.query(coluna).filter(coluna.isnot(None)).distinct()]
    existentes_lower = {v.lower() for v in existentes}
    extras = [op for op in opcoes_curadas if op.lower() not in existentes_lower]
    return sorted(existentes + extras, key=lambda s: s.lower())


def _linhas_rnc_qualidade(args):
    page = args.get("page", 1, type=int)
    query, filtros = _filtrar_rnc_qualidade(args)
    total_filtrado = query.count()
    total_paginas = max(1, (total_filtrado + PAGE_SIZE - 1) // PAGE_SIZE)
    pagina = query.offset((page - 1) * PAGE_SIZE).limit(PAGE_SIZE).all()
    return pagina, page, total_paginas, total_filtrado, filtros


def _dashboard_rnc_qualidade():
    """Recalcula ao vivo os mesmos KPIs/quebras da aba "Dashboard" da planilha
    do Bruno, direto de RncQualidade — nada é guardado pré-calculado, então
    fica sempre em dia com o que estiver cadastrado (diferente da planilha
    original, que só atualizava quando alguém reabria o arquivo)."""
    rncs = RncQualidade.query.all()
    total = len(rncs)
    abertas = [r for r in rncs if r.esta_aberto]
    fechadas = [r for r in rncs if not r.esta_aberto]

    com_reincidencia = [r for r in rncs if (r.reincidencia or "").strip().lower() == "sim"]
    pct_reincidencia = round((len(com_reincidencia) / total) * 100, 1) if total else 0

    acoes_nao_eficazes = sum(1 for r in rncs if r.eficacia_acao == "Não Eficaz")
    criticas_abertas = sum(1 for r in rncs if r.severidade == "Crítica" and r.esta_aberto)

    dias_abertos = [r.dias_em_aberto for r in rncs if r.dias_em_aberto is not None]
    tempo_medio_aberto = round(sum(dias_abertos) / len(dias_abertos), 1) if dias_abertos else 0

    custo_total = sum(r.custo_estimado or 0 for r in rncs)

    def _quebra_por(atributo, opcoes_ordem=None):
        contagem = {}
        for r in rncs:
            chave = getattr(r, atributo) or "—"
            contagem[chave] = contagem.get(chave, 0) + 1
        if opcoes_ordem:
            chaves = list(opcoes_ordem) + sorted(k for k in contagem if k not in opcoes_ordem and k != "—")
            if "—" in contagem:
                chaves.append("—")
            return [{"chave": k, "total": contagem.get(k, 0)} for k in chaves if k in contagem or k in opcoes_ordem]
        return sorted(({"chave": k, "total": v} for k, v in contagem.items()), key=lambda d: -d["total"])

    return {
        "total": total,
        "abertas": len(abertas),
        "fechadas": len(fechadas),
        "pct_reincidencia": pct_reincidencia,
        "acoes_nao_eficazes": acoes_nao_eficazes,
        "criticas_abertas": criticas_abertas,
        "tempo_medio_aberto": tempo_medio_aberto,
        "custo_total": custo_total,
        "por_origem": _quebra_por("origem", RNC_ORIGEM_OPCOES),
        "por_tipo_nc": _quebra_por("tipo_nc", RNC_TIPO_NC_OPCOES),
        "por_severidade": _quebra_por("severidade", RNC_SEVERIDADE_OPCOES),
        "por_status_geral": _quebra_por("status_geral", RNC_STATUS_GERAL_OPCOES),
        "por_eficacia": _quebra_por("eficacia_acao", RNC_EFICACIA_OPCOES),
        "por_status_acao": _quebra_por("status_acao_corretiva", RNC_STATUS_ACAO_OPCOES),
    }


def _campos_form_rnc(f):
    """Lê e converte todos os campos do formulário de RNC (novo/editar) — uma
    função só, reaproveitada pelas duas rotas, pra não duplicar a conversão
    de cada um dos ~35 campos editáveis."""
    def _txt(nome):
        return f.get(nome, "").strip() or None

    def _num_int(nome):
        valor = f.get(nome, "").strip()
        try:
            return int(valor) if valor else None
        except ValueError:
            return None

    def _num_float(nome):
        valor = f.get(nome, "").strip().replace(",", ".")
        try:
            return float(valor) if valor else None
        except ValueError:
            return None

    return dict(
        numero_rnc=_txt("numero_rnc"),
        revisao=_num_int("revisao") or 0,
        data_emissao=_parse_data_form(f.get("data_emissao")),
        emitente=_txt("emitente"),
        setor=_txt("setor"),
        origem=_txt("origem"),
        cliente_projeto=_txt("cliente_projeto"),
        numero_pedido_contrato=_txt("numero_pedido_contrato"),
        produto_equipamento=_txt("produto_equipamento"),
        numero_op=_txt("numero_op"),
        local_setor=_txt("local_setor"),
        data_identificacao=_parse_data_form(f.get("data_identificacao")),
        responsavel_identificacao=_txt("responsavel_identificacao"),
        descricao_nc=_txt("descricao_nc"),
        qtd_nao_conforme=_num_int("qtd_nao_conforme"),
        requisito_nao_atendido=_txt("requisito_nao_atendido"),
        tipo_nc=_txt("tipo_nc"),
        severidade=_txt("severidade"),
        acao_contencao_imediata=_txt("acao_contencao_imediata"),
        porque_1=_txt("porque_1"),
        porque_2=_txt("porque_2"),
        porque_3=_txt("porque_3"),
        porque_4=_txt("porque_4"),
        porque_5=_txt("porque_5"),
        causa_raiz=_txt("causa_raiz"),
        ferramenta_analise=_txt("ferramenta_analise"),
        disposicao_produto=_txt("disposicao_produto"),
        acao_corretiva_descricao=_txt("acao_corretiva_descricao"),
        responsavel_acao_corretiva=_txt("responsavel_acao_corretiva"),
        prazo_acao_corretiva=_parse_data_form(f.get("prazo_acao_corretiva")),
        data_realizacao=_parse_data_form(f.get("data_realizacao")),
        status_acao_corretiva=_txt("status_acao_corretiva"),
        data_verificacao_eficacia=_parse_data_form(f.get("data_verificacao_eficacia")),
        eficacia_acao=_txt("eficacia_acao"),
        obs_verificacao=_txt("obs_verificacao"),
        reincidencia=_txt("reincidencia"),
        numero_rnc_relacionada=_txt("numero_rnc_relacionada"),
        custo_estimado=_num_float("custo_estimado"),
        status_geral=_txt("status_geral") or "Aberto",
        responsavel_qualidade=_txt("responsavel_qualidade"),
        data_fechamento=_parse_data_form(f.get("data_fechamento")),
        evidencias_anexos=_txt("evidencias_anexos"),
        observacoes_gerais=_txt("observacoes_gerais"),
    )


_CAMPOS_DATA_RNC = [
    "data_emissao", "data_identificacao", "prazo_acao_corretiva", "data_realizacao",
    "data_verificacao_eficacia", "data_fechamento",
]


def _rnc_para_form_dict(rnc):
    """Converte um RncQualidade em dict de strings prontas pra repopular o
    formulário HTML (mesmo formato que os <input> mandam de volta) — usado
    na tela de edição (GET). Datas em ISO (yyyy-mm-dd, o que <input type=date>
    espera); os demais campos, string direta ou vazio."""
    campos = [
        "numero_rnc", "revisao", "data_emissao", "emitente", "setor", "origem",
        "cliente_projeto", "numero_pedido_contrato", "produto_equipamento", "numero_op",
        "local_setor", "data_identificacao", "responsavel_identificacao", "descricao_nc",
        "qtd_nao_conforme", "requisito_nao_atendido", "tipo_nc", "severidade",
        "acao_contencao_imediata", "porque_1", "porque_2", "porque_3", "porque_4", "porque_5",
        "causa_raiz", "ferramenta_analise", "disposicao_produto", "acao_corretiva_descricao",
        "responsavel_acao_corretiva", "prazo_acao_corretiva", "data_realizacao",
        "status_acao_corretiva", "data_verificacao_eficacia", "eficacia_acao", "obs_verificacao",
        "reincidencia", "numero_rnc_relacionada", "custo_estimado", "status_geral",
        "responsavel_qualidade", "data_fechamento", "evidencias_anexos", "observacoes_gerais",
    ]
    valores = {}
    for campo in campos:
        v = getattr(rnc, campo)
        if v is None:
            valores[campo] = ""
        elif campo in _CAMPOS_DATA_RNC:
            valores[campo] = v.isoformat()
        else:
            valores[campo] = str(v)
    return valores


# ----------------------------------------------------------------------
# Relatórios (fase 12) — exportações CSV/Excel sob demanda, sem agendamento
# automático (geradas na hora, a partir dos mesmos dados já calculados
# pelas telas de Listagem, Faturamento e Gargalos).
# ----------------------------------------------------------------------
def _linhas_export_listagem(pedidos):
    cabecalho = [
        "Pedido", "Cliente", "Cidade", "UF", "Vendedor", "Produto(s)",
        "Valor total (R$)", "Prioridade", "Estação", "Status",
        "Semáforo", "Dias (negativo = atrasado)", "Data de inclusão",
    ]
    linhas = []
    for p in pedidos:
        cor, dias = p.semaforo
        linhas.append([
            p.pedido_venda or "",
            p.cliente,
            p.cidade or "",
            p.estado or "",
            p.vendedor or "",
            p.descricao_resumo,
            round(p.valor_total, 2),
            p.prioridade or "",
            p.estacao_resumo,
            p.status_producao,
            SEMAFORO_LABELS.get(cor, cor),
            dias if dias is not None else "",
            p.data_inclusao_pedido.isoformat() if p.data_inclusao_pedido else "",
        ])
    return cabecalho, linhas


def _linhas_export_faturamento(itens):
    cabecalho = ["Pedido", "Cliente", "UF", "Vendedor", "Produto", "Valor faturado (R$)", "Nº nota fiscal", "Data de faturamento"]
    linhas = []
    for i in itens:
        p = i.pedido
        linhas.append([
            p.pedido_venda if p else "",
            p.cliente if p else "—",
            p.estado if p else "",
            p.vendedor if p else "",
            i.descricao_produto,
            round(i.valor_faturamento_realizado, 2),
            i.numero_nota_fiscal or "",
            i.liberacao_faturamento.isoformat() if i.liberacao_faturamento else "",
        ])
    return cabecalho, linhas


def _linhas_export_gargalos(linhas_gargalo):
    cabecalho = ["Estação", "Fila", "Atrasados", "Tempo de espera médio (dias)", "Lead time médio (dias)", "Valor parado (R$)"]
    linhas = [
        [
            g["estacao"],
            g["fila"],
            g["atraso"],
            g["tempo_espera_medio"] if g["tempo_espera_medio"] is not None else "",
            g["lt_medio"] if g["lt_medio"] is not None else "",
            g["valor_parado"],
        ]
        for g in linhas_gargalo
    ]
    return cabecalho, linhas


def _responder_csv(nome_arquivo, cabecalho, linhas):
    """Gera um CSV com ';' como separador (padrão do Excel em pt-BR) e um BOM
    UTF-8 no início, pra acentos abrirem certo direto no Excel."""
    buffer = io.StringIO()
    writer = csv.writer(buffer, delimiter=";")
    writer.writerow(cabecalho)
    writer.writerows(linhas)
    conteudo = "﻿" + buffer.getvalue()
    resposta = Response(conteudo, mimetype="text/csv; charset=utf-8")
    resposta.headers["Content-Disposition"] = f"attachment; filename={nome_arquivo}"
    return resposta


def _responder_xlsx(nome_arquivo, cabecalho, linhas, titulo="Relatório"):
    wb = Workbook()
    ws = wb.active
    ws.title = titulo[:31] or "Relatório"
    ws.append(cabecalho)
    for celula in ws[1]:
        celula.font = Font(bold=True)
    for linha in linhas:
        ws.append(linha)
    for coluna in ws.columns:
        valores = [len(str(c.value)) for c in coluna if c.value is not None]
        largura = max(valores) if valores else 10
        ws.column_dimensions[coluna[0].column_letter].width = min(largura + 2, 45)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    resposta = Response(
        buffer.getvalue(),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    resposta.headers["Content-Disposition"] = f"attachment; filename={nome_arquivo}"
    return resposta


def _construir_backup_pedidos_wb():
    """Monta um Workbook com TODAS as colunas de Pedido, ItemPedido e
    PedidoOperacao (uma aba cada), lendo as colunas direto do mapeamento do
    SQLAlchemy (não uma lista escrita à mão) — assim não corre o risco de
    esquecer um campo novo que apareça no futuro. Usado como rede de
    segurança antes de "/admin/zerar-dados" apagar tudo de verdade."""
    wb = Workbook()
    wb.remove(wb.active)

    def _add_sheet(nome, modelo, linhas):
        ws = wb.create_sheet(nome[:31])
        colunas = [c.name for c in modelo.__table__.columns]
        ws.append(colunas)
        for celula in ws[1]:
            celula.font = Font(bold=True)
        for obj in linhas:
            linha = []
            for nome_col in colunas:
                valor = getattr(obj, nome_col)
                if isinstance(valor, (datetime, date)):
                    valor = valor.isoformat()
                linha.append(valor)
            ws.append(linha)
        for coluna in ws.columns:
            valores = [len(str(c.value)) for c in coluna if c.value is not None]
            largura = max(valores) if valores else 10
            ws.column_dimensions[coluna[0].column_letter].width = min(largura + 2, 40)

    _add_sheet("Pedidos", Pedido, Pedido.query.order_by(Pedido.id).all())
    _add_sheet("Itens", ItemPedido, ItemPedido.query.order_by(ItemPedido.id).all())
    _add_sheet("Gestao Operacao", PedidoOperacao, PedidoOperacao.query.order_by(PedidoOperacao.id).all())
    return wb


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

        # Quadro mensal de projeção PCP (soma das semanas dentro de cada mês) —
        # filtro de período em <input type=month>, com um intervalo padrão de
        # 3 meses atrás até 6 meses à frente (dá pra ver "quanto foi entregue"
        # nos meses passados e "quanto já está projetado" nos meses seguintes).
        pcp_de_padrao = _somar_meses(hoje.year, hoje.month, -3)
        pcp_ate_padrao = _somar_meses(hoje.year, hoje.month, 6)
        pcp_de = _parse_mes_ano_form(request.args.get("pcp_de"), pcp_de_padrao)
        pcp_ate = _parse_mes_ano_form(request.args.get("pcp_ate"), pcp_ate_padrao)
        if pcp_de > pcp_ate:
            pcp_de, pcp_ate = pcp_ate, pcp_de

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
            projecao_pcp=_projecao_pcp(),
            projecao_pcp_mensal=_projecao_pcp_mensal(pcp_de, pcp_ate),
            pcp_de_str=f"{pcp_de[0]:04d}-{pcp_de[1]:02d}",
            pcp_ate_str=f"{pcp_ate[0]:04d}-{pcp_ate[1]:02d}",
        )

    @app.route("/kpis")
    @login_required
    def kpis():
        meses = request.args.get("meses", 6, type=int)
        if meses not in (3, 6, 12):
            meses = 6
        desde = date.today() - timedelta(days=30 * meses)

        return render_template(
            "kpis.html",
            meses=meses,
            lead_time_estacao=_lead_time_por_estacao(desde=desde),
            otd_vendedor=_otd_por_vendedor(desde=desde),
            tendencia=_tendencia_kpis(meses=meses),
        )

    @app.route("/gargalos")
    @login_required
    def gargalos():
        return render_template("gargalos.html", linhas=_gargalos_por_estacao())

    @app.route("/faturamento")
    @login_required
    def faturamento():
        hoje = date.today()
        ano = request.args.get("ano", hoje.year, type=int)
        mes = request.args.get("mes", hoje.month, type=int)
        if not (1 <= mes <= 12):
            mes = hoje.month

        cliente = request.args.get("cliente", "").strip()
        regiao = request.args.get("regiao", "").strip()
        vendedor = request.args.get("vendedor", "").strip()

        dados = _faturamento_detalhado(ano, mes, cliente=cliente or None, regiao=regiao or None, vendedor=vendedor or None)

        mes_anterior_ano, mes_anterior_mes = (ano, mes - 1) if mes > 1 else (ano - 1, 12)
        mes_seguinte_ano, mes_seguinte_mes = (ano, mes + 1) if mes < 12 else (ano + 1, 1)

        return render_template(
            "faturamento.html",
            ano=ano,
            mes=mes,
            mes_label=f"{MESES_PT[mes - 1]}/{ano}",
            mes_anterior=dict(ano=mes_anterior_ano, mes=mes_anterior_mes),
            mes_seguinte=dict(ano=mes_seguinte_ano, mes=mes_seguinte_mes),
            filtros=dict(cliente=cliente, regiao=regiao, vendedor=vendedor),
            **dados,
        )

    @app.route("/logistica")
    @login_required
    def logistica():
        query = ItemPedido.query.join(Pedido).options(selectinload(ItemPedido.pedido))

        regiao = request.args.get("regiao", "").strip()
        estado = request.args.get("estado", "").strip()
        cidade = request.args.get("cidade", "").strip()
        frete = request.args.get("frete", "").strip()
        transportadora_id = request.args.get("transportadora_id", "").strip()
        em_risco = request.args.get("em_risco", "").strip()
        mostrar_finalizados = request.args.get("mostrar_finalizados", "").strip()

        if regiao:
            ufs_da_regiao = [uf for uf, r in REGIAO_POR_UF.items() if r == regiao]
            if ufs_da_regiao:
                query = query.filter(Pedido.estado.in_(ufs_da_regiao))
        if estado:
            query = query.filter(Pedido.estado == estado)
        if cidade:
            query = query.filter(Pedido.cidade.ilike(f"%{cidade}%"))
        if frete:
            query = query.filter(Pedido.frete == frete)
        if transportadora_id.isdigit():
            query = query.filter(ItemPedido.transportadora_id == int(transportadora_id))
        if not mostrar_finalizados:
            query = query.filter(ItemPedido.status_producao != "FINALIZADO")

        query = query.order_by(Pedido.data_inclusao_pedido.desc().nullslast(), ItemPedido.id.desc())
        itens = query.all()

        # semáforo é calculado em Python (depende de "hoje"), então o filtro
        # "em risco" é aplicado depois da consulta — mas como os filtros acima
        # já reduzem bastante a lista (nunca é a tabela inteira), não há o
        # mesmo problema de performance que o dashboard antigo tinha.
        if em_risco:
            itens = [i for i in itens if i.semaforo[0] in ("amarelo", "vermelho")]

        transportadoras = Transportadora.query.filter_by(ativo=True).order_by(Transportadora.nome).all()
        estados_disponiveis = [
            r[0]
            for r in db.session.query(Pedido.estado)
            .filter(Pedido.estado.isnot(None), Pedido.estado != "")
            .distinct()
            .order_by(Pedido.estado)
            .all()
        ]

        return render_template(
            "logistica.html",
            itens=itens,
            transportadoras=transportadoras,
            estados_disponiveis=estados_disponiveis,
            filtros=dict(
                regiao=regiao,
                estado=estado,
                cidade=cidade,
                frete=frete,
                transportadora_id=transportadora_id,
                em_risco=em_risco,
                mostrar_finalizados=mostrar_finalizados,
            ),
        )

    @app.route("/cadastros/transportadoras")
    @requer_role("ADMIN", "PCP")
    def cadastros_transportadoras():
        transportadoras = Transportadora.query.order_by(Transportadora.nome).all()
        return render_template("cadastros_transportadoras.html", transportadoras=transportadoras)

    @app.route("/cadastros/transportadoras/novo", methods=["GET", "POST"])
    @requer_role("ADMIN", "PCP")
    def cadastros_transportadoras_novo():
        if request.method == "POST":
            f = request.form
            nome = f.get("nome", "").strip()
            if not nome:
                flash("Informe o nome da transportadora.", "danger")
                return render_template("cadastros_transportadoras_form.html", transportadora=None, form=f)
            if Transportadora.query.filter_by(nome=nome).first():
                flash("Já existe uma transportadora com esse nome.", "danger")
                return render_template("cadastros_transportadoras_form.html", transportadora=None, form=f)

            nova = Transportadora(nome=nome, ativo=True)
            db.session.add(nova)
            db.session.commit()
            flash(f"Transportadora {nome} cadastrada com sucesso.", "success")
            return redirect(url_for("cadastros_transportadoras"))

        return render_template("cadastros_transportadoras_form.html", transportadora=None, form={})

    @app.route("/cadastros/transportadoras/<int:transportadora_id>/editar", methods=["GET", "POST"])
    @requer_role("ADMIN", "PCP")
    def cadastros_transportadoras_editar(transportadora_id):
        transportadora = db.session.get(Transportadora, transportadora_id)
        if transportadora is None:
            flash("Transportadora não encontrada.", "danger")
            return redirect(url_for("cadastros_transportadoras"))

        if request.method == "POST":
            f = request.form
            nome = f.get("nome", "").strip()
            if not nome:
                flash("Informe o nome da transportadora.", "danger")
                return render_template("cadastros_transportadoras_form.html", transportadora=transportadora, form=f)
            outra = Transportadora.query.filter(Transportadora.nome == nome, Transportadora.id != transportadora.id).first()
            if outra:
                flash("Já existe outra transportadora com esse nome.", "danger")
                return render_template("cadastros_transportadoras_form.html", transportadora=transportadora, form=f)

            transportadora.nome = nome
            transportadora.ativo = bool(f.get("ativo"))
            db.session.commit()
            flash(f"Transportadora {transportadora.nome} atualizada com sucesso.", "success")
            return redirect(url_for("cadastros_transportadoras"))

        return render_template("cadastros_transportadoras_form.html", transportadora=transportadora, form={})

    @app.route("/alertas")
    @login_required
    def alertas():
        pedidos_atrasados = (
            Pedido.query.options(selectinload(Pedido.itens))
            .filter(_predicado_atrasado())
            .order_by(Pedido.data_inclusao_pedido.desc().nullslast())
            .all()
        )
        pedidos_vencendo = (
            Pedido.query.options(selectinload(Pedido.itens))
            .filter(_predicado_vencendo())
            .order_by(Pedido.data_inclusao_pedido.desc().nullslast())
            .all()
        )
        gargalos_criticos = [g for g in _gargalos_por_estacao() if g["fila"] > 0 or g["atraso"] > 0][:5]
        faturamento_pendente = _faturamento_previsto_nao_realizado()

        return render_template(
            "alertas.html",
            pedidos_atrasados=pedidos_atrasados,
            pedidos_vencendo=pedidos_vencendo,
            gargalos_criticos=gargalos_criticos,
            faturamento_pendente=faturamento_pendente,
        )

    # ------------------------------------------------------------------
    # Relatórios (fase 12) — exportações sob demanda, sem agendamento
    # automático. Cada relatório tem uma versão .csv e uma .xlsx.
    # ------------------------------------------------------------------
    @app.route("/relatorios")
    @login_required
    def relatorios():
        hoje = date.today()
        return render_template("relatorios.html", hoje=hoje)

    @app.route("/relatorios/listagem.csv")
    @login_required
    def relatorio_listagem_csv():
        query, _ = _filtrar_pedidos(request.args)
        cabecalho, linhas = _linhas_export_listagem(query.all())
        return _responder_csv("listagem_pedidos.csv", cabecalho, linhas)

    @app.route("/relatorios/listagem.xlsx")
    @login_required
    def relatorio_listagem_xlsx():
        query, _ = _filtrar_pedidos(request.args)
        cabecalho, linhas = _linhas_export_listagem(query.all())
        return _responder_xlsx("listagem_pedidos.xlsx", cabecalho, linhas, titulo="Listagem")

    @app.route("/relatorios/faturamento.csv")
    @login_required
    def relatorio_faturamento_csv():
        hoje = date.today()
        ano = request.args.get("ano", hoje.year, type=int)
        mes = request.args.get("mes", hoje.month, type=int)
        if not (1 <= mes <= 12):
            mes = hoje.month
        cliente = request.args.get("cliente", "").strip()
        regiao = request.args.get("regiao", "").strip()
        vendedor = request.args.get("vendedor", "").strip()
        dados = _faturamento_detalhado(ano, mes, cliente=cliente or None, regiao=regiao or None, vendedor=vendedor or None)
        cabecalho, linhas = _linhas_export_faturamento(dados["itens_realizados_lista"])
        return _responder_csv(f"faturamento_{ano}_{mes:02d}.csv", cabecalho, linhas)

    @app.route("/relatorios/faturamento.xlsx")
    @login_required
    def relatorio_faturamento_xlsx():
        hoje = date.today()
        ano = request.args.get("ano", hoje.year, type=int)
        mes = request.args.get("mes", hoje.month, type=int)
        if not (1 <= mes <= 12):
            mes = hoje.month
        cliente = request.args.get("cliente", "").strip()
        regiao = request.args.get("regiao", "").strip()
        vendedor = request.args.get("vendedor", "").strip()
        dados = _faturamento_detalhado(ano, mes, cliente=cliente or None, regiao=regiao or None, vendedor=vendedor or None)
        cabecalho, linhas = _linhas_export_faturamento(dados["itens_realizados_lista"])
        return _responder_xlsx(f"faturamento_{ano}_{mes:02d}.xlsx", cabecalho, linhas, titulo="Faturamento")

    @app.route("/relatorios/gargalos.csv")
    @login_required
    def relatorio_gargalos_csv():
        cabecalho, linhas = _linhas_export_gargalos(_gargalos_por_estacao())
        return _responder_csv("gargalos.csv", cabecalho, linhas)

    @app.route("/relatorios/gargalos.xlsx")
    @login_required
    def relatorio_gargalos_xlsx():
        cabecalho, linhas = _linhas_export_gargalos(_gargalos_por_estacao())
        return _responder_xlsx("gargalos.xlsx", cabecalho, linhas, titulo="Gargalos")

    @app.route("/")
    @login_required
    def dashboard():
        page = request.args.get("page", 1, type=int)
        sort = request.args.get("sort", SORT_PADRAO)
        dir_ordenacao = request.args.get("dir", DIR_PADRAO)
        if sort not in SORT_KEYS_LISTAGEM_GERAL:
            sort = SORT_PADRAO
        if dir_ordenacao not in ("asc", "desc"):
            dir_ordenacao = DIR_PADRAO

        # A Listagem Geral mostra 1 linha por PRODUTO (item), não por pedido —
        # então a ordenação/paginação não dá mais pra fazer em SQL direto (várias
        # colunas, como prioridade/status/prazo, têm ordem própria calculada em
        # Python). Busca tudo que passou pelo filtro, achata em linhas por item,
        # ordena e pagina em Python — mesmo padrão já usado em Gestão Operação e
        # no Painel pra esse tipo de coluna.
        query, filtros = _filtrar_pedidos(request.args)
        pedidos = query.all()
        linhas = _linhas_listagem_geral(pedidos, request.args)
        linhas = _ordenar_com_nulos_no_fim(linhas, SORT_KEYS_LISTAGEM_GERAL[sort], reverse=(dir_ordenacao == "desc"))

        # A pedido de Bruno: a Listagem Geral mostra todos os itens filtrados
        # numa página só, sem paginação (nada de "página 2, 3, 4...").
        total_filtrado = len(linhas)
        total_paginas = 1
        linhas_pagina = linhas

        resumo = _calcular_resumo()

        filtros_paginacao = dict(filtros, sort=sort, dir=dir_ordenacao)

        return render_template(
            "dashboard.html",
            linhas=linhas_pagina,
            resumo=resumo,
            page=page,
            total_paginas=total_paginas,
            total_filtrado=total_filtrado,
            filtros=filtros,
            filtros_paginacao=filtros_paginacao,
            sort=sort,
            dir_ordenacao=dir_ordenacao,
        )

    @app.route("/pedidos/novo", methods=["GET", "POST"])
    @requer_role("ADMIN", "PCP")
    def novo_pedido():
        if request.method == "POST":
            f = request.form
            descricoes = f.getlist("item_descricao[]")
            quantidades = f.getlist("item_quantidade[]")
            custos = f.getlist("item_custo[]")
            estacoes = f.getlist("item_estacao[]")

            itens = []
            itens_recarregados = []
            for desc, qtd, custo, estacao_item in zip(descricoes, quantidades, custos, estacoes):
                itens_recarregados.append({"descricao": desc, "quantidade": qtd, "custo": custo, "estacao": estacao_item})
                desc = desc.strip()
                if not desc:
                    continue
                itens.append(
                    ItemPedido(
                        descricao_produto=desc,
                        quantidade=_parse_float_form(qtd),
                        custo_unitario=_parse_float_form(custo),
                        estacao=estacao_item or None,
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

        transportadoras = Transportadora.query.filter_by(ativo=True).order_by(Transportadora.nome).all()

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

            # Planejamento semanal / liberação prevista / liberação real (pedido do
            # Bruno, 31/08/2026): deixaram de ser editáveis por item — agora só
            # existe um campo de cada no topo do pedido, e o valor é replicado pra
            # TODOS os itens ao salvar (por isso lidos uma vez só, fora do loop de
            # itens, em vez de via getlist("item_...[]") como os campos por item).
            lib_prevista_pedido = _parse_data_form(f.get("liberacao_prevista"))
            lib_real_pedido = _parse_data_form(f.get("liberacao_real"))
            planejamento_semanal_pedido = f.get("planejamento_semanal", "").strip() or None

            # ---- itens do pedido (vários produtos, cada um com sua própria estação/status/produção) ----
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
            notas_fiscais = f.getlist("item_numero_nota_fiscal[]")
            valores_faturados = f.getlist("item_valor_faturado[]")
            transportadoras_ids = f.getlist("item_transportadora_id[]")
            datas_envio = f.getlist("item_data_envio[]")

            itens_originais = {item.id: item for item in pedido.itens}
            # snapshot ANTES de qualquer alteração, só dos itens que já existiam
            # (itens novos não têm "antes" pra comparar — são uma inclusão, não uma mudança)
            itens_antes = {iid: {c: getattr(item, c) for c in CAMPOS_HISTORICO_ITEM} for iid, item in itens_originais.items()}
            ids_mantidos = set()

            linhas = zip(
                item_ids, descricoes, quantidades, custos, estacoes, status_itens, rncs,
                inicios_producao, inicios_inspecao, terminos_inspecao,
                liberacoes_faturamento, notas_fiscais, valores_faturados,
                transportadoras_ids, datas_envio,
            )
            for (item_id, desc, qtd, custo, estacao_item, status_item, rnc,
                 ini_prod, ini_insp, term_insp, lib_fat, nf, valor_faturado,
                 transportadora_id, data_envio) in linhas:
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
                item.liberacao_prevista = lib_prevista_pedido
                item.liberacao_real = lib_real_pedido
                item.planejamento_semanal = planejamento_semanal_pedido
                item.numero_nota_fiscal = nf.strip() or None
                item.valor_faturado = _parse_float_form(valor_faturado, default=None) if valor_faturado.strip() else None
                item.transportadora_id = int(transportadora_id) if transportadora_id.strip().isdigit() else None
                item.data_envio = _parse_data_form(data_envio)

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
                return render_template("editar_pedido.html", pedido=pedido, transportadoras=transportadoras)

            db.session.commit()
            flash("Pedido atualizado com sucesso.", "success")
            return redirect(url_for("editar_pedido", pedido_id=pedido.id))

        return render_template("editar_pedido.html", pedido=pedido, transportadoras=transportadoras)

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
    # Gestão Operação (Fase 13) — sub-abas coloridas (PCP/Logística/
    # Resultados), uma linha por PEDIDO, reaproveitando os mesmos filtros e
    # paginação da Listagem Geral (_filtrar_pedidos) — só muda o conjunto de
    # colunas mostrado em cada template. A sub-aba "Comercial" (lista) foi
    # removida (pedido do Bruno, 01/09/2026): a Listagem Geral já mostra as
    # mesmas informações comerciais, então a lista separada virou
    # redundante — o FORMULÁRIO de edição da seção Comercial continua
    # existindo normalmente em gestao_operacao_editar (ver GO_SECAO_ENDPOINT).
    # ------------------------------------------------------------------
    @app.route("/gestao-operacao/pcp")
    @login_required
    def gestao_operacao_pcp():
        pedidos, page, total_paginas, total_filtrado, filtros = _linhas_gestao_operacao(request.args)
        return render_template(
            "gestao_operacao_pcp.html",
            pedidos=pedidos, page=page, total_paginas=total_paginas,
            total_filtrado=total_filtrado, filtros=filtros,
        )

    @app.route("/gestao-operacao/logistica")
    @login_required
    def gestao_operacao_logistica():
        pedidos, page, total_paginas, total_filtrado, filtros = _linhas_gestao_operacao(request.args)
        return render_template(
            "gestao_operacao_logistica.html",
            pedidos=pedidos, page=page, total_paginas=total_paginas,
            total_filtrado=total_filtrado, filtros=filtros,
        )

    @app.route("/gestao-operacao/resultados")
    @login_required
    def gestao_operacao_resultados():
        pedidos, page, total_paginas, total_filtrado, filtros = _linhas_gestao_operacao(request.args)
        otd = _resumo_otd()

        # Faturamento por Semana (pedido do Bruno, 28/08/2026) — mês
        # selecionável, sem hardcode: default é o mês atual (mesmo padrão da
        # tela de Faturamento de Gestão Produção).
        hoje = date.today()
        ano = request.args.get("ano", hoje.year, type=int)
        mes = request.args.get("mes", hoje.month, type=int)
        if not (1 <= mes <= 12):
            mes = hoje.month
        faturamento_semanal = _faturamento_semanal_pcp(ano, mes)
        mes_anterior_ano, mes_anterior_mes = (ano, mes - 1) if mes > 1 else (ano - 1, 12)
        mes_seguinte_ano, mes_seguinte_mes = (ano, mes + 1) if mes < 12 else (ano + 1, 1)

        return render_template(
            "gestao_operacao_resultados.html",
            pedidos=pedidos, page=page, total_paginas=total_paginas,
            total_filtrado=total_filtrado, filtros=filtros, otd=otd,
            faturamento_semanal=faturamento_semanal,
            ano=ano, mes=mes, mes_label=f"{MESES_PT[mes - 1]}/{ano}",
            mes_anterior=dict(ano=mes_anterior_ano, mes=mes_anterior_mes),
            mes_seguinte=dict(ano=mes_seguinte_ano, mes=mes_seguinte_mes),
        )

    @app.route("/gestao-operacao/listagem-geral")
    @login_required
    def gestao_operacao_listagem_geral():
        """"Listagem Geral" de Gestão Operação (pedido do Bruno) — 1 linha por
        PEDIDO (não por produto, diferente da Listagem Geral de Gestão
        Produção), com as colunas comerciais principais; passar o mouse (ou
        clicar, no touch) sobre "Itens" mostra os produtos/quantidades já
        preenchidos em Gestão Produção pelo PCP, casando pelo nº de pedido de
        venda — sem criar nenhum vínculo real entre as duas tabelas."""
        pedidos, page, total_paginas, total_filtrado, filtros = _linhas_gestao_operacao(request.args)
        itens_por_pedido_venda = _itens_producao_por_pedido_venda([p.pedido_venda for p in pedidos])
        return render_template(
            "gestao_operacao_listagem_geral.html",
            pedidos=pedidos, page=page, total_paginas=total_paginas,
            total_filtrado=total_filtrado, filtros=filtros,
            itens_por_pedido_venda=itens_por_pedido_venda,
        )

    @app.route("/gestao-operacao/novo", methods=["GET", "POST"])
    @requer_role("ADMIN", "PCP")
    def gestao_operacao_novo():
        """Cria um pedido novo DENTRO da Gestão Operação — 100% independente da
        Gestão Produção (não cria Pedido/ItemPedido nenhum). Só o bloco
        Comercial; depois de salvar, redireciona pra edição pra preencher PCP/
        Logística/Resultados aos poucos, conforme o pedido avança."""
        if request.method == "POST":
            f = request.form
            cliente = f.get("cliente", "").strip()
            if not cliente:
                flash("Cliente é obrigatório.", "danger")
                return render_template("gestao_operacao_novo.html", form=f)

            novo = PedidoOperacao(
                cliente=cliente,
                vendedor=f.get("vendedor", "").strip() or None,
                pedido_venda=f.get("pedido_venda", "").strip() or None,
                data_inclusao_pedido=_parse_data_form(f.get("data_inclusao_pedido")),
                prioridade=f.get("prioridade") or "MÉDIA",
                frete=f.get("frete") or None,
                pais=f.get("pais", "").strip() or "Brasil",
                estado=f.get("estado", "").strip() or None,
                cidade=f.get("cidade", "").strip() or None,
                go_tipo_pedido=f.get("go_tipo_pedido", "").strip() or None,
                go_contrato=f.get("go_contrato", "").strip() or None,
                go_pedido_compra_cliente=f.get("go_pedido_compra_cliente", "").strip() or None,
                go_proposta=f.get("go_proposta", "").strip() or None,
                go_data_solicitada_entrega=_parse_data_form(f.get("go_data_solicitada_entrega")),
                go_status_pedido_info=f.get("go_status_pedido_info", "").strip() or None,
                go_valor_pedido_operacao=(
                    _parse_float_form(f.get("go_valor_pedido_operacao"), default=None)
                    if f.get("go_valor_pedido_operacao", "").strip()
                    else None
                ),
            )
            db.session.add(novo)
            db.session.commit()
            flash(f"Pedido de {novo.cliente} incluído na Gestão Operação com sucesso.", "success")
            return redirect(url_for("gestao_operacao_editar", pedido_id=novo.id))

        return render_template("gestao_operacao_novo.html", form={})

    @app.route("/gestao-operacao/<int:pedido_id>/editar", methods=["GET", "POST"])
    @requer_role("ADMIN", "PCP")
    def gestao_operacao_editar(pedido_id):
        pedido = db.session.get(PedidoOperacao, pedido_id)
        if pedido is None:
            flash("Pedido não encontrado.", "danger")
            return redirect(url_for("gestao_operacao_listagem_geral"))

        # Pedido do Bruno (31/08/2026): cada aba só edita os campos da
        # própria área (PCP só mostra/edita campos de PCP, Logística só os
        # de Logística...) — nunca lê nem sobrescreve campos de outra seção,
        # mesmo que o form de outra aba tivesse ficado aberto em outra guia.
        secao = request.args.get("secao", "comercial")
        if secao not in GO_SECOES:
            secao = "comercial"
        campos_secao = GO_CAMPOS_POR_SECAO[secao]

        transportadoras = Transportadora.query.filter_by(ativo=True).order_by(Transportadora.nome).all()

        if request.method == "POST":
            f = request.form
            campos_historico = [c for c in campos_secao if c in CAMPOS_HISTORICO_GESTAO_OPERACAO]
            antes = {c: getattr(pedido, c) for c in campos_historico}

            for campo in campos_secao:
                setattr(pedido, campo, _parse_campo_go(campo, f))

            depois = {c: getattr(pedido, c) for c in campos_historico}
            # pedido_id fica None de propósito: essa coluna tem FK de verdade pra
            # `pedidos` (Gestão Produção) — PedidoOperacao é uma tabela totalmente
            # separada, então gravar o id dela ali violaria a FK. entidade_id já
            # guarda o id certo.
            _registrar_alteracoes("pedido_operacao", pedido.id, None, antes, depois, campos_historico)

            db.session.commit()
            flash(f"Gestão Operação ({GO_SECAO_LABEL[secao]}) do pedido atualizada com sucesso.", "success")
            return redirect(url_for("gestao_operacao_editar", pedido_id=pedido.id, secao=secao))

        return render_template(
            "gestao_operacao_editar.html", pedido=pedido, transportadoras=transportadoras,
            secao=secao, GO_SECOES=GO_SECOES, GO_SECAO_ENDPOINT=GO_SECAO_ENDPOINT, GO_SECAO_LABEL=GO_SECAO_LABEL,
        )

    # ------------------------------------------------------------------
    # Qualidade — RNC (Relatório de Não Conformidade). Área nova (31/08/2026),
    # independente de Gestão Produção/Operação. Pedido do Bruno: controle
    # totalmente manual e intuitivo, aberto a todos os usuários autenticados
    # (view + edição) — sem role dedicado, confirmado com ele antes de
    # implementar (ver decisão em AskUserQuestion).
    # ------------------------------------------------------------------
    @app.route("/qualidade/dashboard")
    @login_required
    def qualidade_dashboard():
        dados = _dashboard_rnc_qualidade()
        return render_template("qualidade_dashboard.html", dados=dados)

    @app.route("/qualidade")
    @login_required
    def qualidade_lista():
        rncs, page, total_paginas, total_filtrado, filtros = _linhas_rnc_qualidade(request.args)
        opcoes_filtro = dict(
            status_geral=_rnc_opcoes_filtro("status_geral", RNC_STATUS_GERAL_OPCOES),
            severidade=_rnc_opcoes_filtro("severidade", RNC_SEVERIDADE_OPCOES),
            origem=_rnc_opcoes_filtro("origem", RNC_ORIGEM_OPCOES),
            tipo_nc=_rnc_opcoes_filtro("tipo_nc", RNC_TIPO_NC_OPCOES),
            setor=_rnc_opcoes_filtro("setor", RNC_SETOR_OPCOES),
        )
        return render_template(
            "qualidade_lista.html",
            rncs=rncs, page=page, total_paginas=total_paginas,
            total_filtrado=total_filtrado, filtros=filtros,
            opcoes_filtro=opcoes_filtro,
        )

    @app.route("/qualidade/novo", methods=["GET", "POST"])
    @login_required
    def qualidade_novo():
        if request.method == "POST":
            f = request.form
            cliente_projeto = f.get("cliente_projeto", "").strip()
            descricao_nc = f.get("descricao_nc", "").strip()
            if not cliente_projeto or not descricao_nc:
                flash("Cliente/Projeto e Descrição da Não Conformidade são obrigatórios.", "danger")
                return render_template("qualidade_novo.html", valores=f)

            valores = _campos_form_rnc(f)
            novo = RncQualidade(criado_por_id=current_user.id, **valores)
            db.session.add(novo)
            db.session.commit()
            flash(f"RNC {novo.numero_rnc or ('#' + str(novo.id))} cadastrado com sucesso.", "success")
            return redirect(url_for("qualidade_editar", rnc_id=novo.id))

        return render_template("qualidade_novo.html", valores={})

    @app.route("/qualidade/<int:rnc_id>/editar", methods=["GET", "POST"])
    @login_required
    def qualidade_editar(rnc_id):
        rnc = db.session.get(RncQualidade, rnc_id)
        if rnc is None:
            flash("RNC não encontrado.", "danger")
            return redirect(url_for("qualidade_lista"))

        if request.method == "POST":
            f = request.form
            cliente_projeto = f.get("cliente_projeto", "").strip()
            descricao_nc = f.get("descricao_nc", "").strip()
            if not cliente_projeto or not descricao_nc:
                flash("Cliente/Projeto e Descrição da Não Conformidade são obrigatórios.", "danger")
                return render_template("qualidade_editar.html", rnc=rnc, valores=f, historico=[])

            antes = {c: getattr(rnc, c) for c in CAMPOS_HISTORICO_RNC}
            valores = _campos_form_rnc(f)
            for campo, valor in valores.items():
                setattr(rnc, campo, valor)
            depois = {c: getattr(rnc, c) for c in CAMPOS_HISTORICO_RNC}
            _registrar_alteracoes("rnc_qualidade", rnc.id, None, antes, depois, CAMPOS_HISTORICO_RNC)

            db.session.commit()
            flash("RNC atualizado com sucesso.", "success")
            return redirect(url_for("qualidade_editar", rnc_id=rnc.id))

        historico = (
            HistoricoAlteracao.query
            .filter_by(entidade_tipo="rnc_qualidade", entidade_id=rnc.id)
            .order_by(HistoricoAlteracao.criado_em.desc())
            .all()
        )
        return render_template("qualidade_editar.html", rnc=rnc, valores=_rnc_para_form_dict(rnc), historico=historico)

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
    # Zerar dados de teste (pedido do Bruno, 28/08/2026) — apaga TODOS os
    # Pedido/ItemPedido (Gestão Produção) e PedidoOperacao (Gestão Operação)
    # de uma vez, pra ele testar manualmente a inclusão de itens do zero. Só
    # ADMIN, com confirmação por texto digitado — e sempre com um backup
    # (.xlsx) disponível antes, já que é uma exclusão permanente.
    # ------------------------------------------------------------------
    _FRASE_CONFIRMACAO_ZERAR = "ZERAR TUDO"

    @app.route("/admin/zerar-dados")
    @requer_role("ADMIN")
    def admin_zerar_dados():
        return render_template(
            "admin_zerar_dados.html",
            total_pedidos=Pedido.query.count(),
            total_itens=ItemPedido.query.count(),
            total_pedidos_operacao=PedidoOperacao.query.count(),
            frase_confirmacao=_FRASE_CONFIRMACAO_ZERAR,
        )

    @app.route("/admin/zerar-dados/backup.xlsx")
    @requer_role("ADMIN")
    def admin_zerar_dados_backup():
        """Só gera e baixa o backup — não apaga nada. Pode ser clicado quantas
        vezes quiser antes (ou até sem intenção de zerar depois)."""
        wb = _construir_backup_pedidos_wb()
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        nome = f"backup_pedidos_{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}.xlsx"
        resposta = Response(
            buffer.getvalue(),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        resposta.headers["Content-Disposition"] = f"attachment; filename={nome}"
        return resposta

    @app.route("/admin/zerar-dados", methods=["POST"])
    @requer_role("ADMIN")
    def admin_zerar_dados_confirmar():
        confirmacao = request.form.get("confirmacao", "").strip()
        if confirmacao != _FRASE_CONFIRMACAO_ZERAR:
            flash(f'Digite exatamente "{_FRASE_CONFIRMACAO_ZERAR}" pra confirmar. Nada foi apagado.', "danger")
            return redirect(url_for("admin_zerar_dados"))

        total_pedidos = Pedido.query.count()
        total_itens = ItemPedido.query.count()
        total_pedidos_operacao = PedidoOperacao.query.count()

        # Ordem importa: filhos antes dos pais, por causa das foreign keys
        # (Programacao -> ItemPedido; HistoricoAlteracao -> Pedido).
        Programacao.query.delete(synchronize_session=False)
        HistoricoAlteracao.query.delete(synchronize_session=False)
        ItemPedido.query.delete(synchronize_session=False)
        Pedido.query.delete(synchronize_session=False)
        PedidoOperacao.query.delete(synchronize_session=False)

        # Marca que foi de propósito — sem isso, _seed_inicial e
        # _importar_gestao_operacao reimportariam a planilha antiga sozinhos
        # no próximo deploy, assim que virem as tabelas vazias.
        if ControleSistema.query.filter_by(chave=_CHAVE_DADOS_ZERADOS_MANUALMENTE).first() is None:
            db.session.add(ControleSistema(chave=_CHAVE_DADOS_ZERADOS_MANUALMENTE))

        db.session.commit()
        app.logger.warning(
            "ZERAR DADOS: %s executou a limpeza manual — %d pedidos, %d itens e %d pedidos de "
            "Gestão Operação apagados.",
            current_user.nome, total_pedidos, total_itens, total_pedidos_operacao,
        )
        flash(
            f"Pronto: {total_pedidos} pedido(s), {total_itens} item(ns) e {total_pedidos_operacao} "
            "pedido(s) de Gestão Operação foram apagados. Pode começar a testar do zero.",
            "success",
        )
        return redirect(url_for("admin_zerar_dados"))

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

        # Pedido do Bruno (25/08/2026): em toda coluna, o item mais novo/mais
        # recém movimentado fica no topo. "atualizado_em" cobre os dois casos
        # de uma vez só — item novo nasce com esse carimbo, e "Avançar" no
        # Kanban (ou qualquer edição) atualiza sozinho (ver ItemPedido.atualizado_em).
        itens = ItemPedido.query.options(selectinload(ItemPedido.pedido)).filter(ItemPedido.estacao == nome).all()
        itens.sort(key=lambda i: (i.atualizado_em or datetime.min, i.id), reverse=True)

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

        # Kanban simplificado (3 colunas): "avançar" carimba a data que falta
        # pra status_producao concluir sozinho a próxima etapa (mesma regra
        # de atualizar_status_automatico, sem duplicar lógica aqui).
        if etapa_atual == "PENDENTE":
            item.inicio_producao = item.inicio_producao or hoje
        elif etapa_atual == "EM_PRODUCAO":
            item.termino_inspecao = item.termino_inspecao or hoje
            item.liberacao_faturamento = item.liberacao_faturamento or hoje
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

    # ------------------------------------------------------------------
    # Programação — calendário mensal do PCP (pedido do Bruno, 31/08/2026):
    # 1 linha por semana (domingo a sábado) do mês, item já entra no dia certo
    # sozinho pela liberação prevista (azul) ou liberação real (verde), sem
    # precisar programar nada manualmente. Substitui o board antigo de
    # segunda-sexta com "+"/modal (ver _semanas_calendario_pcp acima).
    # ------------------------------------------------------------------
    @app.route("/programacao")
    @login_required
    def programacao_semana():
        hoje = date.today()
        ano = request.args.get("ano", hoje.year, type=int)
        mes = request.args.get("mes", hoje.month, type=int)
        if not (1 <= mes <= 12):
            ano, mes = hoje.year, hoje.month

        semanas = _semanas_calendario_pcp(ano, mes)
        inicio_range = semanas[0]["inicio"]
        fim_range = semanas[-1]["fim"]

        itens = (
            ItemPedido.query.options(selectinload(ItemPedido.pedido))
            .filter(
                or_(
                    ItemPedido.liberacao_real.between(inicio_range, fim_range),
                    and_(
                        ItemPedido.liberacao_real.is_(None),
                        ItemPedido.liberacao_prevista.between(inicio_range, fim_range),
                    ),
                )
            )
            .all()
        )

        for semana in semanas:
            # cada dia guarda um GRUPO por pedido, não por item — desde que
            # Liberação prevista/real virou campo único do pedido (cascateia
            # pra todos os itens ao salvar, pedido do Bruno 31/08/2026), todo
            # item do mesmo pedido cai sempre no mesmo dia, então o quadrante
            # mostra 1 card por pedido (resumo), com os itens dentro do
            # hover/clique — em vez de 1 card por item repetindo cliente/frete/etc.
            semana["dias"] = [{} for _ in range(7)]  # 0=domingo ... 6=sábado, cada um {pedido_id: grupo}
            semana["dias_datas"] = [semana["inicio"] + timedelta(days=i) for i in range(7)]

        for item in itens:
            data_efetiva = item.liberacao_real or item.liberacao_prevista
            for semana in semanas:
                if semana["inicio"] <= data_efetiva <= semana["fim"]:
                    coluna = (data_efetiva.weekday() + 1) % 7  # weekday(): segunda=0 -> domingo vira 0
                    grupos = semana["dias"][coluna]
                    grupo = grupos.get(item.pedido_id)
                    if grupo is None:
                        grupo = {"pedido": item.pedido, "itens": [], "data": data_efetiva}
                        grupos[item.pedido_id] = grupo
                    grupo["itens"].append(item)
                    break

        for semana in semanas:
            for coluna_idx, grupos in enumerate(semana["dias"]):
                lista = sorted(grupos.values(), key=lambda g: ((g["pedido"].cliente or ""), g["pedido"].pedido_venda or ""))
                for g in lista:
                    # se os itens desse pedido não estiverem 100% sincronizados
                    # (dado antigo, de antes do campo virar único por pedido), o
                    # card só fica verde quando TODOS já tiverem liberação real —
                    # senão fica azul (mais conservador).
                    g["confirmado"] = all(it.liberacao_real is not None for it in g["itens"])
                semana["dias"][coluna_idx] = lista

        total_sem_liberacao_prevista = (
            ItemPedido.query
            .filter(ItemPedido.liberacao_prevista.is_(None), ItemPedido.liberacao_real.is_(None))
            .filter(ItemPedido.status_producao != "FINALIZADO")
            .count()
        )

        mes_anterior_ano, mes_anterior_mes = _somar_meses(ano, mes, -1)
        mes_seguinte_ano, mes_seguinte_mes = _somar_meses(ano, mes, 1)

        return render_template(
            "programacao_semana.html",
            semanas=semanas,
            ano=ano, mes=mes,
            mes_label=f"{MESES_PT[mes - 1]} / {ano}",
            mes_anterior=dict(ano=mes_anterior_ano, mes=mes_anterior_mes),
            mes_seguinte=dict(ano=mes_seguinte_ano, mes=mes_seguinte_mes),
            hoje=hoje,
            total_sem_liberacao_prevista=total_sem_liberacao_prevista,
            DIAS_SEMANA_LABELS=["Domingo", "Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado"],
        )

    @app.route("/programacao/novo", methods=["POST"])
    @requer_role("ADMIN", "PCP")
    def programacao_novo():
        f = request.form
        item_id = f.get("item_pedido_id", type=int)
        data_programada = _parse_data_form(f.get("data_programada"))
        item = db.session.get(ItemPedido, item_id) if item_id else None

        if item is None or data_programada is None:
            flash("Selecione um item e uma data válidos para programar.", "danger")
            return redirect(url_for("programacao_semana"))

        estacao_escolhida = f.get("estacao") or item.estacao or "—"

        nova = Programacao(
            item_pedido_id=item.id,
            data_programada=data_programada,
            estacao=estacao_escolhida,
            prioridade_producao=f.get("prioridade_producao") or None,
            observacao=f.get("observacao", "").strip() or None,
            responsavel_id=f.get("responsavel_id", type=int) or None,
            criado_por_id=current_user.id,
            status="ATIVA",
        )
        db.session.add(nova)

        # o Kanban de Estações mostra o item pela estação DELE (item.estacao), não
        # pela estação gravada na programação — então programar pra uma estação
        # diferente da atual também reatribui o item, pra ele aparecer no board certo
        if estacao_escolhida and estacao_escolhida != "—" and item.estacao != estacao_escolhida:
            antes = {c: getattr(item, c) for c in CAMPOS_HISTORICO_ITEM}
            item.estacao = estacao_escolhida
            depois = {c: getattr(item, c) for c in CAMPOS_HISTORICO_ITEM}
            _registrar_alteracoes("item_pedido", item.id, item.pedido_id, antes, depois, CAMPOS_HISTORICO_ITEM)

        db.session.commit()
        flash(f"\"{item.descricao_produto}\" programado para {data_programada.strftime('%d/%m/%Y')}.", "success")
        return redirect(url_for("programacao_semana", semana=data_programada.isoformat()))

    @app.route("/programacao/<int:programacao_id>/reprogramar", methods=["POST"])
    @requer_role("ADMIN", "PCP")
    def programacao_reprogramar(programacao_id):
        atual = db.session.get(Programacao, programacao_id)
        if atual is None or atual.status != "ATIVA":
            flash("Programação não encontrada ou já não está mais ativa.", "danger")
            return redirect(url_for("programacao_semana"))

        nova_data = _parse_data_form(request.form.get("data_programada"))
        if nova_data is None:
            flash("Informe uma nova data válida para reprogramar.", "danger")
            return redirect(url_for("programacao_semana"))

        atual.status = "REPROGRAMADA"
        nova = Programacao(
            item_pedido_id=atual.item_pedido_id,
            data_programada=nova_data,
            estacao=atual.estacao,
            prioridade_producao=atual.prioridade_producao,
            observacao=atual.observacao,
            responsavel_id=atual.responsavel_id,
            criado_por_id=current_user.id,
            status="ATIVA",
        )
        db.session.add(nova)
        db.session.commit()
        flash(f"Reprogramado para {nova_data.strftime('%d/%m/%Y')}.", "success")
        return redirect(url_for("programacao_semana", semana=nova_data.isoformat()))

    @app.route("/programacao/<int:programacao_id>/cancelar", methods=["POST"])
    @requer_role("ADMIN", "PCP")
    def programacao_cancelar(programacao_id):
        atual = db.session.get(Programacao, programacao_id)
        if atual is not None and atual.status == "ATIVA":
            atual.status = "CANCELADA"
            db.session.commit()
            flash("Programação cancelada.", "info")
        return redirect(url_for("programacao_semana"))


app = create_app()

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
