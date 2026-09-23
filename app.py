import csv
import io
import math
import os
import re
import unicodedata
from calendar import monthrange
from datetime import date, datetime, timedelta
from itertools import zip_longest

from flask import Flask, Response, abort, flash, jsonify, redirect, render_template, request, session, url_for
from flask_login import current_user, login_required, login_user, logout_user
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from sqlalchemy import and_, case, extract, false, func, inspect, not_, or_, text
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import selectinload

from extensions import db, login_manager
from models import (
    ESTACOES,
    ESTACOES_GRUPOS_MONITORAMENTO,
    FRETE_OPCOES,
    GO_OTD_META_PERCENTUAL,
    GO_STATUS_PEDIDO_INFO_CORES,
    GO_STATUS_PEDIDO_INFO_OPCOES,
    GO_TIPO_PEDIDO_OPCOES,
    LEAD_TIME_MODALIDADE_OPCOES,
    LEAD_TIME_UNIDADE_OPCOES,
    PD_CATEGORIA_OPCOES,
    PD_ETAPA_CORES,
    PD_ETAPA_OPCOES,
    PD_RESULTADO_ESPERADO_OPCOES,
    PD_TESTE_RESULTADO_INFO,
    PD_TESTE_RESULTADO_OPCOES,
    PD_TIPO_EVENTO_OPCOES,
    PRAZO_ALERTA_DIAS,
    PRIORIDADE_CORES,
    PRIORIDADE_OPCOES,
    RDIM_CATEGORIA_DESVIO_OPCOES,
    RDIM_COMPONENTE_LBD_OPCOES,
    RDIM_ESTACOES_OPCOES,
    RDIM_GRANDEZAS_PADRAO,
    RDIM_INSPECAO_VISUAL_OPCOES,
    RDIM_RESULTADO_CORES,
    RDIM_RESULTADO_LABELS,
    RDIM_RESULTADO_OPCOES,
    RDIM_SUBCATEGORIA_DESVIO_OPCOES,
    RDIM_TIPO_PRODUTO_LABELS,
    RDIM_TIPO_PRODUTO_OPCOES,
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
    UFS_BRASIL,
    ControleSistema,
    CorrespondenciaManualCusto,
    Estacao,
    EstruturaProduto,
    EstruturaProdutoItem,
    HistoricoAlteracao,
    InspecaoFinal,
    ItemPedido,
    KpiGerencialMensal,
    LeadTimeProducao,
    LeadTimeProducaoHistorico,
    LeadTimeTransportadora,
    MateriaPrima,
    MateriaPrimaHistorico,
    ParametroHoraHomem,
    ParametroHoraHomemHistorico,
    Pedido,
    PedidoOperacao,
    Produto,
    Programacao,
    ProjetoPD,
    RdimComponenteDesvio,
    RdimMedicao,
    RdimPecaDesvio,
    RncQualidade,
    TesteProjetoPD,
    Transportadora,
    Usuario,
    VisitaReuniaoPD,
    gerar_semanas_pcp,
    rotulo_estacao,
)
from permissoes import ROLES, ROLES_LABELS, pode_acessar_endpoint, pode_editar_estacao, requer_role

MESES_PT = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
MESES_PT_EXTENSO = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
]

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

# Campos de P&D (Fase 14) que geram histórico de alteração — em especial
# etapa_atual, pra que toda troca de etapa (inclusive "andar pra trás", ex.:
# Teste -> Reprovado -> Desenvolvimento) fique registrada, como o Bruno pediu.
CAMPOS_HISTORICO_PD = [
    "etapa_atual",
    "percentual_conclusao",
    "prioridade",
    "responsavel",
    "data_prevista_conclusao",
    "data_real_conclusao",
    "custo_realizado",
    "investimento_realizado",
    "economia_realizada",
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
    # "pcp" apontava pra rota própria da aba PCP, apagada a pedido do Bruno
    # (16/09/2026: "já tenho todo o grupo Gestão Produção completa" —
    # redundante). O FORMULÁRIO de edição da seção PCP continua existindo
    # (fallback manual pra pedido ainda não lançado em Produção), só o botão
    # "Cancelar" dele volta pra Operação 360 agora.
    "pcp": "gestao_operacao_listagem_geral",
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
        _atualizar_estacoes_planilha_03_09_2026(app)
        # Depende da reconciliação acima já ter rodado (garante que Reforma/
        # Revenda existem no cadastro pra serem excluídas, e que Manutenção /
        # Devolução já existe pra ser renomeada).
        _organizar_estacoes_03_09_2026(app)
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
        _backfill_go_data_solicitada_cliente_retira(app)
        _migrar_rdim_inspecao_final(app)
        _migrar_rdim_pecas_desvio(app)
        _migrar_rdim_tipo_produto(app)
        # Roda por último de todos: depende de tudo acima (Pedido/ItemPedido
        # com todas as colunas migradas, PedidoOperacao já existindo).
        _sincronizar_planilha_producao_03_09_2026(app)
        # Mesmo motivo do sync 28/08 acima, só que com a planilha mais nova
        # que o Bruno mandou em 03/09 (mesmo layout de coluna, bloco
        # "Resultados" novo) — depende só de PedidoOperacao já existir.
        _sincronizar_gestao_operacao_03_09_2026(app)
        # Correção pontual pedida pelo Bruno (04/09/2026) — precisa rodar
        # depois da sincronização acima (é o bug dela que corrige).
        _corrigir_colisao_barra_gestao_operacao(app)
        _seed_usuario_pd_gustavo(app)
        _seed_usuarios_pcp_fabiano_daniel(app)
        _seed_lead_time_transportadora(app)
        _seed_parametro_hora_homem(app)
        _seed_custos_pig_mandril(app)
        # Depende do seed acima já ter criado os produtos/estruturas/
        # placeholders da família PU CAST.
        _migrar_pu_cast_decompor_quimica(app)
        _seed_custos_espuma(app)
        _seed_custos_superflex_silicone(app)
        _importar_historico_custos_manual(app)
        _seed_custos_pig_alojamento(app)
        _seed_custos_pig_calandra(app)
        # Depende de todos os seeds de matéria-prima acima já terem rodado
        # (precisa do catálogo completo pra classificar) — roda em todo
        # boot, não só uma vez (ver docstring da função).
        _migrar_materia_prima_origem_planilha(app)
        # Depende dos seeds de espuma (H/HS/HL/HDISC/HLR.../HLB/HLCC) acima
        # já terem rodado — roda em todo boot (ver docstring da função).
        _migrar_densidade_estrutura_produto(app)

    # Filtro Jinja "normalizar_pedido_venda" (pedido do Bruno, 10/09/2026):
    # mesma normalização usada no casamento Produção<->Operação em Python
    # (_normalizar_pedido_venda), disponível nos templates pra 2 coisas — (1)
    # montar a MESMA chave usada pelos dicts *_por_pedido_venda passados pro
    # template (ex. status_real_por_pedido_venda.get(p.pedido_venda |
    # normalizar_pedido_venda)) e (2) exibir o nº do pedido sem zero à
    # esquerda (ex. "000872" -> "872") — ele reclamou vendo o número com
    # zero antes na coluna "Pedido" da Operação 360.
    app.jinja_env.filters["normalizar_pedido_venda"] = _normalizar_pedido_venda

    # Filtro Jinja "moeda_brl" (módulo Gestão de Custos, 20/09/2026): mesma
    # formatação R$ 1.234,56 já repetida manualmente em vários templates
    # (ex. gargalos.html) — aqui vira filtro reaproveitável, pros vários
    # valores monetários das telas novas do módulo de custos.
    def _moeda_brl(valor):
        if valor is None:
            return "—"
        return "R$ " + "{:,.2f}".format(valor).replace(",", "X").replace(".", ",").replace("X", ".")

    app.jinja_env.filters["moeda_brl"] = _moeda_brl

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
            # Pedido do Bruno (04/09/2026): no filtro "Planejamento semanal
            # (PCP)" da Listagem Geral de Produção, incluir também as semanas
            # de meses mais antigos — "considerar de janeiro/2026 em diante".
            # gerar_semanas_pcp() por padrão só olha 1 mês pra trás; aqui
            # calculamos quantos meses atrás fica janeiro/2026 (cresce sozinho
            # com o tempo, sempre alcançando jan/2026 mesmo daqui a alguns
            # meses) e mantemos os 6 meses à frente já usados antes. Esse
            # mesmo GO_SEMANAS_PCP também alimenta os dropdowns de
            # "Planejamento semanal"/"Término Semanal PCP" nas telas de
            # editar pedido (Produção e Gestão Operação) — ganhar mais opções
            # antigas ali também é positivo, não só um efeito colateral.
            GO_SEMANAS_PCP=gerar_semanas_pcp(
                meses_atras=max(1, (date.today().year - 2026) * 12 + (date.today().month - 1)),
                meses_frente=6,
            ),
            GO_TIPO_PEDIDO_OPCOES=GO_TIPO_PEDIDO_OPCOES,
            GO_STATUS_PEDIDO_INFO_OPCOES=GO_STATUS_PEDIDO_INFO_OPCOES,
            GO_STATUS_PEDIDO_INFO_CORES=GO_STATUS_PEDIDO_INFO_CORES,
            UFS_BRASIL=UFS_BRASIL,
            LEAD_TIME_MODALIDADE_OPCOES=LEAD_TIME_MODALIDADE_OPCOES,
            LEAD_TIME_UNIDADE_OPCOES=LEAD_TIME_UNIDADE_OPCOES,
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
            PD_CATEGORIA_OPCOES=PD_CATEGORIA_OPCOES,
            PD_ETAPA_OPCOES=PD_ETAPA_OPCOES,
            PD_ETAPA_CORES=PD_ETAPA_CORES,
            PD_RESULTADO_ESPERADO_OPCOES=PD_RESULTADO_ESPERADO_OPCOES,
            PD_TIPO_EVENTO_OPCOES=PD_TIPO_EVENTO_OPCOES,
            PD_TESTE_RESULTADO_OPCOES=PD_TESTE_RESULTADO_OPCOES,
            PD_TESTE_RESULTADO_INFO=PD_TESTE_RESULTADO_INFO,
            RDIM_ESTACOES_OPCOES=RDIM_ESTACOES_OPCOES,
            RDIM_RESULTADO_OPCOES=RDIM_RESULTADO_OPCOES,
            RDIM_RESULTADO_LABELS=RDIM_RESULTADO_LABELS,
            RDIM_RESULTADO_CORES=RDIM_RESULTADO_CORES,
            RDIM_INSPECAO_VISUAL_OPCOES=RDIM_INSPECAO_VISUAL_OPCOES,
            RDIM_CATEGORIA_DESVIO_OPCOES=RDIM_CATEGORIA_DESVIO_OPCOES,
            RDIM_SUBCATEGORIA_DESVIO_OPCOES=RDIM_SUBCATEGORIA_DESVIO_OPCOES,
            RDIM_GRANDEZAS_PADRAO=RDIM_GRANDEZAS_PADRAO,
            RDIM_TIPO_PRODUTO_OPCOES=RDIM_TIPO_PRODUTO_OPCOES,
            RDIM_TIPO_PRODUTO_LABELS=RDIM_TIPO_PRODUTO_LABELS,
            RDIM_COMPONENTE_LBD_OPCOES=RDIM_COMPONENTE_LBD_OPCOES,
            hoje_iso=date.today().isoformat(),
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


_CHAVE_ESTACOES_PLANILHA_03_09_2026 = "atualizacao_estacoes_planilha_03_09_2026"
_ESTACOES_NOVAS_03_09_2026 = ["SOBRESSALENTE METAL MECANICA", "SOBRESSALENTE BORRACHA", "MANUTENÇÃO / DEVOLUÇÃO"]
_ESTACOES_DESATIVADAS_03_09_2026 = ["REFORMA", "REVENDA"]


def _atualizar_estacoes_planilha_03_09_2026(app):
    """Reconcilia a tabela `estacoes` (cadastro) com a nova lista ESTACOES —
    pedido do Bruno (03/09/2026): "considere exatamente as mesmas estações da
    planilha... revenda não vai existir mais, sobressalentes de borracha e
    metal mecânica passarão a existir". `_migrar_estacoes_tabela` só semeia
    a tabela quando ela está VAZIA, então em produção (já populada desde o
    primeiro boot) mudar a constante ESTACOES sozinha não reflete nos
    registros existentes — precisa desta migração à parte, guardada por
    ControleSistema pra rodar exatamente uma vez.

    Desativa (ativo=False) em vez de apagar — mesmo critério já usado no
    cadastro manual de estações (não existe rota de exclusão física, só
    ativar/desativar), então nenhum pedido antigo que porventura já tenha
    usado "Reforma"/"Revenda" fica com referência quebrada."""
    if ControleSistema.query.filter_by(chave=_CHAVE_ESTACOES_PLANILHA_03_09_2026).first() is not None:
        return

    for nome in _ESTACOES_DESATIVADAS_03_09_2026:
        estacao = Estacao.query.filter_by(nome=nome).first()
        if estacao is not None and estacao.ativo:
            estacao.ativo = False

    maior_ordem = db.session.query(func.max(Estacao.ordem_exibicao)).scalar() or 0
    for i, nome in enumerate(_ESTACOES_NOVAS_03_09_2026):
        if Estacao.query.filter_by(nome=nome).first() is None:
            maior_ordem += 1
            db.session.add(Estacao(nome=nome, ordem_exibicao=maior_ordem, ativo=True))

    db.session.add(ControleSistema(chave=_CHAVE_ESTACOES_PLANILHA_03_09_2026))
    db.session.commit()
    app.logger.info(
        "Migração automática: estações atualizadas conforme planilha 03/09/2026 — "
        "%s desativadas, %s criadas.",
        ", ".join(_ESTACOES_DESATIVADAS_03_09_2026), ", ".join(_ESTACOES_NOVAS_03_09_2026),
    )


_CHAVE_ORGANIZAR_ESTACOES_03_09_2026 = "organizar_estacoes_03_09_2026"
_ESTACOES_EXCLUIDAS_03_09_2026 = ["REFORMA", "REVENDA"]
_ESTACAO_RENOMEADA_03_09_2026 = ("MANUTENÇÃO / DEVOLUÇÃO", "MANUTENÇÃO")


def _organizar_estacoes_03_09_2026(app):
    """Segunda rodada de ajuste no cadastro de Estações, mesmo dia da
    reconciliação com a planilha (`_atualizar_estacoes_planilha_03_09_2026`),
    agora a pedido direto do Bruno na tela de Estações:
      1. Exclui de vez "Reforma" e "Revenda" — antes só tinham sido
         desativadas (não existia pedido pra excluir ainda). Como
         `ItemPedido.estacao` é texto solto, sem chave estrangeira pra
         `Estacao` (ver docstring do model), apagar a linha do cadastro é
         seguro — não quebra nenhum pedido antigo que porventura tenha usado
         esses nomes, só deixa de aparecer no cadastro/dropdowns.
      2. Renomeia "Manutenção / Devolução" pra só "Manutenção", tanto no
         cadastro quanto em qualquer ItemPedido já gravado com o nome antigo
         (pra não sobrar pedido usando um nome de estação que não existe
         mais em lugar nenhum do site).

    Guardado por ControleSistema pra rodar exatamente uma vez, mesmo padrão
    de todas as outras migrações de estações."""
    if ControleSistema.query.filter_by(chave=_CHAVE_ORGANIZAR_ESTACOES_03_09_2026).first() is not None:
        return

    excluidas = []
    for nome in _ESTACOES_EXCLUIDAS_03_09_2026:
        estacao = Estacao.query.filter_by(nome=nome).first()
        if estacao is not None:
            db.session.delete(estacao)
            excluidas.append(nome)

    nome_antigo, nome_novo = _ESTACAO_RENOMEADA_03_09_2026
    estacao_manutencao = Estacao.query.filter_by(nome=nome_antigo).first()
    if estacao_manutencao is not None:
        estacao_manutencao.nome = nome_novo

    itens_renomeados = ItemPedido.query.filter_by(estacao=nome_antigo).update(
        {"estacao": nome_novo}, synchronize_session=False
    )

    db.session.add(ControleSistema(chave=_CHAVE_ORGANIZAR_ESTACOES_03_09_2026))
    db.session.commit()
    app.logger.info(
        "Migração automática: estações organizadas (pedido Bruno 03/09/2026) — "
        "excluídas: %s | renomeada: '%s' -> '%s' (%d item(ns) atualizado(s)).",
        ", ".join(excluidas) or "nenhuma", nome_antigo, nome_novo, itens_renomeados,
    )


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


_CHAVE_SINCRONIZACAO_03_09_2026 = "sincronizacao_planilha_producao_03_09_2026"


def _sincronizar_planilha_producao_03_09_2026(app):
    """Sincroniza Gestão Produção (Pedido/ItemPedido) com a aba "GERAL TESTE"
    da nova planilha enviada pelo Bruno em 03/09/2026 ("03_09 CONTROLE
    PRODUCAO_V1.xlsx", layout de colunas diferente da de 25/08 — ver
    sincronizar_planilha_producao.py) e, na sequência, cria registros básicos
    em Gestão Operação (PedidoOperacao) pros pedidos desta planilha que ainda
    não têm nenhum lá — pedido do Bruno (03/09/2026, AskUserQuestion):
    "extraia todos os dados da planilha anexa... distribua em todo o
    aplicativo, devidamente para cada area (produção e operação)", com o
    reforço explícito de NÃO tocar em Qualidade nem P&D (por isso esta função
    só chama os dois sincronizadores de Produção/Operação, nada de RNC/RDIM/
    ProjetoPD).

    Protegida por `ControleSistema` (mesmo padrão de
    `_sincronizar_planilha_producao_25_08_2026`) — roda por cima de dados que
    já existem, precisa rodar exatamente uma vez. Se a validação de
    cabeçalhos da planilha falhar (`stats["cabecalhos_invalidos"]`), NÃO
    marca o ControleSistema como concluído — loga um erro e sai, pra dar
    outra chance de rodar automaticamente depois que o mapeamento de colunas
    (COL) for corrigido, sem precisar mexer manualmente no banco. Nunca
    lança exceção: um erro aqui não pode derrubar o boot do site inteiro."""
    if ControleSistema.query.filter_by(chave=_CHAVE_SINCRONIZACAO_03_09_2026).first() is not None:
        return

    xlsx_path = os.path.join(BASE_DIR, "data", "sincronizacao_producao_03_09_2026.xlsx")
    if not os.path.exists(xlsx_path):
        return

    from sincronizar_planilha_producao import seed_pedidos_operacao_basico, sincronizar_planilha_producao

    try:
        stats = sincronizar_planilha_producao(xlsx_path)
        if stats.get("cabecalhos_invalidos"):
            app.logger.error(
                "Sincronização planilha 03/09/2026 ABORTADA — cabeçalhos da planilha não "
                "batem com o mapeamento esperado (COL): %s",
                " | ".join(stats["cabecalhos_invalidos"]),
            )
            return

        stats_operacao = seed_pedidos_operacao_basico(xlsx_path)
        if stats_operacao.get("cabecalhos_invalidos"):
            # Não deveria acontecer (mesma planilha já validou acima), mas se
            # acontecer não descarta o que a sincronização de Produção já
            # commitou — só loga e segue sem os registros de Operação.
            app.logger.error(
                "Seed de PedidoOperacao (planilha 03/09/2026) ABORTADO — cabeçalhos não bateram: %s",
                " | ".join(stats_operacao["cabecalhos_invalidos"]),
            )
            stats_operacao = {"pedidos_operacao_criados": 0, "pedidos_operacao_ja_existentes": 0}
    except Exception:
        db.session.rollback()
        app.logger.exception("Sincronização planilha 03/09/2026 falhou com erro inesperado.")
        return

    db.session.add(ControleSistema(chave=_CHAVE_SINCRONIZACAO_03_09_2026))
    db.session.commit()
    app.logger.info(
        "Sincronização planilha 03/09/2026: %d linhas | %d pedidos atualizados | "
        "%d pedidos criados | %d itens atualizados | %d itens criados | "
        "%d pedidos sem cliente ignorados | Operação: %d PedidoOperacao criados, "
        "%d já existentes.",
        stats["linhas_lidas"],
        stats["pedidos_atualizados"],
        stats["pedidos_criados"],
        stats["itens_atualizados"],
        stats["itens_criados"],
        len(stats["pedidos_sem_cliente_ignorados"]),
        stats_operacao["pedidos_operacao_criados"],
        stats_operacao["pedidos_operacao_ja_existentes"],
    )
    if stats["valores_nao_reconhecidos"]:
        app.logger.warning(
            "Sincronização planilha 03/09/2026 — valores não reconhecidos (não importados "
            "automaticamente, checar manualmente): %s",
            {campo: list(valores.keys()) for campo, valores in stats["valores_nao_reconhecidos"].items()},
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


_CHAVE_CORRECAO_COLISAO_BARRA_GO_04_09_2026 = "correcao_colisao_barra_gestao_operacao_04_09_2026"


def _corrigir_colisao_barra_gestao_operacao(app):
    """Correção pontual (pedido do Bruno, 04/09/2026: "Valor liberado no mês"
    divergindo da planilha em abril/maio/junho) — reprocessa a MESMA planilha
    "03_09 Gestão de Fluxo Produtivo 2026" (já sincronizada uma vez em
    _sincronizar_gestao_operacao_03_09_2026), agora com o bug de casamento
    aproximado por barra já corrigido em `_match_pedido`
    (importar_gestao_operacao.py — ver comentário lá).

    Por que basta reprocessar o mesmo arquivo: na 1ª sincronização, uma linha
    tipo "492/1" colidia (errado) com o pedido "492" já cadastrado e
    SOBRESCREVIA os campos dele — "492" ficou no banco com os dados de
    "492/1" (valor, término semanal etc.), e "492/1" nunca virou um registro
    próprio. Rodando de novo, na MESMA ordem de linhas, com o casamento por
    barra desativado: a linha "492" (que vem antes na planilha) bate exato de
    novo no pedido "492" e devolve os campos dele ao normal; a linha "492/1"
    não bate mais em nada por aproximação, então vira um PedidoOperacao NOVO,
    com os dados que estavam perdidos. Sem duplicar nada: pedidos que já
    batiam certo na 1ª sincronização continuam batendo exato e não mudam
    (update sem diferença nenhuma).

    Roda só uma vez (ControleSistema), e nunca lança exceção (mesmo padrão de
    _sincronizar_gestao_operacao_03_09_2026 — um erro aqui não pode derrubar
    o boot do site)."""
    if ControleSistema.query.filter_by(chave=_CHAVE_CORRECAO_COLISAO_BARRA_GO_04_09_2026).first() is not None:
        return

    xlsx_path = os.path.join(BASE_DIR, "data", "sincronizacao_gestao_operacao_03_09_2026.xlsx")
    if not os.path.exists(xlsx_path):
        return

    from sincronizar_gestao_operacao import sincronizar_gestao_operacao

    try:
        stats = sincronizar_gestao_operacao(xlsx_path)
    except Exception:
        db.session.rollback()
        app.logger.exception("Correção de colisão por barra (Gestão Operação) falhou com erro inesperado.")
        return

    db.session.add(ControleSistema(chave=_CHAVE_CORRECAO_COLISAO_BARRA_GO_04_09_2026))
    db.session.commit()
    app.logger.info(
        "Correção de colisão por barra (Gestão Operação) 04/09/2026: %d linhas | %d pedidos atualizados | "
        "%d pedidos criados (eram os que tinham sumido, tipo NNN/1) | %d campos atualizados | "
        "%d casamentos exatos | %d aproximados (deve ser 0 pra valores com barra) | %d sem match.",
        stats["linhas_lidas"],
        stats["pedidos_atualizados"],
        stats["pedidos_criados"],
        stats["campos_atualizados"],
        stats["exato"],
        stats["aproximado"],
        stats["sem_match_pedido_venda"],
    )


_CHAVE_SINCRONIZACAO_GO_03_09_2026 = "sincronizacao_gestao_operacao_03_09_2026"


def _sincronizar_gestao_operacao_03_09_2026(app):
    """Sincroniza Gestão Operação (PedidoOperacao) com a planilha "03_09 Gestão
    de Fluxo Produtivo 2026" enviada pelo Bruno — mesmo formato de coluna já
    usado em 28/08/2026 (conferido coluna a coluna antes de reaproveitar
    sincronizar_gestao_operacao.py), com o bloco novo "Resultados" (valor NF/
    custo produção/custo frete, mais completo — ver _coalesce_numero em
    sincronizar_gestao_operacao.py) e rastreio de criticidade/OTD fora do
    esperado. Protegido por `ControleSistema` — roda exatamente uma vez, e
    nunca lança exceção (um erro aqui não pode derrubar o boot do site
    inteiro)."""
    if ControleSistema.query.filter_by(chave=_CHAVE_SINCRONIZACAO_GO_03_09_2026).first() is not None:
        return

    xlsx_path = os.path.join(BASE_DIR, "data", "sincronizacao_gestao_operacao_03_09_2026.xlsx")
    if not os.path.exists(xlsx_path):
        return

    from sincronizar_gestao_operacao import sincronizar_gestao_operacao

    try:
        stats = sincronizar_gestao_operacao(xlsx_path)
    except Exception:
        db.session.rollback()
        app.logger.exception("Sincronização Gestão Operação 03/09/2026 falhou com erro inesperado.")
        return

    db.session.add(ControleSistema(chave=_CHAVE_SINCRONIZACAO_GO_03_09_2026))
    db.session.commit()
    app.logger.info(
        "Sincronização Gestão Operação 03/09/2026: %d linhas | %d pedidos atualizados | "
        "%d pedidos criados | %d campos atualizados | %d casamentos exatos | %d aproximados | "
        "%d sem match de pedido_venda (viraram novos).",
        stats["linhas_lidas"],
        stats["pedidos_atualizados"],
        stats["pedidos_criados"],
        stats["campos_atualizados"],
        stats["exato"],
        stats["aproximado"],
        stats["sem_match_pedido_venda"],
    )
    if stats["criticidades_nao_reconhecidas"]:
        app.logger.warning(
            "Sincronização Gestão Operação 03/09/2026 — valores de CRITICIDADE não reconhecidos "
            "(prioridade não atualizada automaticamente nessas linhas, checar manualmente): %s",
            stats["criticidades_nao_reconhecidas"],
        )
    if stats["otd_nao_reconhecidos"]:
        app.logger.warning(
            "Sincronização Gestão Operação 03/09/2026 — valores de OTD fora do padrão SIM/NÃO "
            "(gravados como texto truncado, checar manualmente): %s",
            stats["otd_nao_reconhecidos"],
        )


_CHAVE_BACKFILL_SOLICITADA_CLIENTE_RETIRA_01_09_2026 = "backfill_go_data_solicitada_cliente_retira_01_09_2026"


def _backfill_go_data_solicitada_cliente_retira(app):
    """Correção pontual pedida pelo Bruno em 01/09/2026: "Solicitada cliente/
    retira" (PCP) tem que acompanhar a mesma data que "Data solicitada
    entrega" (Comercial) — ambas vêm de Pedido.data_cliente na criação
    automática (ver _criar_pedido_operacao_a_partir_de_producao), mas essa
    ligação só foi adicionada agora; os pedidos criados pelo fluxo automático
    ANTES desta correção ficaram com "Solicitada cliente/retira" em branco.
    Preenche só isso — nunca sobrescreve um valor que já tenha sido digitado
    manualmente em PCP (só mexe onde está None)."""
    if ControleSistema.query.filter_by(chave=_CHAVE_BACKFILL_SOLICITADA_CLIENTE_RETIRA_01_09_2026).first() is not None:
        return

    pedidos = PedidoOperacao.query.filter(
        PedidoOperacao.go_data_solicitada_cliente_retira.is_(None),
        PedidoOperacao.go_data_solicitada_entrega.isnot(None),
    ).all()
    for pedido in pedidos:
        pedido.go_data_solicitada_cliente_retira = pedido.go_data_solicitada_entrega

    app.logger.info(
        "Backfill pontual: 'Solicitada cliente/retira' preenchida em %d pedido(s) de Gestão Operação.",
        len(pedidos),
    )
    db.session.add(ControleSistema(chave=_CHAVE_BACKFILL_SOLICITADA_CLIENTE_RETIRA_01_09_2026))
    db.session.commit()


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


def _migrar_rdim_inspecao_final(app):
    """Adiciona os 2 campos novos da Fase 2 do RDIM (pedido do Bruno,
    02/09/2026, depois de já usar a área) na tabela `inspecoes_finais` —
    que já existe em produção desde a Fase 1, então essas colunas não são
    cobertas só por `db.create_all()` (que só cria tabelas novas). Mesmo
    padrão de `_migrar_faturamento_itens`: campos 100% novos e opcionais,
    ficam em branco nas inspeções já registradas."""
    inspector = inspect(db.engine)
    if "inspecoes_finais" not in inspector.get_table_names():
        return

    colunas = {c["name"] for c in inspector.get_columns("inspecoes_finais")}
    faltando = [c for c in ("subcategoria_desvio", "quantidade_com_desvio") if c not in colunas]
    if not faltando:
        return

    with db.engine.begin() as conn:
        if "subcategoria_desvio" not in colunas:
            conn.execute(text("ALTER TABLE inspecoes_finais ADD COLUMN subcategoria_desvio VARCHAR(40)"))
        if "quantidade_com_desvio" not in colunas:
            conn.execute(text("ALTER TABLE inspecoes_finais ADD COLUMN quantidade_com_desvio FLOAT"))
    app.logger.info("Migração automática: campos subcategoria_desvio e quantidade_com_desvio adicionados em inspecoes_finais.")


def _migrar_rdim_pecas_desvio(app):
    """Adiciona os 2 campos novos da Fase 4 do RDIM (pedido do Bruno,
    02/09/2026: contexto do desvio por peça — espec. mín/máx x medido) na
    tabela `rdim_pecas_desvio` — que já existe em produção desde a Fase 3
    (feature "Apontamentos por peça"), então essas colunas não são cobertas
    só por `db.create_all()`. Mesmo padrão de `_migrar_rdim_inspecao_final`:
    campos 100% novos e opcionais, ficam em branco nos apontamentos já
    registrados."""
    inspector = inspect(db.engine)
    if "rdim_pecas_desvio" not in inspector.get_table_names():
        return

    colunas = {c["name"] for c in inspector.get_columns("rdim_pecas_desvio")}
    faltando = [c for c in ("especificado_min", "especificado_max") if c not in colunas]
    if not faltando:
        return

    with db.engine.begin() as conn:
        if "especificado_min" not in colunas:
            conn.execute(text("ALTER TABLE rdim_pecas_desvio ADD COLUMN especificado_min FLOAT"))
        if "especificado_max" not in colunas:
            conn.execute(text("ALTER TABLE rdim_pecas_desvio ADD COLUMN especificado_max FLOAT"))
    app.logger.info("Migração automática: campos especificado_min e especificado_max adicionados em rdim_pecas_desvio.")


def _migrar_rdim_tipo_produto(app):
    """Adiciona o campo novo da Fase 5 do RDIM (pedido do Bruno, 11/09/2026:
    Discos/FlexPig no modelo atual, PIG LBD/LUN/SUPERFLEX no modelo por
    componente) na tabela `inspecoes_finais` — que já existe em produção,
    então essa coluna não é coberta só por `db.create_all()`. Mesmo padrão de
    `_migrar_rdim_inspecao_final`: campo novo e opcional, fica em branco nas
    inspeções já registradas (nenhum backfill)."""
    inspector = inspect(db.engine)
    if "inspecoes_finais" not in inspector.get_table_names():
        return

    colunas = {c["name"] for c in inspector.get_columns("inspecoes_finais")}
    if "tipo_produto_inspecionado" in colunas:
        return

    with db.engine.begin() as conn:
        conn.execute(text("ALTER TABLE inspecoes_finais ADD COLUMN tipo_produto_inspecionado VARCHAR(30)"))
    app.logger.info("Migração automática: campo tipo_produto_inspecionado adicionado em inspecoes_finais.")


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


def _seed_usuario_pd_gustavo(app):
    """Cria (uma única vez) o login do Gustavo Fugita, líder de P&D — pedido
    do Bruno (03/09/2026): acesso completo (visualização + edição) só na
    área de P&D, e só visualização em Gestão Produção > Estações (papel
    novo "PD" — ver permissoes.py, ENDPOINTS_PERMITIDOS_PD). Idempotente
    pelo username, mesmo padrão de _seed_inicial (usuário admin padrão) —
    se já existir (ex.: já rodou numa execução anterior, ou o Bruno já
    trocou a senha depois), não mexe em nada."""
    if Usuario.query.filter_by(username="gustavo.fugita").first() is not None:
        return
    gustavo = Usuario(nome="Gustavo Fugita", username="gustavo.fugita", role="PD", ativo=True)
    gustavo.set_senha("PD@Fugita2026!")
    db.session.add(gustavo)
    db.session.commit()
    app.logger.info("Usuário gustavo.fugita (P&D) criado — troque a senha depois de conferir o acesso.")


def _seed_usuarios_pcp_fabiano_daniel(app):
    """Cria (uma única vez, cada um independente) os logins do Fabiano e do
    Daniel, time de PCP — pedido do Bruno (04/09/2026): "acesso completo".
    Papel "PCP" já existente (não é um papel novo como o "PD" do Gustavo) —
    mesmo acesso que qualquer outro PCP: cadastra/edita pedidos, programação,
    produção de qualquer estação, cadastros de estação/transportadora,
    Gestão Operação (ver requer_role("ADMIN", "PCP") espalhado em app.py).
    Idempotente pelo username, mesmo padrão de _seed_usuario_pd_gustavo."""
    novos = [
        ("Fabiano", "fabiano.pcp", "PCP@Fabiano2026!"),
        ("Daniel", "daniel.pcp", "PCP@Daniel2026!"),
    ]
    for nome, username, senha in novos:
        if Usuario.query.filter_by(username=username).first() is not None:
            continue
        usuario = Usuario(nome=nome, username=username, role="PCP", ativo=True)
        usuario.set_senha(senha)
        db.session.add(usuario)
        db.session.commit()
        app.logger.info("Usuário %s (PCP) criado — troque a senha depois de conferir o acesso.", username)


_CHAVE_SEED_LEAD_TIME_TRANSPORTADORA_11_09_2026 = "seed_lead_time_transportadora_pindamonhangaba_11_09_2026"
_LEAD_TIME_ORIGEM_PADRAO = "Pindamonhangaba - SP"

# Tabela de prazos que o Bruno anexou (11/09/2026) — prazo de entrega por UF/
# modalidade de um parceiro logístico (a tabela original tinha "São Paulo
# (SP)" como referência de saída; ele pediu pra simular saindo de
# Pindamonhangaba-SP com os MESMOS prazos, sem recalcular nada). 1 linha por
# (uf, modalidade, prazo_minimo, prazo_maximo, unidade, observacao) — Bahia
# já entra como exceção própria (7d rodoviário / 4d aéreo), por isso as
# outras linhas do Nordeste não carregam mais a nota "BA: Xd" que a tabela
# original tinha (fica redundante com a linha própria da BA).
#
# Nota: a tabela original não lista Tocantins (TO) na região Norte — apesar
# de TO fazer parte do Norte no restante do sistema (REGIAO_POR_UF), ele NÃO
# foi cadastrado aqui pra não inventar um prazo que a tabela do Bruno não
# informou. Fica faltando de propósito — ver aviso no relatório de entrega.
_LEAD_TIME_TRANSPORTADORA_SEED = [
    # UF,  modalidade,   min, max,  unidade,       observação
    ("RJ", "Rodoviário", 48, 72, "Horas", None),
    ("SP", "Rodoviário", 48, 72, "Horas", None),
    ("MG", "Rodoviário", 5, 7, "Dias úteis", None),
    ("MG", "Aéreo", 4, None, "Dias úteis", None),
    ("ES", "Rodoviário", 5, 7, "Dias úteis", None),
    ("ES", "Aéreo", 4, None, "Dias úteis", None),
    ("SC", "Rodoviário", 5, 7, "Dias úteis", None),
    ("SC", "Aéreo", 4, None, "Dias úteis", None),
    ("PR", "Rodoviário", 5, 7, "Dias úteis", None),
    ("PR", "Aéreo", 4, None, "Dias úteis", None),
    ("RS", "Rodoviário", 5, 7, "Dias úteis", None),
    ("RS", "Aéreo", 4, None, "Dias úteis", None),
    ("GO", "Rodoviário", 10, None, "Dias úteis", None),
    ("GO", "Aéreo", 4, None, "Dias úteis", None),
    ("MT", "Rodoviário", 10, None, "Dias úteis", None),
    ("MT", "Aéreo", 4, None, "Dias úteis", None),
    ("MS", "Rodoviário", 10, None, "Dias úteis", None),
    ("MS", "Aéreo", 4, None, "Dias úteis", None),
    ("DF", "Rodoviário", 10, None, "Dias úteis", None),
    ("DF", "Aéreo", 4, None, "Dias úteis", None),
    ("AC", "Rodoviário", 22, None, "Dias úteis", None),
    ("AC", "Aéreo", 12, None, "Dias úteis", None),
    ("AP", "Rodoviário", 22, None, "Dias úteis", None),
    ("AP", "Aéreo", 12, None, "Dias úteis", None),
    ("AM", "Rodoviário", 22, None, "Dias úteis", None),
    ("AM", "Aéreo", 12, None, "Dias úteis", None),
    ("PA", "Rodoviário", 22, None, "Dias úteis", None),
    ("PA", "Aéreo", 12, None, "Dias úteis", None),
    ("RO", "Rodoviário", 22, None, "Dias úteis", None),
    ("RO", "Aéreo", 12, None, "Dias úteis", None),
    ("RR", "Rodoviário", 22, None, "Dias úteis", None),
    ("RR", "Aéreo", 12, None, "Dias úteis", None),
    ("AL", "Rodoviário", 15, None, "Dias úteis", None),
    ("AL", "Aéreo", 7, None, "Dias úteis", None),
    ("CE", "Rodoviário", 15, None, "Dias úteis", None),
    ("CE", "Aéreo", 7, None, "Dias úteis", None),
    ("MA", "Rodoviário", 15, None, "Dias úteis", None),
    ("MA", "Aéreo", 7, None, "Dias úteis", None),
    ("PB", "Rodoviário", 15, None, "Dias úteis", None),
    ("PB", "Aéreo", 7, None, "Dias úteis", None),
    ("PE", "Rodoviário", 15, None, "Dias úteis", None),
    ("PE", "Aéreo", 7, None, "Dias úteis", None),
    ("PI", "Rodoviário", 15, None, "Dias úteis", None),
    ("PI", "Aéreo", 7, None, "Dias úteis", None),
    ("RN", "Rodoviário", 15, None, "Dias úteis", None),
    ("RN", "Aéreo", 7, None, "Dias úteis", None),
    ("SE", "Rodoviário", 15, None, "Dias úteis", None),
    ("SE", "Aéreo", 7, None, "Dias úteis", None),
    ("BA", "Rodoviário", 7, None, "Dias úteis", "Exceção Nordeste"),
    ("BA", "Aéreo", 4, None, "Dias úteis", "Exceção Nordeste"),
]


def _seed_lead_time_transportadora(app):
    """Importa (uma única vez) a tabela de prazos de entrega por UF/
    modalidade que o Bruno anexou (11/09/2026), pro novo cadastro "Lead time
    Transportadora" — simulando frete saindo de Pindamonhangaba-SP. Guardado
    por `ControleSistema` (mesmo padrão de _seed_rnc_qualidade): roda
    exatamente uma vez — depois disso, os prazos ficam livres pra ele editar
    pela tela de cadastro sem risco de um próximo boot sobrescrever o ajuste."""
    if ControleSistema.query.filter_by(chave=_CHAVE_SEED_LEAD_TIME_TRANSPORTADORA_11_09_2026).first() is not None:
        return

    total = 0
    for uf, modalidade, minimo, maximo, unidade, observacao in _LEAD_TIME_TRANSPORTADORA_SEED:
        db.session.add(
            LeadTimeTransportadora(
                origem=_LEAD_TIME_ORIGEM_PADRAO,
                uf=uf,
                regiao=REGIAO_POR_UF.get(uf, ""),
                modalidade=modalidade,
                prazo_minimo=minimo,
                prazo_maximo=maximo,
                unidade_prazo=unidade,
                observacao=observacao,
                ativo=True,
            )
        )
        total += 1
    db.session.add(ControleSistema(chave=_CHAVE_SEED_LEAD_TIME_TRANSPORTADORA_11_09_2026))
    db.session.commit()
    app.logger.info("Lead time Transportadora: %d linhas cadastradas (origem %s).", total, _LEAD_TIME_ORIGEM_PADRAO)


def _seed_parametro_hora_homem(app):
    """Garante a linha singleton (id=1) de ParametroHoraHomem com o valor
    inicial de R$ 40,00/h pedido pelo Bruno (20/09/2026, módulo Gestão de
    Custos) — idempotente por natureza (só cria se ainda não existir
    nenhuma linha), sem precisar de flag em ControleSistema."""
    if db.session.get(ParametroHoraHomem, 1) is not None:
        return
    db.session.add(ParametroHoraHomem(id=1, valor=40.0))
    db.session.commit()
    app.logger.info("Gestão de Custos: parâmetro de Hora-Homem inicializado em R$ 40,00/h.")


_CHAVE_SEED_CUSTOS_PIG_MANDRIL_20_09_2026 = "seed_custos_pig_mandril_20_09_2026"
_CUSTOS_HH_RATE_SEED = 40.0  # valor vigente na planilha na data da importação — usado só pra DERIVAR ciclo_horas


def _seed_custos_pig_mandril(app):
    """Importa (uma única vez) a planilha de custos que o Bruno anexou
    (20/09/2026, `data/custo_de_producao_20_09_2026.xlsx`) pro novo módulo
    GESTÃO DE CUSTOS — fase 1, grupo PIG MANDRIL (LBD, LUN, PU CAST, CORPO
    MANDRIL, ELC_MG_PC, PIGS EM BORRACHA). Roda exatamente uma vez (mesmo
    padrão de `_seed_lead_time_transportadora`, guardado por
    ControleSistema) — depois disso os cadastros ficam livres pra edição
    manual sem risco de um próximo boot sobrescrever o ajuste.

    As fórmulas de cada família foram mapeadas a fundo (célula a célula, não
    só o texto explicativo da planilha, que tinha discrepâncias confirmadas
    contra a fórmula real) e o resultado foi validado 1:1 contra a aba
    BUSCA DE CUSTO / colunas de total de cada aba antes desta versão ir pro
    ar — ver `scripts/importar_custos_pig_mandril.py` (script irmão usado
    pra iterar/validar localmente) pra o relatório completo de verificação."""
    if ControleSistema.query.filter_by(chave=_CHAVE_SEED_CUSTOS_PIG_MANDRIL_20_09_2026).first() is not None:
        return

    xlsx_path = os.path.join(BASE_DIR, "data", "custo_de_producao_20_09_2026.xlsx")
    if not os.path.exists(xlsx_path):
        app.logger.warning("Gestão de Custos: planilha de importação não encontrada em %s — seed não executado.", xlsx_path)
        return

    import openpyxl

    def _dn_str(v):
        if v is None:
            return None
        if isinstance(v, float) and v == int(v):
            return str(int(v))
        return str(v).strip()

    def _dn_num_str(v):
        """Extrai só a parte numérica do DN, descartando aspas de polegada (ex.: "6''" -> "6") —
        as tabelas de PIGS EM BORRACHA em PARÂMETROS usam esse formato, diferente do resto."""
        if v is None:
            return None
        s = "".join(ch for ch in str(v).strip() if ch.isdigit() or ch == ".")
        if not s:
            return None
        try:
            f = float(s)
            return str(int(f)) if f == int(f) else str(f)
        except ValueError:
            return s

    mp_cache, produto_cache = {}, {}

    def _get_or_create_mp(codigo, descricao, unidade, custo, categoria):
        mp = mp_cache.get(codigo)
        if mp is not None:
            return mp
        mp = MateriaPrima.query.filter_by(codigo=codigo).first()
        if mp is None:
            mp = MateriaPrima(codigo=codigo, descricao=descricao, unidade=unidade, custo_atual=custo or 0, categoria=categoria, ativo=True)
            db.session.add(mp)
            db.session.flush()
        mp_cache[codigo] = mp
        return mp

    def _get_or_create_produto(familia, codigo, descricao=None, categoria=None, chave_busca=None):
        key = (familia, codigo)
        p = produto_cache.get(key)
        if p is not None:
            return p
        p = Produto.query.filter_by(familia=familia, codigo=codigo).first()
        if p is None:
            p = Produto(familia=familia, codigo=codigo, descricao=descricao, categoria=categoria, chave_busca=chave_busca, ativo=True)
            db.session.add(p)
            db.session.flush()
        produto_cache[key] = p
        return p

    def _get_or_create_estrutura(produto, dn, ciclo_horas):
        e = EstruturaProduto.query.filter_by(produto_id=produto.id, dn=dn).first()
        if e is None:
            e = EstruturaProduto(produto_id=produto.id, dn=dn, ciclo_horas=ciclo_horas or 0, ativo=True)
            db.session.add(e)
            db.session.flush()
        return e

    def _add_item(estrutura, ordem, tipo, quantidade, materia_prima=None, subproduto=None, observacao=None):
        if not quantidade:
            return
        db.session.add(EstruturaProdutoItem(
            estrutura_id=estrutura.id, tipo=tipo, quantidade=quantidade,
            materia_prima_id=materia_prima.id if materia_prima else None,
            subproduto_id=subproduto.id if subproduto else None,
            observacao=observacao, ordem=ordem,
        ))

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws_param = wb["PARÂMETROS"]

    # 1. Matérias-primas "químicas" — PARÂMETROS linhas 6-15
    quimicas_por_linha = {}
    for r in range(6, 16):
        desc = ws_param.cell(r, 2).value
        preco = ws_param.cell(r, 5).value
        if not desc:
            continue
        codigo = "QUIM-" + "".join(ch for ch in desc.upper() if ch.isalnum())[:30]
        quimicas_por_linha[r] = _get_or_create_mp(codigo, desc, "kg", preco, "Química")
    MP_PRE_TDI = quimicas_por_linha[6]  # 12-70 A (PRE, sistema TDI) — usado no bumper elastômero de LBD/LUN/CORPO MANDRIL

    # 2. Matérias-primas por DN — PARÂMETROS linhas 82-101 (LBD_REV A: tubo/flange bumper/flange solda/parafuso/arruela/porca)
    tubo_por_dn, flange_bumper_por_dn, flange_solda_por_dn = {}, {}, {}
    parafuso_por_dn, arruela_por_dn, porca_por_dn = {}, {}, {}
    for r in range(83, 102):
        dn = _dn_str(ws_param.cell(r, 1).value)
        if dn is None:
            continue
        tubo_por_dn[dn] = _get_or_create_mp(f"TUBO-DN{dn}", f"Tubo DN {dn}", "un", ws_param.cell(r, 2).value, "Componente DN")
        flange_bumper_por_dn[dn] = _get_or_create_mp(f"FLANGE-BUMPER-DN{dn}", f"Flange bumper DN {dn}", "un", ws_param.cell(r, 3).value, "Componente DN")
        flange_solda_por_dn[dn] = _get_or_create_mp(f"FLANGE-SOLDA-DN{dn}", f"Flange solda DN {dn}", "un", ws_param.cell(r, 4).value, "Componente DN")
        parafuso_por_dn[dn] = _get_or_create_mp(f"PARAFUSO-DN{dn}", f"Parafuso DN {dn}", "un", ws_param.cell(r, 5).value, "Componente DN")
        arruela_por_dn[dn] = _get_or_create_mp(f"ARRUELA-DN{dn}", f"Arruela DN {dn}", "un", ws_param.cell(r, 6).value, "Componente DN")
        porca_por_dn[dn] = _get_or_create_mp(f"PORCA-DN{dn}", f"Porca DN {dn}", "un", ws_param.cell(r, 7).value, "Componente DN")

    # 3. Matérias-primas PIGS EM BORRACHA — PARÂMETROS linhas 65-69 (copo por material) e 73-77 (componentes comuns)
    copo_epdm_por_dn, copo_buna_por_dn, copo_viton_por_dn = {}, {}, {}
    for r in range(65, 70):
        dn = _dn_num_str(ws_param.cell(r, 1).value)
        if dn is None:
            continue
        copo_epdm_por_dn[dn] = _get_or_create_mp(f"COPO-BORRACHA-EPDM-DN{dn}", f"Copo de borracha EPDM DN {dn}", "un", ws_param.cell(r, 2).value, "Componente DN")
        copo_buna_por_dn[dn] = _get_or_create_mp(f"COPO-BORRACHA-BUNA-DN{dn}", f"Copo de borracha BUNA N DN {dn}", "un", ws_param.cell(r, 3).value, "Componente DN")
        copo_viton_por_dn[dn] = _get_or_create_mp(f"COPO-BORRACHA-VITON-DN{dn}", f"Copo de borracha VITON DN {dn}", "un", ws_param.cell(r, 4).value, "Componente DN")

    eixo_por_dn, cabecote_por_dn, nylon_por_dn, porca_bor_por_dn, flange_bor_por_dn = {}, {}, {}, {}, {}
    hh_montagem_borracha = None
    for r in range(73, 78):
        dn = _dn_num_str(ws_param.cell(r, 1).value)
        if dn is None:
            continue
        eixo_por_dn[dn] = _get_or_create_mp(f"EIXO-BORRACHA-DN{dn}", f"Eixo (barra roscada) DN {dn}", "un", ws_param.cell(r, 2).value, "Componente DN")
        cabecote_por_dn[dn] = _get_or_create_mp(f"CABECOTE-BORRACHA-DN{dn}", f"Cabeçote PU DN {dn}", "un", ws_param.cell(r, 3).value, "Componente DN")
        nylon_por_dn[dn] = _get_or_create_mp(f"NYLON-BORRACHA-DN{dn}", f"De nylon preto (2X) DN {dn}", "un", ws_param.cell(r, 4).value, "Componente DN")
        porca_bor_por_dn[dn] = _get_or_create_mp(f"PORCA-BORRACHA-DN{dn}", f"Porca DN {dn}", "un", ws_param.cell(r, 5).value, "Componente DN")
        flange_bor_por_dn[dn] = _get_or_create_mp(f"FLANGE-BORRACHA-DN{dn}", f"Flange (2X) DN {dn}", "un", ws_param.cell(r, 6).value, "Componente DN")
        hh_montagem_borracha = ws_param.cell(r, 7).value

    # 4. Família PU CAST — DS, DG, DE, COPO CONICO, COPO PISTAO, HFLEX, DISCFLEX SD, DISCFLEX SDI.
    #    A coluna "CUSTO MP" de cada linha entra como 1 valor já consolidado por (item, DN) — não é
    #    peso×preço fixo por coluna (varia por linha, confirmado célula a célula) — preserva fidelidade
    #    exata ao valor da planilha em vez de arriscar uma decomposição incorreta. "BUMPER (2X)" fica de
    #    fora (a própria aba referencia a LBD pra esse valor, não tem BOM própria).
    ws_pu = wb["PU CAST"]
    col_map = {"ITEM": 27, "CUSTO_MP": 29, "CUSTO_HH_UNIT": 34, "DN_NUM": 37}
    # chave de busca pra casar com o texto livre dos pedidos do PCP (confirmado por amostragem real:
    # "DISCO SELO DN 12"", "PIG DISCFLEX SDI DN 6"" etc. — vendidos/produzidos como sobressalente avulso).
    chave_busca_pu_cast = {
        "DS": "DISCO SELO", "DG": "DISCO GUIA", "DE": "DISCO ESPAÇADOR",
        "COPO CONICO": "COPO CONICO", "COPO PISTAO": "COPO PIST", "HFLEX": "HFLEX",
        "DISCFLEX SD": "DISCFLEX SD", "DISCFLEX SDI": "DISCFLEX SDI",
    }
    pu_cast_estruturas = {}
    for r in range(10, ws_pu.max_row + 1):
        item = ws_pu.cell(r, col_map["ITEM"]).value
        dn_num = ws_pu.cell(r, col_map["DN_NUM"]).value
        if item is None or item == "BUMPER (2X)" or dn_num in (None, ""):
            continue
        dn = _dn_str(dn_num)
        custo_mp = ws_pu.cell(r, col_map["CUSTO_MP"]).value or 0
        custo_hh_unit = ws_pu.cell(r, col_map["CUSTO_HH_UNIT"]).value or 0
        ciclo_horas = round(custo_hh_unit / _CUSTOS_HH_RATE_SEED, 6) if custo_hh_unit else 0

        produto = _get_or_create_produto("PU CAST", item, descricao=f"PU CAST — {item}", categoria="Sobressalente", chave_busca=chave_busca_pu_cast.get(item))
        estrutura = _get_or_create_estrutura(produto, dn, ciclo_horas)
        slug = "".join(ch for ch in str(item).upper() if ch.isalnum())[:20]
        mp = _get_or_create_mp(f"PUCAST-MP-{slug}-DN{dn}", f"PU CAST {item} — matéria-prima consolidada DN {dn}", "un", custo_mp, "Química (consolidada)")
        _add_item(estrutura, 0, "MATERIA_PRIMA", 1, materia_prima=mp)
        pu_cast_estruturas[(item, dn)] = estrutura

    def _itens_corpo_fixacao(estrutura, ws, r, ordem):
        """Tubo/flanges/bumper/parafuso/arruela/porca de 1 linha da aba LBD — reaproveitado por LBD,
        LUN e CORPO MANDRIL (que usa exatamente os mesmos itens/quantidades da LBD, sem os discos)."""
        tubo_qtd = ws.cell(r, 17).value or 0
        flange_bumper_qtd = 2 if ws.cell(r, 24).value else 0
        bumper_peso_total = ws.cell(r, 27).value or 0
        flange_solda_qtd = 2 if ws.cell(r, 35).value else 0
        parafuso_qtd = ws.cell(r, 38).value or 0
        arruela_qtd = ws.cell(r, 43).value or 0
        porca_qtd = ws.cell(r, 48).value or 0
        if tubo_qtd and dn in tubo_por_dn:
            _add_item(estrutura, ordem, "MATERIA_PRIMA", tubo_qtd, materia_prima=tubo_por_dn[dn]); ordem += 1
        if flange_bumper_qtd and dn in flange_bumper_por_dn:
            _add_item(estrutura, ordem, "MATERIA_PRIMA", flange_bumper_qtd, materia_prima=flange_bumper_por_dn[dn]); ordem += 1
        if bumper_peso_total:
            _add_item(estrutura, ordem, "MATERIA_PRIMA", bumper_peso_total, materia_prima=MP_PRE_TDI, observacao="bumper (elastômero)"); ordem += 1
        if flange_solda_qtd and dn in flange_solda_por_dn:
            _add_item(estrutura, ordem, "MATERIA_PRIMA", flange_solda_qtd, materia_prima=flange_solda_por_dn[dn]); ordem += 1
        if parafuso_qtd and dn in parafuso_por_dn:
            _add_item(estrutura, ordem, "MATERIA_PRIMA", parafuso_qtd, materia_prima=parafuso_por_dn[dn]); ordem += 1
        if arruela_qtd and dn in arruela_por_dn:
            _add_item(estrutura, ordem, "MATERIA_PRIMA", arruela_qtd, materia_prima=arruela_por_dn[dn]); ordem += 1
        if porca_qtd and dn in porca_por_dn:
            _add_item(estrutura, ordem, "MATERIA_PRIMA", porca_qtd, materia_prima=porca_por_dn[dn]); ordem += 1
        return ordem

    # 5. Família LBD — produto único "LBD-DG2-DS4"
    ws_lbd = wb["LBD"]
    produto_lbd = _get_or_create_produto("LBD", "LBD-DG2-DS4", descricao="Mandril LBD-DG2-DS4", categoria="PIG", chave_busca="LBD")
    for r in range(6, 28):
        modelo = ws_lbd.cell(r, 1).value
        dn_val = ws_lbd.cell(r, 2).value
        if modelo is None or dn_val is None:
            continue
        dn = _dn_str(dn_val)
        custo_hh_unit = ws_lbd.cell(r, 59).value or 0  # BG
        ciclo_horas = round(custo_hh_unit / _CUSTOS_HH_RATE_SEED, 6) if custo_hh_unit else 0
        estrutura = _get_or_create_estrutura(produto_lbd, dn, ciclo_horas)
        ordem = 0
        for item_pu, qtd in (("DS", 4), ("DG", 2), ("DE", 6)):  # multiplicador fixo confirmado nas fórmulas
            sub = pu_cast_estruturas.get((item_pu, dn))
            if sub is not None:
                _add_item(estrutura, ordem, "SUBPRODUTO", qtd, subproduto=produto_cache[("PU CAST", item_pu)]); ordem += 1
        _itens_corpo_fixacao(estrutura, ws_lbd, r, ordem)

    # 6. Família LUN — produto único "LUN" (mesma estrutura da LBD, discos com quantidade variável por DN)
    ws_lun = wb["LUN"]
    produto_lun = _get_or_create_produto("LUN", "LUN", descricao="Mandril LUN", categoria="PIG", chave_busca="LUN")
    for r in range(6, 28):
        modelo = ws_lun.cell(r, 1).value
        dn_val = ws_lun.cell(r, 2).value
        if modelo is None or dn_val is None:
            continue
        dn = _dn_str(dn_val)
        custo_hh_unit = ws_lun.cell(r, 59).value or 0
        ciclo_horas = round(custo_hh_unit / _CUSTOS_HH_RATE_SEED, 6) if custo_hh_unit else 0
        estrutura = _get_or_create_estrutura(produto_lun, dn, ciclo_horas)
        ordem = 0
        for item_pu, qtd_col in (("COPO CONICO", 3), ("DS", 6), ("DG", 9), ("DE", 12)):
            qtd = ws_lun.cell(r, qtd_col).value or 0
            sub = pu_cast_estruturas.get((item_pu, dn))
            if qtd and sub is not None:
                _add_item(estrutura, ordem, "SUBPRODUTO", qtd, subproduto=produto_cache[("PU CAST", item_pu)]); ordem += 1
        _itens_corpo_fixacao(estrutura, ws_lun, r, ordem)

    # 7. Família CORPO MANDRIL — "corpo + fixação" (sem discos). Confirmado via fórmula real:
    #    CORPO MANDRIL!D = LBD!AZ (tubo+flange bumper+bumper+flange solda), CORPO MANDRIL!E = LBD!BA
    #    (parafuso+arruela+porca) — sempre lidas da MESMA linha da LBD (mesma DN, quantidades reais,
    #    não fixas). Relê a aba LBD (não a CORPO MANDRIL, que só tem DN+total, sem quantidade); a aba
    #    CORPO MANDRIL só é usada pro HH/ciclo (coluna H), que é independente.
    ws_corpo = wb["CORPO MANDRIL"]
    produto_corpo = _get_or_create_produto("CORPO MANDRIL", "CORPO + FIXAÇÃO", descricao="Corpo do PIG (tubo+flanges+fixação, sem discos)", categoria="Sobressalente", chave_busca="CORPO MANDRIL")
    for r in range(6, 28):
        dn_val = ws_lbd.cell(r, 2).value
        if dn_val is None:
            continue
        dn = _dn_str(dn_val)
        custo_hh_cico = 0
        for rc in range(4, ws_corpo.max_row + 1):
            if _dn_str(ws_corpo.cell(rc, 3).value) == dn:
                custo_hh_cico = ws_corpo.cell(rc, 8).value or 0
                break
        ciclo_horas = round(custo_hh_cico / _CUSTOS_HH_RATE_SEED, 6) if custo_hh_cico else 0
        estrutura = _get_or_create_estrutura(produto_corpo, dn, ciclo_horas)
        _itens_corpo_fixacao(estrutura, ws_lbd, r, 0)

    # 8. Família ELC_MG_PC — ELC (AÇO), ELP (PP), CINTA MAGNÉTICA, PLACA CALIBRADORA — custo MP já
    #    vem pronto por DN (fonte externa não enviada), vira 1 matéria-prima por (item, DN).
    ws_elc = wb["ELC_MG_PC"]
    for r in range(4, ws_elc.max_row + 1):
        item = ws_elc.cell(r, 2).value
        dn_val = ws_elc.cell(r, 3).value
        custo_mp = ws_elc.cell(r, 4).value
        custo_hh_unit = ws_elc.cell(r, 9).value or 0
        if item is None or dn_val is None:
            continue
        dn = _dn_str(dn_val)
        slug = "".join(ch for ch in item.upper() if ch.isalnum())[:20]
        mp = _get_or_create_mp(f"{slug}-MP-DN{dn}", f"{item} — matéria-prima DN {dn}", "un", custo_mp, "Acessório")
        produto = _get_or_create_produto("ELC_MG_PC", item, descricao=item, categoria="Acessório", chave_busca=item.split(" ")[0])
        ciclo_horas = round(custo_hh_unit / _CUSTOS_HH_RATE_SEED, 6) if custo_hh_unit else 0
        estrutura = _get_or_create_estrutura(produto, dn, ciclo_horas)
        _add_item(estrutura, 0, "MATERIA_PRIMA", 1, materia_prima=mp)

    # 9. Família PIGS EM BORRACHA — 3 variantes de material (EPDM, BUNA N, VITON), DN 6/8/10/12/14
    ciclo_borracha = round((hh_montagem_borracha or 0) / _CUSTOS_HH_RATE_SEED, 6)
    variantes_borracha = (
        ("PIG LUN-CP3 EPDM", copo_epdm_por_dn, "LUN-CP3 EPDM"),
        ("PIG LUN-CP3 BUNA N", copo_buna_por_dn, "LUN-CP3 BUNA"),
        ("PIG LUN-CP3 VITON", copo_viton_por_dn, "LUN-CP3 VITON"),
    )
    for codigo, copo_por_dn, chave in variantes_borracha:
        produto = _get_or_create_produto("PIGS EM BORRACHA", codigo, descricao=codigo, categoria="PIG", chave_busca=chave)
        for dn in ("6", "8", "10", "12", "14"):
            estrutura = _get_or_create_estrutura(produto, dn, ciclo_borracha)
            ordem = 0
            if dn in copo_por_dn:
                _add_item(estrutura, ordem, "MATERIA_PRIMA", 3, materia_prima=copo_por_dn[dn], observacao="3X"); ordem += 1  # "COPO BORRACHA ___ (3X)"
            if dn in eixo_por_dn:
                _add_item(estrutura, ordem, "MATERIA_PRIMA", 1, materia_prima=eixo_por_dn[dn]); ordem += 1
            if dn in cabecote_por_dn:
                _add_item(estrutura, ordem, "MATERIA_PRIMA", 1, materia_prima=cabecote_por_dn[dn]); ordem += 1
            if dn in nylon_por_dn:
                _add_item(estrutura, ordem, "MATERIA_PRIMA", 1, materia_prima=nylon_por_dn[dn]); ordem += 1
            if dn in porca_bor_por_dn:
                _add_item(estrutura, ordem, "MATERIA_PRIMA", 1, materia_prima=porca_bor_por_dn[dn]); ordem += 1
            if dn in flange_bor_por_dn:
                _add_item(estrutura, ordem, "MATERIA_PRIMA", 1, materia_prima=flange_bor_por_dn[dn]); ordem += 1

    db.session.add(ControleSistema(chave=_CHAVE_SEED_CUSTOS_PIG_MANDRIL_20_09_2026))
    db.session.commit()
    app.logger.info(
        "Gestão de Custos: importação inicial do grupo PIG MANDRIL concluída (%d matérias-primas, %d produtos, %d estruturas).",
        MateriaPrima.query.count(), Produto.query.count(), EstruturaProduto.query.count(),
    )


_CHAVE_MIGRACAO_PU_CAST_QUIMICA_23_09_2026 = "migracao_pu_cast_decompor_quimica_23_09_2026"


def _migrar_pu_cast_decompor_quimica(app):
    """Decompõe DS/DG/DE/COPO CONICO/COPO PISTAO/HFLEX/DISCFLEX SD/DISCFLEX
    SDI (família PU CAST) na química REAL (pré-polímero + MOCA, em kg) em
    vez do placeholder "matéria-prima consolidada" (1 un, custo já pronto)
    criado por `_seed_custos_pig_mandril`. Pedido do Bruno (23/09/2026), ao
    ver o popup de matéria-prima de um LBD-DG2-DS4: "esta puxando errado...
    eu quero as materias primas completas, inclusve as materias primas
    para produção de disco guia DG, etc" — o placeholder aparecia como
    "PU CAST DG — matéria-prima consolidada DN 18" (1 un) em vez da química
    real que entra na fabricação do disco.

    Por que o placeholder existia: a coluna CUSTO MP da aba PU CAST tem uma
    referência de preço que VARIA por linha (confirmado célula a célula na
    Fase 1 — ver `_seed_custos_pig_mandril`), então decompor por uma coluna
    fixa teria dado valor errado sem reler a fórmula de cada linha. Reli
    agora: das 132 linhas de (item, DN) desta família, 123 seguem o mesmo
    padrão de fórmula (`=T{linha}` = PRE×preço + MOCA×preço, colunas F/H) e
    só o preço do PRE varia entre 3 químicas já cadastradas (12-70 A, ATP
    85, ATS 85 — a MOCA é sempre fixa). As 9 linhas restantes (DS/DG/DE nos
    3 menores DN — 2", 3", 4") não têm peso (kg) nenhum lançado na planilha
    original, só um custo digitado direto — pra essas, sem peso pra
    decompor, o placeholder consolidado continua sendo a única opção fiel
    (sinalizado no log, não escondido).

    Roda uma vez só (`ControleSistema`, mesmo padrão do seed). Não mexe na
    MateriaPrima placeholder em si (`PUCAST-MP-*`) — ela continua existindo
    e ativa, porque `/custos/configurador-pig` lê o custo dela direto por
    código (disco espaçador extra do configurador de acessórios) — só troca
    o item da BOM que apontava pra ela por 2 itens novos apontando pra
    química real. Antes de trocar, confere que peso×preço da química nova
    bate com o custo antigo do placeholder (mesma fórmula, só decomposta) —
    se não bater por algum motivo não previsto, mantém o placeholder
    daquele (item, DN) em vez de arriscar um número errado."""
    if ControleSistema.query.filter_by(chave=_CHAVE_MIGRACAO_PU_CAST_QUIMICA_23_09_2026).first() is not None:
        return

    xlsx_path = os.path.join(BASE_DIR, "data", "custo_de_producao_20_09_2026.xlsx")
    if not os.path.exists(xlsx_path):
        app.logger.warning("Gestão de Custos: planilha de importação não encontrada em %s — decomposição química PU CAST não executada.", xlsx_path)
        return

    import re

    import openpyxl

    wb_formulas = openpyxl.load_workbook(xlsx_path, data_only=False)
    wb_valores = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws_f = wb_formulas["PU CAST"]
    ws_v = wb_valores["PU CAST"]
    ws_param = wb_valores["PARÂMETROS"]

    quimicas_por_linha = {}
    for r in range(6, 16):
        desc = ws_param.cell(r, 2).value
        if not desc:
            continue
        codigo = "QUIM-" + "".join(ch for ch in desc.upper() if ch.isalnum())[:30]
        mp = MateriaPrima.query.filter_by(codigo=codigo).first()
        if mp is not None:
            quimicas_por_linha[r] = mp
    mp_moca = quimicas_por_linha.get(8)  # MOCA CURATIVO TDI — referência fixa (col H sempre =G*PARÂMETROS!$E$8)

    itens_pu_cast = {"DS", "DG", "DE", "COPO CONICO", "COPO PISTAO", "HFLEX", "DISCFLEX SD", "DISCFLEX SDI"}
    n_decompostos = n_sem_peso = n_divergencia = n_sem_referencia = 0

    for r in range(10, ws_f.max_row + 1):
        item = ws_v.cell(r, 27).value
        if item not in itens_pu_cast:
            continue
        dn_num = ws_v.cell(r, 37).value
        if dn_num in (None, ""):
            continue
        dn = str(int(dn_num)) if isinstance(dn_num, float) and dn_num == int(dn_num) else str(dn_num).strip()

        custo_mp_formula = ws_f.cell(r, 29).value
        if not (isinstance(custo_mp_formula, str) and custo_mp_formula.startswith("=")):
            n_sem_peso += 1
            continue  # 9 linhas de DN muito pequeno sem peso cadastrado na planilha original

        peso_pre = ws_v.cell(r, 5).value or 0    # E: PESO PRE (kg)
        peso_moca = ws_v.cell(r, 7).value or 0   # G: PESO MOCA (kg)
        if not peso_pre and not peso_moca:
            n_sem_peso += 1
            continue

        f_formula = ws_f.cell(r, 6).value or ""  # F: CUSTO PRE — referência de preço varia por linha
        m = re.search(r"PAR[ÂA]METROS!\$E\$(\d+)", f_formula)
        mp_pre = quimicas_por_linha.get(int(m.group(1))) if m else None
        if mp_pre is None or mp_moca is None:
            app.logger.warning("Gestão de Custos: não achei a química de referência pra PU CAST %s DN %s (fórmula %r) — mantendo placeholder.", item, dn, f_formula)
            n_sem_referencia += 1
            continue

        produto = Produto.query.filter_by(familia="PU CAST", codigo=item).first()
        estrutura = EstruturaProduto.query.filter_by(produto_id=produto.id, dn=dn).first() if produto else None
        if estrutura is None:
            continue

        slug = "".join(ch for ch in str(item).upper() if ch.isalnum())[:20]
        codigo_placeholder = f"PUCAST-MP-{slug}-DN{dn}"
        item_placeholder = next(
            (i for i in estrutura.itens if i.tipo == "MATERIA_PRIMA" and i.materia_prima and i.materia_prima.codigo == codigo_placeholder),
            None,
        )
        if item_placeholder is None:
            continue  # já decomposto ou estrutura editada manualmente depois do seed — não mexe

        custo_antigo = (item_placeholder.materia_prima.custo_atual or 0) * item_placeholder.quantidade
        custo_novo = peso_pre * (mp_pre.custo_atual or 0) + peso_moca * (mp_moca.custo_atual or 0)
        if abs(custo_novo - custo_antigo) > max(0.05, custo_antigo * 0.01):
            app.logger.warning(
                "Gestão de Custos: decomposição química de PU CAST %s DN %s não bateu com o custo original (novo=%.4f antigo=%.4f) — mantendo placeholder por segurança.",
                item, dn, custo_novo, custo_antigo,
            )
            n_divergencia += 1
            continue

        db.session.delete(item_placeholder)
        db.session.flush()
        ordem = 0
        if peso_pre:
            db.session.add(EstruturaProdutoItem(
                estrutura_id=estrutura.id, tipo="MATERIA_PRIMA", quantidade=peso_pre,
                materia_prima_id=mp_pre.id, ordem=ordem, observacao="química real (decomposta 23/09/2026)",
            ))
            ordem += 1
        if peso_moca:
            db.session.add(EstruturaProdutoItem(
                estrutura_id=estrutura.id, tipo="MATERIA_PRIMA", quantidade=peso_moca,
                materia_prima_id=mp_moca.id, ordem=ordem, observacao="química real (decomposta 23/09/2026)",
            ))
        n_decompostos += 1

    db.session.add(ControleSistema(chave=_CHAVE_MIGRACAO_PU_CAST_QUIMICA_23_09_2026))
    db.session.commit()
    app.logger.info(
        "Gestão de Custos: decomposição química PU CAST concluída — %d decompostos, %d sem peso (placeholder mantido), "
        "%d divergentes (placeholder mantido), %d sem referência de química.",
        n_decompostos, n_sem_peso, n_divergencia, n_sem_referencia,
    )


_CHAVE_SEED_CUSTOS_ESPUMA_20_09_2026 = "seed_custos_espuma_20_09_2026"


def _seed_custos_espuma(app):
    """Importa (uma única vez) a família espuma (fase 2 do módulo Gestão de
    Custos) da mesma planilha da fase 1 (`data/custo_de_producao_20_09_2026.xlsx`):
    abas H, HS, HL, HLR, HLR X, HLR R, HLR V, HLB, HDISC, HLCC, HLCC PC.
    Mesmo padrão idempotente de `_seed_custos_pig_mandril` (guardado por
    ControleSistema), e reaproveita a mesma matéria-prima central da aba
    PARÂMETROS (linhas 18-31: bloco de espuma D26/D45/D60/D80, elastômero
    TECPUR, sistema amino A alta/média + B iso, pigmento, corda, escova
    fina/grossa, velcro, cola) e a MOCA CURATIVO TDI já cadastrada na fase 1
    (linha 8 — mesmo código de matéria-prima, pra não duplicar o mesmo
    insumo em 2 linhas do catálogo).

    A aba H tem uma particularidade: 3 variantes cumulativas por DN/densidade
    ("H", "H COM SELO", "H COM SELO E CORDA" — confirmado célula a célula
    que cada variante = a anterior + uma camada extra), modeladas como 3
    Produtos encadeados por SUBPRODUTO (mesmo mecanismo recursivo já usado
    na fase 1 pra LUN→PU CAST e CORPO MANDRIL→LBD — nada novo no motor de
    cálculo). As outras 9 abas ("sistema A+B" — poliuretano vazado em 2
    componentes) são 1 produto por família com todas as camadas na mesma
    estrutura (não há variante cumulativa nelas, só 1 "R$" final por DN).

    Verificado 1:1 contra a aba BUSCA DE CUSTO antes deste código ir pro ar
    — ver relatório de verificação anexo à entrega."""
    if ControleSistema.query.filter_by(chave=_CHAVE_SEED_CUSTOS_ESPUMA_20_09_2026).first() is not None:
        return

    xlsx_path = os.path.join(BASE_DIR, "data", "custo_de_producao_20_09_2026.xlsx")
    if not os.path.exists(xlsx_path):
        app.logger.warning("Gestão de Custos: planilha de importação não encontrada em %s — seed espuma não executado.", xlsx_path)
        return

    import openpyxl

    def _dn_str(v):
        if v is None:
            return None
        if isinstance(v, float) and v == int(v):
            return str(int(v))
        return str(v).strip()

    def _dn_num(v):
        """Só a parte numérica do DN (pra decidir escova fina/grossa por faixa) — trata
        formatos tipo "9,5''" (vírgula decimal) também."""
        if v is None:
            return None
        s = "".join(ch for ch in str(v).strip() if ch.isdigit() or ch in ",.").replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return None

    mp_cache, produto_cache = {}, {}

    def _get_or_create_mp(codigo, descricao, unidade, custo, categoria):
        mp = mp_cache.get(codigo)
        if mp is not None:
            return mp
        mp = MateriaPrima.query.filter_by(codigo=codigo).first()
        if mp is None:
            mp = MateriaPrima(codigo=codigo, descricao=descricao, unidade=unidade, custo_atual=custo or 0, categoria=categoria, ativo=True)
            db.session.add(mp)
            db.session.flush()
        mp_cache[codigo] = mp
        return mp

    def _get_or_create_produto(familia, codigo, descricao=None, categoria=None, chave_busca=None):
        key = (familia, codigo)
        p = produto_cache.get(key)
        if p is not None:
            return p
        p = Produto.query.filter_by(familia=familia, codigo=codigo).first()
        if p is None:
            p = Produto(familia=familia, codigo=codigo, descricao=descricao, categoria=categoria, chave_busca=chave_busca, ativo=True)
            db.session.add(p)
            db.session.flush()
        produto_cache[key] = p
        return p

    def _get_or_create_estrutura(produto, dn, ciclo_horas):
        e = EstruturaProduto.query.filter_by(produto_id=produto.id, dn=dn).first()
        if e is None:
            e = EstruturaProduto(produto_id=produto.id, dn=dn, ciclo_horas=ciclo_horas or 0, ativo=True)
            db.session.add(e)
            db.session.flush()
        return e

    def _add_item(estrutura, ordem, tipo, quantidade, materia_prima=None, subproduto=None, observacao=None):
        if not quantidade:
            return
        db.session.add(EstruturaProdutoItem(
            estrutura_id=estrutura.id, tipo=tipo, quantidade=quantidade,
            materia_prima_id=materia_prima.id if materia_prima else None,
            subproduto_id=subproduto.id if subproduto else None,
            observacao=observacao, ordem=ordem,
        ))

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws_param = wb["PARÂMETROS"]

    # 0. Catálogo central de matérias-primas da família espuma — PARÂMETROS linhas 18-31.
    #    MOCA CURATIVO TDI (linha 8) já foi cadastrada na fase 1 com este código —
    #    reaproveita a MESMA linha do catálogo (não duplica o mesmo insumo).
    def _mp_parametro(linha, codigo, categoria, unidade):
        desc = ws_param.cell(linha, 2).value
        preco = ws_param.cell(linha, 5).value
        return _get_or_create_mp(codigo, desc, unidade, preco, categoria)

    MP_MOCA = _get_or_create_mp("QUIM-MOCACURATIVOTDI", ws_param.cell(8, 2).value, "kg", ws_param.cell(8, 5).value, "Química")
    MP_BLOCO_D26 = _mp_parametro(18, "ESPUMA-BLOCO-D26", "Bloco de espuma", "m³")
    MP_BLOCO_D45 = _mp_parametro(19, "ESPUMA-BLOCO-D45", "Bloco de espuma", "m³")
    MP_BLOCO_D60 = _mp_parametro(20, "ESPUMA-BLOCO-D60", "Bloco de espuma", "m³")
    MP_BLOCO_D80 = _mp_parametro(21, "ESPUMA-BLOCO-D80", "Bloco de espuma", "m³")
    MP_ELASTOMERO = _mp_parametro(22, "ESPUMA-ELASTOMERO-TECPUR", "Química", "kg")
    MP_AMINO_A_ALTA = _mp_parametro(23, "ESPUMA-AMINO-A-ALTA", "Química", "kg")
    MP_AMINO_A_MEDIA = _mp_parametro(24, "ESPUMA-AMINO-A-MEDIA", "Química", "kg")
    MP_AMINO_B_ISO = _mp_parametro(25, "ESPUMA-AMINO-B-ISO", "Química", "kg")
    MP_PIGMENTO = _mp_parametro(26, "ESPUMA-PIGMENTO", "Química", "kg")
    MP_CORDA = _mp_parametro(27, "ESPUMA-CORDA-OLHAL", "Acessório", "m")
    MP_ESCOVA_FINA = _mp_parametro(28, "ESPUMA-ESCOVA-FINA-HLR", "Acessório", "m")
    MP_ESCOVA_GROSSA = _mp_parametro(29, "ESPUMA-ESCOVA-GROSSA-HLR", "Acessório", "m")
    MP_VELCRO = _mp_parametro(30, "ESPUMA-VELCRO-HLR-V", "Acessório", "m")
    MP_COLA = _mp_parametro(31, "ESPUMA-COLA-SAPATEIRO", "Acessório", "kg")
    # Itens exclusivos de HLCC / HLCC PC — não estão centralizados em PARÂMETROS na planilha
    # original (ficam soltos no cabeçalho da própria aba); centralizamos aqui do mesmo jeito
    # (1 matéria-prima cada, valor tirado do cabeçalho da aba na data da importação).
    #
    # ATENÇÃO — achado na verificação, repassado ao Bruno: o cabeçalho da aba mostra
    # "CUSTO METRO CABO DE AÇO" = R$9,70/m pro olhal, mas o valor REALMENTE aplicado em
    # toda linha (custo ÷ comprimento, conferido em várias linhas de HLCC e HLCC PC) é
    # sempre R$3,00/m — o mesmo preço da CORDA OLHAL (PARÂMETROS!$E$27). O rótulo do
    # cabeçalho parece estar desatualizado/não é o que a fórmula usa de fato; replicamos
    # o valor realmente aplicado (reaproveitando a matéria-prima da corda) pra bater com
    # BUSCA DE CUSTO.
    MP_PRENSA_CABO = _get_or_create_mp("ESPUMA-PRENSA-CABO", 'Prensa cabo 3/16" — HLCC/HLCC PC', "un", 0.95, "Acessório")
    MP_KIT_ARRUELA_PORCA = _get_or_create_mp("ESPUMA-KIT-ARRUELA-PORCA-HLCCPC", "Kit fixação (arruela + porca) — HLCC PC", "un", 35.15, "Acessório")
    MP_BARRA_ROSCADA = _get_or_create_mp("ESPUMA-BARRA-ROSCADA-HLCCPC", 'Barra roscada 3/4" — HLCC PC', "m", 57.0, "Acessório")
    MP_BUMPER_PU = _get_or_create_mp("ESPUMA-BUMPER-PU-HLCCPC", "Bumper PU (MP COIM) — HLCC PC", "kg", 55.0, "Acessório")

    def _amino_a(densidade):
        return MP_AMINO_A_ALTA if "ALTA" in densidade else MP_AMINO_A_MEDIA

    def _bloco_por_densidade_label(label):
        # label ex. "BAIXA D26" / "BAIXA D45" / "BAIXA D60" / "BAIXA D80"
        return {"D26": MP_BLOCO_D26, "D45": MP_BLOCO_D45, "D60": MP_BLOCO_D60, "D80": MP_BLOCO_D80}[label.split()[-1]]

    # ------------------------------------------------------------------
    # 1. Aba H — bloco de espuma, 4 densidades (D26/D45/D60/D80) x 3 variantes
    #    cumulativas (H / H COM SELO / H COM SELO E CORDA), encadeadas por
    #    SUBPRODUTO (H COM SELO = SUBPRODUTO(H) + camada selo; H COM SELO E
    #    CORDA = SUBPRODUTO(H COM SELO) + corda) — confirmado fórmula a
    #    fórmula (col19=D+I+M, col22=col19+G+K, col25=col22+O).
    # ------------------------------------------------------------------
    ws_h = wb["H"]
    produto_h = _get_or_create_produto("H", "H", descricao="PIG H (bloco de espuma)", categoria="PIG", chave_busca=None)
    produto_h_selo = _get_or_create_produto("H", "H-COM-SELO", descricao="PIG H com selo", categoria="PIG", chave_busca=None)
    produto_h_selo_corda = _get_or_create_produto("H", "H-COM-SELO-CORDA", descricao="PIG H com selo e corda", categoria="PIG", chave_busca=None)

    blocos_h = [(6, 29), (30, 54), (55, 79), (80, 104)]
    for r0, r1 in blocos_h:
        for r in range(r0, r1 + 1):
            dn_raw = ws_h.cell(r, 1).value
            dens_label = ws_h.cell(r, 2).value
            if dn_raw is None or dens_label is None:
                continue
            dn = f"{_dn_str(dn_raw)} {dens_label}"  # ex. "6'' BAIXA D26"
            consumo_m3 = ws_h.cell(r, 3).value or 0
            peso_elast_selo = ws_h.cell(r, 6).value or 0
            peso_elast_etiqueta = ws_h.cell(r, 8).value or 0
            peso_moca_selo = ws_h.cell(r, 10).value or 0
            peso_moca_etiqueta = ws_h.cell(r, 12).value or 0
            comprimento_corda = ws_h.cell(r, 14).value or 0
            # "CUSTO HH ESC/LOTE" (não "CUSTO HH CICLO" cru) — a planilha divide o
            # custo do ciclo da máquina pela quantidade de peças do lote escalonado
            # (col "QNT PCS"); CUSTO TOTAL = CUSTO MP + CUSTO HH ESC/LOTE, confirmado
            # fórmula a fórmula (col71=col67+col70, não col67+col68).
            #
            # IMPORTANTE: o bloco de tempos (HH) de "H COM SELO" e "H COM SELO E CORDA"
            # é um recálculo COMPLETO do ciclo (corte+roletagem+chanfro+... de novo, não
            # só o incremento da camada extra) — confirmado comparando os blocos de tempo
            # das 3 variantes célula a célula. Por isso as 3 estruturas abaixo são
            # independentes (cada uma com seu próprio ciclo_horas) — só a matéria-prima é
            # cumulativa (repetida explicitamente em cada variante, sem SUBPRODUTO), senão
            # o motor de cálculo (que soma custo_hh do subproduto inteiro) contaria a hora
            # de máquina da variante anterior de novo.
            custo_hh_base = ws_h.cell(r, 70).value or 0
            custo_hh_selo = ws_h.cell(r, 78).value or 0
            custo_hh_corda = ws_h.cell(r, 86).value or 0

            mp_bloco = _bloco_por_densidade_label(dens_label)

            est_base = _get_or_create_estrutura(produto_h, dn, round(custo_hh_base / _CUSTOS_HH_RATE_SEED, 6))
            _add_item(est_base, 0, "MATERIA_PRIMA", consumo_m3, materia_prima=mp_bloco, observacao="bloco de espuma")
            _add_item(est_base, 1, "MATERIA_PRIMA", peso_elast_etiqueta, materia_prima=MP_ELASTOMERO, observacao="etiqueta")
            _add_item(est_base, 2, "MATERIA_PRIMA", peso_moca_etiqueta, materia_prima=MP_MOCA, observacao="etiqueta")

            est_selo = _get_or_create_estrutura(produto_h_selo, dn, round(custo_hh_selo / _CUSTOS_HH_RATE_SEED, 6))
            _add_item(est_selo, 0, "MATERIA_PRIMA", consumo_m3, materia_prima=mp_bloco, observacao="bloco de espuma")
            _add_item(est_selo, 1, "MATERIA_PRIMA", peso_elast_etiqueta, materia_prima=MP_ELASTOMERO, observacao="etiqueta")
            _add_item(est_selo, 2, "MATERIA_PRIMA", peso_moca_etiqueta, materia_prima=MP_MOCA, observacao="etiqueta")
            _add_item(est_selo, 3, "MATERIA_PRIMA", peso_elast_selo, materia_prima=MP_ELASTOMERO, observacao="selo")
            _add_item(est_selo, 4, "MATERIA_PRIMA", peso_moca_selo, materia_prima=MP_MOCA, observacao="selo")

            est_corda = _get_or_create_estrutura(produto_h_selo_corda, dn, round(custo_hh_corda / _CUSTOS_HH_RATE_SEED, 6))
            _add_item(est_corda, 0, "MATERIA_PRIMA", consumo_m3, materia_prima=mp_bloco, observacao="bloco de espuma")
            _add_item(est_corda, 1, "MATERIA_PRIMA", peso_elast_etiqueta, materia_prima=MP_ELASTOMERO, observacao="etiqueta")
            _add_item(est_corda, 2, "MATERIA_PRIMA", peso_moca_etiqueta, materia_prima=MP_MOCA, observacao="etiqueta")
            _add_item(est_corda, 3, "MATERIA_PRIMA", peso_elast_selo, materia_prima=MP_ELASTOMERO, observacao="selo")
            _add_item(est_corda, 4, "MATERIA_PRIMA", peso_moca_selo, materia_prima=MP_MOCA, observacao="selo")
            _add_item(est_corda, 5, "MATERIA_PRIMA", comprimento_corda, materia_prima=MP_CORDA, observacao="corda")

    # ------------------------------------------------------------------
    # 2. As 9 abas "sistema A+B" (poliuretano vazado, 2 componentes) — núcleo
    #    comum (poliol/isocianato + elastômero + moca, 2 ou 3 camadas conforme
    #    a aba) + bloco de acessórios específico por família. 1 produto só por
    #    família, sem variante cumulativa (diferente da aba H).
    # ------------------------------------------------------------------
    def _core_ab(estrutura, ws, r, densidade, ordem, elast_specs, moca_specs):
        """elast_specs/moca_specs: lista de (col_peso, rótulo). Sempre lê o peso
        (já em kg) direto da célula — a fórmula/derivação de peso varia entre
        abas (algumas usam peso fixo, outras 1% do peso de outra camada), mas
        o valor final (data_only=True) é sempre o número certo a multiplicar
        pelo preço unitário, então não precisamos replicar a fórmula."""
        peso_a = ws.cell(r, 3).value or 0
        peso_b = ws.cell(r, 5).value or 0
        mp_a = _amino_a(densidade)
        _add_item(estrutura, ordem, "MATERIA_PRIMA", peso_a, materia_prima=mp_a, observacao="poliol (A)"); ordem += 1
        _add_item(estrutura, ordem, "MATERIA_PRIMA", peso_b, materia_prima=MP_AMINO_B_ISO, observacao="isocianato (B)"); ordem += 1
        for col, rotulo in elast_specs:
            peso = ws.cell(r, col).value or 0
            _add_item(estrutura, ordem, "MATERIA_PRIMA", peso, materia_prima=MP_ELASTOMERO, observacao=f"elastômero ({rotulo})"); ordem += 1
        for col, rotulo in moca_specs:
            peso = ws.cell(r, col).value or 0
            _add_item(estrutura, ordem, "MATERIA_PRIMA", peso, materia_prima=MP_MOCA, observacao=f"moca ({rotulo})"); ordem += 1
        return ordem

    def _pigmento_ab(estrutura, ws, r, ordem, col_revest, col_selo):
        # ATENÇÃO — achado na verificação, repassado ao Bruno: nas 9 abas "sistema
        # A+B" a fórmula real do custo do pigmento usa o preço da MOCA
        # (PARÂMETROS!$E$8 = R$30,89/kg), não o preço do próprio pigmento
        # (PARÂMETROS!$E$26 = R$89,30/kg, que é o valor mostrado no cabeçalho da
        # aba como referência) — confirmado fórmula a fórmula nas 6 abas de 3
        # camadas (HL, HLR/HLR X/HLR R/HLR V, HLB, HDISC, HLCC, HLCC PC). Parece
        # um copy-paste que ficou preso na fórmula da MOCA; replicamos o cálculo
        # real da planilha (não o rótulo) pra bater com BUSCA DE CUSTO — MP_PIGMENTO
        # fica cadastrada no catálogo mesmo assim, caso o Bruno prefira corrigir a
        # fórmula na planilha numa próxima rodada.
        peso_revest = ws.cell(r, col_revest).value or 0
        peso_selo = ws.cell(r, col_selo).value or 0
        _add_item(estrutura, ordem, "MATERIA_PRIMA", peso_revest, materia_prima=MP_MOCA, observacao="pigmento (revestimento) — preço aplicado: moca, ver observação no código"); ordem += 1
        _add_item(estrutura, ordem, "MATERIA_PRIMA", peso_selo, materia_prima=MP_MOCA, observacao="pigmento (selo) — preço aplicado: moca, ver observação no código"); ordem += 1
        return ordem

    # A coluna 12 (rotulada "PESO ELASTÔMERO ETIQUETA" no cabeçalho) é na verdade
    # precificada com o preço da MOCA na fórmula real da planilha (=L*PARÂMETROS!$E$8,
    # não $E$22) — confirmado célula a célula nas 6 abas de 3 camadas (HL, HLR/HLR X/
    # HLR R/HLR V, HLB, HDISC, HLCC, HLCC PC): rótulo da planilha está errado/copiado,
    # a fórmula manda. Por isso ela entra no grupo "moca", não "elastômero" (só afeta
    # QUAL matéria-prima é debitada — o motor de cálculo soma os itens igual).
    ELAST_2 = [(8, "revestimento"), (10, "selagem")]
    MOCA_4 = [(12, "etiqueta"), (14, "revestimento"), (16, "selo"), (18, "etiqueta")]

    def _processar_sheet_ab(nome_aba, familia, chave_busca, custo_hh_ciclo_col,
                             blocos_densidade, elast_specs, moca_specs,
                             accessorios_fn):
        # "CUSTO HH ESC/LOTE" fica sempre 2 colunas depois de "CUSTO HH CICLO"
        # (CUSTO MP, CUSTO HH CICLO, QNT PCS, CUSTO HH ESC/LOTE, CUSTO TOTAL — padrão
        # confirmado nas 9 abas) — é esse valor escalonado que compõe o CUSTO TOTAL
        # da planilha (CUSTO TOTAL = CUSTO MP + CUSTO HH ESC/LOTE, não + CUSTO HH CICLO cru).
        custo_hh_esc_lote_col = custo_hh_ciclo_col + 2
        ws = wb[nome_aba]
        produto = _get_or_create_produto(familia, familia, descricao=f"PIG {familia}", categoria="PIG", chave_busca=chave_busca)
        for r0, r1 in blocos_densidade:
            for r in range(r0, r1 + 1):
                dn_raw = ws.cell(r, 1).value
                densidade = ws.cell(r, 2).value
                if dn_raw is None or densidade is None:
                    continue
                dn = f"{_dn_str(dn_raw)} {'ALTA' if 'ALTA' in densidade else 'MÉDIA'}"
                custo_hh_esc_lote = ws.cell(r, custo_hh_esc_lote_col).value or 0
                estrutura = _get_or_create_estrutura(produto, dn, round(custo_hh_esc_lote / _CUSTOS_HH_RATE_SEED, 6))
                ordem = _core_ab(estrutura, ws, r, densidade, 0, elast_specs, moca_specs)
                if accessorios_fn is not None:
                    accessorios_fn(estrutura, ws, r, ordem, dn_raw)

    BLOCOS_2_TIER = [(7, 30), (31, 54)]
    BLOCOS_1_TIER = [(7, 13)]

    # HS — sem camada de revestimento (só selo+etiqueta), corda embutida, sem pigmento/escova/cola.
    def _acessorios_hs(estrutura, ws, r, ordem, dn_raw):
        comprimento_corda = ws.cell(r, 16).value or 0
        _add_item(estrutura, ordem, "MATERIA_PRIMA", comprimento_corda, materia_prima=MP_CORDA, observacao="corda"); ordem += 1
    _processar_sheet_ab("HS", "HS", "HS", 35, BLOCOS_2_TIER,
                        [(8, "selo"), (10, "etiqueta")], [(12, "selo"), (14, "etiqueta")],
                        _acessorios_hs)

    # HL e HDISC — 3 camadas + corda + pigmento (revest+selo). Sem escova/cola.
    def _acessorios_corda_pigmento(col_corda, col_pig_r, col_pig_s):
        def _fn(estrutura, ws, r, ordem, dn_raw):
            comprimento_corda = ws.cell(r, col_corda).value or 0
            _add_item(estrutura, ordem, "MATERIA_PRIMA", comprimento_corda, materia_prima=MP_CORDA, observacao="corda")
            ordem += 1
            _pigmento_ab(estrutura, ws, r, ordem, col_pig_r, col_pig_s)
        return _fn

    _processar_sheet_ab("HL", "HL", "HL", 45, BLOCOS_2_TIER, ELAST_2, MOCA_4,
                        _acessorios_corda_pigmento(20, 22, 24))
    _processar_sheet_ab("HDISC", "HDISC", "HDISC", 46, BLOCOS_2_TIER, ELAST_2, MOCA_4,
                        _acessorios_corda_pigmento(20, 22, 24))

    # HLR / HLR X / HLR R — corda + escova de aço (fina até DN4, grossa DN6+) + cola + pigmento.
    def _acessorios_escova(estrutura, ws, r, ordem, dn_raw):
        comprimento_corda = ws.cell(r, 20).value or 0
        _add_item(estrutura, ordem, "MATERIA_PRIMA", comprimento_corda, materia_prima=MP_CORDA, observacao="corda"); ordem += 1
        comprimento_escova = ws.cell(r, 22).value or 0
        dn_num = _dn_num(dn_raw)
        mp_escova = MP_ESCOVA_FINA if (dn_num is not None and dn_num <= 4) else MP_ESCOVA_GROSSA
        _add_item(estrutura, ordem, "MATERIA_PRIMA", comprimento_escova, materia_prima=mp_escova, observacao="escova"); ordem += 1
        peso_cola = ws.cell(r, 24).value or 0
        _add_item(estrutura, ordem, "MATERIA_PRIMA", peso_cola, materia_prima=MP_COLA, observacao="cola (fixação escova)"); ordem += 1
        ordem = _pigmento_ab(estrutura, ws, r, ordem, 26, 28)

    for nome_aba, chave in (("HLR", "HLR"), ("HLR X", "HLR X"), ("HLR R", "HLR R")):
        _processar_sheet_ab(nome_aba, nome_aba, chave, 50, BLOCOS_2_TIER, ELAST_2, MOCA_4, _acessorios_escova)

    # HLR V / HLB — corda + velcro (preço único, sem faixa por DN) + cola + pigmento.
    def _acessorios_velcro(estrutura, ws, r, ordem, dn_raw):
        comprimento_corda = ws.cell(r, 20).value or 0
        _add_item(estrutura, ordem, "MATERIA_PRIMA", comprimento_corda, materia_prima=MP_CORDA, observacao="corda"); ordem += 1
        comprimento_velcro = ws.cell(r, 22).value or 0
        _add_item(estrutura, ordem, "MATERIA_PRIMA", comprimento_velcro, materia_prima=MP_VELCRO, observacao="velcro"); ordem += 1
        peso_cola = ws.cell(r, 24).value or 0
        _add_item(estrutura, ordem, "MATERIA_PRIMA", peso_cola, materia_prima=MP_COLA, observacao="cola (fixação velcro)"); ordem += 1
        ordem = _pigmento_ab(estrutura, ws, r, ordem, 26, 28)

    for nome_aba, chave in (("HLR V", "HLR V"), ("HLB", "HLB")):
        _processar_sheet_ab(nome_aba, nome_aba, chave, 50, BLOCOS_2_TIER, ELAST_2, MOCA_4, _acessorios_velcro)

    # HLCC — olhal (cabo de aço) + prensa cabo + pigmento. Sem corda/escova/cola.
    def _acessorios_hlcc(estrutura, ws, r, ordem, dn_raw):
        comprimento_olhal = ws.cell(r, 20).value or 0
        _add_item(estrutura, ordem, "MATERIA_PRIMA", comprimento_olhal, materia_prima=MP_CORDA, observacao="olhal (cabo de aço) — preço aplicado: mesmo da corda, ver observação no código"); ordem += 1
        qnt_prensa = ws.cell(r, 22).value or 0
        _add_item(estrutura, ordem, "MATERIA_PRIMA", qnt_prensa, materia_prima=MP_PRENSA_CABO, observacao="prensa cabo"); ordem += 1
        ordem = _pigmento_ab(estrutura, ws, r, ordem, 24, 26)

    _processar_sheet_ab("HLCC", "HLCC", "HLCC", 46, BLOCOS_1_TIER, ELAST_2, MOCA_4, _acessorios_hlcc)

    # HLCC PC — kit arruela+porca + barra roscada + placa calibradora + bumper PU +
    # olhal (cabo de aço) + prensa cabo + pigmento. A mais "acessorizada" das 11.
    def _acessorios_hlcc_pc(estrutura, ws, r, ordem, dn_raw):
        _add_item(estrutura, ordem, "MATERIA_PRIMA", 1, materia_prima=MP_KIT_ARRUELA_PORCA, observacao="kit fixação (1x)"); ordem += 1
        comprimento_barra = ws.cell(r, 22).value or 0
        _add_item(estrutura, ordem, "MATERIA_PRIMA", comprimento_barra, materia_prima=MP_BARRA_ROSCADA, observacao="barra roscada"); ordem += 1
        custo_placa = ws.cell(r, 24).value or 0
        if custo_placa:
            mp_placa = _get_or_create_mp(f"ESPUMA-PLACA-CALIBRADORA-HLCCPC-DN{_dn_str(dn_raw)}", f"Placa calibradora (usinagem+alumínio) — HLCC PC DN {_dn_str(dn_raw)}", "un", custo_placa, "Acessório")
            _add_item(estrutura, ordem, "MATERIA_PRIMA", 1, materia_prima=mp_placa, observacao="placa calibradora"); ordem += 1
        peso_bumper = ws.cell(r, 25).value or 0
        _add_item(estrutura, ordem, "MATERIA_PRIMA", peso_bumper, materia_prima=MP_BUMPER_PU, observacao="bumper PU"); ordem += 1
        comprimento_olhal = ws.cell(r, 27).value or 0
        _add_item(estrutura, ordem, "MATERIA_PRIMA", comprimento_olhal, materia_prima=MP_CORDA, observacao="olhal (cabo de aço) — preço aplicado: mesmo da corda, ver observação no código"); ordem += 1
        qnt_prensa = ws.cell(r, 29).value or 0
        _add_item(estrutura, ordem, "MATERIA_PRIMA", qnt_prensa, materia_prima=MP_PRENSA_CABO, observacao="prensa cabo"); ordem += 1
        ordem = _pigmento_ab(estrutura, ws, r, ordem, 31, 33)

    _processar_sheet_ab("HLCC PC", "HLCC PC", "HLCC PC", 57, BLOCOS_1_TIER, ELAST_2, MOCA_4, _acessorios_hlcc_pc)

    db.session.add(ControleSistema(chave=_CHAVE_SEED_CUSTOS_ESPUMA_20_09_2026))
    db.session.commit()
    app.logger.info(
        "Gestão de Custos: importação da família espuma (fase 2) concluída (%d matérias-primas, %d produtos, %d estruturas).",
        MateriaPrima.query.count(), Produto.query.count(), EstruturaProduto.query.count(),
    )

_CHAVE_SEED_CUSTOS_SUPERFLEX_SILICONE_20_09_2026 = "seed_custos_superflex_silicone_20_09_2026"


def _seed_custos_superflex_silicone(app):
    """Importa (uma única vez) SUPERFLEX e SILICONE (fase 3 do módulo Gestão de
    Custos, junto com a tela de Simulação/Histórico que fica numa função à parte)
    da mesma planilha das fases 1 e 2 (`data/custo_de_producao_20_09_2026.xlsx`).
    Mesmo padrão idempotente de `_seed_custos_pig_mandril`/`_seed_custos_espuma`
    (guardado por ControleSistema).

    SUPERFLEX (2 produtos, CS3 e CS4) tem um layout bem diferente das famílias já
    importadas: cada variante é montada a partir de 3 "sub-tabelas" por DN (COPO
    PISTÃO EM PU, DISCO ESPAÇADORES EM PU, EIXO EM PU — cada uma com seu próprio
    peso/tempo/matéria-prima), multiplicadas por uma quantidade fixa (3x pra CS3,
    4x pra CS4, confirmado nas fórmulas B6=P6*3 / B15=P6*4), mais um item de custo
    fixo ("ANEL DE TRAVAMENTO + CABO DE AÇO" = R$45 literal, igual nas 2 variantes
    e em todos os DNs) e um bloco fixo de HH de montagem (3h, também igual nas 2
    variantes — `=$N$2*$N$1`). O CS4 tem ainda um acréscimo literal por DN no
    custo do eixo (confirmado na fórmula `=B8+27`/`+38`/`+43`/`+73`/`+114` — não é
    proporcional a nada, é um valor fixo por DN, tratado aqui como mais uma
    matéria-prima "consolidada" por DN, mesmo padrão já usado na fase 1 pra
    ELC_MG_PC). Modelado com o mesmo mecanismo de SUBPRODUTO já usado desde a fase
    1 (LUN→PU CAST): cada sub-tabela vira um Produto próprio (categoria
    "Componente", não aparece como PIG vendável) e CS3/CS4 referenciam essas
    3 sub-tabelas com a quantidade certa — o motor de cálculo resolve tudo
    recursivamente sem precisar de nenhuma mudança.

    SUPERFLEX não está indexado na aba BUSCA DE CUSTO (nenhuma linha "SUPERFLEX"
    lá) — verificado célula a célula contra o próprio TOTAL de cada coluna DN na
    aba SUPERFLEX (linhas 11 e 20) em vez disso.

    SILICONE (aba com só a família DISCFLEX) é bem mais simples e autocontida:
    cada linha é 1 DN com peso de manta de silicone + qtde de ímãs (0 pro
    DISCFLEX "puro", 2 ou 4 pras variantes "COM IMÃ"), sem nenhuma dependência de
    outra aba. 3 produtos: DISCFLEX (8 linhas — nota: 2 delas têm o mesmo rótulo
    de DN "3"" só que com pesos diferentes, 0,2834kg e 0,2995kg — divergência sem
    explicação na própria planilha, mesma que a aba BUSCA DE CUSTO também carrega
    sem diferenciar; aqui é preciso separar em 2 DNs distintos pra caber no
    modelo, marcados "3\" (A)"/"3\" (B)" com o peso de cada um na observação, e
    fica sinalizado pro Bruno revisar a origem dessa duplicidade), DISCFLEX COM
    IMÃ (2 linhas, ímã 8x5mm — confirmado na fórmula `=H14*$B$3`) e HSC AZUL COM
    IMÃ (1 linha, ímã 22x10mm — fórmula `=H16*$B$2`, catálogo de ímã diferente do
    das outras 2 variantes). Verificado 1:1 contra BUSCA DE CUSTO (as 3 famílias
    estão indexadas lá)."""
    if ControleSistema.query.filter_by(chave=_CHAVE_SEED_CUSTOS_SUPERFLEX_SILICONE_20_09_2026).first() is not None:
        return

    xlsx_path = os.path.join(BASE_DIR, "data", "custo_de_producao_20_09_2026.xlsx")
    if not os.path.exists(xlsx_path):
        app.logger.warning("Gestão de Custos: planilha de importação não encontrada em %s — seed SUPERFLEX/SILICONE não executado.", xlsx_path)
        return

    import openpyxl

    mp_cache, produto_cache = {}, {}

    def _get_or_create_mp(codigo, descricao, unidade, custo, categoria):
        mp = mp_cache.get(codigo)
        if mp is not None:
            return mp
        mp = MateriaPrima.query.filter_by(codigo=codigo).first()
        if mp is None:
            mp = MateriaPrima(codigo=codigo, descricao=descricao, unidade=unidade, custo_atual=custo or 0, categoria=categoria, ativo=True)
            db.session.add(mp)
            db.session.flush()
        mp_cache[codigo] = mp
        return mp

    def _get_or_create_produto(familia, codigo, descricao=None, categoria=None, chave_busca=None):
        key = (familia, codigo)
        p = produto_cache.get(key)
        if p is not None:
            return p
        p = Produto.query.filter_by(familia=familia, codigo=codigo).first()
        if p is None:
            p = Produto(familia=familia, codigo=codigo, descricao=descricao, categoria=categoria, chave_busca=chave_busca, ativo=True)
            db.session.add(p)
            db.session.flush()
        produto_cache[key] = p
        return p

    def _get_or_create_estrutura(produto, dn, ciclo_horas):
        e = EstruturaProduto.query.filter_by(produto_id=produto.id, dn=dn).first()
        if e is None:
            e = EstruturaProduto(produto_id=produto.id, dn=dn, ciclo_horas=ciclo_horas or 0, ativo=True)
            db.session.add(e)
            db.session.flush()
        return e

    def _add_item(estrutura, ordem, tipo, quantidade, materia_prima=None, subproduto=None, observacao=None):
        if not quantidade:
            return
        db.session.add(EstruturaProdutoItem(
            estrutura_id=estrutura.id, tipo=tipo, quantidade=quantidade,
            materia_prima_id=materia_prima.id if materia_prima else None,
            subproduto_id=subproduto.id if subproduto else None,
            observacao=observacao, ordem=ordem,
        ))

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws_param = wb["PARÂMETROS"]

    def _dn_str(v):
        if v is None:
            return None
        if isinstance(v, float) and v == int(v):
            return str(int(v))
        return str(v).strip()

    # ------------------------------------------------------------------
    # SUPERFLEX
    # ------------------------------------------------------------------
    def _mp_quimica(linha):
        """Reaproveita a MESMA matéria-prima "química" já cadastrada na fase 1
        (PARÂMETROS linhas 6-15, catálogo `quimicas_por_linha` do
        `_seed_custos_pig_mandril`) — mesmo código, não duplica o insumo."""
        desc = ws_param.cell(linha, 2).value
        preco = ws_param.cell(linha, 5).value
        codigo = "QUIM-" + "".join(ch for ch in desc.upper() if ch.isalnum())[:30]
        return _get_or_create_mp(codigo, desc, "kg", preco, "Química")

    MP_TDI_1270A = _mp_quimica(6)   # "12-70 A PRÉ-POLÍMERO TDI" — copo pistão
    MP_ATS85 = _mp_quimica(15)      # "ATS 85 PRÉ-POLÍMERO (AMINO) TDI" — eixo e disco espaçador

    ws_sf = wb["SUPERFLEX"]
    dn_labels = [_dn_str(ws_sf.cell(5, c).value) for c in range(2, 7)]  # "6''".."14''"

    produto_copo_pistao = _get_or_create_produto(
        "SUPERFLEX", "SUPERFLEX-COPO-PISTAO",
        descricao="SUPERFLEX — Copo pistão em PU (70-80 Shore A)", categoria="Componente")
    produto_eixo_pu = _get_or_create_produto(
        "SUPERFLEX", "SUPERFLEX-EIXO-PU",
        descricao="SUPERFLEX — Eixo em PU (80-90 Shore A)", categoria="Componente")
    produto_disco_espacador = _get_or_create_produto(
        "SUPERFLEX", "SUPERFLEX-DISCO-ESPACADOR",
        descricao="SUPERFLEX — Disco espaçador em PU (80-90 Shore A)", categoria="Componente")

    # Sub-tabela "COPO PISTAO: 70-80 SHORE A" (linhas 6-10, colunas L/N — peso/tempo)
    for i, r in enumerate(range(6, 11)):
        dn = dn_labels[i]
        peso = ws_sf.cell(r, 12).value or 0
        tempo_h = ws_sf.cell(r, 14).value or 0
        est = _get_or_create_estrutura(produto_copo_pistao, dn, tempo_h)
        _add_item(est, 0, "MATERIA_PRIMA", peso, materia_prima=MP_TDI_1270A, observacao="copo pistão (12-70 A)")

    # Sub-tabela "EIXO PU: 80-90 SHORE A" (linhas 6-10, colunas S/U)
    for i, r in enumerate(range(6, 11)):
        dn = dn_labels[i]
        peso = ws_sf.cell(r, 19).value or 0
        tempo_h = ws_sf.cell(r, 21).value or 0
        est = _get_or_create_estrutura(produto_eixo_pu, dn, tempo_h)
        _add_item(est, 0, "MATERIA_PRIMA", peso, materia_prima=MP_ATS85, observacao="eixo (ATS 85)")

    # Sub-tabela "DISCO ESPAÇADOR: 80-90 SHORE A" (linhas 14-18, colunas L/N)
    for i, r in enumerate(range(14, 19)):
        dn = dn_labels[i]
        peso = ws_sf.cell(r, 12).value or 0
        tempo_h = ws_sf.cell(r, 14).value or 0
        est = _get_or_create_estrutura(produto_disco_espacador, dn, tempo_h)
        _add_item(est, 0, "MATERIA_PRIMA", peso, materia_prima=MP_ATS85, observacao="disco espaçador (ATS 85)")

    MP_ANEL_TRAVAMENTO = _get_or_create_mp(
        "SUPERFLEX-ANEL-TRAVAMENTO-CABO-ACO", "Anel de travamento + cabo de aço (SUPERFLEX)",
        "un", 45, "Acessório")

    # Acréscimo literal do eixo no CS4 (fórmula "=B8+27"/"+38"/"+43"/"+73"/"+114" —
    # valor fixo por DN, sem decomposição em peso/preço na planilha; centralizado
    # como matéria-prima "consolidada" por DN, mesmo padrão já usado na fase 1
    # pra ELC_MG_PC).
    acrescimos_eixo_cs4 = [27, 38, 43, 73, 114]
    mp_acrescimo_cs4 = []
    for i, dn in enumerate(dn_labels):
        dn_num = _dn_str(ws_sf.cell(5, i + 2).value).replace("''", "").strip()
        mp = _get_or_create_mp(
            f"SUPERFLEX-CS4-EIXO-ACRESCIMO-DN{dn_num}",
            f"SUPERFLEX CS4 — acréscimo de eixo DN {dn} (valor fixo da planilha, sem decomposição)",
            "un", acrescimos_eixo_cs4[i], "Componente DN")
        mp_acrescimo_cs4.append(mp)

    produto_cs3 = _get_or_create_produto("SUPERFLEX", "SUPERFLEX-CS3", descricao="PIG SUPERFLEX-CS3", categoria="PIG", chave_busca="SUPERFLEX-CS3")
    produto_cs4 = _get_or_create_produto("SUPERFLEX", "SUPERFLEX-CS4", descricao="PIG SUPERFLEX-CS4", categoria="PIG", chave_busca="SUPERFLEX-CS4")

    for i, dn in enumerate(dn_labels):
        est3 = _get_or_create_estrutura(produto_cs3, dn, 3)  # HH MONTAGEM+QUALIDADE+PCP = 3h fixo, igual nas 2 variantes
        ordem = 0
        _add_item(est3, ordem, "SUBPRODUTO", 3, subproduto=produto_copo_pistao, observacao="copo pistão (3x)"); ordem += 1
        _add_item(est3, ordem, "SUBPRODUTO", 1, subproduto=produto_disco_espacador, observacao="disco espaçador"); ordem += 1
        _add_item(est3, ordem, "SUBPRODUTO", 1, subproduto=produto_eixo_pu, observacao="eixo"); ordem += 1
        _add_item(est3, ordem, "MATERIA_PRIMA", 1, materia_prima=MP_ANEL_TRAVAMENTO, observacao="anel de travamento + cabo de aço"); ordem += 1

        est4 = _get_or_create_estrutura(produto_cs4, dn, 3)
        ordem = 0
        _add_item(est4, ordem, "SUBPRODUTO", 4, subproduto=produto_copo_pistao, observacao="copo pistão (4x)"); ordem += 1
        _add_item(est4, ordem, "SUBPRODUTO", 1, subproduto=produto_disco_espacador, observacao="disco espaçador"); ordem += 1
        _add_item(est4, ordem, "SUBPRODUTO", 1, subproduto=produto_eixo_pu, observacao="eixo (base)"); ordem += 1
        _add_item(est4, ordem, "MATERIA_PRIMA", 1, materia_prima=mp_acrescimo_cs4[i], observacao="eixo — acréscimo CS4 (valor fixo da planilha)"); ordem += 1
        _add_item(est4, ordem, "MATERIA_PRIMA", 1, materia_prima=MP_ANEL_TRAVAMENTO, observacao="anel de travamento + cabo de aço"); ordem += 1

    # ------------------------------------------------------------------
    # SILICONE (família DISCFLEX)
    # ------------------------------------------------------------------
    MP_MANTA_SILICONE = _get_or_create_mp(
        "SILICONE-MANTA", ws_param.cell(32, 2).value, "kg", ws_param.cell(32, 5).value, "Química")
    MP_IMA_22X10 = _get_or_create_mp(
        "SILICONE-IMA-22X10MM", ws_param.cell(33, 2).value, "un", ws_param.cell(33, 5).value, "Acessório")
    MP_IMA_8X5 = _get_or_create_mp(
        "SILICONE-IMA-8X5MM", ws_param.cell(34, 2).value, "un", ws_param.cell(34, 5).value, "Acessório")

    ws_sil = wb["SILICONE"]

    def _tempo_para_horas(t):
        if t is None:
            return 0
        try:
            return t.hour + t.minute / 60 + t.second / 3600
        except AttributeError:
            return float(t) * 24  # fallback se vier como fração de dia (float)

    produto_discflex = _get_or_create_produto("SILICONE", "DISCFLEX", descricao="Disco DISCFLEX (manta silicone)", categoria="Sobressalente", chave_busca="DISCFLEX")
    produto_discflex_ima = _get_or_create_produto("SILICONE", "DISCFLEX-COM-IMA", descricao="Disco DISCFLEX com ímã (manta silicone)", categoria="Sobressalente", chave_busca="DISCFLEX COM IMA")
    produto_hsc_azul = _get_or_create_produto("SILICONE", "HSC-AZUL-COM-IMA", descricao="HSC azul com ímã (manta silicone)", categoria="Sobressalente", chave_busca="HSC AZUL COM IMA")

    dn_vistos_discflex = {}
    for r in range(6, 14):  # DISCFLEX "puro" — sem ímã
        dn_raw = _dn_str(ws_sil.cell(r, 5).value)
        peso = ws_sil.cell(r, 6).value or 0
        tempo_h = _tempo_para_horas(ws_sil.cell(r, 12).value)
        n = dn_vistos_discflex.get(dn_raw, 0)
        dn_vistos_discflex[dn_raw] = n + 1
        if n == 0:
            dn = dn_raw
        else:
            # 2 linhas com o mesmo rótulo de DN e pesos diferentes (planilha original,
            # sem explicação — linhas 9 e 10, 0,2834kg vs 0,2995kg) — desambiguado aqui
            # pra caber no modelo (UniqueConstraint produto+dn); sinalizado pro Bruno revisar
            # na observação da estrutura (o campo `dn` é VARCHAR(20), por isso o sufixo
            # curto aqui — o detalhe completo vai pra `observacao`, que é texto livre).
            dn = f"{dn_raw} (v{n + 1})"
        est = _get_or_create_estrutura(produto_discflex, dn, tempo_h)
        if n > 0 and not est.observacao:
            est.observacao = (
                f"Variante {n + 1} do DN {dn_raw} — a planilha original (aba SILICONE) tem 2 linhas "
                f"DISCFLEX com o mesmo rótulo de DN ({dn_raw}) e pesos de manta diferentes, sem "
                f"explicação (linhas 9 e 10: 0,2834kg vs 0,2995kg). Peso desta variante: {peso}kg. "
                "Sinalizado pro Bruno revisar a origem dessa duplicidade na planilha."
            )
        _add_item(est, 0, "MATERIA_PRIMA", peso, materia_prima=MP_MANTA_SILICONE, observacao="manta de silicone")

    for r in range(14, 16):  # DISCFLEX COM IMÃ — ímã 8x5mm (fórmula usa $B$3)
        dn = _dn_str(ws_sil.cell(r, 5).value)
        peso = ws_sil.cell(r, 6).value or 0
        qnt_ima = ws_sil.cell(r, 8).value or 0
        tempo_h = _tempo_para_horas(ws_sil.cell(r, 12).value)
        est = _get_or_create_estrutura(produto_discflex_ima, dn, tempo_h)
        _add_item(est, 0, "MATERIA_PRIMA", peso, materia_prima=MP_MANTA_SILICONE, observacao="manta de silicone")
        _add_item(est, 1, "MATERIA_PRIMA", qnt_ima, materia_prima=MP_IMA_8X5, observacao="ímã 8x5mm")

    for r in range(16, 17):  # HSC AZUL COM IMÃ — ímã 22x10mm (fórmula usa $B$2)
        dn = _dn_str(ws_sil.cell(r, 5).value)
        peso = ws_sil.cell(r, 6).value or 0
        qnt_ima = ws_sil.cell(r, 8).value or 0
        tempo_h = _tempo_para_horas(ws_sil.cell(r, 12).value)
        est = _get_or_create_estrutura(produto_hsc_azul, dn, tempo_h)
        _add_item(est, 0, "MATERIA_PRIMA", peso, materia_prima=MP_MANTA_SILICONE, observacao="manta de silicone")
        _add_item(est, 1, "MATERIA_PRIMA", qnt_ima, materia_prima=MP_IMA_22X10, observacao="ímã 22x10mm")

    db.session.add(ControleSistema(chave=_CHAVE_SEED_CUSTOS_SUPERFLEX_SILICONE_20_09_2026))
    db.session.commit()
    app.logger.info(
        "Gestão de Custos: importação de SUPERFLEX/SILICONE (fase 3) concluída (%d matérias-primas, %d produtos, %d estruturas).",
        MateriaPrima.query.count(), Produto.query.count(), EstruturaProduto.query.count(),
    )

_CHAVE_IMPORTAR_HISTORICO_CUSTOS_MP_20_09_2026 = "importar_historico_custos_mp_20_09_2026"

# Alias explícito pra itens do log manual da aba "EVOLUÇÃO DE CUSTOS - MP" cujo texto
# não bate 1:1 com a `descricao` já cadastrada no catálogo (a maioria bate exata, por
# vir literalmente de PARÂMETROS — estes 4 são os acessórios exclusivos de HLCC/HLCC PC,
# cuja descrição no catálogo foi escrita a partir do cabeçalho da aba, não de PARÂMETROS;
# confirmados pelo preço batendo exato entre o log e o cadastro, não só pelo nome).
_ALIAS_HISTORICO_MP = {
    'PRENSA CABO 3/16" (unidade)': "ESPUMA-PRENSA-CABO",
    "KIT FIXAÇÃO HLCC-PC (conjunto)": "ESPUMA-KIT-ARRUELA-PORCA-HLCCPC",
    'BARRA ROSCADA 3/4" (metro)': "ESPUMA-BARRA-ROSCADA-HLCCPC",
    "BUMPER PU - COIM (kg)": "ESPUMA-BUMPER-PU-HLCCPC",
}


def _importar_historico_custos_manual(app):
    """Importa (uma única vez) o log manual de mudanças de preço que o Bruno já vinha
    mantendo à mão na aba "EVOLUÇÃO DE CUSTOS - MP" da planilha (tabela "REGISTRO DE
    ATUALIZAÇÕES DE PREÇO", colunas F-N) pra dentro de `MateriaPrimaHistorico` — assim a
    tela de Histórico (fase 3) mostra a trajetória real de preço desde que ele começou a
    registrar (nov/2024), não só as mudanças feitas a partir de agora dentro do app.

    NÃO importa o log da aba "EVOLUÇÃO DE CUSTOS - PRODUTOS": checado célula a célula, esse
    log é só 1 fotografia única (811 linhas, todas com observação "Carga inicial (snapshot
    do índice BUSCA DE CUSTO)", datadas do mesmo dia 07/08/2026, zero mudança real registrada
    depois disso) — e custo de produto no app nunca é armazenado, é sempre recalculado ao
    vivo a partir da matéria-prima/hora-homem correntes (`_custo_estrutura_produto`), então
    guardar essa fotografia como "histórico" seria só duplicar um número já derivável e
    ficaria descolado do valor real assim que qualquer matéria-prima mudasse de preço depois
    dela. Se Bruno passar a registrar mudanças de custo de produto de verdade na planilha
    (histórico com mais de 1 ponto por produto), dá pra reconsiderar numa fase futura.

    Casamento log→catálogo: por `descricao` exata (cobre a maioria — vem literalmente da
    aba PARÂMETROS nos dois lados) + um alias explícito pra 4 itens exclusivos de HLCC/HLCC
    PC cuja descrição no catálogo foi escrita a partir do cabeçalho da aba, não de
    PARÂMETROS (confirmados pelo preço batendo exato, não só pelo nome — ver
    `_ALIAS_HISTORICO_MP`). 3 linhas do log não têm correspondência confiável no catálogo
    atual (texto diferente E preço não bate com nada) — não importadas, só logadas como aviso
    pro Bruno decidir se são itens obsoletos ou se falta cadastrar: "8086 BUMPER PRÉ-POLÍMERO
    TDI", "MOCA / CURATIVO (kg)" (grafia antiga, distinta de "MOCA CURATIVO TDI" já
    cadastrada) e "CABO DE AÇO (metro)" (o app usa a MESMA matéria-prima da corda pro olhal
    de cabo de aço da HLCC/HLCC PC — achado da fase 2, revisão pendente com o Bruno — não um
    item de catálogo próprio)."""
    if ControleSistema.query.filter_by(chave=_CHAVE_IMPORTAR_HISTORICO_CUSTOS_MP_20_09_2026).first() is not None:
        return

    xlsx_path = os.path.join(BASE_DIR, "data", "custo_de_producao_20_09_2026.xlsx")
    if not os.path.exists(xlsx_path):
        app.logger.warning("Gestão de Custos: planilha de importação não encontrada em %s — histórico manual não importado.", xlsx_path)
        return

    import openpyxl
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb["EVOLUÇÃO DE CUSTOS - MP"]

    mps_por_descricao = {mp.descricao: mp for mp in MateriaPrima.query.all()}

    importados = 0
    nao_encontrados = []
    for r in range(21, ws.max_row + 1):
        item = ws.cell(r, 8).value
        if not item:
            continue
        item = str(item).strip()
        preco = ws.cell(r, 10).value
        data_registro = ws.cell(r, 6).value
        observacao = (ws.cell(r, 13).value or "").strip()
        if preco is None or data_registro is None:
            continue

        mp = mps_por_descricao.get(item)
        if mp is None and item in _ALIAS_HISTORICO_MP:
            mp = MateriaPrima.query.filter_by(codigo=_ALIAS_HISTORICO_MP[item]).first()
        if mp is None:
            nao_encontrados.append(item)
            continue

        db.session.add(MateriaPrimaHistorico(
            materia_prima_id=mp.id, custo_anterior=None, custo_novo=preco,
            motivo=(observacao or "Carga inicial (migrado do log manual da planilha)")[:300],
            usuario_nome="Importação (planilha)", criado_em=data_registro,
        ))
        importados += 1

    db.session.add(ControleSistema(chave=_CHAVE_IMPORTAR_HISTORICO_CUSTOS_MP_20_09_2026))
    db.session.commit()
    app.logger.info(
        "Gestão de Custos: histórico manual de matéria-prima importado (%d registros; %d itens do log sem correspondência no catálogo: %s).",
        importados, len(nao_encontrados), ", ".join(nao_encontrados) if nao_encontrados else "nenhum",
    )


_CHAVE_SEED_CUSTOS_PIG_ALOJAMENTO_20_09_2026 = "seed_custos_pig_alojamento_20_09_2026"


def _seed_custos_pig_alojamento(app):
    """Importa (uma única vez) o custo de "ALOJAMENTO" (embalagem/caixa do PIG) que já
    existe como coluna própria nas abas LBD e LUN (coluna BI, "CUSTO TOTAL + ALOJAMENTO"
    em BJ) mas nunca tinha sido trazido pro app nas fases 1-3 — um valor literal (não
    fórmula) por DN, R$100 até DN 8 e R$120 de DN 10 em diante, IDÊNTICO nas duas abas
    (LBD e LUN), conferido linha a linha antes de escrever este seed.

    Motivado pelo pedido do Bruno (21/09/2026) de um configurador de acessórios pro
    LBD/LUN — ao montar esse configurador ficou claro que Alojamento é a única coluna
    de custo que existe na planilha do LBD/LUN e nunca tinha entrado no app. Cadastrado
    aqui como matéria-prima de catálogo (`ALOJAMENTO-DN{dn}`, 1 por DN, compartilhada
    entre LBD e LUN já que o valor é o mesmo) — assim fica editável/com histórico igual
    qualquer outra matéria-prima. Importante: isso NÃO é adicionado automaticamente na
    EstruturaProduto do LBD/LUN (não quero mudar em silêncio o custo total que já é
    mostrado hoje em Custos dos Produtos/Necessidades do PCP) — fica só como opção no
    novo configurador de acessórios, pro Bruno decidir se/quando incluir. Sinalizar pra
    ele: talvez faça sentido esse valor entrar sempre por padrão no custo do PIG (parece
    ser custo de embalagem obrigatório, não um acessório opcional de verdade) — decisão
    dele, não assumida aqui."""
    if ControleSistema.query.filter_by(chave=_CHAVE_SEED_CUSTOS_PIG_ALOJAMENTO_20_09_2026).first() is not None:
        return

    xlsx_path = os.path.join(BASE_DIR, "data", "custo_de_producao_20_09_2026.xlsx")
    if not os.path.exists(xlsx_path):
        app.logger.warning("Gestão de Custos: planilha de importação não encontrada em %s — seed de Alojamento não executado.", xlsx_path)
        return

    import openpyxl

    def _dn_str(v):
        if v is None:
            return None
        if isinstance(v, float) and v == int(v):
            return str(int(v))
        return str(v).strip()

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws_lbd = wb["LBD"]

    criados = 0
    for r in range(6, 28):
        dn_val = ws_lbd.cell(r, 2).value
        alojamento = ws_lbd.cell(r, 61).value  # BI
        if dn_val is None or alojamento is None:
            continue
        dn = _dn_str(dn_val)
        codigo = f"ALOJAMENTO-DN{dn}"
        if MateriaPrima.query.filter_by(codigo=codigo).first() is None:
            db.session.add(MateriaPrima(
                codigo=codigo, descricao=f"Alojamento (embalagem) — PIG DN {dn}",
                unidade="un", custo_atual=alojamento, categoria="Acessório", ativo=True,
            ))
            criados += 1

    db.session.add(ControleSistema(chave=_CHAVE_SEED_CUSTOS_PIG_ALOJAMENTO_20_09_2026))
    db.session.commit()
    app.logger.info("Gestão de Custos: matérias-primas de Alojamento importadas (%d criadas).", criados)


_CHAVE_SEED_CUSTOS_PIG_CALANDRA_20_09_2026 = "seed_custos_pig_calandra_20_09_2026"


def _seed_custos_pig_calandra(app):
    """Importa (uma única vez) o custo de "CALANDRA" — achado ao investigar a pergunta
    do Bruno sobre o Disco Espaçador somar automaticamente no configurador de acessórios
    (21/09/2026): a aba LBD tem, a partir da linha 30, um CONFIGURADOR DE ACESSÓRIOS
    PRÓPRIO da planilha (colunas A-S, que eu tinha lido errado na fase 1 como "bloco
    vazio" — só tinha checado as colunas BH/BI/BJ daquele intervalo de linhas, que
    realmente ficam vazias ali; os dados reais estão nas colunas A-S, num layout
    diferente da tabela principal) — e ele mostra que marcar ELC (AÇO) ou ELP (PP) soma,
    além do custo do próprio item (ELC_MG_PC), um custo extra de "CALANDRA" (processo de
    calandragem, tabela própria "CUSTOS CALANDRA" em LBD!A76:B99 / LUN!A77:B100 — só
    ELC/ELP, CINTA MAGNÉTICA e PLACA CALIBRADORA não somam calandra, confirmado nas
    fórmulas D32/F32 vs H32/J32). Valor idêntico nas abas LBD e LUN (conferido linha a
    linha), então uma matéria-prima só por DN (`CALANDRA-DN{dn}`), reaproveitada pelas
    duas. R$0 pra DN 2/3/4 (não se aplica nesse tamanho).

    Não mexe na EstruturaProduto "oficial" do LBD/LUN, mesmo raciocínio do Alojamento —
    fica só disponível pro configurador de acessórios somar quando ELC/ELP for marcado."""
    if ControleSistema.query.filter_by(chave=_CHAVE_SEED_CUSTOS_PIG_CALANDRA_20_09_2026).first() is not None:
        return

    xlsx_path = os.path.join(BASE_DIR, "data", "custo_de_producao_20_09_2026.xlsx")
    if not os.path.exists(xlsx_path):
        app.logger.warning("Gestão de Custos: planilha de importação não encontrada em %s — seed de Calandra não executado.", xlsx_path)
        return

    import openpyxl

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws_lbd = wb["LBD"]

    criados = 0
    for r in range(77, 100):
        rotulo = ws_lbd.cell(r, 1).value  # "DN2", "DN6", ...
        valor = ws_lbd.cell(r, 2).value
        if not rotulo or not str(rotulo).startswith("DN") or valor is None:
            continue
        dn = str(rotulo)[2:].strip()
        codigo = f"CALANDRA-DN{dn}"
        if MateriaPrima.query.filter_by(codigo=codigo).first() is None:
            db.session.add(MateriaPrima(
                codigo=codigo, descricao=f"Calandra (processo) — acessório ELC/ELP PIG DN {dn}",
                unidade="un", custo_atual=valor, categoria="Acessório", ativo=True,
            ))
            criados += 1

    db.session.add(ControleSistema(chave=_CHAVE_SEED_CUSTOS_PIG_CALANDRA_20_09_2026))
    db.session.commit()
    app.logger.info("Gestão de Custos: matérias-primas de Calandra importadas (%d criadas).", criados)


# Prefixos e códigos exatos de MateriaPrima que correspondem a uma linha
# validada de verdade na aba PARÂMETROS da planilha original (as 2 tabelas
# de "químicas" — PU CAST/LBD e Espuma/Silicone — e as famílias de
# componente-por-DN da tabela LBD_REV A e do bloco PIGS EM BORRACHA). Tudo
# que NÃO cai aqui é um "wrapper" que criei pra representar um produto
# acabado ou um custo já consolidado (Placa Calibradora, Cinta Magnética,
# ELC, ELP, Alojamento, Calandra, os discos "PU CAST consolidado", os
# acessórios exclusivos de HLCC/HLCC PC e os valores literais de SUPERFLEX)
# — não é matéria-prima de verdade, mesmo tendo entrado no catálogo pra
# poder participar das fórmulas de custo.
_ORIGEM_MP_PREFIXOS_PARAMETROS = (
    "QUIM-",
    "TUBO-DN", "FLANGE-BUMPER-DN", "FLANGE-SOLDA-DN", "PARAFUSO-DN", "ARRUELA-DN", "PORCA-DN",
    "COPO-BORRACHA-EPDM-DN", "COPO-BORRACHA-BUNA-DN", "COPO-BORRACHA-VITON-DN",
    "EIXO-BORRACHA-DN", "CABECOTE-BORRACHA-DN", "NYLON-BORRACHA-DN", "PORCA-BORRACHA-DN", "FLANGE-BORRACHA-DN",
)
_ORIGEM_MP_CODIGOS_PARAMETROS = {
    "ESPUMA-BLOCO-D26", "ESPUMA-BLOCO-D45", "ESPUMA-BLOCO-D60", "ESPUMA-BLOCO-D80",
    "ESPUMA-ELASTOMERO-TECPUR", "ESPUMA-AMINO-A-ALTA", "ESPUMA-AMINO-A-MEDIA", "ESPUMA-AMINO-B-ISO",
    "ESPUMA-PIGMENTO", "ESPUMA-CORDA-OLHAL", "ESPUMA-ESCOVA-FINA-HLR", "ESPUMA-ESCOVA-GROSSA-HLR",
    "ESPUMA-VELCRO-HLR-V", "ESPUMA-COLA-SAPATEIRO",
    "SILICONE-MANTA", "SILICONE-IMA-22X10MM", "SILICONE-IMA-8X5MM",
}

# Legenda amigável por seção — mesmo agrupamento da aba PARÂMETROS, na
# mesma ordem em que as seções aparecem lá (usado só na tela, pra montar os
# grupos exibidos ao Bruno).
_ORIGEM_MP_SECOES = (
    ("Matérias-primas químicas — PU CAST / LBD (poliuretano fundido)", ("QUIM-",)),
    ("Matérias-primas químicas e insumos — Espuma (blocos / envase A+B)", (
        "ESPUMA-BLOCO-", "ESPUMA-ELASTOMERO-TECPUR", "ESPUMA-AMINO-", "ESPUMA-PIGMENTO",
        "ESPUMA-CORDA-OLHAL", "ESPUMA-ESCOVA-", "ESPUMA-VELCRO-", "ESPUMA-COLA-SAPATEIRO",
    )),
    ("Matéria-prima — Silicone (manta e ímãs)", ("SILICONE-",)),
    ("Componentes do corpo do PIG por DN — tubo, flange, parafuso, arruela, porca (LBD_REV A)", (
        "TUBO-DN", "FLANGE-BUMPER-DN", "FLANGE-SOLDA-DN", "PARAFUSO-DN", "ARRUELA-DN", "PORCA-DN",
    )),
    ("Copo de borracha por DN e material — PIGS EM BORRACHA", ("COPO-BORRACHA-",)),
    ("Componentes comuns por DN — PIGS EM BORRACHA (eixo, cabeçote, nylon, porca, flange)", (
        "EIXO-BORRACHA-DN", "CABECOTE-BORRACHA-DN", "NYLON-BORRACHA-DN", "PORCA-BORRACHA-DN", "FLANGE-BORRACHA-DN",
    )),
)


def _classificar_origem_materia_prima(codigo):
    """"PARAMETROS" = corresponde a uma linha validada na aba PARÂMETROS da
    planilha original — é o que o Bruno pediu pra ver na tela principal de
    Matéria-Prima. "DERIVADO" = tudo o resto: produto acabado ou custo já
    consolidado modelado como matéria-prima só pra poder entrar nas
    fórmulas (Placa Calibradora, Cinta Magnética, ELC/ELP, Alojamento,
    Calandra, discos "PU CAST consolidado", acessórios HLCC/HLCC PC,
    valores literais de SUPERFLEX) — continua no banco (o motor de custo
    depende disso pra calcular o total dos produtos), só não aparece mais
    misturado na tela principal de Matéria-Prima."""
    if codigo in _ORIGEM_MP_CODIGOS_PARAMETROS:
        return "PARAMETROS"
    if codigo.startswith(_ORIGEM_MP_PREFIXOS_PARAMETROS):
        return "PARAMETROS"
    return "DERIVADO"


def _secao_materia_prima(codigo):
    """Nome da seção da aba PARÂMETROS que esta matéria-prima pertence, pra
    agrupar a tela igual à planilha original. None pras não-PARAMETROS."""
    for nome, prefixos in _ORIGEM_MP_SECOES:
        if codigo.startswith(prefixos):
            return nome
    return None


def _migrar_materia_prima_origem_planilha(app):
    """Pedido do Bruno (21/09/2026, com print anexo mostrando Placa
    Calibradora/Cinta Magnética/Escova dentro da tela de Matéria-Prima):
    "dentro da aba Matéria Prima, você está misturando os custos de
    produto e custo de matéria prima [...] quero que você valide as
    matérias-primas somente os itens dentro da aba parâmetro [...] deixe
    mais organizado essa aba, agrupados conforme está exatamente na aba
    parâmetros [...] citando o fornecedor, detalhando o nome".

    Adiciona a coluna `origem_planilha` (ALTER TABLE idempotente, mesmo
    padrão de `_migrar_usuarios_role`) e classifica CADA matéria-prima do
    catálogo entre "PARAMETROS" (aparece na tela principal, agrupada por
    seção) e "DERIVADO" (fica de fora da tela principal — só acessível via
    o link de edição direto, ex. a partir da composição do produto — mas
    continua ativa no banco, porque o motor de custo depende dela).

    Roda em TODO boot (não só uma vez): a classificação é 100% derivada do
    código da matéria-prima (barato, ~500 linhas), então isso também
    classifica automaticamente qualquer matéria-prima nova que um seed
    futuro venha a criar, sem precisar lembrar de manter uma lista à parte
    sincronizada."""
    inspector = inspect(db.engine)
    if "materias_primas" not in inspector.get_table_names():
        return  # banco novo — db.create_all() já cuidou de tudo

    colunas = {c["name"] for c in inspector.get_columns("materias_primas")}
    if "origem_planilha" not in colunas:
        with db.engine.begin() as conn:
            conn.execute(text("ALTER TABLE materias_primas ADD COLUMN origem_planilha VARCHAR(20)"))

    reclassificadas = 0
    for mp in MateriaPrima.query.all():
        nova_origem = _classificar_origem_materia_prima(mp.codigo)
        if mp.origem_planilha != nova_origem:
            mp.origem_planilha = nova_origem
            reclassificadas += 1
    if reclassificadas:
        db.session.commit()
        app.logger.info("Gestão de Custos: %d matérias-primas (re)classificadas por origem_planilha.", reclassificadas)

    # Backfill do fornecedor/data de atualização pros itens "PARAMETROS" que
    # a planilha original já trazia em coluna própria (FORNECEDOR / REF. e
    # DATA ATUALIZAÇÃO, aba PARÂMETROS) — nenhum seed anterior capturou isso.
    # Só preenche onde ainda está vazio, pra nunca sobrescrever uma edição
    # manual feita depois pelo Bruno na tela de edição.
    pendentes = MateriaPrima.query.filter(
        MateriaPrima.origem_planilha == "PARAMETROS",
        MateriaPrima.fornecedor.is_(None),
    ).all()
    if not pendentes:
        return

    xlsx_path = os.path.join(BASE_DIR, "data", "custo_de_producao_20_09_2026.xlsx")
    if not os.path.exists(xlsx_path):
        app.logger.warning("Gestão de Custos: planilha de importação não encontrada em %s — fornecedor/data não preenchidos.", xlsx_path)
        return

    import openpyxl

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws_param = wb["PARÂMETROS"]

    fornecedor_por_codigo, data_por_codigo = {}, {}
    for r in range(6, 16):  # químicas PU CAST / LBD — PARÂMETROS linhas 6-15
        desc = ws_param.cell(r, 2).value
        if not desc:
            continue
        codigo = "QUIM-" + "".join(ch for ch in desc.upper() if ch.isalnum())[:30]
        fornecedor_por_codigo[codigo] = ws_param.cell(r, 3).value
        data_por_codigo[codigo] = ws_param.cell(r, 4).value

    mapa_linha_espuma_silicone = {
        18: "ESPUMA-BLOCO-D26", 19: "ESPUMA-BLOCO-D45", 20: "ESPUMA-BLOCO-D60", 21: "ESPUMA-BLOCO-D80",
        22: "ESPUMA-ELASTOMERO-TECPUR", 23: "ESPUMA-AMINO-A-ALTA", 24: "ESPUMA-AMINO-A-MEDIA", 25: "ESPUMA-AMINO-B-ISO",
        26: "ESPUMA-PIGMENTO", 27: "ESPUMA-CORDA-OLHAL", 28: "ESPUMA-ESCOVA-FINA-HLR", 29: "ESPUMA-ESCOVA-GROSSA-HLR",
        30: "ESPUMA-VELCRO-HLR-V", 31: "ESPUMA-COLA-SAPATEIRO",
        32: "SILICONE-MANTA", 33: "SILICONE-IMA-22X10MM", 34: "SILICONE-IMA-8X5MM",
    }  # PARÂMETROS linhas 18-34
    for r, codigo in mapa_linha_espuma_silicone.items():
        fornecedor_por_codigo[codigo] = ws_param.cell(r, 3).value
        data_por_codigo[codigo] = ws_param.cell(r, 4).value

    atualizadas = 0
    for mp in pendentes:
        fornecedor = fornecedor_por_codigo.get(mp.codigo)
        data_valor = data_por_codigo.get(mp.codigo)
        mudou = False
        if fornecedor:
            mp.fornecedor = str(fornecedor).strip()
            mudou = True
        if isinstance(data_valor, datetime):
            mp.data_atualizacao_fornecedor = data_valor.date()
            mudou = True
        elif isinstance(data_valor, date):
            mp.data_atualizacao_fornecedor = data_valor
            mudou = True
        if mudou:
            atualizadas += 1
    if atualizadas:
        db.session.commit()
        app.logger.info("Gestão de Custos: fornecedor/data preenchidos pra %d matérias-primas (aba PARÂMETROS).", atualizadas)


# Famílias em que a MESMA dn tem mais de uma EstruturaProduto, diferindo só
# pela densidade — hoje embutida como texto dentro do próprio `dn` (ex.
# "10'' MÉDIA", "1'' BAIXA D26"). Mapeia produto.codigo -> regex que extrai
# (numero_dn, densidade) do texto de `dn` já cadastrado. Usado só por
# `_migrar_densidade_estrutura_produto` pra preencher o classificador
# redundante `EstruturaProduto.densidade` — nunca reescreve `dn` (evita
# qualquer risco de colidir com a UniqueConstraint(produto_id, dn) já
# existente).
_RE_DN_DENSIDADE_COMPOSTO = re.compile(r"^\s*[\d.,]+\s*''?\s*(ALTA|MÉDIA|MEDIA|BAIXA(?:\s+D\d+)?)\s*$", re.I)


def _migrar_densidade_estrutura_produto(app):
    """Pedido do Bruno (22/09/2026, junto com o pedido de mostrar kg de
    matéria-prima no Kanban): "Modelo produto: HLR, DENSIDADE: ALTA, DN:
    4''" — ele percebeu (confirmado investigando o matching real do PCP)
    que hoje o casamento ItemPedido -> Produto pra família de espuma ou
    ACEITA QUALQUER densidade pro mesmo DN (`_matching_produto_pcp` não
    desambiguava ALTA x MÉDIA — risco real de casar com a estrutura ERRADA
    e reportar kg errado) ou nem casa (família "H" tinha chave_busca NULL).

    Adiciona a coluna `densidade` (ALTER TABLE idempotente, mesmo padrão de
    `_migrar_materia_prima_origem_planilha`) e classifica TODA
    EstruturaProduto das famílias com essa ambiguidade, extraindo o texto
    de densidade de dentro do `dn` já cadastrado (nunca altera `dn` em si —
    evita qualquer risco na UniqueConstraint(produto_id, dn)). Roda em todo
    boot, 100% determinístico a partir de `dn`, então também classifica
    qualquer estrutura nova que um seed futuro venha a criar."""
    inspector = inspect(db.engine)
    if "custos_estruturas_produto" not in inspector.get_table_names():
        return  # banco novo — db.create_all() já cuidou de tudo

    colunas = {c["name"] for c in inspector.get_columns("custos_estruturas_produto")}
    if "densidade" not in colunas:
        with db.engine.begin() as conn:
            conn.execute(text("ALTER TABLE custos_estruturas_produto ADD COLUMN densidade VARCHAR(30)"))
        app.logger.info("Migração automática: coluna densidade adicionada em custos_estruturas_produto.")

    familias_com_densidade = (
        "H", "HS", "HL", "HDISC", "HLR", "HLR X", "HLR R", "HLR V", "HLB", "HLCC", "HLCC PC",
    )
    produtos = Produto.query.filter(Produto.familia.in_(familias_com_densidade)).all()
    reclassificadas = 0
    for produto in produtos:
        for estrutura in produto.estruturas:
            m = _RE_DN_DENSIDADE_COMPOSTO.match(estrutura.dn or "")
            nova_densidade = m.group(1).upper().replace("MEDIA", "MÉDIA") if m else None
            if estrutura.densidade != nova_densidade:
                estrutura.densidade = nova_densidade
                reclassificadas += 1
    if reclassificadas:
        db.session.commit()
        app.logger.info("Gestão de Custos: %d EstruturaProduto (re)classificadas por densidade.", reclassificadas)

    # As 3 variantes do produto "H" (H, H-COM-SELO, H-COM-SELO-CORDA) nasceram
    # com chave_busca NULL (nunca tinham `''` no meio do código pra virar
    # chave, diferente de "HLR", "HS" etc.) — isso fazia `_matching_produto_pcp`
    # NUNCA casar a família inteira, mesmo com dn/densidade corretos.
    # Backfill único, idempotente (só preenche se ainda estiver vazio — nunca
    # sobrescreve uma edição manual feita depois pelo Bruno).
    # "HFLEX" (sem separador nenhum) nunca casava com o texto real dos
    # pedidos ("H-FLEX", "H- FLEX") porque `_normalizar_texto_matching_custos`
    # troca hífen por ESPAÇO dos dois lados — "H-FLEX" normaliza pra "H FLEX"
    # (com espaço), que não é a mesma string que "HFLEX" (sem espaço).
    # Corrigido pra "H FLEX" — idempotente, sempre reescreve pro valor
    # canônico certo (diferente do backfill de família H acima, que só
    # preenche se ainda estiver NULL).
    chaves_corrigidas = {"H": "H", "H-COM-SELO": "H COM SELO", "H-COM-SELO-CORDA": "H COM SELO CORDA", "HFLEX": "H FLEX"}
    mudou_chave = False
    for codigo, chave in chaves_corrigidas.items():
        p = Produto.query.filter_by(codigo=codigo).first()
        if p and p.chave_busca != chave:
            p.chave_busca = chave
            mudou_chave = True
    if mudou_chave:
        db.session.commit()
        app.logger.info("Gestão de Custos: chave_busca corrigida pra família H (H, H-COM-SELO, H-COM-SELO-CORDA) e HFLEX (-> \"H FLEX\").")


def _pagina_inicial(usuario):
    """Pra onde mandar o usuário logo após o login (e se ele visitar /login
    já autenticado) — normalmente a Listagem Geral de Produção, mas o papel
    PD (Líder de P&D, restrito só a P&D + Estações — ver permissoes.py)
    bateria num 403 ali, então vai direto pro Dashboard de P&D."""
    if usuario.role == "PD":
        return url_for("pd_dashboard")
    return url_for("dashboard")


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


def _calcular_resumo_filtrado(linhas):
    """Mesmos cards de resumo de _calcular_resumo (topo da Listagem Geral),
    mas recalculados em cima do que está de fato filtrado/visível na tela —
    pedido do Bruno (03/09/2026): "quero que TODO esse painel seja
    totalmente dinâmico... se eu filtrar algo como mês, região, estação,
    status, quero que automaticamente ele atualize". Antes esses 6 números
    vinham sempre de _calcular_resumo() (banco inteiro, sem filtro nenhum) —
    por isso nunca mudavam ao filtrar a tabela abaixo.

    `linhas` é a lista já achatada (1 linha por item, ver _LinhaListagemGeral)
    depois de TODOS os filtros da Listagem Geral já aplicados. Total de
    pedidos/Pendentes/Em tratativa/Em andamento/Finalizados contam PEDIDOS
    distintos entre essas linhas — usando o status "rollup" de cada pedido
    (Pedido.status_producao, o mesmo critério já usado pelo próprio filtro de
    Status em _predicado_status), não o status do item individual. Valor
    total soma só os ITENS que aparecem nas linhas (bate exatamente com a
    tabela logo abaixo — se um pedido tem 5 itens e só 2 passaram no filtro
    de estação, por exemplo, conta só esses 2, não o pedido inteiro).

    Calculado em Python sobre objetos que a rota já carregou (nenhuma query
    nova ao banco) — os mesmos `pedido`/`item` que _linhas_listagem_geral
    usou pra montar as linhas."""
    pedidos_vistos = {}
    valor_total = 0.0
    for l in linhas:
        pedidos_vistos[l.pedido_id] = l.pedido
        valor_total += l.venda_total or 0

    contagem = {"PENDENTE": 0, "EM TRATATIVA": 0, "ANDAMENTO": 0, "FINALIZADO": 0}
    for pedido in pedidos_vistos.values():
        status = pedido.status_producao
        if status in contagem:
            contagem[status] += 1

    return {
        "total": len(pedidos_vistos),
        "pendente": contagem["PENDENTE"],
        "em_tratativa": contagem["EM TRATATIVA"],
        "andamento": contagem["ANDAMENTO"],
        "finalizado": contagem["FINALIZADO"],
        "valor_total": round(valor_total, 2),
    }


def _resumo_otd(query=None):
    """Estatísticas de OTD (On-Time Delivery) da Gestão Operação — usa o campo
    go_otd_realizado (SIM/NÃO preenchido manualmente na planilha/tela), bem
    mais confiável que o _otd_percentual() antigo (que depende de datas quase
    nunca preenchidas no histórico). Só considera pedidos com OTD preenchido —
    quem ainda não tem essa informação fica de fora do percentual (não conta
    como "não cumpriu").

    `query` (opcional): uma query de PedidoOperacao já filtrada (ver
    _filtrar_pedidos_operacao) — pedido do Bruno (03/09/2026): filtrar
    Resultados/OTD por mês/segmento (planejamento semanal PCP / faturados) e
    ver o OTD só desse recorte. Sem argumento, mantém o comportamento antigo
    (todos os pedidos).

    Gestão Operação tem tabela própria (PedidoOperacao) — cada linha já é um
    pedido comercial, sem precisar agrupar nada em tempo de execução."""
    base = query if query is not None else PedidoOperacao.query
    pedidos = base.filter(PedidoOperacao.go_otd_realizado.isnot(None)).all()

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


def _media_dias(valores):
    """Média simples de uma lista de inteiros/None, ignorando os None — mesmo
    critério do resto do resumo de Resultados/OTD: quem não tem o dado
    calculável fica de fora da média (não conta como zero). Devolve
    (média arredondada em 1 casa, quantidade de pedidos que entraram na
    conta) — a quantidade é sempre mostrada junto do número no card (pedido
    do Bruno: "bem claro, didático"), pra deixar claro quando a média vem de
    poucos pedidos."""
    validos = [v for v in valores if v is not None]
    if not validos:
        return None, 0
    return round(sum(validos) / len(validos), 1), len(validos)


def _resumo_lead_times(query):
    """Lead times médios da Gestão Operação pra tela Resultados/OTD — pedido
    do Bruno (03/09/2026): "preciso ver os resultados detalhados de cada
    mês, como otd, lead time operação, lead time chão de fábrica (produção),
    lead time operação cif, lead time operação fob". Recebe a MESMA query já
    filtrada por mês/segmento usada em _resumo_otd (ver
    _filtrar_pedidos_operacao) — cada card do resumo reflete o mesmo recorte.

      - lt_operacao: PedidoOperacao.go_lead_time_operacao_dias (inclusão do
        pedido até entrega efetiva no cliente) — geral, depois quebrado por
        modalidade de frete (CIF/FOB, ver FRETE_OPCOES).
      - lt_frete: PedidoOperacao.go_lead_time_frete_dias (expedição até
        entrega/coleta real).
      - lt_producao: "chão de fábrica" — vem de Gestão Produção (ItemPedido.
        lt_producao_dias: início produção até término inspeção), cruzado
        pelo mesmo casamento por pedido_venda (trim, sem FK, nunca
        aproximado) já usado em _itens_producao_por_pedido_venda. Pedidos
        que ainda não foram lançados em Gestão Produção, ou cujos itens
        ainda não têm as duas datas, ficam de fora da média."""
    pedidos = query.all()

    lt_operacao_media, lt_operacao_n = _media_dias(p.go_lead_time_operacao_dias for p in pedidos)
    lt_operacao_cif_media, lt_operacao_cif_n = _media_dias(
        p.go_lead_time_operacao_dias for p in pedidos if p.frete == "CIF"
    )
    lt_operacao_fob_media, lt_operacao_fob_n = _media_dias(
        p.go_lead_time_operacao_dias for p in pedidos if p.frete == "FOB"
    )
    lt_frete_media, lt_frete_n = _media_dias(p.go_lead_time_frete_dias for p in pedidos)

    itens_por_pedido = _itens_producao_por_pedido_venda([p.pedido_venda for p in pedidos])
    valores_lt_producao = [
        item.lt_producao_dias for itens in itens_por_pedido.values() for item in itens
    ]
    lt_producao_media, lt_producao_n = _media_dias(valores_lt_producao)

    return {
        "lt_operacao": {"media": lt_operacao_media, "n": lt_operacao_n},
        "lt_operacao_cif": {"media": lt_operacao_cif_media, "n": lt_operacao_cif_n},
        "lt_operacao_fob": {"media": lt_operacao_fob_media, "n": lt_operacao_fob_n},
        "lt_frete": {"media": lt_frete_media, "n": lt_frete_n},
        "lt_producao": {"media": lt_producao_media, "n": lt_producao_n},
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


def _pedidos_pcp_por_semana(rotulos):
    """Agrupa ItemPedido por (semana do Planejamento Semanal PCP da Listagem
    Geral, pedido) pros rótulos dados — base compartilhada da prévia semanal
    do Painel (`_preview_semanal_pcp_painel`) e do resumo do mês seguinte
    (`_resumo_mes_seguinte_pcp`). Devolve dict rotulo -> {pedido_id: {...}}.

    "Data prevista PCP" e "Data efetiva de liberação" usam a mais recente
    entre os itens do pedido naquela semana, mesmo critério já usado em
    _liberacao_pcp_por_pedido_venda. `finalizado` = TODOS os itens do pedido
    naquela semana já têm "Liberação real" preenchida — mesmo critério
    "verde bandeira" já corrigido na Listagem Geral/relatório PDF (pedido
    835/Morken, 09/2026); antes esta função usava status_producao ==
    FINALIZADO, que pode divergir da data real preenchida (auditoria do
    Painel, 16/09/2026)."""
    por_semana = {rotulo: {} for rotulo in rotulos}
    if not rotulos:
        return por_semana

    itens = (
        ItemPedido.query.options(selectinload(ItemPedido.pedido))
        .join(Pedido, ItemPedido.pedido_id == Pedido.id)
        .filter(ItemPedido.planejamento_semanal.in_(rotulos))
        .all()
    )

    for item in itens:
        if not item.pedido or item.planejamento_semanal not in por_semana:
            continue
        grupo = por_semana[item.planejamento_semanal].setdefault(
            item.pedido_id,
            {
                "pedido_id": item.pedido_id,
                "pedido_venda": item.pedido.pedido_venda,
                "cliente": item.pedido.cliente,
                "valor": 0.0,
                "data_solicitada_cliente": item.pedido.data_cliente,
                "data_prevista_pcp": None,
                "data_efetiva_liberacao": None,
                "frete": item.pedido.frete,
                "estado": item.pedido.estado,
                "finalizado": True,
            },
        )
        grupo["valor"] += item.valor_total
        if not item.liberacao_real:
            grupo["finalizado"] = False
        if item.liberacao_prevista and (
            grupo["data_prevista_pcp"] is None or item.liberacao_prevista > grupo["data_prevista_pcp"]
        ):
            grupo["data_prevista_pcp"] = item.liberacao_prevista
        if item.liberacao_real and (
            grupo["data_efetiva_liberacao"] is None or item.liberacao_real > grupo["data_efetiva_liberacao"]
        ):
            grupo["data_efetiva_liberacao"] = item.liberacao_real

    return por_semana


def _preview_semanal_pcp_painel(hoje=None):
    """Prévia horizontal do Planejamento Semanal PCP pro Painel (pedido do
    Bruno, 09/09/2026): "de bate pronto", como gestor, ele quer ver resumido
    o que o PCP tem planejado nas semanas próximas — pedido, valor, data
    solicitada pelo cliente, data prevista PCP, data efetiva de liberação,
    frete e estado — os mesmos dados que já existem espalhados em Gestão
    Produção/Consulta Pedido, só que resumidos aqui na tela principal, com
    link direto pro contexto completo do pedido (Detalhe do Pedido).

    Uma coluna por semana, um card por PEDIDO dentro de cada semana — soma
    dos itens daquele pedido planejados pra aquela semana especificamente
    (um pedido com itens em semanas diferentes aparece em cada uma delas,
    com a soma da semana em questão).

    Ajuste do Bruno (16/09/2026): antes as colunas eram uma janela ROLANTE
    de poucas semanas (1 atrás + atual + 3 à frente), e a Semana 01 do mês
    "saía" da tela conforme o mês avançava. Agora mostra SEMPRE todas as
    semanas do mês atual (`gerar_semanas_pcp` com meses_atras=meses_frente=0
    centrado no mês de `hoje`) — o resumo do mês seguinte (Outubro) tem seu
    próprio quadrante à parte (`_resumo_mes_seguinte_pcp`).

    Ajuste do Bruno (09/09/2026, depois de ver a 1ª versão ao vivo): cards
    ordenados por "Data solicitada cliente" (prazo de entrega mais curto
    primeiro) em vez de valor — sem data fica por último."""
    hoje = hoje or date.today()
    rotulos = gerar_semanas_pcp(meses_atras=0, meses_frente=0, hoje=hoje)
    if not rotulos:
        return {"semanas": []}

    semana_atual_num = -(-hoje.day // 7)
    chave_atual = (hoje.year, hoje.month, semana_atual_num)

    por_semana = _pedidos_pcp_por_semana(rotulos)

    semanas = []
    for rotulo in rotulos:
        pedidos = sorted(
            por_semana[rotulo].values(),
            key=lambda p: (p["data_solicitada_cliente"] is None, p["data_solicitada_cliente"]),
        )
        for p in pedidos:
            p["valor"] = round(p["valor"], 2)
        semanas.append(
            {
                "rotulo": rotulo,
                "rotulo_curto": rotulo.replace("SEMANA ", "S").replace(" / ", "/"),
                "atual": _chave_semana_pcp(rotulo) == chave_atual,
                "pedidos": pedidos,
                "total": round(sum(p["valor"] for p in pedidos), 2),
            }
        )
    return {"semanas": semanas}


def _resumo_mes_pcp(ano, mes):
    """Resumo consolidado do planejamento PCP (Listagem Geral) de UM mês —
    base compartilhada dos cards "Faturamento previsto" (mês atual) e
    "Backlog mês seguinte" + o quadrante de resumo (pedido do Bruno,
    16/09/2026). Mesma base de dados de `_preview_semanal_pcp_painel`
    (`_pedidos_pcp_por_semana`), só que consolidando todas as semanas do mês
    num resumo único em vez de uma coluna por semana — inclui o Top 5
    pedidos por valor (pedido do Bruno, 16/09/2026: "top 5 pedidos melhores
    faturamento, com nome, pv e valor pedido").

    `valor_total` alimenta tanto o card KPI quanto o quadrante de resumo —
    mesmo número nos dois lugares, pra nunca divergir. `mes_str` (formato
    "AAAA-MM") serve pra linkar direto pra Listagem Geral já filtrada nesse
    mês (`dashboard(planejamento_mensal=...)`)."""
    rotulos = gerar_semanas_pcp(meses_atras=0, meses_frente=0, hoje=date(ano, mes, 1))
    por_semana = _pedidos_pcp_por_semana(rotulos)

    pedidos = {}
    for grupo in por_semana.values():
        for pid, dados in grupo.items():
            alvo = pedidos.setdefault(pid, {**dados, "valor": 0.0, "finalizado": True})
            alvo["valor"] += dados["valor"]
            alvo["finalizado"] = alvo["finalizado"] and dados["finalizado"]

    lista = list(pedidos.values())
    for p in lista:
        p["valor"] = round(p["valor"], 2)
    finalizados = [p for p in lista if p["finalizado"]]
    em_aberto = [p for p in lista if not p["finalizado"]]
    top5 = sorted(lista, key=lambda p: p["valor"], reverse=True)[:5]

    return {
        "mes_label": f"{MESES_PT[mes - 1]}/{ano}",
        "mes_str": f"{ano:04d}-{mes:02d}",
        "pedidos_total": len(lista),
        "valor_total": round(sum(p["valor"] for p in lista), 2),
        "pedidos_finalizados": len(finalizados),
        "valor_finalizado": round(sum(p["valor"] for p in finalizados), 2),
        "pedidos_em_aberto": len(em_aberto),
        "valor_em_aberto": round(sum(p["valor"] for p in em_aberto), 2),
        "top5": top5,
    }


def _resumo_mes_seguinte_pcp(hoje=None):
    """Wrapper de `_resumo_mes_pcp` pro mês SEGUINTE ao atual — "Backlog mês
    seguinte" (pedido do Bruno, 16/09/2026: "um resumo em um único
    quadrante de todo o planejamento do mês seguinte (Outubro)")."""
    hoje = hoje or date.today()
    ano_seg, mes_seg = _somar_meses(hoje.year, hoje.month, 1)
    return _resumo_mes_pcp(ano_seg, mes_seg)


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
    """Devolve as semanas (domingo a sábado) do mês/ano dado, cada uma como
    {"numero", "inicio", "fim"} — regra de calendário impresso comum: uma
    semana pertence ao mês que tem a MAIORIA dos seus 7 dias (equivalente a
    olhar em que mês cai a quarta-feira daquela semana — o dia central de
    domingo a sábado; como são 7 dias, nunca empata 3x4). Cada semana
    aparece em UM mês só, nunca repetida em dois meses vizinhos.

    Pedido do Bruno (09/09/2026): a 1ª versão desta função (só domingo a
    sábado "cobrindo" o mês) fazia a semana de fronteira aparecer igual nos
    dois meses vizinhos — Bruno testou e não é isso que ele quer. Confirmado
    com 2 exemplos reais que ele mandou: Setembro/2026 tem que começar em
    30/08 (a semana de 26/07-01/08 é de Julho, não de Agosto nem Setembro,
    então nem aparece em Agosto) e Agosto/2026 tem só 4 semanas, 02/08 a
    29/08 (a semana de 30/08-05/09 já é de Setembro). Essa regra de maioria
    bate exatamente com os dois exemplos ao mesmo tempo."""
    primeiro_dia = date(ano, mes, 1)
    dias_desde_domingo = (primeiro_dia.weekday() + 1) % 7  # weekday(): segunda=0 ... domingo=6
    domingo = primeiro_dia - timedelta(days=dias_desde_domingo)

    def _quarta_e_do_mes(domingo_semana):
        quarta = domingo_semana + timedelta(days=3)
        return quarta.year == ano and quarta.month == mes

    if not _quarta_e_do_mes(domingo):
        domingo += timedelta(days=7)

    semanas = []
    numero = 1
    while _quarta_e_do_mes(domingo):
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


_PERIODO_TIPOS = ("mes", "tri", "sem", "ano")


def _parse_periodo(valor_str):
    """Decodifica o parâmetro `periodo` da tela Resultados/OTD — pedido do
    Bruno (03/09/2026): "inclua em formato de lista... além do mês, inclua
    também 1º trimestre, 2, 3, 4 trimestre, 1º e segundo semestre... e o
    filtro geral (incluindo todos os períodos/meses)". Formatos aceitos:
    "todos", "mes-AAAA-M", "tri-AAAA-T" (T 1-4), "sem-AAAA-S" (S 1-2),
    "ano-AAAA". Sem valor reconhecível, cai no mês atual (mesmo default de
    sempre). Devolve (tipo, ano, valor, label) — `valor` é None pra
    tipo "ano" e pra "todos"."""
    hoje = date.today()
    padrao_mes_atual = ("mes", hoje.year, hoje.month, f"{MESES_PT[hoje.month - 1]}/{hoje.year}")

    if valor_str == "todos":
        return "todos", None, None, "Todos os períodos"

    m = re.match(r"^(mes|tri|sem|ano)-(\d{4})(?:-(\d+))?$", valor_str or "")
    if not m:
        return padrao_mes_atual
    tipo, ano_s, valor_s = m.groups()
    ano = int(ano_s)
    valor = int(valor_s) if valor_s else None

    if tipo == "mes" and valor and 1 <= valor <= 12:
        return "mes", ano, valor, f"{MESES_PT[valor - 1]}/{ano}"
    if tipo == "tri" and valor and 1 <= valor <= 4:
        return "tri", ano, valor, f"{valor}º Trimestre/{ano}"
    if tipo == "sem" and valor and 1 <= valor <= 2:
        return "sem", ano, valor, f"{valor}º Semestre/{ano}"
    if tipo == "ano":
        return "ano", ano, None, f"Ano {ano}"
    return padrao_mes_atual


def _periodo_para_str(tipo, ano, valor):
    """Inverso de _parse_periodo (sem o label) — monta a string `periodo`
    a partir de (tipo, ano, valor), pros links/campos ocultos do
    formulário."""
    if tipo == "todos":
        return "todos"
    if tipo == "ano":
        return f"ano-{ano}"
    return f"{tipo}-{ano}-{valor}"


def _periodo_vizinho(tipo, ano, valor, direcao):
    """String `periodo` do período vizinho (direcao -1 ou +1), na MESMA
    granularidade — usado nos botões "Período anterior"/"Próximo período"
    da tela Resultados/OTD, que agora navegam mês a mês, trimestre a
    trimestre, semestre a semestre ou ano a ano, dependendo do que está
    selecionado. "todos" não tem vizinho (botões ficam escondidos)."""
    if tipo == "mes":
        m, a = valor + direcao, ano
        if m < 1:
            m, a = 12, ano - 1
        elif m > 12:
            m, a = 1, ano + 1
        return _periodo_para_str("mes", a, m)
    if tipo == "tri":
        t, a = valor + direcao, ano
        if t < 1:
            t, a = 4, ano - 1
        elif t > 4:
            t, a = 1, ano + 1
        return _periodo_para_str("tri", a, t)
    if tipo == "sem":
        s, a = valor + direcao, ano
        if s < 1:
            s, a = 2, ano - 1
        elif s > 2:
            s, a = 1, ano + 1
        return _periodo_para_str("sem", a, s)
    if tipo == "ano":
        return _periodo_para_str("ano", ano + direcao, None)
    return None


def _opcoes_periodo():
    """Lista de opções pro dropdown de período da tela Resultados/OTD,
    agrupada por ano (Ano atual - 1 até Ano atual + 1) — cada grupo com Ano
    inteiro, os 2 semestres, os 4 trimestres e os 12 meses, mais "Todos os
    períodos" no topo, fora de qualquer grupo."""
    hoje = date.today()
    grupos = []
    for ano in range(hoje.year - 1, hoje.year + 2):
        itens = [{"valor": _periodo_para_str("ano", ano, None), "label": f"Ano {ano}"}]
        itens += [
            {"valor": _periodo_para_str("sem", ano, s), "label": f"{s}º Semestre {ano}"} for s in (1, 2)
        ]
        itens += [
            {"valor": _periodo_para_str("tri", ano, t), "label": f"{t}º Trimestre {ano}"} for t in (1, 2, 3, 4)
        ]
        itens += [
            {"valor": _periodo_para_str("mes", ano, m), "label": f"{MESES_PT[m - 1]}/{ano}"} for m in range(1, 13)
        ]
        grupos.append({"ano": ano, "itens": itens})
    return grupos


def _semanas_do_periodo(tipo, ano, valor):
    """Lista de rótulos de semana PCP (formato de gerar_semanas_pcp) cobertos
    por um período mes/tri/sem/ano — não trata "todos" (ver
    _filtrar_por_periodo_pcp, que pra "todos" não filtra nada)."""
    if tipo == "mes":
        meses = [valor]
    elif tipo == "tri":
        meses = list(range((valor - 1) * 3 + 1, (valor - 1) * 3 + 4))
    elif tipo == "sem":
        meses = list(range(1, 7)) if valor == 1 else list(range(7, 13))
    else:  # "ano"
        meses = list(range(1, 13))
    semanas = []
    for m in meses:
        semanas.extend(gerar_semanas_pcp(meses_atras=0, meses_frente=0, hoje=date(ano, m, 1)))
    return semanas


def _filtrar_por_periodo_pcp(query, tipo, ano, valor):
    """Aplica (ou não) o filtro de Término Semanal PCP num query de
    PedidoOperacao, de acordo com o período — "todos" não filtra nada
    (literalmente todos os pedidos, tenham ou não Término Semanal PCP
    definido; pedido do Bruno, 03/09/2026: "o filtro geral, incluindo
    todos os períodos/meses")."""
    if tipo == "todos":
        return query
    semanas = _semanas_do_periodo(tipo, ano, valor)
    return query.filter(PedidoOperacao.go_termino_semanal_pcp.in_(semanas))


def _pedidos_operacao_do_periodo(tipo, ano, valor):
    """Query de PedidoOperacao filtrada pelo período (ver
    _filtrar_por_periodo_pcp), direta (sem passar pelos filtros de
    cliente/vendedor/busca da tela). Usada pelo resumo fixo do topo da tela
    Resultados/OTD (pedido do Bruno, 03/09/2026) — sempre mostra o período
    selecionado, independente do dropdown de segmento mais abaixo na
    página."""
    return _filtrar_por_periodo_pcp(PedidoOperacao.query, tipo, ano, valor)


def _faturamento_por_periodo(tipo, ano, valor):
    """Generaliza o antigo "Faturamento por Semana" (só um mês) pra
    qualquer período — mês, trimestre, semestre, ano ou "todos" — pedido do
    Bruno em 28/08/2026 (base) e 03/09/2026 (dropdown de período). Uma
    linha por semana de PCP dentro do período; pra "todos", usa toda semana
    que realmente tem algum pedido (em vez de gerar um calendário sem fim).

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
    junto com a entrega.

    Cada linha também carrega a lista de `pedidos` daquela semana (pedido do
    Bruno, 03/09/2026: clicar na semana e abrir os pedidos dela, sem sair da
    página) — já vem pronta daqui (sem N+1) porque os pedidos da semana já
    são carregados de qualquer forma pra somar qtd/valor."""
    if tipo == "todos":
        pedidos_todos = (
            PedidoOperacao.query.filter(PedidoOperacao.go_termino_semanal_pcp.isnot(None))
            .order_by(PedidoOperacao.pedido_venda)
            .all()
        )
        semanas = sorted({p.go_termino_semanal_pcp for p in pedidos_todos}, key=_chave_semana_pcp)
        pedidos = pedidos_todos
    else:
        semanas = _semanas_do_periodo(tipo, ano, valor)
        pedidos = (
            PedidoOperacao.query.filter(PedidoOperacao.go_termino_semanal_pcp.in_(semanas))
            .order_by(PedidoOperacao.pedido_venda)
            .all()
        )

    baldes = {
        s: {"qtd_liberada": 0, "valor_liberado": 0.0, "qtd_faturada": 0, "valor_faturado": 0.0, "pedidos": []}
        for s in semanas
    }
    for p in pedidos:
        b = baldes[p.go_termino_semanal_pcp]
        b["qtd_liberada"] += 1
        b["valor_liberado"] += p.go_valor_pedido_operacao or 0.0
        b["pedidos"].append(p)
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
            "pedidos": baldes[s]["pedidos"],
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


def _otd_mensal_ano(ano):
    """OTD (Gestão Operação, go_otd_realizado) mês a mês, Jan-Dez de um ano —
    pedido do Bruno (03/09/2026): gráfico de OTD mensal no Painel, mesma
    meta mínima da tela Resultados/OTD (GO_OTD_META_PERCENTUAL, 78%). Usa o
    mesmo recorte por Término Semanal PCP já usado em Resultados/OTD
    (_pedidos_operacao_do_periodo/_resumo_otd), só resumido mês a mês pro
    ano inteiro."""
    resultado = []
    for mes in range(1, 13):
        query = _pedidos_operacao_do_periodo("mes", ano, mes)
        otd = _resumo_otd(query)
        resultado.append({
            "mes": MESES_PT[mes - 1],
            "percentual": otd["percentual"],
            "total": otd["total"],
            "atinge_meta": otd["atinge_meta"],
        })
    return resultado


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


def _sugestao_risco_prazo(pedido, data_prevista_pcp, mapa_lead_time):
    """Sugestão de ação pro card do mini risco (pedido do Bruno, 16/09/2026):
    "melhor caminho pensando no frete" pra manter o OTD positivo. Pra CIF,
    soma o lead time de transporte cadastrado (MESMO cadastro único de
    Cadastros > Lead time Transportadora já usado na Torre de Controle de
    OTD/Gestão de Risco — nunca duplicado, só consultado). Pra FOB o frete é
    por conta do cliente, então não existe trecho de transporte sob
    responsabilidade da 4PIPE pra somar — a sugestão vira só a própria data
    solicitada pelo cliente."""
    prazo_cliente = pedido.data_cliente
    if not prazo_cliente:
        return "Sem data solicitada pelo cliente cadastrada — não dá pra calcular um prazo seguro."

    if pedido.frete == "FOB":
        return f"Frete por conta do cliente (FOB) — libere até {prazo_cliente.strftime('%d/%m/%Y')} pra cumprir o prazo solicitado."

    uf = (pedido.estado or "").strip().upper()
    linha = mapa_lead_time.get((uf, "Rodoviário"))
    if not linha:
        return "Sem lead time de transporte cadastrado pra esse Estado (Cadastros > Lead time Transportadora)."

    dias_transporte = _lead_time_transporte_dias(linha)
    data_limite = prazo_cliente - timedelta(days=dias_transporte)
    if data_prevista_pcp and data_prevista_pcp <= data_limite:
        return (
            f"Dentro da margem: previsto PCP em {data_prevista_pcp.strftime('%d/%m/%Y')}, prazo limite pra "
            f"liberar é {data_limite.strftime('%d/%m/%Y')} (frete de {dias_transporte}d até {uf})."
        )
    return (
        f"Melhor caminho: libere até {data_limite.strftime('%d/%m/%Y')} — somado o frete rodoviário "
        f"({dias_transporte}d até {uf}), a entrega fica dentro do prazo solicitado ({prazo_cliente.strftime('%d/%m/%Y')})."
    )


def _linha_risco_prazo(pedido, mapa_lead_time):
    """Monta 1 card do mini risco (pedido do Bruno, 16/09/2026) com o
    contexto que ele pediu: data solicitada cliente, estado, data prevista
    PCP (a mais urgente entre os itens do pedido ainda não finalizados — a
    mesma que está causando o atraso/vencimento) e a sugestão de ação."""
    itens_abertos = [
        i for i in pedido.itens if i.status_producao != "FINALIZADO" and i.liberacao_prevista
    ]
    data_prevista_pcp = min((i.liberacao_prevista for i in itens_abertos), default=None)
    return {
        "id": pedido.id,
        "pedido_venda": pedido.pedido_venda,
        "cliente": pedido.cliente,
        "frete": pedido.frete,
        "estado": pedido.estado,
        "data_solicitada_cliente": pedido.data_cliente,
        "data_prevista_pcp": data_prevista_pcp,
        "sugestao": _sugestao_risco_prazo(pedido, data_prevista_pcp, mapa_lead_time),
    }


def _mini_risco_prazos_painel(limite=5):
    """Mini gestão de risco de prazos pro Painel, separada por frete FOB x
    CIF (pedido do Bruno, 16/09/2026) — reaproveita a MESMA regra já usada e
    confiável dos cards Atrasados/Vencendo (_predicado_atrasado /
    _predicado_vencendo), só filtrada por Pedido.frete e organizada em 2
    colunas. Recalculado ao vivo a cada carregamento da tela — nada fica
    gravado, então já sai "atualizando diariamente" sem precisar de job
    nenhum (mesma filosofia da Gestão de Risco/Torre de Controle de OTD já
    existente em Gestão Operação, só que aqui, "mini", puxando Gestão
    Produção/Listagem Geral).

    Cada card também traz data solicitada cliente, estado, data prevista PCP
    e uma sugestão de ação (_linha_risco_prazo/_sugestao_risco_prazo) —
    ajuste do Bruno (16/09/2026) depois de ver a 1ª versão ao vivo."""
    mapa_lead_time = _mapa_lead_time_transportadora()
    resultado = {}
    for frete in ("FOB", "CIF"):
        base = Pedido.query.options(selectinload(Pedido.itens)).filter(Pedido.frete == frete)
        atrasados = (
            base.filter(_predicado_atrasado())
            .order_by(Pedido.data_inclusao_pedido.desc().nullslast())
            .all()
        )
        vencendo = (
            base.filter(_predicado_vencendo())
            .order_by(Pedido.data_inclusao_pedido.desc().nullslast())
            .all()
        )
        resultado[frete] = {
            "atrasados_total": len(atrasados),
            "vencendo_total": len(vencendo),
            "atrasados": [_linha_risco_prazo(p, mapa_lead_time) for p in atrasados[:limite]],
            "vencendo": [_linha_risco_prazo(p, mapa_lead_time) for p in vencendo[:limite]],
        }
    return resultado


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


def _faturamento_mensal_ano(ano):
    """Faturamento REALIZADO mês a mês, Jan-Dez de um ano — pedido do Bruno
    (03/09/2026): trocar o gráfico "Faturamento previsto × realizado (6
    meses)" do Painel por uma visão anual, só do realizado (o previsto
    continua disponível no card "Faturamento previsto (mês)" acima e na tela
    de Faturamento). Reaproveita _faturamento_mes (mesma fonte/critério de
    sempre: liberação de faturamento dentro do mês, valor faturado real
    quando preenchido, senão quantidade × custo)."""
    resultado = []
    for mes in range(1, 13):
        _, realizado = _faturamento_mes(ano, mes)
        resultado.append({"mes": MESES_PT[mes - 1], "realizado": realizado})
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


def _entrega_cliente_por_pedido_venda(pedidos_venda):
    """dict pedido_venda (trim, sem FK) -> data de entrega no cliente,
    olhando pra Gestão Operação (PedidoOperacao.go_data_entregue_cliente,
    com go_data_real_entrega como fallback) — mesmo casamento por
    pedido_venda já usado em _liberacao_pcp_por_pedido_venda/
    _data_cliente_por_pedido_venda. Só leitura, nunca grava nada. Usado
    exclusivamente pelo "Lead time total" do detalhamento de Lead Time do
    Painel (pedido do Bruno, 16/09/2026) — Gestão Produção não tem campo
    próprio de entrega no cliente, só Gestão Operação tem."""
    valores = sorted({_normalizar_pedido_venda(v) for v in pedidos_venda if v and v.strip()})
    if not valores:
        return {}
    pedidos_go = (
        PedidoOperacao.query
        .filter(_pedido_venda_normalizado_sql(PedidoOperacao.pedido_venda).in_(valores))
        .filter(
            or_(
                PedidoOperacao.go_data_entregue_cliente.isnot(None),
                PedidoOperacao.go_data_real_entrega.isnot(None),
            )
        )
        .all()
    )
    mapa = {}
    for p in pedidos_go:
        chave = _normalizar_pedido_venda(p.pedido_venda)
        if not chave:
            continue
        data = p.go_data_entregue_cliente or p.go_data_real_entrega
        anterior = mapa.get(chave)
        if data and (anterior is None or data > anterior):
            mapa[chave] = data
    return mapa


def _lead_time_detalhado_painel():
    """Detalhamento de Lead Time do Painel (pedido do Bruno, 16/09/2026),
    baseado em Gestão Produção/Listagem Geral — todas as médias, cada uma
    com o "n" (quantidade que entrou na conta) igual ao resto do sistema
    (_media_dias):

      - chao_fabrica: Inclusão do pedido -> Liberação efetiva, por item.
      - fila_espera: Inclusão do pedido -> Início de produção, mesmo campo
        já existente ItemPedido.tempo_espera_dias.
      - prazo_comercial: Inclusão do pedido -> Data solicitada pelo
        cliente, por pedido — quanto prazo a empresa se comprometeu, na
        média (não é o realizado, é o prometido).
      - total: Inclusão do pedido -> Entrega no cliente. Gestão Produção
        não tem esse campo próprio — cruza com Gestão Operação por
        pedido_venda (_entrega_cliente_por_pedido_venda), só leitura."""
    pedidos = Pedido.query.filter(Pedido.data_inclusao_pedido.isnot(None)).all()

    chao_fabrica_valores = []
    fila_espera_valores = []
    for pedido in pedidos:
        for item in pedido.itens:
            if item.liberacao_real:
                chao_fabrica_valores.append((item.liberacao_real - pedido.data_inclusao_pedido).days)
            if item.tempo_espera_dias is not None:
                fila_espera_valores.append(item.tempo_espera_dias)

    prazo_comercial_valores = [
        (p.data_cliente - p.data_inclusao_pedido).days for p in pedidos if p.data_cliente
    ]

    pedidos_venda = [p.pedido_venda for p in pedidos if p.pedido_venda]
    entrega_por_pedido = _entrega_cliente_por_pedido_venda(pedidos_venda)
    total_valores = []
    for p in pedidos:
        data_entrega = entrega_por_pedido.get(_normalizar_pedido_venda(p.pedido_venda))
        if data_entrega:
            total_valores.append((data_entrega - p.data_inclusao_pedido).days)

    chao_fabrica_media, chao_fabrica_n = _media_dias(chao_fabrica_valores)
    fila_espera_media, fila_espera_n = _media_dias(fila_espera_valores)
    prazo_comercial_media, prazo_comercial_n = _media_dias(prazo_comercial_valores)
    total_media, total_n = _media_dias(total_valores)

    return {
        "chao_fabrica": {"media": chao_fabrica_media, "n": chao_fabrica_n},
        "fila_espera": {"media": fila_espera_media, "n": fila_espera_n},
        "prazo_comercial": {"media": prazo_comercial_media, "n": prazo_comercial_n},
        "total": {"media": total_media, "n": total_n},
    }


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
    """Quantidade de itens não finalizados por estação, com uma amostra dos
    próprios pedidos parados ali — usado no gráfico "Backlog por estação" do
    Painel (única tela que chama esta função). Pedido do Bruno (03/09/2026):

      1) não contar "OUTROS"/"PROJETO ESPECIAL" nesse gráfico específico —
         são "estações" genéricas demais aqui; a tela /estacoes continua
         mostrando as duas normalmente (estacoes_lista() tem sua própria
         consulta, independente desta função, então não é afetada);
      2) ao passar o mouse, já ver quais pedidos estão parados naquela
         estação, com link direto pra abrir o Kanban dela — por isso cada
         linha já sai com uma amostra de itens (pedido/cliente/produto) e a
         própria URL da estação, prontas pro tooltip/clique do gráfico (ver
         painel.html)."""
    ESTACOES_FORA_DO_GRAFICO_PAINEL = {"OUTROS", "PROJETO ESPECIAL"}
    itens = (
        ItemPedido.query.options(selectinload(ItemPedido.pedido))
        .filter(
            ItemPedido.status_producao != "FINALIZADO",
            ItemPedido.estacao.isnot(None),
            ~ItemPedido.estacao.in_(ESTACOES_FORA_DO_GRAFICO_PAINEL),
        )
        .all()
    )

    agrupado = {}
    for item in itens:
        agrupado.setdefault(item.estacao, []).append(item)

    resultado = []
    for estacao, lista in agrupado.items():
        lista.sort(key=lambda i: i.pedido.data_inclusao_pedido if (i.pedido and i.pedido.data_inclusao_pedido) else date.min)
        amostra = [
            {
                "pedido_venda": i.pedido.pedido_venda if i.pedido else None,
                "cliente": i.pedido.cliente if i.pedido else "—",
                "produto": i.descricao_produto,
            }
            for i in lista[:8]
        ]
        resultado.append({
            "estacao": estacao,
            "quantidade": len(lista),
            "itens_amostra": amostra,
            "restante": max(0, len(lista) - len(amostra)),
            "url": url_for("estacao_kanban", nome=estacao),
        })
    resultado.sort(key=lambda r: -r["quantidade"])
    return resultado


def _tempo_relativo(quando):
    """"agora mesmo" / "há Xmin" / "há Xh" / "há X dias" a partir de um
    datetime (UTC, mesmo padrão de todo `criado_em` no app) — usado só no
    feed de apontamentos de Qualidade do Painel, pra dar uma ideia rápida de
    "quando" sem cravar hora exata. Itens com mais de 30 dias caem pra data
    cheia (dd/mm/aaaa), já que "há 47 dias" deixa de ser uma leitura útil."""
    if not quando:
        return "—"
    segundos = (datetime.utcnow() - quando).total_seconds()
    if segundos < 60:
        return "agora mesmo"
    if segundos < 3600:
        return f"há {int(segundos // 60)}min"
    if segundos < 86400:
        return f"há {int(segundos // 3600)}h"
    dias = int(segundos // 86400)
    if dias == 1:
        return "há 1 dia"
    if dias < 30:
        return f"há {dias} dias"
    return quando.strftime("%d/%m/%Y")


def _fmt_num_rdim(v):
    """Formata um float de medição RDIM sem casas decimais penduradas — 5.0
    vira "5", 5.25 vira "5,25" (vírgula, padrão BR) — usado só no texto do
    feed de apontamentos do Painel, pra não sair "5.0mm" feio na notificação."""
    if v is None:
        return None
    texto = f"{v:.3f}".rstrip("0").rstrip(".")
    if texto in ("", "-"):
        texto = "0"
    return texto.replace(".", ",")


def _contexto_desvio_rdim(i):
    """Texto curto com o DETALHE do desvio (não só o resultado) — pedido do
    Bruno (03/09/2026): "cite com detalhes o desvio, ex: desvio dimensional,
    desvio espessura, mínimo e máximo era X, e ficou com Y". Monta a partir
    de categoria_desvio/subcategoria_desvio + o pior apontamento peça a peça
    (pecas_desvio) quando existir; se a inspeção não tiver detalhamento
    peça a peça, cai pra pior medição fora de tolerância (RdimMedicao).
    None quando a inspeção foi aprovada sem nenhum desvio."""
    if i.resultado == "APROVADO":
        return None

    partes = []
    if i.categoria_desvio:
        rotulo_cat = f"Desvio {i.categoria_desvio}"
        if i.subcategoria_desvio:
            rotulo_cat += f" · {i.subcategoria_desvio}"
        partes.append(rotulo_cat)
    elif i.subcategoria_desvio:
        partes.append(f"Desvio {i.subcategoria_desvio}")

    pior_peca = i.pior_apontamento_peca
    if pior_peca and pior_peca.valor_medido is not None:
        if pior_peca.especificado_min is not None and pior_peca.especificado_max is not None:
            espec_txt = f"mín. {_fmt_num_rdim(pior_peca.especificado_min)} e máx. {_fmt_num_rdim(pior_peca.especificado_max)}"
        elif pior_peca.especificado_max is not None:
            espec_txt = f"máx. {_fmt_num_rdim(pior_peca.especificado_max)}"
        elif pior_peca.especificado_min is not None:
            espec_txt = f"mín. {_fmt_num_rdim(pior_peca.especificado_min)}"
        else:
            espec_txt = None
        if espec_txt:
            partes.append(f"especificação era {espec_txt} e ficou com {_fmt_num_rdim(pior_peca.valor_medido)}")
        else:
            partes.append(f"medido {_fmt_num_rdim(pior_peca.valor_medido)}")
    else:
        pior_medicao = None
        for m in i.medicoes:
            if m.dentro_da_tolerancia is False:
                pior_medicao = m
                break
        if pior_medicao:
            medido = pior_medicao.medido_max if pior_medicao.medido_max is not None else pior_medicao.medido_min
            if pior_medicao.especificado_min is not None and pior_medicao.especificado_max is not None:
                espec_txt = f"mín. {_fmt_num_rdim(pior_medicao.especificado_min)} e máx. {_fmt_num_rdim(pior_medicao.especificado_max)}"
            elif pior_medicao.especificado_max is not None:
                espec_txt = f"máx. {_fmt_num_rdim(pior_medicao.especificado_max)}"
            elif pior_medicao.especificado_min is not None:
                espec_txt = f"mín. {_fmt_num_rdim(pior_medicao.especificado_min)}"
            else:
                espec_txt = None
            if espec_txt and medido is not None:
                partes.append(f"{pior_medicao.grandeza}: especificação era {espec_txt} e ficou com {_fmt_num_rdim(medido)}")
            elif medido is not None:
                partes.append(f"{pior_medicao.grandeza}: medido {_fmt_num_rdim(medido)}")

    if not partes and i.desvio_encontrado:
        texto = i.desvio_encontrado.strip()
        partes.append(texto[:140] + ("…" if len(texto) > 140 else ""))

    if i.quantidade_com_desvio and i.item and i.item.quantidade:
        partes.append(f"{_fmt_num_rdim(i.quantidade_com_desvio)} de {_fmt_num_rdim(i.item.quantidade)} peças")

    return " — ".join(partes) if partes else None


def _contexto_rnc(r):
    """Texto curto de contexto pro apontamento de RNC no feed do Painel —
    mesmo espírito do detalhe de desvio do RDIM acima: dar uma ideia real do
    que aconteceu sem precisar abrir o RNC. Usa a descrição da não
    conformidade (texto livre já preenchido no formulário); cai pro
    requisito não atendido quando a descrição está vazia."""
    texto = (r.descricao_nc or "").strip()
    if texto:
        return texto[:140] + ("…" if len(texto) > 140 else "")
    if r.requisito_nao_atendido:
        return f"Requisito não atendido: {r.requisito_nao_atendido}"
    return None


def _apontamentos_recentes_qualidade(desde=None, limite=15):
    """Últimos apontamentos de Qualidade (RDIM + RNC), mais recentes primeiro
    — pedido do Bruno (03/09/2026): "todo apontamento diário da qualidade...
    tenha um campo de aviso ou notificação dentro do painel", pra ele, como
    gestor, ter ciência breve de todo apontamento sem precisar entrar na
    área de Qualidade todo dia — cada linha já sai com link direto pro
    registro (RDIM -> rdim_editar, RNC -> qualidade_editar).

    `desde` (opcional): datetime da última vez que o Painel foi aberto (ver
    sessão 'ultima_visita_painel' na rota `painel()`) — marca quais entram
    como "novo desde a última visita", sem precisar de tabela/coluna nova de
    controle de leitura (guardado na sessão do próprio navegador)."""
    rdims = (
        InspecaoFinal.query
        .options(
            selectinload(InspecaoFinal.item).selectinload(ItemPedido.pedido),
            selectinload(InspecaoFinal.medicoes),
            selectinload(InspecaoFinal.pecas_desvio),
        )
        .order_by(InspecaoFinal.criado_em.desc())
        .limit(limite)
        .all()
    )
    rncs = RncQualidade.query.order_by(RncQualidade.criado_em.desc()).limit(limite).all()

    eventos = []
    for i in rdims:
        eventos.append({
            "tipo": "RDIM",
            "criado_em": i.criado_em,
            "titulo": f"{i.pedido_venda or 'Pedido —'} · {i.cliente or 'Sem cliente'}",
            "detalhe": RDIM_RESULTADO_LABELS.get(i.resultado, i.resultado or "Sem resultado"),
            "contexto": _contexto_desvio_rdim(i),
            "cor": RDIM_RESULTADO_CORES.get(i.resultado, "secondary"),
            "link": url_for("rdim_editar", inspecao_id=i.id),
        })
    for r in rncs:
        eventos.append({
            "tipo": "RNC",
            "criado_em": r.criado_em,
            "titulo": f"{r.numero_rnc or ('RNC #' + str(r.id))} · {r.cliente_projeto or 'Sem cliente/projeto'}",
            "detalhe": r.tipo_nc or r.status_geral or "Sem tipo",
            "contexto": _contexto_rnc(r),
            "cor": RNC_SEVERIDADE_CORES.get(r.severidade, "secondary"),
            "link": url_for("qualidade_editar", rnc_id=r.id),
        })

    eventos.sort(key=lambda e: e["criado_em"] or datetime.min, reverse=True)
    eventos = eventos[:limite]
    for e in eventos:
        e["ha_quanto_tempo"] = _tempo_relativo(e["criado_em"])
        e["novo"] = bool(desde and e["criado_em"] and e["criado_em"] > desde)
    return eventos


def _pedidos_recentes_producao(desde=None, limite=10):
    """Últimos pedidos incluídos em Gestão Produção (+Novo pedido), mais
    recentes primeiro — pedido do Bruno (03/09/2026), no mesmo formato do
    feed de apontamentos de Qualidade acima: "todo pedido incluído... com
    cliente, número, data de inclusão, data solicitada pelo cliente e valor,
    passar o mouse pra ver mais informações". Ordenado por `criado_em`
    (timestamp real de quando o registro foi gravado) — `data_inclusao_pedido`
    é só um campo de texto/data digitado no formulário, pode não bater com a
    ordem real de cadastro. O tooltip (atributo title, mesmo padrão já usado
    em qualidade_rdim_lista.html) cobre o que não cabe na linha visível:
    produtos, quantidade, vendedor, cidade/UF, prioridade."""
    pedidos = (
        Pedido.query
        .options(selectinload(Pedido.itens))
        .order_by(Pedido.criado_em.desc())
        .limit(limite)
        .all()
    )

    eventos = []
    for p in pedidos:
        tooltip_linhas = [
            f"{p.quantidade_total:g} peça(s) · {p.descricao_resumo}",
            f"Vendedor: {p.vendedor or '—'}",
            f"Cidade/UF: {p.cidade or '—'}/{p.estado or '—'}",
            f"Prioridade: {p.prioridade or '—'}",
        ]
        eventos.append({
            "criado_em": p.criado_em,
            "titulo": f"{p.pedido_venda or ('Pedido #' + str(p.id))} · {p.cliente}",
            "data_inclusao": p.data_inclusao_pedido,
            "data_cliente": p.data_cliente,
            "valor": p.valor_total,
            "tooltip": "\n".join(tooltip_linhas),
            "link": url_for("detalhe_pedido", pedido_id=p.id),
        })

    eventos.sort(key=lambda e: e["criado_em"] or datetime.min, reverse=True)
    eventos = eventos[:limite]
    for e in eventos:
        e["ha_quanto_tempo"] = _tempo_relativo(e["criado_em"])
        e["novo"] = bool(desde and e["criado_em"] and e["criado_em"] > desde)
    return eventos


def _janela_utc_do_dia_brt(dia_brt):
    """Início/fim (UTC, intervalo [início, fim)) do dia `dia_brt` no fuso de
    Brasília (UTC-3, sem horário de verão) — todo `criado_em`/`atualizado_em`
    do banco é gravado em UTC, então filtrar "o dia de ontem" (no fuso do
    Bruno) precisa dessa conversão. Usado só pelo Relatório Diário
    automático abaixo (pedido do Bruno, 09/09/2026)."""
    inicio_utc = datetime(dia_brt.year, dia_brt.month, dia_brt.day) + timedelta(hours=3)
    return inicio_utc, inicio_utc + timedelta(days=1)


def _relatorio_diario_dados(dia_brt=None):
    """Relatório Diário automático (pedido do Bruno, 09/09/2026): "todo dia
    às 8h, sem eu precisar colocar a mão no site" — desvios de Qualidade
    (RDIM com desvio + RNC) e pedidos incluídos/finalizados de UM dia
    específico (por padrão, ontem no fuso de Brasília). Alimenta a rota
    /api/relatorio-diario, consumida por uma tarefa agendada no Claude que
    manda a notificação pro Bruno — não tem link em nenhuma tela do site.

    "Finalizado no dia" = item com status_producao FINALIZADO cujo
    atualizado_em (carimbo automático de última alteração) caiu dentro do
    dia — não existe campo próprio de "quando finalizou", então esse é o
    melhor proxy disponível (mesmo espírito do Kanban de Estações, que já
    trata status_producao como o campo "confiável" de progresso)."""
    dia_brt = dia_brt or (date.today() - timedelta(days=1))
    inicio_utc, fim_utc = _janela_utc_do_dia_brt(dia_brt)

    rdims = (
        InspecaoFinal.query
        .options(selectinload(InspecaoFinal.pecas_desvio), selectinload(InspecaoFinal.medicoes))
        .filter(InspecaoFinal.criado_em >= inicio_utc, InspecaoFinal.criado_em < fim_utc)
        .filter(InspecaoFinal.resultado != "APROVADO")
        .order_by(InspecaoFinal.criado_em)
        .all()
    )
    rncs = (
        RncQualidade.query
        .filter(RncQualidade.criado_em >= inicio_utc, RncQualidade.criado_em < fim_utc)
        .order_by(RncQualidade.criado_em)
        .all()
    )
    pedidos_incluidos = (
        Pedido.query
        .filter(Pedido.criado_em >= inicio_utc, Pedido.criado_em < fim_utc)
        .order_by(Pedido.criado_em)
        .all()
    )
    itens_finalizados = (
        ItemPedido.query
        .options(selectinload(ItemPedido.pedido))
        .join(Pedido, ItemPedido.pedido_id == Pedido.id)
        .filter(
            ItemPedido.status_producao == "FINALIZADO",
            ItemPedido.atualizado_em >= inicio_utc,
            ItemPedido.atualizado_em < fim_utc,
        )
        .order_by(ItemPedido.atualizado_em)
        .all()
    )
    pedidos_finalizados = {}
    for item in itens_finalizados:
        if not item.pedido:
            continue
        grupo = pedidos_finalizados.setdefault(
            item.pedido_id,
            {
                "pedido_id": item.pedido_id,
                "pedido_venda": item.pedido.pedido_venda,
                "cliente": item.pedido.cliente,
                "itens": [],
            },
        )
        grupo["itens"].append(item.descricao_produto)

    return {
        "dia": dia_brt.isoformat(),
        "desvios_rdim": [
            {
                "pedido_venda": i.pedido_venda,
                "cliente": i.cliente,
                "resultado": RDIM_RESULTADO_LABELS.get(i.resultado, i.resultado),
                "detalhe": _contexto_desvio_rdim(i),
                "link": url_for("rdim_editar", inspecao_id=i.id, _external=True),
            }
            for i in rdims
        ],
        "rncs": [
            {
                "numero_rnc": r.numero_rnc,
                "cliente_projeto": r.cliente_projeto,
                "tipo_nc": r.tipo_nc,
                "severidade": r.severidade,
                "detalhe": _contexto_rnc(r),
                "link": url_for("qualidade_editar", rnc_id=r.id, _external=True),
            }
            for r in rncs
        ],
        "pedidos_incluidos": [
            {
                "pedido_venda": p.pedido_venda,
                "cliente": p.cliente,
                "valor": p.valor_total,
                "vendedor": p.vendedor,
                "link": url_for("detalhe_pedido", pedido_id=p.id, _external=True),
            }
            for p in pedidos_incluidos
        ],
        "pedidos_finalizados": [
            {
                "pedido_venda": g["pedido_venda"],
                "cliente": g["cliente"],
                "itens": g["itens"],
                "link": url_for("detalhe_pedido", pedido_id=g["pedido_id"], _external=True),
            }
            for g in pedidos_finalizados.values()
        ],
    }


# Rótulos amigáveis pros campos de CAMPOS_HISTORICO_PD, usados só na
# descrição compacta do feed de atualizações de P&D no Painel (abaixo) — o
# histórico em si (tela de edição de cada projeto) já mostra o nome técnico
# do campo, este dicionário é só pra deixar o feed do Painel mais legível.
CAMPO_LABELS_PD = {
    "etapa_atual": "Etapa",
    "percentual_conclusao": "% concluído",
    "prioridade": "Prioridade",
    "responsavel": "Responsável",
    "data_prevista_conclusao": "Previsão de conclusão",
    "data_real_conclusao": "Conclusão real",
    "custo_realizado": "Custo realizado",
    "investimento_realizado": "Investimento realizado",
    "economia_realizada": "Economia realizada",
}


def _descricao_mudanca_pd(h):
    rotulo = CAMPO_LABELS_PD.get(h.campo, h.campo)
    antigo = h.valor_anterior or "—"
    novo = h.valor_novo or "—"
    return f"{rotulo}: {antigo} → {novo}"


def _atualizacoes_recentes_pd(desde=None, limite=15):
    """Últimas atualizações de P&D (projeto novo cadastrado + avanço de
    etapa/status ou qualquer outro campo acompanhado em CAMPOS_HISTORICO_PD)
    — pedido do Bruno (03/09/2026): "qualquer atualização de cadastro ou
    avanço de status/etapa do P&D também apareça pra mim", mesmo formato do
    feed de Qualidade. Reaproveita HistoricoAlteracao (já gravado em todo
    pd_editar/pd_mover_etapa) em vez de criar tabela nova — várias linhas de
    histórico gravadas na mesma edição (ex.: mudou etapa E prioridade de
    uma vez) são agrupadas num único evento no feed, pra não poluir a lista
    com uma linha por campo."""
    projetos_novos = ProjetoPD.query.order_by(ProjetoPD.criado_em.desc()).limit(limite).all()

    historico = (
        HistoricoAlteracao.query
        .filter_by(entidade_tipo="projeto_pd")
        .order_by(HistoricoAlteracao.criado_em.desc())
        .limit(limite * 6)
        .all()
    )

    eventos = []
    for p in projetos_novos:
        eventos.append({
            "tipo": "NOVO",
            "criado_em": p.criado_em,
            "titulo": f"{p.codigo or ('#' + str(p.id))} · {p.nome}",
            "detalhe": "Novo projeto cadastrado",
            "etapa": p.etapa_atual,
            "cor": PD_ETAPA_CORES.get(p.etapa_atual, "secondary"),
            "link": url_for("pd_editar", projeto_id=p.id),
        })

    if historico:
        ids_projetos = {h.entidade_id for h in historico}
        projetos_por_id = {p.id: p for p in ProjetoPD.query.filter(ProjetoPD.id.in_(ids_projetos)).all()}

        # Agrupa linhas de histórico da MESMA edição: mesmo projeto + mesmo
        # segundo (uma edição só grava tudo dentro de um único commit, então
        # os `criado_em` ficam praticamente idênticos — arredondar pro
        # segundo é o bastante pra juntar sem precisar de um "id de edição").
        grupos = {}
        for h in historico:
            chave = (h.entidade_id, h.criado_em.replace(microsecond=0) if h.criado_em else None)
            grupos.setdefault(chave, []).append(h)

        for (projeto_id, _), linhas in grupos.items():
            projeto = projetos_por_id.get(projeto_id)
            if projeto is None:
                continue
            linhas_ordenadas = sorted(linhas, key=lambda h: h.criado_em or datetime.min)
            partes = [_descricao_mudanca_pd(h) for h in linhas_ordenadas[:3]]
            if len(linhas_ordenadas) > 3:
                partes.append(f"+{len(linhas_ordenadas) - 3} campo(s)")
            eventos.append({
                "tipo": "ATUALIZAÇÃO",
                "criado_em": max((h.criado_em or datetime.min) for h in linhas_ordenadas),
                "titulo": f"{projeto.codigo or ('#' + str(projeto.id))} · {projeto.nome}",
                "detalhe": " · ".join(partes),
                "etapa": projeto.etapa_atual,
                "cor": PD_ETAPA_CORES.get(projeto.etapa_atual, "secondary"),
                "link": url_for("pd_editar", projeto_id=projeto.id),
            })

    eventos.sort(key=lambda e: e["criado_em"] or datetime.min, reverse=True)
    eventos = eventos[:limite]
    for e in eventos:
        e["ha_quanto_tempo"] = _tempo_relativo(e["criado_em"])
        e["novo"] = bool(desde and e["criado_em"] and e["criado_em"] > desde)
    return eventos


# ----------------------------------------------------------------------
# Tela de KPIs — reconstrução completa (pedido do Bruno, 17-18/09/2026):
# "zere todos que já existe" (os 4 widgets antigos: Pedidos finalizados por
# mês, Lead time médio + OTD por mês, Lead time médio por estação, OTD por
# vendedor — nenhum é usado por mais nenhuma tela, removidos junto) e
# substituir por 9 indicadores gerenciais novos. As funções abaixo cobrem os
# itens automáticos (1, 2, 3, 4-ranking e 9); os itens manuais (6, 7, 8 e as
# observações de 4/5) ficam em KpiGerencialMensal (models.py) e
# _kpi_gerencial_mensal, mais abaixo.
# ----------------------------------------------------------------------


def _lead_time_fila_por_estacao(desde=None):
    """Fila + processamento + lead time total, por estação — itens 1+2+3 da
    tela de KPIs numa tabela só, porque são a MESMA base de dados
    (ItemPedido.tempo_espera_dias = fila, ItemPedido.lt_producao_dias =
    processamento, soma dos dois = lead time total), só quebrados de jeitos
    diferentes: "Lead time das OPs (por estação, por período e no total)",
    "Tempo de fila das OPs (por estação, por período e no total)" e "Lead
    time por setor: tempo de fila e tempo de processamento por setor".

    Substitui a antiga _lead_time_por_estacao (só tinha o lead time total).
    Percorre o catálogo de Estacao (mesmo padrão de _gargalos_por_estacao),
    não os valores distintos de ItemPedido.estacao — garante que toda
    estação ativa apareça (mesmo sem dado no período) e já sai na ordem
    canônica (Estacao.ordem_exibicao). Filtra por `termino_inspecao >=
    desde` quando informado (mesmo seletor de período 3/6/12 meses que a
    tela já tinha). Devolve 1 linha por estação + 1 linha "TOTAL" no fim."""
    estacoes = Estacao.query.filter_by(ativo=True).order_by(Estacao.ordem_exibicao).all()
    resultado = []
    todos_fila, todos_proc, todos_total = [], [], []

    for e in estacoes:
        query = ItemPedido.query.filter(
            ItemPedido.estacao == e.nome,
            ItemPedido.inicio_producao.isnot(None),
            ItemPedido.termino_inspecao.isnot(None),
        )
        if desde:
            query = query.filter(ItemPedido.termino_inspecao >= desde)
        itens = query.all()

        fila_vals = [i.tempo_espera_dias for i in itens if i.tempo_espera_dias is not None]
        proc_vals = [i.lt_producao_dias for i in itens if i.lt_producao_dias is not None]
        total_vals = [
            i.tempo_espera_dias + i.lt_producao_dias
            for i in itens
            if i.tempo_espera_dias is not None and i.lt_producao_dias is not None
        ]
        todos_fila += fila_vals
        todos_proc += proc_vals
        todos_total += total_vals

        fila_media, _ = _media_dias(fila_vals)
        proc_media, _ = _media_dias(proc_vals)
        total_media, total_n = _media_dias(total_vals)
        resultado.append({
            "estacao": e.nome,
            "fila_media": fila_media,
            "processamento_medio": proc_media,
            "lead_time_medio": total_media,
            "quantidade": total_n,
        })

    fila_media, _ = _media_dias(todos_fila)
    proc_media, _ = _media_dias(todos_proc)
    total_media, total_n = _media_dias(todos_total)
    resultado.append({
        "estacao": "TOTAL (todas as estações)",
        "fila_media": fila_media,
        "processamento_medio": proc_media,
        "lead_time_medio": total_media,
        "quantidade": total_n,
    })
    return resultado


def _tendencia_fila_lead_time(meses=6):
    """Tendência mensal de fila/lead time (gráfico do Bloco A da tela de
    KPIs) + ranking de variação por estação entre os 2 últimos meses com
    dado — é o "ranking automático" do item 4 ("Análise das principais
    causas do aumento do tempo de fila e do lead time das OPs", pedido do
    Bruno, 17-18/09/2026): aponta EM QUAL estação a fila/lead time mais
    cresceu; o "porquê" fica no campo de observação manual
    (KpiGerencialMensal.obs_causas_fila_lead_time), preenchido à parte.

    Substitui a antiga _tendencia_kpis (que trazia finalizados/lt_medio/otd
    — otd saiu, não faz parte dos 9 itens pedidos)."""
    hoje = date.today()
    ano, mes = hoje.year, hoje.month
    pontos = []
    for _ in range(meses):
        pontos.append((ano, mes))
        mes -= 1
        if mes == 0:
            mes, ano = 12, ano - 1
    pontos.reverse()

    tendencia = []
    por_mes_estacao = {}
    for ano_p, mes_p in pontos:
        inicio = date(ano_p, mes_p, 1)
        fim = date(ano_p + 1, 1, 1) if mes_p == 12 else date(ano_p, mes_p + 1, 1)
        itens = ItemPedido.query.filter(
            ItemPedido.termino_inspecao.isnot(None),
            ItemPedido.termino_inspecao >= inicio,
            ItemPedido.termino_inspecao < fim,
        ).all()

        fila_vals, total_vals = [], []
        por_estacao = {}
        for i in itens:
            fila = i.tempo_espera_dias
            proc = i.lt_producao_dias
            total = fila + proc if (fila is not None and proc is not None) else None
            estacao = i.estacao or "Sem estação"
            grupo = por_estacao.setdefault(estacao, {"fila": [], "total": []})
            if fila is not None:
                fila_vals.append(fila)
                grupo["fila"].append(fila)
            if total is not None:
                total_vals.append(total)
                grupo["total"].append(total)

        fila_media, _ = _media_dias(fila_vals)
        total_media, _ = _media_dias(total_vals)
        tendencia.append({
            "mes": f"{MESES_PT[mes_p - 1]}/{ano_p}",
            "fila_media": fila_media,
            "lead_time_medio": total_media,
            "finalizados": len(itens),
        })
        por_mes_estacao[(ano_p, mes_p)] = por_estacao

    # Ranking de variação: os 2 últimos pontos que já têm QUALQUER dado
    # (não necessariamente os 2 últimos de `pontos` — o mês corrente costuma
    # estar incompleto/vazio ainda).
    meses_com_dado = [p for p in pontos if por_mes_estacao.get(p)]
    variacoes = []
    if len(meses_com_dado) >= 2:
        atual_key, anterior_key = meses_com_dado[-1], meses_com_dado[-2]
        atual, anterior = por_mes_estacao[atual_key], por_mes_estacao[anterior_key]
        for estacao in set(atual) | set(anterior):
            fila_atual, _ = _media_dias(atual.get(estacao, {}).get("fila", []))
            fila_anterior, _ = _media_dias(anterior.get(estacao, {}).get("fila", []))
            total_atual, _ = _media_dias(atual.get(estacao, {}).get("total", []))
            total_anterior, _ = _media_dias(anterior.get(estacao, {}).get("total", []))
            delta_fila = (
                round(fila_atual - fila_anterior, 1) if (fila_atual is not None and fila_anterior is not None) else None
            )
            delta_lead_time = (
                round(total_atual - total_anterior, 1) if (total_atual is not None and total_anterior is not None) else None
            )
            if delta_fila is None and delta_lead_time is None:
                continue
            variacoes.append({
                "estacao": estacao,
                "mes_atual": f"{MESES_PT[atual_key[1] - 1]}/{atual_key[0]}",
                "mes_anterior": f"{MESES_PT[anterior_key[1] - 1]}/{anterior_key[0]}",
                "fila_atual": fila_atual,
                "fila_anterior": fila_anterior,
                "delta_fila": delta_fila,
                "lead_time_atual": total_atual,
                "lead_time_anterior": total_anterior,
                "delta_lead_time": delta_lead_time,
            })
        variacoes.sort(key=lambda v: (v["delta_lead_time"] is None, -(v["delta_lead_time"] or 0)))

    return {"tendencia": tendencia, "variacoes": variacoes}


_RE_DIAMETRO_POL = re.compile(r"(\d+(?:[.,]\d+)?)\s*(?:''|\"|[’”]|'|POL(?:EGADAS)?\b)", re.I)
_RE_DIAMETRO_MM = re.compile(r"(\d+(?:[.,]\d+)?)\s*MM\b", re.I)


def _classificar_dn_mm(descricao):
    """Best-effort: tenta achar o DN/diâmetro na descrição do produto, pro
    item 9 da tela de KPIs (pedido do Bruno, 17-18/09/2026: "organizar por
    DN ou mm") — testado contra a base real: ~95% de cobertura. Tenta
    polegadas primeiro (padrão mais comum na base: 8'', 10", 12" etc.),
    depois mm; o que não bater cai em "Sem DN/mm identificado" (mostrado
    explicitamente na tela, não escondido — mesma transparência já usada
    noutros relatórios quando um dado não é 100% capturável)."""
    if not descricao:
        return "Sem DN/mm identificado"
    m = _RE_DIAMETRO_POL.search(descricao)
    if m:
        return f'{m.group(1)}"'
    m = _RE_DIAMETRO_MM.search(descricao)
    if m:
        return f"{m.group(1)}MM"
    return "Sem DN/mm identificado"


def _categoria_produto(descricao):
    """PIG × Sobressalente (item 9 da tela de KPIs, pedido do Bruno,
    17-18/09/2026) — critério "começa com PIG", não "contém PIG": itens tipo
    "DISCO GUIA PARA PIG...", "PLACA CALIBRADORA... PARA PIG..." contêm a
    palavra PIG mas são acessórios/sobressalentes, não o PIG em si
    (confirmado contra a base real antes de implementar)."""
    if descricao and descricao.strip().upper().startswith("PIG"):
        return "PIG"
    return "Sobressalente"


def _produtividade_por_setor(ano, mes):
    """Item 9 da tela de KPIs (pedido do Bruno, 17-18/09/2026): "Produtividade
    por setor, considerando o volume de PIGs e sobressalentes produzidos...
    controle mensal de todos os produtos produzidos, organizar por DN ou
    mm" — AUTOMÁTICO (confirmado por ele), a partir dos pedidos já
    lançados: soma ItemPedido.quantidade dos itens cujo término de inspeção
    caiu no mês pedido (mesmo critério de "produção concluída" do resto da
    tela), agrupado por estação + categoria (PIG/Sobressalente) + DN/mm."""
    inicio = date(ano, mes, 1)
    fim = date(ano + 1, 1, 1) if mes == 12 else date(ano, mes + 1, 1)
    itens = ItemPedido.query.filter(
        ItemPedido.termino_inspecao.isnot(None),
        ItemPedido.termino_inspecao >= inicio,
        ItemPedido.termino_inspecao < fim,
    ).all()

    por_estacao = {}
    for item in itens:
        estacao = item.estacao or "Sem estação"
        categoria = _categoria_produto(item.descricao_produto)
        dn = _classificar_dn_mm(item.descricao_produto)
        grupo = por_estacao.setdefault(
            estacao, {"PIG": {}, "Sobressalente": {}, "total_pig": 0.0, "total_sobressalente": 0.0}
        )
        grupo[categoria][dn] = grupo[categoria].get(dn, 0.0) + (item.quantidade or 0)
        chave_total = "total_pig" if categoria == "PIG" else "total_sobressalente"
        grupo[chave_total] += item.quantidade or 0

    resultado = []
    for estacao, dados in por_estacao.items():
        resultado.append({
            "estacao": estacao,
            "total_pig": round(dados["total_pig"], 2),
            "total_sobressalente": round(dados["total_sobressalente"], 2),
            "pig_por_dn": sorted(
                ({"dn": dn, "quantidade": round(q, 2)} for dn, q in dados["PIG"].items()),
                key=lambda r: -r["quantidade"],
            ),
            "sobressalente_por_dn": sorted(
                ({"dn": dn, "quantidade": round(q, 2)} for dn, q in dados["Sobressalente"].items()),
                key=lambda r: -r["quantidade"],
            ),
        })
    resultado.sort(key=lambda r: -(r["total_pig"] + r["total_sobressalente"]))
    return resultado


def _kpi_gerencial_mensal(ano, mes):
    """Busca a linha de KpiGerencialMensal do mês pedido (itens 6/7/8 +
    observações manuais de 4/5) — devolve um objeto "vazio" em memória (sem
    salvar no banco) quando o mês ainda não foi preenchido nenhuma vez, pra
    tela e formulário sempre terem algo pra mostrar/editar."""
    linha = KpiGerencialMensal.query.filter_by(ano=ano, mes=mes).first()
    if linha is None:
        linha = KpiGerencialMensal(ano=ano, mes=mes)
    return linha


def _gerar_pdf_kpis(meses, ano, mes, lead_time_fila_estacao, tendencia, variacoes, gargalos, produtividade, kpi_mensal):
    """Relatório PDF da tela de KPIs (pedido do Bruno, 18/09/2026: "incluir
    para gerar relatório em PDF... totalmente intuitivo e didático") —
    paisagem A4, reaproveita EXATAMENTE os mesmos dados já calculados pra
    tela (mesmos parâmetros que a rota kpis() já usa: `desde`/`meses` pro
    Bloco A, `ano`/`mes` pro Bloco D), então o PDF nunca diverge do que
    aparece no navegador. Usa reportlab (mesma lib dos outros 3 PDFs do
    sistema — _gerar_pdf_risco_otd/_gerar_pdf_estacao/
    _gerar_pdf_planejamento_mensal_pcp — pura Python, sem dependência de
    pacote de sistema, importante porque o deploy no Render não dá controle
    sobre isso).

    "Didático": nada de emoji (fontes padrão do reportlab não têm os glyphs
    coloridos — vira quadrado preto, mesmo problema já resolvido antes no
    relatório de Estações); em vez disso, cor de fundo de célula pra
    piora/melhora (mesmo truque do relatório de Risco OTD) e um gráfico de
    linha NATIVO do reportlab (reportlab.graphics, sem dependência nova)
    pra tendência de fila/lead time — o mesmo gráfico que já aparece na
    tela, só que desenhado com as primitivas do próprio reportlab."""
    from reportlab.graphics.charts.legends import Legend
    from reportlab.graphics.charts.lineplots import LinePlot
    from reportlab.graphics.shapes import Drawing
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    COR_PIORA = colors.HexColor("#f8d7da")
    COR_MELHORA = colors.HexColor("#d1e7dd")
    COR_CABECALHO_BG = colors.HexColor("#d3e0f2")
    COR_CABECALHO_TEXTO = colors.HexColor("#1b2a4a")
    COR_FILA = colors.HexColor("#f59f00")
    COR_LEAD_TIME = colors.HexColor("#6ea8fe")

    def _moeda(valor):
        return "R$ " + "{:,.2f}".format(valor or 0).replace(",", "X").replace(".", ",").replace("X", ".")

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=landscape(A4),
        leftMargin=10 * mm, rightMargin=10 * mm, topMargin=12 * mm, bottomMargin=12 * mm,
        title="Relatório de KPIs — Gestão da Produção",
    )
    estilos = getSampleStyleSheet()
    estilo_celula = ParagraphStyle("celula", parent=estilos["Normal"], fontSize=8.5, leading=10.5)
    estilo_celula_bold = ParagraphStyle("celula_bold", parent=estilo_celula, fontName="Helvetica-Bold")
    estilo_cabecalho_tabela = ParagraphStyle(
        "cabecalho_tabela", parent=estilo_celula_bold, fontSize=9, leading=11, textColor=COR_CABECALHO_TEXTO,
    )
    estilo_secao = ParagraphStyle("secao", parent=estilos["Heading2"], spaceBefore=4, spaceAfter=6)
    estilo_obs = ParagraphStyle("obs", parent=estilos["Normal"], fontSize=9, leading=12, textColor=colors.HexColor("#495057"))

    largura_pagina = landscape(A4)[0] - doc.leftMargin - doc.rightMargin

    def _kpi_box(valor, rotulo):
        return [
            Paragraph(str(valor), ParagraphStyle("kpi_valor", parent=estilos["Normal"], fontSize=17, fontName="Helvetica-Bold", alignment=1)),
            Paragraph(rotulo, ParagraphStyle("kpi_rotulo", parent=estilos["Normal"], fontSize=8.5, alignment=1)),
        ]

    def _tabela_kpis(kpis):
        largura = largura_pagina / len(kpis)
        t = Table([[k[0] for k in kpis], [k[1] for k in kpis]], colWidths=[largura] * len(kpis))
        t.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#dee2e6")),
            ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#dee2e6")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        return t

    mes_label = f"{MESES_PT[mes - 1]}/{ano}"
    elementos = [
        Paragraph("Relatório de KPIs — Gestão da Produção", estilos["Title"]),
        Paragraph(
            f'Gerado em {_agora_brt().strftime("%d/%m/%Y %H:%M")} · '
            f'Bloco de lead time/fila: últimos {meses} meses · Controle mensal: {mes_label}',
            estilos["Normal"],
        ),
        Spacer(1, 5 * mm),
    ]

    # ---------------- Bloco A — Lead time e fila das OPs ----------------
    elementos.append(Paragraph("Lead time e fila das OPs", estilo_secao))
    total_linha = lead_time_fila_estacao[-1] if lead_time_fila_estacao else None
    kpis_a = [
        _kpi_box((f'{total_linha["lead_time_medio"]}d' if total_linha and total_linha["lead_time_medio"] is not None else "—"), "Lead time médio (total)"),
        _kpi_box((f'{total_linha["fila_media"]}d' if total_linha and total_linha["fila_media"] is not None else "—"), "Tempo de fila médio (total)"),
        _kpi_box((f'{total_linha["processamento_medio"]}d' if total_linha and total_linha["processamento_medio"] is not None else "—"), "Tempo de processamento (total)"),
        _kpi_box(sum(t["finalizados"] for t in tendencia), f"OPs concluídas (últimos {meses} meses)"),
    ]
    elementos.append(_tabela_kpis(kpis_a))
    elementos.append(Spacer(1, 5 * mm))

    pontos_grafico = [t for t in tendencia if t["fila_media"] is not None and t["lead_time_medio"] is not None]
    if len(pontos_grafico) >= 2:
        d = Drawing(largura_pagina, 62 * mm)
        lp = LinePlot()
        lp.x, lp.y = 15 * mm, 10 * mm
        lp.width, lp.height = largura_pagina - 25 * mm, 42 * mm
        lp.data = [
            list(enumerate(p["fila_media"] for p in pontos_grafico)),
            list(enumerate(p["lead_time_medio"] for p in pontos_grafico)),
        ]
        lp.lines[0].strokeColor, lp.lines[0].strokeWidth = COR_FILA, 2
        lp.lines[1].strokeColor, lp.lines[1].strokeWidth = COR_LEAD_TIME, 2
        lp.xValueAxis.valueMin, lp.xValueAxis.valueMax = 0, len(pontos_grafico) - 1
        lp.xValueAxis.valueSteps = list(range(len(pontos_grafico)))
        rotulos_x = [p["mes"] for p in pontos_grafico]
        lp.xValueAxis.labelTextFormat = lambda x, rotulos=rotulos_x: rotulos[int(x)]
        lp.xValueAxis.labels.fontSize = 7
        lp.yValueAxis.valueMin = 0
        lp.yValueAxis.labels.fontSize = 7
        d.add(lp)
        legenda = Legend()
        legenda.x, legenda.y = 15 * mm, 58 * mm
        legenda.colorNamePairs = [(COR_FILA, "Tempo de fila médio (d)"), (COR_LEAD_TIME, "Lead time médio (d)")]
        legenda.fontSize, legenda.alignment = 8, "left"
        d.add(legenda)
        elementos.append(Paragraph("Tendência mensal — fila e lead time", estilos["Heading3"]))
        elementos.append(d)
    else:
        elementos.append(Paragraph("Ainda não há meses suficientes com dado pra desenhar o gráfico de tendência.", estilo_obs))
    elementos.append(Spacer(1, 4 * mm))

    cabecalho_lt = ["Estação", "Fila", "Processamento", "Total", "OPs"]
    dados_lt = [[Paragraph(c, estilo_cabecalho_tabela) for c in cabecalho_lt]]
    for l in lead_time_fila_estacao:
        negrito = l["estacao"].startswith("TOTAL")
        estilo = estilo_celula_bold if negrito else estilo_celula
        dados_lt.append([
            Paragraph(l["estacao"], estilo),
            Paragraph(f'{l["fila_media"]}d' if l["fila_media"] is not None else "—", estilo),
            Paragraph(f'{l["processamento_medio"]}d' if l["processamento_medio"] is not None else "—", estilo),
            Paragraph(f'{l["lead_time_medio"]}d' if l["lead_time_medio"] is not None else "—", estilo),
            Paragraph(str(l["quantidade"]), estilo),
        ])
    pesos_lt = [30, 15, 18, 15, 12]
    larguras_lt = [p / sum(pesos_lt) * largura_pagina for p in pesos_lt]
    tabela_lt = Table(dados_lt, colWidths=larguras_lt, repeatRows=1)
    tabela_lt.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), COR_CABECALHO_BG),
        ("LINEBELOW", (0, 0), (-1, 0), 1, colors.HexColor("#8fa8cc")),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#ced4da")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#f1f3f5")),
    ]))
    elementos.append(tabela_lt)
    elementos.append(Spacer(1, 7 * mm))

    # ---------------- Bloco B — Causas do aumento ----------------
    elementos.append(Paragraph("Causas do aumento do tempo de fila e do lead time", estilo_secao))
    if variacoes:
        elementos.append(Paragraph(
            f'Comparação entre {variacoes[0]["mes_anterior"]} e {variacoes[0]["mes_atual"]} '
            "(últimos 2 meses com dado fechado) — maior piora primeiro.", estilos["Normal"],
        ))
        elementos.append(Spacer(1, 3 * mm))
        cab_v = ["Estação", "Fila (d)", "Variação fila", "Lead time (d)", "Variação lead time"]
        dados_v = [[Paragraph(c, estilo_cabecalho_tabela) for c in cab_v]]
        cores_v = [None]
        for v in variacoes:
            fila_txt = f'{v["fila_anterior"]}d → {v["fila_atual"]}d' if v["fila_anterior"] is not None and v["fila_atual"] is not None else "—"
            lt_txt = f'{v["lead_time_anterior"]}d → {v["lead_time_atual"]}d' if v["lead_time_anterior"] is not None and v["lead_time_atual"] is not None else "—"
            delta_fila_txt = f'{"+" if v["delta_fila"] and v["delta_fila"] > 0 else ""}{v["delta_fila"]}d' if v["delta_fila"] is not None else "—"
            delta_lt_txt = f'{"+" if v["delta_lead_time"] and v["delta_lead_time"] > 0 else ""}{v["delta_lead_time"]}d' if v["delta_lead_time"] is not None else "—"
            dados_v.append([
                Paragraph(v["estacao"], estilo_celula),
                Paragraph(fila_txt, estilo_celula),
                Paragraph(delta_fila_txt, estilo_celula_bold),
                Paragraph(lt_txt, estilo_celula),
                Paragraph(delta_lt_txt, estilo_celula_bold),
            ])
            if v["delta_lead_time"] is None:
                cores_v.append(None)
            elif v["delta_lead_time"] > 0:
                cores_v.append(COR_PIORA)
            elif v["delta_lead_time"] < 0:
                cores_v.append(COR_MELHORA)
            else:
                cores_v.append(None)
        tabela_v = Table(dados_v, colWidths=[p / 100 * largura_pagina for p in (22, 20, 18, 20, 20)], repeatRows=1)
        estilo_v = [
            ("BACKGROUND", (0, 0), (-1, 0), COR_CABECALHO_BG),
            ("LINEBELOW", (0, 0), (-1, 0), 1, colors.HexColor("#8fa8cc")),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#ced4da")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 3.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
        ]
        for i, cor in enumerate(cores_v):
            if cor:
                estilo_v.append(("BACKGROUND", (0, i), (-1, i), cor))
        tabela_v.setStyle(TableStyle(estilo_v))
        elementos.append(tabela_v)
    else:
        elementos.append(Paragraph("Ainda não há 2 meses fechados com dado suficiente pra comparar.", estilo_obs))
    elementos.append(Spacer(1, 4 * mm))
    elementos.append(Paragraph("Observação (causas do aumento):", estilo_celula_bold))
    elementos.append(Paragraph(kpi_mensal.obs_causas_fila_lead_time or "Nenhuma observação registrada.", estilo_obs))
    elementos.append(Spacer(1, 6 * mm))

    # ---------------- Bloco C — Gargalos ----------------
    elementos.append(Paragraph("Gargalos do processo produtivo — setores críticos", estilo_secao))
    cab_g = ["Estação", "Fila", "Atrasados", "Tempo de espera médio", "Lead time médio", "Valor parado"]
    dados_g = [[Paragraph(c, estilo_cabecalho_tabela) for c in cab_g]]
    for l in gargalos:
        dados_g.append([
            Paragraph(l["estacao"], estilo_celula_bold),
            Paragraph(str(l["fila"]), estilo_celula),
            Paragraph(str(l["atraso"]) if l["atraso"] else "—", estilo_celula),
            Paragraph(f'{l["tempo_espera_medio"]}d' if l["tempo_espera_medio"] is not None else "—", estilo_celula),
            Paragraph(f'{l["lt_medio"]}d' if l["lt_medio"] is not None else "—", estilo_celula),
            Paragraph(_moeda(l["valor_parado"]), estilo_celula),
        ])
    tabela_g = Table(dados_g, colWidths=[p / 100 * largura_pagina for p in (24, 12, 14, 18, 16, 16)], repeatRows=1)
    tabela_g.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), COR_CABECALHO_BG),
        ("LINEBELOW", (0, 0), (-1, 0), 1, colors.HexColor("#8fa8cc")),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#ced4da")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
    ]))
    elementos.append(tabela_g)
    elementos.append(Spacer(1, 4 * mm))
    elementos.append(Paragraph("Observação (causas e plano de ação):", estilo_celula_bold))
    elementos.append(Paragraph(kpi_mensal.obs_gargalos_plano_acao or "Nenhuma observação registrada.", estilo_obs))
    elementos.append(PageBreak())

    # ---------------- Bloco D — Controle mensal ----------------
    elementos.append(Paragraph(f"Controle mensal — {mes_label}", estilo_secao))

    aderencia_txt = "—"
    if kpi_mensal.aderencia_planejado:
        pct = 100 * (kpi_mensal.aderencia_realizado or 0) / kpi_mensal.aderencia_planejado
        aderencia_txt = f"{pct:.1f}%"
    kpis_d = [
        _kpi_box(kpi_mensal.aderencia_planejado if kpi_mensal.aderencia_planejado is not None else "—", "Aderência — Planejado"),
        _kpi_box(kpi_mensal.aderencia_realizado if kpi_mensal.aderencia_realizado is not None else "—", "Aderência — Realizado"),
        _kpi_box(aderencia_txt, "Aderência ao planejamento"),
        _kpi_box(kpi_mensal.consumo_materia_prima if kpi_mensal.consumo_materia_prima is not None else "—", "Consumo total de matéria-prima"),
        _kpi_box(kpi_mensal.indice_perdas if kpi_mensal.indice_perdas is not None else "—", "Perdas"),
        _kpi_box(kpi_mensal.indice_refugos if kpi_mensal.indice_refugos is not None else "—", "Refugos"),
        _kpi_box(kpi_mensal.indice_descartes if kpi_mensal.indice_descartes is not None else "—", "Descartes"),
    ]
    elementos.append(_tabela_kpis(kpis_d))
    elementos.append(Spacer(1, 6 * mm))

    elementos.append(Paragraph(f"Produtividade por setor — {mes_label}", estilos["Heading3"]))
    elementos.append(Paragraph(
        "Volume de PIGs e sobressalentes com produção concluída no mês, por estação, agrupado por DN/mm "
        "identificado na descrição do produto (best-effort — itens sem diâmetro identificável entram em "
        '"Sem DN/mm identificado").', estilo_obs,
    ))
    elementos.append(Spacer(1, 3 * mm))
    cab_p = ["Estação", "Total PIGs", "PIGs por DN/mm", "Total sobressalentes", "Sobressalentes por DN/mm"]
    dados_p = [[Paragraph(c, estilo_cabecalho_tabela) for c in cab_p]]
    for p in produtividade:
        pig_dn = ", ".join(f'{d["dn"]}: {d["quantidade"]}' for d in p["pig_por_dn"]) or "—"
        sobra_dn = ", ".join(f'{d["dn"]}: {d["quantidade"]}' for d in p["sobressalente_por_dn"]) or "—"
        dados_p.append([
            Paragraph(p["estacao"], estilo_celula_bold),
            Paragraph(str(p["total_pig"]), estilo_celula),
            Paragraph(pig_dn, estilo_celula),
            Paragraph(str(p["total_sobressalente"]), estilo_celula),
            Paragraph(sobra_dn, estilo_celula),
        ])
    if not produtividade:
        dados_p.append([Paragraph("Nenhuma produção concluída no mês.", estilo_celula), "", "", "", ""])
    tabela_p = Table(dados_p, colWidths=[p / 100 * largura_pagina for p in (16, 12, 30, 12, 30)], repeatRows=1)
    tabela_p.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), COR_CABECALHO_BG),
        ("LINEBELOW", (0, 0), (-1, 0), 1, colors.HexColor("#8fa8cc")),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#ced4da")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
    ]))
    elementos.append(tabela_p)

    doc.build(elementos)
    buffer.seek(0)
    resposta = Response(buffer.getvalue(), mimetype="application/pdf")
    nome_arquivo = f"kpis_{ano}-{mes:02d}.pdf"
    resposta.headers["Content-Disposition"] = f"attachment; filename={nome_arquivo}"
    return resposta


def _lead_times_estacao(nome, meses_historico=3):
    """Lead time médio de 1 estação, quebrado em 3 etapas (pedido do Bruno,
    11/09/2026): "fila" (inclusão do pedido -> início da OP), "chão de
    fábrica" (início -> conclusão da OP) e "total" (inclusão -> conclusão).
    Olha TODO item que já começou nessa estação (não só os ainda abertos),
    porque lead time é uma métrica histórica de desempenho, não de fila
    atual — mesma fonte de dados de _gargalos_por_estacao. Devolve também o
    histórico mensal (últimos N meses, pelo mês de CONCLUSÃO da OP), mesmo
    padrão de _tendencia_kpis, com os 3 lead times médios (fila/chão de
    fábrica/total) de cada mês — pra ver se a estação está melhorando ou
    piorando com o tempo — e a produção mais longa/mais curta (lead time
    total) de cada mês, pedido do Bruno (11/09/2026)."""
    itens = (
        ItemPedido.query.options(selectinload(ItemPedido.pedido))
        .filter(ItemPedido.estacao == nome, ItemPedido.inicio_producao.isnot(None))
        .all()
    )

    def _media(valores):
        return round(sum(valores) / len(valores), 1) if valores else None

    fila = _media([
        (i.inicio_producao - i.pedido.data_inclusao_pedido).days
        for i in itens if i.pedido and i.pedido.data_inclusao_pedido
    ])
    chao = _media([(i.termino_inspecao - i.inicio_producao).days for i in itens if i.termino_inspecao])
    total = _media([
        (i.termino_inspecao - i.pedido.data_inclusao_pedido).days
        for i in itens if i.termino_inspecao and i.pedido and i.pedido.data_inclusao_pedido
    ])

    # Pedido do Bruno (11/09/2026): o histórico começa no mês ANTERIOR, não
    # no atual — o mês corrente já aparece "ao vivo" nos 3 cartões de cima
    # (fila/chão/total), então repeti-lo aqui na lista de meses seria
    # redundante e ainda ficaria incompleto (mês em andamento).
    hoje = date.today()
    ano, mes = hoje.year, hoje.month
    mes -= 1
    if mes == 0:
        mes, ano = 12, ano - 1
    pontos = []
    for _ in range(meses_historico):
        pontos.append((ano, mes))
        mes -= 1
        if mes == 0:
            mes, ano = 12, ano - 1
    pontos.reverse()

    def _producao(i, dias):
        """Dados de 1 OP/produto pra identificar a produção mais longa/curta
        do mês — pedido do Bruno (11/09/2026), com cliente e pedido pra dar
        contexto de quem é (pedido dele, 11/09/2026 seguinte)."""
        return {
            "pedido_venda": i.pedido.pedido_venda if i.pedido else None,
            "cliente": i.pedido.cliente if i.pedido else None,
            "produto": i.descricao_produto,
            "dias": dias,
        }

    historico = []
    for ano, mes in pontos:
        inicio = date(ano, mes, 1)
        fim = date(ano + 1, 1, 1) if mes == 12 else date(ano, mes + 1, 1)
        itens_mes = [
            i for i in itens
            if i.termino_inspecao and inicio <= i.termino_inspecao < fim and i.pedido and i.pedido.data_inclusao_pedido
        ]
        valores_fila = [(i.inicio_producao - i.pedido.data_inclusao_pedido).days for i in itens_mes]
        valores_chao = [(i.termino_inspecao - i.inicio_producao).days for i in itens_mes]
        producoes = [_producao(i, (i.termino_inspecao - i.pedido.data_inclusao_pedido).days) for i in itens_mes]

        mais_longa = max(producoes, key=lambda p: p["dias"]) if producoes else None
        mais_curta = min(producoes, key=lambda p: p["dias"]) if producoes else None
        # Se só tem 1 produção no mês, "mais longa" e "mais curta" seriam a
        # mesma linha duas vezes — mostra só uma vez pra não confundir.
        if producoes and len(producoes) == 1:
            mais_curta = None

        historico.append({
            "mes": f"{MESES_PT[mes - 1]}/{ano}",
            "lead_fila_medio": _media(valores_fila),
            "lead_chao_medio": _media(valores_chao),
            "lead_total_medio": _media([p["dias"] for p in producoes]),
            "finalizados": len(itens_mes),
            "mais_longa": mais_longa,
            "mais_curta": mais_curta,
        })

    return {"fila": fila, "chao": chao, "total": total, "historico": historico}


# ----------------------------------------------------------------------
# Lead Time de Produção — parametrização por Produto + Estação (pedido do
# Bruno, 11/09/2026): base OFICIAL de "quanto tempo uma estação normalmente
# leva pra produzir um produto", cadastrada em Cadastros > Lead time
# Produção, e confrontada aqui com o histórico REAL das OPs finalizadas.
#
# 3 conceitos, pra não confundir (mesmo vocabulário usado no cadastro e na
# aba Simulação da Gestão de Risco):
#   LT PADRÃO    = parâmetro oficial, cadastrado manualmente
#                   (LeadTimeProducao.lt_padrao_dias).
#   LT HISTÓRICO = comportamento real das OPs finalizadas, calculado AO
#                   VIVO a cada carregamento (nunca guardado) — ver
#                   _estatisticas_lead_time_producao.
#   LT PROJETADO = o que usar pra prever prazo HOJE: por padrão é o próprio
#                   LT padrão, a não ser que o histórico recente esteja
#                   indicando desvio relevante (ver "indicador"/"tendencia"
#                   abaixo) — quem decide revisar é sempre um humano, o
#                   sistema só sugere (nunca altera o LT padrão sozinho).
#
# Preparado pra ser reaproveitado, sem alterações, por: a aba Simulação da
# Gestão de Risco (_simulacao_otd_linha, mais abaixo), o widget "Simulado A"
# na tela do pedido (rota editar_pedido) e, no futuro, um simulador "E SE" e
# a própria Torre de Controle OTD — nenhuma dessas funções depende de rota
# nenhuma, só de dados, então dá pra chamar de qualquer lugar novo depois.
# ----------------------------------------------------------------------

LT_PRODUCAO_LIMITE_ATENCAO = 1.10    # média de referência até 10% acima do padrão = 🟢
LT_PRODUCAO_LIMITE_VERMELHO = 1.25   # até 25% acima = 🟡; acima disso = 🔴
LT_PRODUCAO_DESVIO_SUGESTAO = 0.15   # desvio mínimo (15%) pra sugerir revisão do LT padrão
LT_PRODUCAO_OPS_MINIMAS_SUGESTAO = 3  # nº mínimo de OPs (últimos 3 meses) pra sugerir revisão


def _itens_finalizados_para_lt(produto, estacao_nome):
    """Todo ItemPedido FINALIZADO, com início e término de produção
    preenchidos, cuja descrição CONTENHA `produto` (ILIKE — mesmo padrão de
    busca já usado no resto do sistema, não exige catálogo nem digitação
    idêntica) e esteja na estação `estacao_nome`. Base de todas as
    estatísticas de lead time de produção — reaproveita
    ItemPedido.lt_producao_dias, que já existe."""
    if not produto or not estacao_nome:
        return []
    return (
        ItemPedido.query.filter(
            ItemPedido.estacao == estacao_nome,
            ItemPedido.status_producao == "FINALIZADO",
            ItemPedido.inicio_producao.isnot(None),
            ItemPedido.termino_inspecao.isnot(None),
            ItemPedido.descricao_produto.ilike(f"%{produto.strip()}%"),
        )
        .order_by(ItemPedido.termino_inspecao.asc())
        .all()
    )


def _estatisticas_lead_time_producao(produto, estacao_nome, lt_padrao_dias):
    """Estatísticas AO VIVO (nunca guardadas) de 1 combinação produto+estação
    — pedido do Bruno (11/09/2026, item 2): "não tratar o LT como número
    estático". Médias em janelas corridas (30d/3m/6m — aqui é "últimos N
    dias/meses", diferente do histórico por MÊS FECHADO que a tela de
    Estações usa), melhor/pior LT realizado, tendência e um indicador
    🟢🟡🔴 comparando a média recente contra o LT padrão — e, quando o
    desvio é relevante, uma sugestão de revisão (nunca aplicada sozinha)."""
    hoje = date.today()
    itens = _itens_finalizados_para_lt(produto, estacao_nome)
    lts_por_item = [(i, i.lt_producao_dias) for i in itens if i.lt_producao_dias is not None]

    def _media_desde(dias_atras):
        limite = hoje - timedelta(days=dias_atras)
        valores = [lt for i, lt in lts_por_item if i.termino_inspecao >= limite]
        return round(sum(valores) / len(valores), 1) if valores else None

    media_30d = _media_desde(30)
    media_3m = _media_desde(90)
    media_6m = _media_desde(180)
    qtd_3m = len([lt for i, lt in lts_por_item if i.termino_inspecao >= hoje - timedelta(days=90)])
    todos_lts = [lt for _, lt in lts_por_item]
    ultimo_lt = lts_por_item[-1][1] if lts_por_item else None

    # Tendência: compara os 30 dias mais recentes com a média dos últimos 3
    # meses — se os dias mais recentes estão puxando a média pra cima/baixo,
    # é sinal de tendência (não só ruído de 1 OP fora da curva).
    tendencia = "sem_dado"
    if media_30d is not None and media_3m is not None and media_3m > 0:
        variacao = media_30d / media_3m
        if variacao >= 1.1:
            tendencia = "aumento"
        elif variacao <= 0.9:
            tendencia = "queda"
        else:
            tendencia = "estavel"

    # Indicador visual (item 9 do pedido do Bruno) — compara a média mais
    # confiável disponível (3M, caindo pra 30d se ainda não tiver 3M de
    # histórico) contra o LT padrão cadastrado.
    referencia = media_3m if media_3m is not None else media_30d
    indicador = "sem_dado"
    if referencia is not None and lt_padrao_dias:
        razao = referencia / lt_padrao_dias
        if razao <= LT_PRODUCAO_LIMITE_ATENCAO:
            indicador = "verde"
        elif razao <= LT_PRODUCAO_LIMITE_VERMELHO:
            indicador = "amarelo"
        else:
            indicador = "vermelho"

    sugestao_revisao = None
    if (
        referencia is not None
        and lt_padrao_dias
        and qtd_3m >= LT_PRODUCAO_OPS_MINIMAS_SUGESTAO
        and abs(referencia - lt_padrao_dias) / lt_padrao_dias >= LT_PRODUCAO_DESVIO_SUGESTAO
    ):
        sugestao_revisao = {
            "valor_sugerido": round(referencia),
            "sentido": "aumento" if referencia > lt_padrao_dias else "reducao",
            "texto": (
                f"Média dos últimos {'3 meses' if media_3m is not None else '30 dias'} "
                f"({referencia}d) está {'acima' if referencia > lt_padrao_dias else 'abaixo'} "
                f"do LT padrão ({lt_padrao_dias}d) — considerar revisar para ~{round(referencia)}d."
            ),
        }

    return {
        "media_30d": media_30d,
        "media_3m": media_3m,
        "media_6m": media_6m,
        "melhor_lt": min(todos_lts) if todos_lts else None,
        "pior_lt": max(todos_lts) if todos_lts else None,
        "qtd_ops_consideradas": qtd_3m,
        "ultimo_lt": ultimo_lt,
        "tendencia": tendencia,
        "indicador": indicador,
        "sugestao_revisao": sugestao_revisao,
    }


def _mapa_lead_time_producao():
    """Todas as LeadTimeProducao ativas, com a Estacao já resolvida — lista
    pequena (cadastro gerencial, não um catálogo de produto), iterada em
    Python porque o casamento é por "contém" (não dá pra indexar por chave
    exata como o mapa de lead time de transporte)."""
    linhas = LeadTimeProducao.query.filter_by(ativo=True).all()
    estacoes = {e.id: e for e in Estacao.query.all()}
    return [(linha, estacoes.get(linha.estacao_id)) for linha in linhas]


def _lt_producao_parametrizado_item(item, mapa=None):
    """LT de produção parametrizado pra 1 ItemPedido: acha a
    LeadTimeProducao cujo `produto` está contido na descrição do item E cuja
    estação bate com item.estacao. Quando não acha nenhuma, cai pro fallback
    já existente e hoje adormecido Estacao.meta_lead_time_dias (evita
    duplicar o conceito de "lead time por estação" — reaproveita o campo já
    cadastrado em Cadastros > Estações). Retorna None quando nem isso
    existe (sem dado pra este item)."""
    if not item.estacao:
        return None
    mapa = mapa if mapa is not None else _mapa_lead_time_producao()
    descricao = (item.descricao_produto or "").upper()
    for linha, estacao in mapa:
        if estacao and estacao.nome == item.estacao and linha.produto.upper() in descricao:
            return linha.lt_padrao_dias
    estacao_cadastro = Estacao.query.filter_by(nome=item.estacao).first()
    if estacao_cadastro and estacao_cadastro.meta_lead_time_dias:
        return estacao_cadastro.meta_lead_time_dias
    return None


def _lt_producao_parametrizado_pedido(itens):
    """LT de produção parametrizado pra um PEDIDO inteiro: aplica
    _lt_producao_parametrizado_item em cada item aberto e usa o PIOR caso
    (máximo) — mesmo critério de "quem manda é o item mais lento" já usado
    em _liberacao_pcp_por_pedido_venda pro prazo real. Retorna
    (lt_dias_ou_none, itens_sem_parametro) pra quem exibe poder avisar
    quantos itens ficaram sem dado, sem esconder a lacuna."""
    if not itens:
        return None, []
    mapa = _mapa_lead_time_producao()
    valores = []
    sem_parametro = []
    for item in itens:
        lt = _lt_producao_parametrizado_item(item, mapa=mapa)
        if lt is None:
            sem_parametro.append(item)
        else:
            valores.append(lt)
    return (max(valores) if valores else None), sem_parametro


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
    quebrado por cliente / região / vendedor — usado na tela de Faturamento
    Projetado.

    "Previsto" (pedido do Bruno, 09/09/2026): passou a espelhar o
    Planejamento semanal/mensal (PCP) de cada item — campo
    `planejamento_semanal` (texto "SEMANA NN / MÊS / ANO", mesma convenção
    usada em Listagem Geral / editar pedido, ver GO_SEMANAS_PCP e
    _mes_ano_da_semana_pcp) — em vez da Liberação prevista (uma data). Bruno
    confirmou que o número que ele acompanha de verdade pra bater o
    faturamento realizado contra a meta é o planejamento PCP, "diante do
    planejamento mensal/semanal PCP" — os dois podiam divergir bastante,
    porque Liberação prevista é preenchida item a item e nem sempre
    acompanha o planejamento semanal que o PCP realmente definiu.

    Análises novas (pedido do Bruno, 09/09/2026: "seja criativo... áreas
    voltadas para faturamentos regionais, estações (tipo de produto),
    principais clientes (regra 80/20)") — todas calculadas em cima do mesmo
    `itens_previstos` acima (o planejamento PCP do mês em vista), porque o
    pedido foi explícito: "considere os números voltados para o mês de
    Setembro (planejamento PCP)" — ou seja, essas quebras são sobre o que
    está PLANEJADO pra faturar no mês, não sobre o que já foi faturado até
    agora (que a esta altura do mês ainda é baixo e não conta a história
    toda).

    Atualização (pedido do Bruno, 09/09/2026): as 3 tabelas antigas
    "Realizado por cliente/região/vendedor" saíram da tela — misturavam, sem
    nenhum cruzamento entre si, o Realizado (poucos itens já faturados até
    agora no mês) ao lado do Previsto (planejamento PCP do mês inteiro), o
    que Bruno achou confuso porque os dois conjuntos de itens não
    "confrontam" um com o outro. Ficam só os 3 números-resumo do topo
    (Previsto/Realizado/Itens faturados) e as análises novas do Previsto."""
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

    itens_com_planejamento_pcp = base.filter(ItemPedido.planejamento_semanal.isnot(None)).all()
    itens_previstos = [
        i for i in itens_com_planejamento_pcp
        if _mes_ano_da_semana_pcp(i.planejamento_semanal) == (ano, mes)
    ]
    itens_realizados = base.filter(
        ItemPedido.liberacao_faturamento >= inicio, ItemPedido.liberacao_faturamento < fim
    ).all()

    def _agrupar(itens, chave_fn, valor_fn=lambda i: i.valor_faturamento_realizado):
        agrupado = {}
        for i in itens:
            chave = chave_fn(i) or "—"
            agrupado[chave] = agrupado.get(chave, 0.0) + valor_fn(i)
        linhas = [{"chave": k, "valor": round(v, 2)} for k, v in agrupado.items()]
        linhas.sort(key=lambda l: l["valor"], reverse=True)
        return linhas

    def _curva_pareto_clientes(itens):
        """Curva de Pareto (regra 80/20) do Previsto por cliente: ordena do
        maior pro menor e marca como "principal" cada cliente até o ponto em
        que o acumulado atinge 80% do total — o grupo enxuto de clientes que
        concentra a maior parte do faturamento previsto do mês."""
        linhas = _agrupar(itens, lambda i: i.pedido.cliente if i.pedido else None, valor_fn=lambda i: i.valor_total)
        total = sum(l["valor"] for l in linhas)
        curva = []
        acumulado = 0.0
        ja_atingiu_80 = False
        for l in linhas:
            principal = not ja_atingiu_80
            acumulado += l["valor"]
            pct = round(l["valor"] / total * 100, 1) if total else 0.0
            pct_acumulado = round(acumulado / total * 100, 1) if total else 0.0
            if pct_acumulado >= 80:
                ja_atingiu_80 = True
            curva.append({**l, "pct": pct, "pct_acumulado": pct_acumulado, "principal": principal})
        n_principais = sum(1 for c in curva if c["principal"])
        return {
            "linhas": curva,
            "n_clientes": len(curva),
            "n_principais": n_principais,
            "pct_clientes_principais": round(n_principais / len(curva) * 100, 1) if curva else 0.0,
        }

    return {
        "previsto_total": round(sum(i.valor_total for i in itens_previstos), 2),
        "realizado_total": round(sum(i.valor_faturamento_realizado for i in itens_realizados), 2),
        "itens_realizados": len(itens_realizados),
        # --- Análises do Previsto (Planejamento PCP) — ver docstring acima ---
        "previsto_por_regiao": _agrupar(
            itens_previstos,
            lambda i: REGIAO_POR_UF.get(i.pedido.estado) if i.pedido and i.pedido.estado else None,
            valor_fn=lambda i: i.valor_total,
        ),
        "previsto_por_estacao": _agrupar(
            itens_previstos,
            lambda i: rotulo_estacao(i.estacao) if i.estacao else None,
            valor_fn=lambda i: i.valor_total,
        ),
        "previsto_pareto_clientes": _curva_pareto_clientes(itens_previstos),
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
    sem_planejamento_semanal = args.get("sem_planejamento_semanal", "").strip()

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
    # Quadrante "Sem planejamento PCP" (pedido do Bruno, 21/09/2026, ao lado do
    # quadrante do mês seguinte): pedidos com pelo menos 1 item cujo
    # Planejamento semanal (PCP) ainda não foi preenchido — mesmo espírito dos
    # outros quadrantes (conta PEDIDOS, a listagem abaixo estreita pra só os
    # ITENS sem planejamento, ver _linhas_listagem_geral).
    if sem_planejamento_semanal:
        query = query.filter(Pedido.itens.any(ItemPedido.planejamento_semanal.is_(None)))
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
        sem_planejamento_semanal=sem_planejamento_semanal,
    )
    return query, filtros


def _quadrantes_planejamento_semanal(filtros, hoje=None):
    """Quadrantes clicáveis no topo da Listagem Geral de Produção (pedido do
    Bruno, 10/09/2026): 1 quadrante pro mês corrente inteiro + 1 por semana
    dele, no espírito do "Faturamento por Semana" que já existe em Gestão
    Operação — clicar já atualiza a listagem (mesmo link de sempre, sem
    JS/AJAX, só reaproveita o próprio filtro de Planejamento Semanal/Mensal
    PCP que a tela já tinha).

    Sempre o mês ATUAL (não fica preso a setembro — troca sozinho quando o
    mês virar, sem precisar mexer em nada aqui). O número de semanas também
    é dinâmico: usa a mesma lista de rótulos "SEMANA NN / MÊS / ANO" de
    gerar_semanas_pcp — a maioria dos meses tem 5 semanas nesse critério,
    alguns têm só 4 (nunca mais que 5, já que ceil(31/7)=5). O que CONTA pra
    cada quadrante (e o que o clique filtra) continua sendo exatamente esse
    rótulo — "baseado no planejamento PCP" como o Bruno pediu (10/09/2026) —
    então nada aqui muda quais pedidos aparecem, só como o card é rotulado.

    O texto do período mostrado em cada card de semana, porém, é a semana de
    CALENDÁRIO de verdade (domingo a sábado) que contém aquele bloco de dias
    — ajustado a pedido do Bruno (10/09/2026): "quero as datas de cada
    quadrante assim: SEMANA 01: 30/08 A 05/09..." — mesmo espírito calendário
    já usado em Programação (_semanas_calendario_pcp), só que aqui é sempre
    exatamente 1 semana por rótulo de gerar_semanas_pcp (não a grade cheia do
    mês), ancorada no domingo igual ou anterior ao dia 1 do mês. Cada card
    de semana também carrega `atual` — True só pro card cuja semana de
    calendário contém a data de hoje — pro pisca-pisca visual (pedido do
    Bruno, 10/09/2026) que mostra em qual semana estamos agora.

    Cada quadrante já mostra quantos PEDIDOS distintos caem naquele período,
    considerando os OUTROS filtros já ativos na tela (cliente, vendedor,
    status, estação, busca etc.) — só ignora o Planejamento Semanal/Mensal
    atual, senão a contagem de cada quadrante ficaria igual à do que já
    estiver selecionado, em vez do total real daquele período. Reaproveita
    _filtrar_pedidos (mesma regra de sempre) pra nunca divergir da lógica
    que a tabela abaixo usa.

    Quadrante "mes_seguinte" (pedido do Bruno, 17/09/2026: "ao lado do
    quadrante SEMANA 05, o quadrante OUTUBRO... completo todo o
    planejamento do mês... número de pedidos que consta no mês... cor azul
    claro, diferenciando dos demais") — mesmo padrão exato do "mes_atual"
    acima (mesmo `contar()`/`filtros_link` via `planejamento_mensal`, nunca
    diverge da tabela), só que pro mês SEGUINTE em vez do atual. Não usa
    `_resumo_mes_pcp`/`_resumo_mes_seguinte_pcp` (que contam por
    ItemPedido.planejamento_semanal agrupado) de propósito: esse quadrante
    fica lado a lado com mes_atual/semanas na mesma linha da Listagem
    Geral, então precisa usar a MESMA contagem/mesmo clique-pra-filtrar
    (_filtrar_pedidos) que os vizinhos, pra nunca mostrar um número
    diferente do que a tabela mostra ao clicar. `fundo="primary-subtle"` é
    o que o template usa pra pintar de azul claro (diferente dos outros,
    que só têm borda)."""
    hoje = hoje or date.today()
    ano, mes = hoje.year, hoje.month
    dias_no_mes = monthrange(ano, mes)[1]
    rotulos_semana = gerar_semanas_pcp(meses_atras=0, meses_frente=0, hoje=hoje)

    # Domingo igual ou anterior ao dia 1 do mês — âncora da "semana 01" no
    # calendário (weekday(): 0=segunda ... 6=domingo).
    primeiro_dia_mes = date(ano, mes, 1)
    domingo_semana_01 = primeiro_dia_mes - timedelta(days=(primeiro_dia_mes.weekday() + 1) % 7)

    filtros_outros = dict(filtros, planejamento_semanal="", planejamento_mensal="", sem_planejamento_semanal="")

    def contar(**override):
        query, _ = _filtrar_pedidos(dict(filtros_outros, **override))
        return query.count()

    valor_mes = f"{ano}-{mes:02d}"
    mes_atual = {
        "titulo": MESES_PT_EXTENSO[mes - 1].upper(),
        "subtitulo": f"01/{mes:02d} – {dias_no_mes:02d}/{mes:02d}",
        "total": contar(planejamento_mensal=valor_mes),
        "ativo": filtros.get("planejamento_mensal") == valor_mes,
        "filtros_link": dict(filtros_outros, planejamento_mensal=valor_mes),
    }

    semanas = []
    for n, rotulo in enumerate(rotulos_semana, start=1):
        inicio_semana = domingo_semana_01 + timedelta(days=7 * (n - 1))
        fim_semana = inicio_semana + timedelta(days=6)
        semanas.append(
            {
                "titulo": f"SEMANA {n:02d}",
                "subtitulo": f"{inicio_semana.strftime('%d/%m')} a {fim_semana.strftime('%d/%m')}",
                "total": contar(planejamento_semanal=rotulo),
                "ativo": filtros.get("planejamento_semanal") == rotulo,
                "filtros_link": dict(filtros_outros, planejamento_semanal=rotulo),
                "atual": inicio_semana <= hoje <= fim_semana,
            }
        )

    ano_seg, mes_seg = _somar_meses(ano, mes, 1)
    dias_no_mes_seg = monthrange(ano_seg, mes_seg)[1]
    valor_mes_seguinte = f"{ano_seg}-{mes_seg:02d}"
    mes_seguinte = {
        "titulo": MESES_PT_EXTENSO[mes_seg - 1].upper(),
        "subtitulo": f"01/{mes_seg:02d} – {dias_no_mes_seg:02d}/{mes_seg:02d}",
        "total": contar(planejamento_mensal=valor_mes_seguinte),
        "ativo": filtros.get("planejamento_mensal") == valor_mes_seguinte,
        "filtros_link": dict(filtros_outros, planejamento_mensal=valor_mes_seguinte),
    }

    # Quadrante "Sem planejamento PCP" (pedido do Bruno, 21/09/2026, "ao lado
    # do quadrante de outubro"): pedidos com pelo menos 1 item cujo
    # Planejamento semanal (PCP) ainda não foi preenchido — mesmo mecanismo
    # de clique-pra-filtrar dos outros (`sem_planejamento_semanal=1`, ver
    # _filtrar_pedidos/_linhas_listagem_geral), sem recorte de mês (mostra
    # TODOS os pendentes, não só os do mês atual/seguinte, já que o ponto é
    # justamente achar quem ainda não entrou em nenhum planejamento).
    sem_planejamento = {
        "titulo": "SEM PLANEJAMENTO PCP",
        "subtitulo": "ainda não planejados por semana",
        "total": contar(sem_planejamento_semanal="1"),
        "ativo": filtros.get("sem_planejamento_semanal") == "1",
        "filtros_link": dict(filtros_outros, sem_planejamento_semanal="1"),
    }

    return {"mes_atual": mes_atual, "semanas": semanas, "mes_seguinte": mes_seguinte, "sem_planejamento": sem_planejamento}


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
    sem_planejamento_semanal = args.get("sem_planejamento_semanal", "").strip()
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
            # Quadrante "Sem planejamento PCP" — estreita pra só os itens sem
            # Planejamento semanal (PCP) preenchido, mesmo raciocínio do
            # filtro de planejamento_semanal exato acima (não mistura com os
            # itens já planejados de um pedido que tem os dois casos).
            if sem_planejamento_semanal and item.planejamento_semanal is not None:
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


def _normalizar_pedido_venda(valor):
    """Normaliza um texto de "nº pedido de venda" pro casamento entre Gestão
    Produção e Gestão Operação (tabelas sem FK, ligadas só por esse texto
    digitado em cada lado) — pedido do Bruno (10/09/2026): "tenho vários
    pedidos com divergência dos status pra realidade" + "não quero número
    zero antes" (reparou pedidos aparecendo como "000872" na listagem).
    Rastreado até aqui: até agora o casamento só tirava espaço (func.trim/
    .strip()) — quando o mesmo pedido é digitado como "872" em Gestão
    Produção e "000872" em Gestão Operação (ou vice-versa), os dois textos
    NUNCA batiam, e o pedido ficava "órfão" pro outro lado — daí o status
    mostrado (calculado só com os dados do lado que bateu) divergindo do
    real. Corrigido tirando também os zeros à esquerda dos dois lados,
    sempre, em todo lugar que casa as duas tabelas — nunca só de um lado.
    "0"/"000"/"" viram "0" (nunca string vazia, pra não virar chave "" e
    colidir com pedidos sem número nenhum)."""
    if not valor:
        return ""
    base = valor.strip()
    if not base:
        return ""
    return base.lstrip("0") or "0"


def _pedido_venda_normalizado_sql(coluna):
    """Equivalente em SQL de _normalizar_pedido_venda (mesmo critério, pros
    dois lados nunca divergirem) — usado em TODO filtro/JOIN "manual" entre
    Pedido/PedidoOperacao no lugar de func.trim() sozinho. ltrim(texto,
    caracteres) existe tanto em SQLite (dev local) quanto em Postgres
    (produção) com esse segundo argumento de conjunto de caracteres a
    remover — portátil nos dois bancos."""
    base = func.trim(coluna)
    sem_zeros = func.ltrim(base, "0")
    return case(
        (base == "", ""),
        (sem_zeros == "", "0"),
        else_=sem_zeros,
    )


def _pedidos_venda_com_planejamento_semanal(rotulos):
    """Lista de `pedido_venda` (trim) de Gestão Produção cujos itens têm
    `planejamento_semanal` dentro de `rotulos` (1 rótulo, ou vários — ex. os
    rótulos de um mês inteiro). Pedido do Bruno (10/09/2026): "quero que
    todo o grupo gestão operação esteja 100% sincronizado com o gestão
    produção... principalmente listagem e pcp" — leva os mesmos quadrantes/
    filtro de Planejamento Semanal (PCP) da Listagem Geral de Produção pra
    Gestão Operação. PedidoOperacao é tabela independente (sem FK) — o único
    jeito de aplicar esse filtro lá é via IN nesse texto, seguindo o MESMO
    casamento por pedido_venda (trim, exato, nunca aproximado) já usado por
    _itens_producao_por_pedido_venda e as outras funções "ao vivo" desta
    seção."""
    if isinstance(rotulos, str):
        rotulos = [rotulos]
    rotulos = [r for r in rotulos if r]
    if not rotulos:
        return []
    linhas = (
        db.session.query(_pedido_venda_normalizado_sql(Pedido.pedido_venda))
        .join(ItemPedido, ItemPedido.pedido_id == Pedido.id)
        .filter(ItemPedido.planejamento_semanal.in_(rotulos), Pedido.pedido_venda.isnot(None))
        .distinct()
        .all()
    )
    return [linha[0] for linha in linhas if linha[0]]


def _pedidos_venda_finalizados_producao():
    """Conjunto de `pedido_venda` (trim) de Gestão Produção cujo
    status_producao é FINALIZADO — usado só por
    _aplicar_filtro_status_pedido_operacao (etapa 3 do "Status pedido"),
    pra reproduzir em SQL a mesma condição que _indice_etapa_pedido já
    calcula em Python (`pedido.status_producao == "FINALIZADO"`)."""
    linhas = (
        db.session.query(_pedido_venda_normalizado_sql(Pedido.pedido_venda))
        .filter(Pedido.status_producao == "FINALIZADO", Pedido.pedido_venda.isnot(None))
        .distinct()
        .all()
    )
    return {linha[0] for linha in linhas if linha[0]}


def _pedidos_venda_em_producao():
    """Conjunto de `pedido_venda` (trim) de Gestão Produção com pelo menos 1
    item cujo `inicio_producao` está preenchido — usado só por
    _aplicar_filtro_status_pedido_operacao (etapa 2 do "Status pedido"),
    reproduzindo em SQL a mesma condição de _indice_etapa_pedido
    (`any(item.inicio_producao for item in pedido.itens)`)."""
    linhas = (
        db.session.query(_pedido_venda_normalizado_sql(Pedido.pedido_venda))
        .join(ItemPedido, ItemPedido.pedido_id == Pedido.id)
        .filter(ItemPedido.inicio_producao.isnot(None), Pedido.pedido_venda.isnot(None))
        .distinct()
        .all()
    )
    return {linha[0] for linha in linhas if linha[0]}


def _condicoes_etapa_pedido_operacao():
    """As 5 condições SQL (mutuamente exclusivas, MESMA ordem de precedência
    de _indice_etapa_pedido/_metricas_operacao_360) usadas pelo filtro
    "Status produção" (lista multi-seleção) da Operação 360 — pedido do
    Bruno (10/09/2026). Reimplementada em SQL (em vez de reaproveitar a
    função Python direto) porque aqui precisamos FILTRAR e paginar no banco,
    não só calcular a etapa de pedidos já carregados; a precedência é a
    mesma (5 entregue > 4 expedido > 3 liberado PCP/produção finalizada > 2
    em produção > 1 pendente produção — cada etapa exclui as de número
    maior, igual ao if/elif em cadeia de _indice_etapa_pedido)."""
    finalizados_pv = _pedidos_venda_finalizados_producao()
    em_producao_pv = _pedidos_venda_em_producao()
    pv_trim = _pedido_venda_normalizado_sql(PedidoOperacao.pedido_venda)

    cond5 = or_(PedidoOperacao.go_data_entregue_cliente.isnot(None), PedidoOperacao.go_data_real_entrega.isnot(None))
    cond4 = and_(PedidoOperacao.go_data_pedido_expedido.isnot(None), not_(cond5))
    cond3_base = or_(
        pv_trim.in_(finalizados_pv) if finalizados_pv else false(),
        PedidoOperacao.go_data_efetiva_liberacao_pcp.isnot(None),
    )
    cond3 = and_(cond3_base, not_(cond4), not_(cond5))
    cond2_base = pv_trim.in_(em_producao_pv) if em_producao_pv else false()
    cond2 = and_(cond2_base, not_(cond3_base), not_(cond4), not_(cond5))
    # etapa 1 ("Pendente produção"): nenhuma das condições acima bate.
    cond1 = and_(not_(cond2_base), not_(cond3_base), not_(cond4), not_(cond5))
    return {1: cond1, 2: cond2, 3: cond3, 4: cond4, 5: cond5}


def _aplicar_filtro_status_pedido_operacao(query, etapas_idx):
    """Filtra `query` (PedidoOperacao) por 1 OU MAIS etapas selecionadas
    (multi-seleção, pedido do Bruno 10/09/2026: "quero que essas listas
    seja possível selecionar mais de uma opção... ex: selecionar status
    produção e inspeção/expedição") — OR entre as condições de cada etapa
    marcada, ver _condicoes_etapa_pedido_operacao."""
    condicoes = _condicoes_etapa_pedido_operacao()
    selecionadas = [condicoes[i] for i in etapas_idx if i in condicoes]
    if not selecionadas:
        return query
    return query.filter(or_(*selecionadas))


def _getlist_seguro(args, chave):
    """Mesmo que `args.getlist(chave)`, mas também aceita um dict Python
    comum no lugar de request.args (MultiDict) — necessário porque
    _quadrantes_planejamento_semanal_operacao chama _filtrar_pedidos_
    operacao com um dict puro (`dict(filtros_outros, **override)`), que não
    tem método `.getlist()`. Um valor já em lista (como `filtros["status_
    pedido"]`, que vem de outra chamada a este mesmo filtro) passa direto;
    um valor único vira lista de 1 item."""
    if hasattr(args, "getlist"):
        return args.getlist(chave)
    valor = args.get(chave)
    if valor is None:
        return []
    if isinstance(valor, (list, tuple, set)):
        return list(valor)
    return [valor]


def _filtrar_pedidos_operacao(args):
    """Filtros das 4 sub-abas de Gestão Operação (Comercial/PCP/Logística/
    Resultados). Independente de _filtrar_pedidos (Gestão Produção) — opera só
    em PedidoOperacao, sem nenhum join com Pedido/ItemPedido/estação — com 2
    exceções pontuais abaixo (planejamento_semanal/mensal), que fazem uma
    consulta à parte só pra achar QUAIS pedido_venda bater, sem criar
    nenhuma relação/FK real entre as tabelas.

    `segmento` + `periodo` (pedido do Bruno, 03/09/2026 — tela Resultados/
    OTD: "quero ver todos os pedidos de julho faturados ou dentro do
    planejamento semanal do pcp, com isso ver o otd"; ampliado no mesmo dia
    pra aceitar não só mês, mas trimestre/semestre/ano/"todos" — ver
    _parse_periodo) — reaproveita a MESMA definição de "período" já usada em
    _faturamento_por_periodo: agrupa pelo Término Semanal PCP
    (go_termino_semanal_pcp), não por nenhuma data de calendário.
      - "planejamento": todo pedido cujo Término Semanal PCP cai no período
        escolhido (equivalente ao "Qtd/Valor liberado" da tabela semanal).
      - "faturados": o mesmo conjunto acima, restrito a quem já tem Valor NF
        Emitida preenchido (equivalente ao "Qtd/Valor faturado").
    Sem `segmento`, nenhum filtro de período é aplicado — comportamento
    antigo, todos os pedidos.

    `planejamento_semanal`/`planejamento_mensal` (pedido do Bruno,
    10/09/2026): MESMO filtro/convenção de Gestão Produção
    (ItemPedido.planejamento_semanal, "SEMANA NN / MÊS / ANO"), usado pelos
    quadrantes de PCP em Listagem Geral e PCP de Gestão Operação — ver
    _quadrantes_planejamento_semanal_operacao. Deliberadamente NÃO usa
    PedidoOperacao.go_termino_semanal_pcp (campo próprio, digitado/
    importado à parte) — o objetivo aqui é filtrar pelo dado real de
    planejamento da fábrica, o mesmo que já aparece ao vivo na coluna
    "Término semanal" da tela PCP (ver _liberacao_pcp_por_pedido_venda).

    `data_inicio`/`data_fim` (mesmo pedido): intervalo de Data de inclusão,
    mesmo campo/rótulo do filtro equivalente em Gestão Produção.

    `status_pedido`/`otd`/`frete` (pedido do Bruno, 10/09/2026, na 1ª volta
    como quadrantes; ampliado no mesmo dia pra "quero em formato de listas...
    seja possível selecionar mais de uma opção" — agora são listas
    multi-seleção, mesmo padrão de `resultado`/`categoria_desvio` em
    _filtrar_inspecoes_finais): `status_pedido` é uma lista de "1".."5"
    (mesma classificação de _indice_etapa_pedido, OR entre as selecionadas —
    ver _aplicar_filtro_status_pedido_operacao); `otd` é uma lista com
    "SIM"/"NAO"/"PENDENTE" (PENDENTE = go_otd_realizado ainda vazio); `frete`
    é uma lista com valores de FRETE_OPCOES e/ou "NAO_INFORMADO" (vazio ou
    fora da lista). `nf_mes_atual` continua um quadrante simples (não virou
    lista — não fazia parte do pedido de mudança) — ="1" filtra
    go_data_emissao_nf dentro do mês corrente.

    `nf_mes` (pedido do Bruno, 10/09/2026, ao lado do Planejamento semanal/
    mensal de PCP): igual "Planejamento mensal (PCP)", mas pra Emissão de
    NF — <input type="month"> ("AAAA-MM"), filtra go_data_emissao_nf dentro
    de QUALQUER mês escolhido (não só o mês corrente, diferente do quadrante
    `nf_mes_atual` acima — os dois convivem, um é atalho rápido pro mês de
    hoje, o outro escolhe qualquer mês)."""
    query = PedidoOperacao.query

    cliente = args.get("cliente", "").strip()
    vendedor = args.get("vendedor", "").strip()
    busca = args.get("busca", "").strip()
    segmento = args.get("segmento", "").strip()
    periodo_str = args.get("periodo", "").strip()
    planejamento_semanal = args.get("planejamento_semanal", "").strip()
    planejamento_mensal = args.get("planejamento_mensal", "").strip()
    data_inicio = args.get("data_inicio", "").strip()
    data_fim = args.get("data_fim", "").strip()
    status_pedido = [v for v in _getlist_seguro(args, "status_pedido") if v in {"1", "2", "3", "4", "5"}]
    otd = [v.upper() for v in _getlist_seguro(args, "otd") if v.upper() in {"SIM", "NAO", "PENDENTE"}]
    frete_filtro = [v for v in _getlist_seguro(args, "frete") if v]
    nf_mes_atual = args.get("nf_mes_atual", "").strip()
    nf_mes = args.get("nf_mes", "").strip()

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
    if data_inicio:
        data_inicio_parsed = _parse_data_form(data_inicio)
        if data_inicio_parsed:
            query = query.filter(PedidoOperacao.data_inclusao_pedido >= data_inicio_parsed)
    if data_fim:
        data_fim_parsed = _parse_data_form(data_fim)
        if data_fim_parsed:
            query = query.filter(PedidoOperacao.data_inclusao_pedido <= data_fim_parsed)
    if planejamento_semanal:
        pedidos_venda_match = _pedidos_venda_com_planejamento_semanal(planejamento_semanal)
        if pedidos_venda_match:
            query = query.filter(_pedido_venda_normalizado_sql(PedidoOperacao.pedido_venda).in_(pedidos_venda_match))
        else:
            query = query.filter(false())
    if planejamento_mensal:
        mes_ano = _parse_mes_ano_form(planejamento_mensal, None)
        if mes_ano:
            semanas_do_mes = [
                s for (s,) in db.session.query(ItemPedido.planejamento_semanal)
                .filter(ItemPedido.planejamento_semanal.isnot(None))
                .distinct()
                if _mes_ano_da_semana_pcp(s) == mes_ano
            ]
            pedidos_venda_match = _pedidos_venda_com_planejamento_semanal(semanas_do_mes)
            if pedidos_venda_match:
                query = query.filter(_pedido_venda_normalizado_sql(PedidoOperacao.pedido_venda).in_(pedidos_venda_match))
            else:
                # Mês escolhido não tem nenhum planejamento semanal preenchido
                # ainda (ou nenhum pedido de Produção bate) — não deve
                # mostrar nada, mesmo comportamento de _filtrar_pedidos.
                query = query.filter(false())
        else:
            query = query.filter(false())

    if segmento in ("planejamento", "faturados"):
        tipo_p, ano_p, valor_p, _ = _parse_periodo(periodo_str)
        query = _filtrar_por_periodo_pcp(query, tipo_p, ano_p, valor_p)
        if segmento == "faturados":
            query = query.filter(PedidoOperacao.go_valor_nf_emitida.isnot(None))
    else:
        segmento = ""

    if status_pedido:
        query = _aplicar_filtro_status_pedido_operacao(query, [int(v) for v in status_pedido])

    if otd:
        condicoes_otd = []
        if "SIM" in otd:
            condicoes_otd.append(PedidoOperacao.go_otd_realizado == "SIM")
        if "NAO" in otd:
            condicoes_otd.append(PedidoOperacao.go_otd_realizado == "NÃO")
        if "PENDENTE" in otd:
            condicoes_otd.append(or_(PedidoOperacao.go_otd_realizado.is_(None), PedidoOperacao.go_otd_realizado == ""))
        query = query.filter(or_(*condicoes_otd))

    if frete_filtro:
        condicoes_frete = []
        valores_frete_opcoes = [v.upper() for v in frete_filtro if v != "NAO_INFORMADO"]
        if valores_frete_opcoes:
            condicoes_frete.append(func.upper(func.trim(PedidoOperacao.frete)).in_(valores_frete_opcoes))
        if "NAO_INFORMADO" in frete_filtro:
            condicoes_frete.append(
                or_(
                    PedidoOperacao.frete.is_(None),
                    PedidoOperacao.frete == "",
                    not_(func.upper(func.trim(PedidoOperacao.frete)).in_(FRETE_OPCOES)),
                )
            )
        query = query.filter(or_(*condicoes_frete))

    if nf_mes_atual == "1":
        hoje_filtro = date.today()
        primeiro_dia_filtro = hoje_filtro.replace(day=1)
        ultimo_dia_filtro = date(hoje_filtro.year, hoje_filtro.month, monthrange(hoje_filtro.year, hoje_filtro.month)[1])
        query = query.filter(PedidoOperacao.go_data_emissao_nf.between(primeiro_dia_filtro, ultimo_dia_filtro))
    else:
        nf_mes_atual = ""

    if nf_mes:
        mes_ano_nf = _parse_mes_ano_form(nf_mes, None)
        if mes_ano_nf:
            ano_nf, mes_nf = mes_ano_nf
            primeiro_dia_nf = date(ano_nf, mes_nf, 1)
            ultimo_dia_nf = date(ano_nf, mes_nf, monthrange(ano_nf, mes_nf)[1])
            query = query.filter(PedidoOperacao.go_data_emissao_nf.between(primeiro_dia_nf, ultimo_dia_nf))
        else:
            query = query.filter(false())

    query = query.order_by(PedidoOperacao.data_inclusao_pedido.desc().nullslast(), PedidoOperacao.id.desc())

    filtros = dict(
        cliente=cliente, vendedor=vendedor, busca=busca, segmento=segmento,
        planejamento_semanal=planejamento_semanal, planejamento_mensal=planejamento_mensal,
        data_inicio=data_inicio, data_fim=data_fim,
        status_pedido=status_pedido, otd=otd, frete=frete_filtro, nf_mes_atual=nf_mes_atual,
        nf_mes=nf_mes,
    )
    return query, filtros


def _linhas_gestao_operacao(args):
    """Usado pelas 4 sub-abas de Gestão Operação: pagina o resultado de
    _filtrar_pedidos_operacao. Cada linha já é 1 pedido comercial (tabela
    própria PedidoOperacao) — sem duplicidade legada, sem precisar agrupar
    nada em Python. Devolve também a `query` (filtrada, sem paginação) —
    usada pela tela Resultados/OTD pra calcular o resumo de OTD só sobre o
    mesmo conjunto filtrado (ver _resumo_otd)."""
    page = args.get("page", 1, type=int)
    query, filtros = _filtrar_pedidos_operacao(args)
    total_filtrado = query.count()
    total_paginas = max(1, (total_filtrado + PAGE_SIZE - 1) // PAGE_SIZE)
    pagina = query.offset((page - 1) * PAGE_SIZE).limit(PAGE_SIZE).all()
    return pagina, page, total_paginas, total_filtrado, filtros, query


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
    valores = sorted({_normalizar_pedido_venda(v) for v in pedidos_venda if v and v.strip()})
    if not valores:
        return {}
    pedidos = (
        Pedido.query.options(selectinload(Pedido.itens))
        # func.trim() nos dois lados: o texto salvo em Pedido.pedido_venda às
        # vezes tem espaço a mais (import antigo de planilha) — sem isso, a
        # igualdade exata falharia por causa só do espaço, escondendo itens
        # que na prática são do mesmo pedido.
        .filter(_pedido_venda_normalizado_sql(Pedido.pedido_venda).in_(valores))
        .all()
    )
    mapa = {}
    for pedido in pedidos:
        chave = _normalizar_pedido_venda(pedido.pedido_venda)
        if not chave:
            continue
        mapa.setdefault(chave, []).extend(pedido.itens)
    return mapa


def _status_producao_por_pedido_venda(pedidos_venda):
    """Pedido do Bruno (01/09/2026): a coluna "Status produção" das telas de
    Gestão Operação (PCP e a tela de editar) tem que refletir o andamento
    REAL na fábrica — mesmo casamento por pedido_venda (trim, sem FK, nunca
    aproximado) já usado pra mostrar os itens na Listagem Geral — em vez do
    status calculado só a partir dos próprios dados da Operação (que
    continua existindo como fallback em PedidoOperacao.status_producao, pra
    quando o pedido ainda não foi lançado em Gestão Produção). Reaproveita a
    mesma regra de agregação de Pedido.status_producao (EM TRATATIVA vence
    tudo; só FINALIZADO se todos os itens estiverem; só PENDENTE se todos
    estiverem; senão ANDAMENTO)."""
    itens_por_pedido = _itens_producao_por_pedido_venda(pedidos_venda)
    status_por_pedido = {}
    for chave, itens in itens_por_pedido.items():
        if not itens:
            continue
        status_itens = {item.status_producao for item in itens}
        if "EM TRATATIVA" in status_itens:
            status_por_pedido[chave] = "EM TRATATIVA"
        elif status_itens == {"FINALIZADO"}:
            status_por_pedido[chave] = "FINALIZADO"
        elif status_itens == {"PENDENTE"}:
            status_por_pedido[chave] = "PENDENTE"
        else:
            status_por_pedido[chave] = "ANDAMENTO"
    return status_por_pedido


def _liberacao_pcp_por_pedido_venda(pedidos_venda):
    """Pedido do Bruno (01/09/2026, ampliado 03/09/2026): as colunas
    "Previsão liberação PCP", "Data efetiva liberação" e "Término semanal"
    da tela PCP de Gestão Operação têm que acompanhar automaticamente os
    dados equivalentes já preenchidos em Gestão Produção ("Liberação
    prevista"/"Liberação real"/"Planejamento semanal" de cada item) — mesmo
    casamento por pedido_venda (trim, sem FK, nunca aproximado) já usado pra
    Status produção/itens da Listagem Geral. Enquanto o pedido não tiver
    sido lançado em Gestão Produção (sem match), os campos próprios de
    PedidoOperacao (go_previsao_liberacao_pcp/go_data_efetiva_liberacao_pcp/
    go_termino_semanal_pcp) continuam valendo como fallback editável
    manualmente — mesmo espírito de status_producao/
    _status_producao_por_pedido_venda.

    Previsão/Efetiva são preenchidas em Produção por um único bloco no topo
    do formulário de edição que aplica o MESMO valor a todos os itens do
    pedido de uma vez (decisão de 31/08/2026) — então, quando por algum
    motivo os itens de um mesmo pedido têm valores diferentes entre si (dado
    legado, de antes dessa decisão existir), usamos a data mais recente
    entre eles. "Término semanal" é texto (ex. "SEMANA 03 / AGO / 2026"), não
    dado — usamos _chave_semana_pcp (mesma função que já ordena esse rótulo
    em outros lugares do sistema) pra achar o rótulo cronologicamente mais
    recente entre os itens."""
    itens_por_pedido = _itens_producao_por_pedido_venda(pedidos_venda)
    liberacao_por_pedido = {}
    for chave, itens in itens_por_pedido.items():
        previstas = [i.liberacao_prevista for i in itens if i.liberacao_prevista]
        reais = [i.liberacao_real for i in itens if i.liberacao_real]
        semanas = [i.planejamento_semanal for i in itens if i.planejamento_semanal]
        if not previstas and not reais and not semanas:
            continue
        liberacao_por_pedido[chave] = {
            "previsao": max(previstas) if previstas else None,
            "efetiva": max(reais) if reais else None,
            "termino_semanal": max(semanas, key=_chave_semana_pcp) if semanas else None,
        }
    return liberacao_por_pedido


def _data_cliente_por_pedido_venda(pedidos_venda):
    """Pedido.data_cliente por pedido_venda (trim, sem FK) — pedido do Bruno
    (03/09/2026): "quero que todos os dados dentro da gestão operação seja
    extraída automaticamente da gestão produção". "Solicitada cliente/
    retira" (PCP) e "Data solicitada entrega" (Comercial) de Gestão Operação
    são o MESMO dado que "Data do cliente" de Gestão Produção (já
    documentado em _criar_pedido_operacao_a_partir_de_producao) — até aqui
    elas só copiavam esse valor UMA VEZ na inclusão do pedido em Operação;
    agora acompanham ao vivo, mesmo espírito de _liberacao_pcp_por_pedido_
    venda. Os campos próprios de PedidoOperacao continuam valendo como
    fallback editável manualmente enquanto o pedido não tiver sido lançado
    em Gestão Produção."""
    valores = sorted({_normalizar_pedido_venda(v) for v in pedidos_venda if v and v.strip()})
    if not valores:
        return {}
    pedidos = (
        Pedido.query
        .filter(_pedido_venda_normalizado_sql(Pedido.pedido_venda).in_(valores), Pedido.data_cliente.isnot(None))
        .all()
    )
    mapa = {}
    for pedido in pedidos:
        chave = _normalizar_pedido_venda(pedido.pedido_venda)
        if chave:
            mapa[chave] = pedido.data_cliente
    return mapa


def _pedidos_producao_por_pedido_venda(pedidos_venda):
    """dict pedido_venda (trim) -> Pedido (Gestão Produção), com os itens já
    carregados. Usado quando precisamos do objeto Pedido inteiro (não só os
    itens/status já resumidos por _itens_producao_por_pedido_venda) — ex.
    pra reaproveitar _indice_etapa_pedido sem duplicar sua lógica em outro
    lugar. Mesmo casamento por texto (trim, exato, nunca aproximado) de
    sempre."""
    valores = sorted({_normalizar_pedido_venda(v) for v in pedidos_venda if v and v.strip()})
    if not valores:
        return {}
    pedidos = (
        Pedido.query.options(selectinload(Pedido.itens))
        .filter(_pedido_venda_normalizado_sql(Pedido.pedido_venda).in_(valores))
        .all()
    )
    mapa = {}
    for pedido in pedidos:
        chave = _normalizar_pedido_venda(pedido.pedido_venda)
        if chave and chave not in mapa:
            mapa[chave] = pedido
    return mapa


# Emoji por etapa do "Acompanhamento do pedido" (pedido do Bruno, 10/09/2026,
# coluna "Status pedido" da Operação 360: "use emoções pra sinalizar") — na
# mesma ordem/índice de _ETAPAS_ACOMPANHAMENTO_PEDIDO (definida mais abaixo,
# junto com _indice_etapa_pedido — reaproveitados aqui, não duplicados).
_ETAPA_EMOJI = ["📥", "⚙️", "📋", "🚚", "✅"]


def _metricas_operacao_360(pedidos, liberacao_pcp_por_pedido_venda, data_cliente_por_pedido_venda, pedidos_producao_por_pedido_venda):
    """"Operação 360" — Listagem Geral de Gestão Operação (pedido do Bruno,
    10/09/2026): "quero que contemple as principais informações dos
    pedidos... data inclusão, data solicitada, data conclusão produção,
    expedido em, real entrega, lead time comercial, lead time produção
    (inclusão x liberação pcp), lead time operação (inclusão x real
    entrega), otd, valor pedido, frete, estado, qualidade, semanal
    planejamento pcp"; ampliada no mesmo dia com "nº NF, data emissão NF,
    status pedido (Pedido Recebido/Produção/Inspeção-Expedição/Em
    transporte/Entrega Realizada), com emoji e descrição no hover"; e de
    novo, ainda no mesmo dia, com "expectativa PCP" entre Data solicitada e
    Conclusão produção (Previsão de liberação do PCP — a data PREVISTA,
    complementando a Conclusão produção que é a data REAL). Reúne,
    por PedidoOperacao.id, tudo que a nova tabela precisa além do que o
    próprio objeto já expõe — reaproveitando os MESMOS dados ao vivo de
    Produção já usados no resto de Gestão Operação (liberação PCP, data do
    cliente — ver _liberacao_pcp_por_pedido_venda/_data_cliente_por_pedido_
    venda), sem nenhuma sincronização nova.

    "Real entrega"/lead time operação usam go_data_entregue_cliente (seção
    Resultados/OTD) — a MESMA data que já alimenta o OTD e a property
    go_lead_time_operacao_dias do próprio modelo — não go_data_real_entrega
    (Logística/NF), que é um passo anterior no processo (confirmação de
    coleta/entrega pela transportadora, não necessariamente o recebimento
    pelo cliente).

    Cada lead time é em dias corridos, sempre a partir de Data de inclusão
    (campo próprio de PedidoOperacao — mesma referência que a coluna "Data
    inclusão" já mostrava nesta tela antes); None quando falta uma das duas
    datas (o template mostra "—").

    "Status pedido" reaproveita EXATAMENTE a mesma lógica de 5 etapas já
    usada no painel "Acompanhamento do pedido" de Consulta Pedido
    (_indice_etapa_pedido/_ETAPAS_ACOMPANHAMENTO_PEDIDO) — pra nunca
    divergir do que aquela tela mostra pro mesmo pedido — só que aqui
    compactado num badge com emoji (a trilha visual completa de lá não cabe
    numa célula de tabela)."""
    metricas = {}
    for p in pedidos:
        chave = _normalizar_pedido_venda(p.pedido_venda)
        liberacao_p = liberacao_pcp_por_pedido_venda.get(chave) or {}
        data_cliente_p = data_cliente_por_pedido_venda.get(chave)
        pedido_producao = pedidos_producao_por_pedido_venda.get(chave)

        data_inclusao = p.data_inclusao_pedido
        solicitada = data_cliente_p or p.go_data_solicitada_entrega
        # "Expectativa PCP" (pedido do Bruno, 10/09/2026, entre "Data
        # solicitada" e "Conclusão produção"): mesma Previsão de liberação
        # do PCP já mostrada na tela PCP de Gestão Operação — ao vivo de
        # Gestão Produção (liberacao_prevista dos itens), com o campo
        # próprio de PedidoOperacao como fallback enquanto o pedido não foi
        # lançado em Produção. Data PREVISTA, não a real (essa é a
        # "Conclusão produção" logo ao lado).
        expectativa_pcp = liberacao_p.get("previsao") or p.go_previsao_liberacao_pcp
        conclusao_producao = liberacao_p.get("efetiva") or p.go_data_efetiva_liberacao_pcp
        termino_semanal = liberacao_p.get("termino_semanal") or p.go_termino_semanal_pcp

        etapa_idx = _indice_etapa_pedido(pedido_producao, p)
        etapa_base = _ETAPAS_ACOMPANHAMENTO_PEDIDO[etapa_idx - 1]

        metricas[p.id] = {
            "solicitada": solicitada,
            "solicitada_automatica": bool(data_cliente_p),
            "expectativa_pcp": expectativa_pcp,
            "expectativa_pcp_automatica": bool(liberacao_p.get("previsao")),
            "conclusao_producao": conclusao_producao,
            "conclusao_producao_automatica": bool(liberacao_p.get("efetiva")),
            "termino_semanal": termino_semanal,
            "termino_semanal_automatico": bool(liberacao_p.get("termino_semanal")),
            "lead_comercial_dias": (solicitada - data_inclusao).days if (data_inclusao and solicitada) else None,
            "lead_producao_dias": (conclusao_producao - data_inclusao).days if (data_inclusao and conclusao_producao) else None,
            "lead_operacao_dias": p.go_lead_time_operacao_dias,
            "status_pedido_idx": etapa_idx,
            "status_pedido_emoji": _ETAPA_EMOJI[etapa_idx - 1],
            "status_pedido_label": etapa_base["label"],
            "status_pedido_descricao": etapa_base["descricao"],
        }
    return metricas


def _painel_operacao_360(filtros, pedidos_filtrados, metricas_filtrados):
    """Quadrante "NFs emitidas no mês" + painel dinâmico de valores/
    faturamento/lead time médio (pedido do Bruno, 10/09/2026) — no lugar dos
    quadrantes semanais/mensais de PCP removidos desta tela (continuam só
    na Listagem Geral de Gestão Produção). "Status produção"/"OTD"/
    "Modalidade de frete" deixaram de ser quadrante nesta mesma tarefa
    (viraram listas multi-seleção no formulário de filtro, ver template —
    pedido do Bruno: "não quero como quadrante, quero em formato de
    listas... selecionar mais de uma opção"), por isso não têm mais
    contagem calculada aqui.

    Calculado sobre o conjunto TOTAL filtrado (`pedidos_filtrados`, a query
    inteira sem paginação) — não só a página atual — pra sempre refletir o
    total real, e reage a cada mudança de filtro (por isso "dinâmico").
    `metricas_filtrados` é o retorno de _metricas_operacao_360 pro MESMO
    conjunto — reaproveitado aqui pra não duplicar o cálculo de lead time
    (só soma o que já foi calculado lá)."""
    def link(**overrides):
        base = dict(filtros)
        base.update(overrides)
        return base

    hoje = date.today()
    primeiro_dia_mes = hoje.replace(day=1)
    ultimo_dia_mes = date(hoje.year, hoje.month, monthrange(hoje.year, hoje.month)[1])
    nfs_mes = 0

    valor_total, n_valor = 0.0, 0
    faturamento_total, n_faturamento = 0.0, 0
    soma_lead_comercial = soma_lead_producao = soma_lead_operacao = 0
    n_lead_comercial = n_lead_producao = n_lead_operacao = 0

    for p in pedidos_filtrados:
        m = metricas_filtrados.get(p.id, {})

        if p.go_data_emissao_nf and primeiro_dia_mes <= p.go_data_emissao_nf <= ultimo_dia_mes:
            nfs_mes += 1

        if p.go_valor_pedido_operacao is not None:
            valor_total += p.go_valor_pedido_operacao
            n_valor += 1
        if p.go_valor_nf_emitida is not None:
            faturamento_total += p.go_valor_nf_emitida
            n_faturamento += 1

        lc = m.get("lead_comercial_dias")
        if lc is not None:
            soma_lead_comercial += lc
            n_lead_comercial += 1
        lp = m.get("lead_producao_dias")
        if lp is not None:
            soma_lead_producao += lp
            n_lead_producao += 1
        lo = m.get("lead_operacao_dias")
        if lo is not None:
            soma_lead_operacao += lo
            n_lead_operacao += 1

    card_nf_mes = {
        "titulo": f"NFs emitidas em {MESES_PT_EXTENSO[hoje.month - 1]}",
        "total": nfs_mes,
        "ativo": filtros.get("nf_mes_atual") == "1",
        "filtros_link": link(nf_mes_atual="1"),
    }

    dinamico = {
        "total_pedidos": len(pedidos_filtrados),
        "valor_total": valor_total,
        "n_valor": n_valor,
        "faturamento_total": faturamento_total,
        "n_faturamento": n_faturamento,
        "lead_comercial_medio": round(soma_lead_comercial / n_lead_comercial, 1) if n_lead_comercial else None,
        "lead_producao_medio": round(soma_lead_producao / n_lead_producao, 1) if n_lead_producao else None,
        "lead_operacao_medio": round(soma_lead_operacao / n_lead_operacao, 1) if n_lead_operacao else None,
    }

    return {
        "nf_mes": card_nf_mes,
        "dinamico": dinamico,
    }


def _pedidos_kanban_expedicao():
    """Kanban Expedição, dentro de Logística/Expedição (pedido do Bruno,
    16/09/2026): "pedidos finalizados PCP e já aos cuidados da logística e
    parados na expedição... só sai do kanban quando o material for
    expedição". Reaproveita a MESMA régua de 5 etapas já usada em Consulta
    Pedido e na coluna "Status pedido" da Operação 360
    (_indice_etapa_pedido) — etapa 3 ("Inspeção / Expedição") já significa
    exatamente isso: produção finalizada (Pedido.status_producao ==
    "FINALIZADO" ou go_data_efetiva_liberacao_pcp preenchido) e AINDA sem
    go_data_pedido_expedido. Assim que go_data_pedido_expedido é
    preenchido a etapa vira 4 ("Em transporte") e o pedido sai do kanban
    sozinho — não precisa de nenhum campo/estado novo.

    Casamento com Produção pelo mesmo padrão de sempre (pedido_venda, trim,
    sem FK, nunca aproximado — ver _liberacao_pcp_por_pedido_venda/
    _pedidos_producao_por_pedido_venda)."""
    candidatos = PedidoOperacao.query.filter(PedidoOperacao.go_data_pedido_expedido.is_(None)).all()
    pedidos_venda = [go.pedido_venda for go in candidatos]
    liberacao_pcp = _liberacao_pcp_por_pedido_venda(pedidos_venda)
    pedidos_producao = _pedidos_producao_por_pedido_venda(pedidos_venda)

    linhas = []
    for go in candidatos:
        chave = _normalizar_pedido_venda(go.pedido_venda)
        pedido_producao = pedidos_producao.get(chave)
        if _indice_etapa_pedido(pedido_producao, go) != 3:
            continue
        conclusao = (liberacao_pcp.get(chave) or {}).get("efetiva") or go.go_data_efetiva_liberacao_pcp
        dias_esperando = (date.today() - conclusao).days if conclusao else None
        linhas.append(
            {
                "id": go.id,
                "pedido_venda": chave or go.pedido_venda,
                "cliente": go.cliente,
                "frete": go.frete,
                "estado": go.estado,
                "valor": go.go_valor_pedido_operacao,
                "conclusao_producao": conclusao,
                "dias_esperando": dias_esperando,
                "numero_nf": go.go_numero_nf,
                "transportadora": go.go_transportadora.nome if go.go_transportadora else None,
            }
        )
    # Quem espera há mais tempo primeiro; sem data de conclusão (caso raro,
    # pedido nunca lançado em Produção mas com liberação efetiva manual em
    # Operação) fica por último, não no topo.
    linhas.sort(key=lambda l: (l["dias_esperando"] is None, -(l["dias_esperando"] or 0)))
    return linhas


def _prazos_pedido(pedido, go, liberacao_pcp, data_cliente_producao):
    """Pedido do Bruno (09/09/2026, tela Consulta Pedido): 2 lead times em
    dias corridos, sempre calculados a partir da Data de inclusão do
    pedido — "lead time comercial" (inclusão -> data solicitada pelo
    cliente, o prazo de entrega combinado) e "lead time de operação
    completa" (inclusão -> liberação efetiva do PCP, quando o material de
    fato ficou pronto/liberado). Usa exatamente as mesmas fontes de dado já
    mostradas mais abaixo nessa tela (bloco Gestão Operação), com Produção
    tendo prioridade sobre o campo digitado à mão em Operação quando os
    dois existem — ver _liberacao_pcp_por_pedido_venda/_data_cliente_por_
    pedido_venda logo acima. Qualquer conta que não dá pra fazer (falta uma
    das duas datas) vem como None — o template mostra "—"."""
    data_inclusao = pedido.data_inclusao_pedido if pedido else (go.data_inclusao_pedido if go else None)
    data_solicitada_cliente = data_cliente_producao or (go.go_data_solicitada_cliente_retira if go else None)
    data_liberacao_efetiva = (liberacao_pcp or {}).get("efetiva") or (go.go_data_efetiva_liberacao_pcp if go else None)

    lead_time_comercial_dias = None
    if data_inclusao and data_solicitada_cliente:
        lead_time_comercial_dias = (data_solicitada_cliente - data_inclusao).days

    lead_time_operacao_dias = None
    if data_inclusao and data_liberacao_efetiva:
        lead_time_operacao_dias = (data_liberacao_efetiva - data_inclusao).days

    return {
        "data_inclusao": data_inclusao,
        "data_solicitada_cliente": data_solicitada_cliente,
        "data_liberacao_efetiva": data_liberacao_efetiva,
        "lead_time_comercial_dias": lead_time_comercial_dias,
        "lead_time_operacao_dias": lead_time_operacao_dias,
    }


def _buscar_pedidos_para_status(termo, limite=12):
    """Pedido do Bruno (01/09/2026): canal único de busca no topo do Painel —
    "sou o PCP, comercial me cobrou de um pedido" — digita o nº do pedido de
    venda OU o cliente e recebe sugestões pra abrir o status completo
    (produção + operação) daquele pedido específico.

    Casa em Pedido (Gestão Produção) e em PedidoOperacao (Gestão Operação)
    separadamente — cada tabela pode ter pedidos que a outra não tem — e
    devolve no máximo 1 sugestão por pedido_venda (nunca duplicada), usando
    o próprio pedido_venda como identificador da busca de detalhe. Por isso
    só considera pedidos com pedido_venda preenchido: sem esse número não dá
    pra cruzar as duas tabelas de forma confiável — mesma regra de "match
    exato, nunca aproximado" já usada em toda outra cross-referência do
    sistema (ver _itens_producao_por_pedido_venda)."""
    termo = (termo or "").strip()
    if not termo:
        return []
    padrao = f"%{termo}%"

    candidatos = {}  # pedido_venda (trim) -> dict de exibição

    pedidos = (
        Pedido.query.filter(
            Pedido.pedido_venda.isnot(None),
            _pedido_venda_normalizado_sql(Pedido.pedido_venda) != "",
            or_(Pedido.pedido_venda.ilike(padrao), Pedido.cliente.ilike(padrao)),
        )
        .order_by(Pedido.data_inclusao_pedido.desc().nullslast())
        .limit(limite * 3)
        .all()
    )
    for p in pedidos:
        chave = _normalizar_pedido_venda(p.pedido_venda)
        if not chave or chave in candidatos:
            continue
        candidatos[chave] = {"pedido_venda": chave, "cliente": p.cliente, "tem_producao": True}

    if len(candidatos) < limite:
        pedidos_operacao = (
            PedidoOperacao.query.filter(
                PedidoOperacao.pedido_venda.isnot(None),
                _pedido_venda_normalizado_sql(PedidoOperacao.pedido_venda) != "",
                or_(PedidoOperacao.pedido_venda.ilike(padrao), PedidoOperacao.cliente.ilike(padrao)),
            )
            .order_by(PedidoOperacao.data_inclusao_pedido.desc().nullslast())
            .limit(limite * 3)
            .all()
        )
        for go in pedidos_operacao:
            chave = _normalizar_pedido_venda(go.pedido_venda)
            if not chave or chave in candidatos:
                continue
            candidatos[chave] = {"pedido_venda": chave, "cliente": go.cliente, "tem_producao": False}

    termo_upper = termo.upper()

    def _relevancia(c):
        pv = c["pedido_venda"].upper()
        cli = (c["cliente"] or "").upper()
        if pv == termo_upper:
            return (0, pv)
        if pv.startswith(termo_upper):
            return (1, pv)
        if termo_upper in pv:
            return (2, pv)
        if cli.startswith(termo_upper):
            return (3, cli)
        return (4, cli)

    return sorted(candidatos.values(), key=_relevancia)[:limite]


def _situacao_entrega_go(go, pedido=None):
    """Resumo em UMA frase só do que mais se pergunta pro PCP/Comercial: "cadê
    esse pedido? já foi expedido? já chegou no cliente?" — a informação
    "principal" que o Bruno pediu (pedido de 01/09/2026), consultando a aba
    Expedição/Logística de Gestão Operação. Prioriza a data de entrega mais
    confiável entre as duas que a Operação guarda (Logística x
    Resultados/OTD — historicamente nem sempre as duas são preenchidas
    juntas)."""
    if go is None:
        if pedido is not None and pedido.status_producao == "FINALIZADO":
            return {
                "texto": "Produção finalizada, mas o pedido ainda não foi lançado em Gestão Operação — sem dado de expedição.",
                "cor": "warning",
                "icone": "bi-exclamation-triangle",
            }
        return {
            "texto": "Pedido ainda não lançado em Gestão Operação — sem dados de expedição/logística.",
            "cor": "secondary",
            "icone": "bi-question-circle",
        }

    data_entrega = go.go_data_entregue_cliente or go.go_data_real_entrega
    if data_entrega:
        return {
            "texto": f"Entregue ao cliente em {data_entrega.strftime('%d/%m/%Y')}",
            "cor": "success",
            "icone": "bi-check-circle-fill",
        }
    if go.go_data_pedido_expedido:
        return {
            "texto": f"Expedido em {go.go_data_pedido_expedido.strftime('%d/%m/%Y')} — aguardando confirmação de entrega/coleta",
            "cor": "info",
            "icone": "bi-truck",
        }
    if go.go_data_efetiva_liberacao_pcp:
        return {
            "texto": f"Liberado pelo PCP em {go.go_data_efetiva_liberacao_pcp.strftime('%d/%m/%Y')} — ainda não expedido",
            "cor": "warning",
            "icone": "bi-hourglass-split",
        }
    return {"texto": "Ainda não expedido.", "cor": "secondary", "icone": "bi-hourglass"}


def _otd_do_pedido(go):
    """OTD (On-Time Delivery) de UM pedido específico — pedido do Bruno
    (03/09/2026, tela Consulta Pedido): "incluir o OTD do pedido, se atendeu
    ou não, bem didático e informativo", ao lado da "Situação de entrega".
    Mesma fonte de verdade da tela Resultados/OTD (go.go_otd_realizado,
    preenchido manualmente na aba Resultados/OTD de Gestão Operação) — nunca
    recalculado por conta própria a partir de datas, pra não divergir do
    número que já aparece agregado em Resultados/OTD.

    go_dias_atraso_antecipacao (solicitado x entregue de verdade) entra só
    como detalhe complementar, quando disponível — o "atendeu ou não" em si
    sempre vem do campo manual."""
    if go is None:
        return {
            "texto": "OTD não disponível — pedido ainda não lançado em Gestão Operação.",
            "cor": "secondary",
            "icone": "bi-question-circle",
        }
    if not go.go_otd_realizado:
        return {
            "texto": "OTD ainda não registrado para este pedido.",
            "cor": "secondary",
            "icone": "bi-hourglass",
        }

    atraso = go.go_dias_atraso_antecipacao
    if atraso is None:
        detalhe = ""
    elif atraso > 0:
        detalhe = f" — entregue {atraso} dia(s) após o prazo solicitado."
    elif atraso < 0:
        detalhe = f" — entregue {abs(atraso)} dia(s) antes do prazo solicitado."
    else:
        detalhe = " — entregue exatamente no dia solicitado."

    if go.go_otd_realizado == "SIM":
        return {
            "texto": f"Atendeu o OTD (dentro do prazo solicitado){detalhe}",
            "cor": "success",
            "icone": "bi-check-circle-fill",
        }
    return {
        "texto": f"Não atendeu o OTD (fora do prazo solicitado){detalhe}",
        "cor": "danger",
        "icone": "bi-x-circle-fill",
    }


# "Acompanhamento do pedido" — trilha de 5 etapas (recebido -> produção ->
# inspeção/expedição -> transporte -> entrega) no painel de status de um
# pedido. Pedido do Bruno (01/09/2026), a partir de uma imagem de referência
# que ele anexou (infográfico com o mesmo formato/rótulos), além do resumo
# em uma frase que já existia (_situacao_entrega_go) — este aqui é o
# complemento visual "em que pé exatamente está".
_ETAPAS_ACOMPANHAMENTO_PEDIDO = [
    {
        # Rótulo trocado de "Pedido recebido" pra "Pendente produção" a
        # pedido do Bruno (10/09/2026, junto com a troca dos quadrantes de
        # Status pedidos da Operação 360 por uma lista multi-seleção) — vale
        # em TODO lugar que reaproveita esta etapa (Consulta Pedido, coluna
        # "Status pedido" da Operação 360, filtro "Status produção"), de
        # propósito, pra nunca divergir.
        "label": "Pendente produção",
        "descricao": "Pedido registrado, aguardando início da produção.",
        "icone": "bi-receipt",
        "cor": "#0d6efd",
        "cor_fraca": "rgba(13, 110, 253, .18)",
    },
    {
        "label": "Produção",
        "descricao": "Pedido em produção. Materiais separados e processos em execução.",
        "icone": "bi-gear-wide-connected",
        "cor": "#12b886",
        "cor_fraca": "rgba(18, 184, 134, .18)",
    },
    {
        "label": "Inspeção / Expedição",
        "descricao": "Pedido finalizado. Inspeção realizada e liberado para expedição.",
        "icone": "bi-clipboard2-check",
        "cor": "#f59f00",
        "cor_fraca": "rgba(245, 159, 0, .18)",
    },
    {
        "label": "Em transporte",
        "descricao": "Pedido coletado e em transporte até o destino final.",
        "icone": "bi-truck",
        "cor": "#7048e8",
        "cor_fraca": "rgba(112, 72, 232, .18)",
    },
    {
        "label": "Entrega realizada",
        "descricao": "Pedido entregue ao cliente com sucesso.",
        "icone": "bi-box-seam",
        "cor": "#198754",
        "cor_fraca": "rgba(25, 135, 84, .18)",
    },
]


def _indice_etapa_pedido(pedido, go):
    """Em que das 5 etapas do "Acompanhamento do pedido" ele está agora.
    Mesmos sinais já usados em _situacao_entrega_go, só que granulares em 5
    passos em vez de só "expedido/entregue" — chamada só quando pedido ou go
    existem (nunca os dois None), por isso sempre devolve pelo menos 1
    ("Pedido recebido")."""
    if go is not None and (go.go_data_entregue_cliente or go.go_data_real_entrega):
        return 5
    if go is not None and go.go_data_pedido_expedido:
        return 4
    producao_finalizada = pedido is not None and pedido.status_producao == "FINALIZADO"
    if producao_finalizada or (go is not None and go.go_data_efetiva_liberacao_pcp):
        return 3
    em_producao = pedido is not None and any(item.inicio_producao for item in pedido.itens)
    if em_producao:
        return 2
    return 1


def _etapas_acompanhamento_pedido(pedido, go):
    etapa_atual = _indice_etapa_pedido(pedido, go)
    etapas = []
    for i, base in enumerate(_ETAPAS_ACOMPANHAMENTO_PEDIDO, start=1):
        if i < etapa_atual:
            estado = "concluida"
        elif i == etapa_atual:
            estado = "atual"
        else:
            estado = "pendente"
        etapas.append({**base, "numero": i, "estado": estado, "linha_concluida": i <= etapa_atual})
    return etapas


# Campos "comercial" que só existem em PedidoOperacao (não têm equivalente em
# Pedido) — pedido do Bruno, 01/09/2026: quem inclui um pedido novo em Gestão
# Produção (PCP) passa a preencher também essas informações do pedido como um
# todo (não só do produto), e ao salvar isso alimenta automaticamente a
# Listagem Geral de Gestão Operação.
CAMPOS_GO_COMERCIAL_NOVO_PEDIDO = [
    "go_tipo_pedido", "go_contrato", "go_proposta", "go_pedido_compra_cliente",
    "go_status_pedido_info", "go_valor_pedido_operacao",
]


def _criar_pedido_operacao_a_partir_de_producao(pedido, f):
    """Cria automaticamente 1 PedidoOperacao a partir de um Pedido recém-
    incluído em Gestão Produção (pedido do Bruno, 01/09/2026) — SÓ na
    inclusão (não em edição futura), e SÓ isso: um cópia inicial dos campos,
    não um vínculo permanente. PedidoOperacao continua sendo uma tabela
    INDEPENDENTE (sem FK) — dali em diante os dois são editados cada um na
    sua própria tela, sem sincronização automática nenhuma (é o próprio
    Bruno quem vai "manusear manualmente entre PCP/Logística/Resultados").

    "Data solicitada entrega" (Comercial) E "Solicitada cliente/retira" (PCP)
    de Gestão Operação são o mesmo dado que "Data do cliente" de Gestão
    Produção (já documentado em _LinhaListagemGeral) — pedido do Bruno
    (01/09/2026): "Solicitada cliente/retira" também tem que acompanhar essa
    data desde a inclusão do pedido, não é um campo novo — por isso não
    existe um campo novo pra ela no formulário de Produção, só reaproveita
    pedido.data_cliente pros dois."""
    novo = PedidoOperacao(
        cliente=pedido.cliente,
        vendedor=pedido.vendedor,
        pedido_venda=pedido.pedido_venda,
        data_inclusao_pedido=pedido.data_inclusao_pedido,
        prioridade=pedido.prioridade,
        frete=pedido.frete,
        pais=pedido.pais,
        estado=pedido.estado,
        cidade=pedido.cidade,
        go_data_solicitada_entrega=pedido.data_cliente,
        go_data_solicitada_cliente_retira=pedido.data_cliente,
        go_tipo_pedido=f.get("go_tipo_pedido", "").strip() or None,
        go_contrato=f.get("go_contrato", "").strip() or None,
        go_proposta=f.get("go_proposta", "").strip() or None,
        go_pedido_compra_cliente=f.get("go_pedido_compra_cliente", "").strip() or None,
        go_status_pedido_info=f.get("go_status_pedido_info", "").strip() or None,
        go_valor_pedido_operacao=(
            _parse_float_form(f.get("go_valor_pedido_operacao"), default=None)
            if f.get("go_valor_pedido_operacao", "").strip()
            else None
        ),
    )
    db.session.add(novo)
    return novo


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
# Qualidade — Inspeção Final / RDIM (pedido do Bruno, 02/09/2026). Ao
# contrário da RNC, usa FK real pra ItemPedido — "OP" nesta base é o próprio
# ItemPedido. Mesmo padrão de filtro/paginação de _filtrar_rnc_qualidade.
# ----------------------------------------------------------------------
CAMPOS_HISTORICO_INSPECAO_FINAL = [
    "resultado", "categoria_desvio", "subcategoria_desvio", "desvio_encontrado",
    "observacao", "inspecao_visual", "quantidade_com_desvio", "tipo_produto_inspecionado",
]


def _itens_rdim_disponiveis(termo, limite=20):
    """Itens de pedido inspecionáveis pelo RDIM — busca em TODO o banco de
    Gestão Produção, sem restrição de estação (pedido do Bruno, 03/09/2026:
    "quero que os campos qualidade puxe todo o banco de dados da produção" —
    ele foi apontar um desvio no RDIM e a busca da OP não encontrou o
    pedido, porque a busca antiga só olhava itens das estações MANDRIL/PU/
    SILICONE. RDIM_ESTACOES_OPCOES continua existindo só como ordem de
    exibição no dashboard/filtro — não é mais um filtro de elegibilidade).
    Busca por descrição do produto, cliente ou nº do pedido de venda, igual
    ao padrão já usado em _buscar_pedidos_para_status."""
    termo = (termo or "").strip()
    query = ItemPedido.query.join(Pedido)
    if termo:
        like = f"%{termo}%"
        query = query.filter(
            or_(
                ItemPedido.descricao_produto.ilike(like),
                Pedido.cliente.ilike(like),
                Pedido.pedido_venda.ilike(like),
            )
        )
    itens = query.order_by(ItemPedido.atualizado_em.desc()).limit(limite).all()
    return [
        {
            "item_id": item.id,
            "cliente": item.pedido.cliente if item.pedido else "",
            "pedido_venda": item.pedido.pedido_venda if item.pedido else "",
            "produto": item.descricao_produto,
            "quantidade": item.quantidade,
            "estacao": item.estacao,
        }
        for item in itens
    ]


def _filtrar_inspecoes_finais(args):
    """`args` é sempre `request.args`. Mesmo espírito de _filtrar_rnc_qualidade:
    multi-seleção (Ctrl+clique) pra resultado/categoria de desvio/estação,
    texto livre pra cliente/produto/pedido/OP, período por data de inspeção."""
    query = InspecaoFinal.query.join(ItemPedido).join(Pedido)

    busca = args.get("busca", "").strip()
    dn_polegada = args.get("dn_polegada", "").strip()
    resultado = [v for v in args.getlist("resultado") if v]
    categoria_desvio = [v for v in args.getlist("categoria_desvio") if v]
    subcategoria_desvio = [v for v in args.getlist("subcategoria_desvio") if v]
    estacao = [v for v in args.getlist("estacao") if v]
    responsavel_id = args.get("responsavel_id", "").strip()
    data_de = args.get("data_de", "").strip()
    data_ate = args.get("data_ate", "").strip()
    mes = args.get("mes", "").strip()  # "AAAA-MM", do <input type="month"> — mesmo padrão de mes_emissao (RNC)

    if busca:
        like = f"%{busca}%"
        query = query.filter(
            or_(
                Pedido.cliente.ilike(like),
                Pedido.pedido_venda.ilike(like),
                ItemPedido.descricao_produto.ilike(like),
                InspecaoFinal.numero_rif.ilike(like),
            )
        )
    if dn_polegada:
        # Filtro separado do "busca" geral (pedido do Bruno, 02/09/2026,
        # RDIM Fase 4) — não existe campo estruturado de DN/polegada no
        # sistema, a medida do produto vem embutida em texto livre dentro de
        # descricao_produto (ex.: "DISCO SELO 18''"). Casamento só contra a
        # descrição do produto, de propósito: se entrasse no OR do "busca"
        # geral, um número de pedido ou RIF que por coincidência contivesse
        # os mesmos dígitos do DN daria falso positivo.
        query = query.filter(ItemPedido.descricao_produto.ilike(f"%{dn_polegada}%"))
    if resultado:
        query = query.filter(InspecaoFinal.resultado.in_(resultado))
    if categoria_desvio:
        query = query.filter(InspecaoFinal.categoria_desvio.in_(categoria_desvio))
    if subcategoria_desvio:
        query = query.filter(InspecaoFinal.subcategoria_desvio.in_(subcategoria_desvio))
    if estacao:
        query = query.filter(InspecaoFinal.estacao.in_(estacao))
    if responsavel_id:
        try:
            query = query.filter(InspecaoFinal.responsavel_id == int(responsavel_id))
        except ValueError:
            responsavel_id = ""
    if data_de:
        try:
            query = query.filter(InspecaoFinal.data_inspecao >= date.fromisoformat(data_de))
        except ValueError:
            data_de = ""
    if data_ate:
        try:
            query = query.filter(InspecaoFinal.data_inspecao <= date.fromisoformat(data_ate))
        except ValueError:
            data_ate = ""
    if mes:
        try:
            ano_m, mes_m = (int(p) for p in mes.split("-"))
            inicio = date(ano_m, mes_m, 1)
            fim = date(ano_m, mes_m, monthrange(ano_m, mes_m)[1])
            query = query.filter(InspecaoFinal.data_inspecao.between(inicio, fim))
        except (ValueError, TypeError):
            mes = ""  # valor incompreensível — ignora o filtro em vez de quebrar a busca

    query = query.order_by(InspecaoFinal.data_inspecao.desc().nullslast(), InspecaoFinal.id.desc())

    filtros = dict(
        busca=busca, dn_polegada=dn_polegada, resultado=resultado, categoria_desvio=categoria_desvio,
        subcategoria_desvio=subcategoria_desvio, estacao=estacao,
        responsavel_id=responsavel_id, data_de=data_de, data_ate=data_ate, mes=mes,
    )
    return query, filtros


def _linhas_inspecoes_finais(args):
    page = args.get("page", 1, type=int)
    query, filtros = _filtrar_inspecoes_finais(args)
    total_filtrado = query.count()
    total_paginas = max(1, (total_filtrado + PAGE_SIZE - 1) // PAGE_SIZE)
    pagina = query.offset((page - 1) * PAGE_SIZE).limit(PAGE_SIZE).all()
    return pagina, page, total_paginas, total_filtrado, filtros


def _dashboard_rdim():
    """Recalcula ao vivo os indicadores da Inspeção Final — mesmo espírito de
    _dashboard_rnc_qualidade, nada pré-calculado/guardado."""
    inspecoes = InspecaoFinal.query.options(selectinload(InspecaoFinal.item).selectinload(ItemPedido.pedido)).all()
    total = len(inspecoes)

    aprovadas = sum(1 for i in inspecoes if i.resultado == "APROVADO")
    reprovadas = sum(1 for i in inspecoes if i.resultado == "REPROVADO")
    aprovadas_desvio = sum(1 for i in inspecoes if i.resultado == "APROVADO_COM_DESVIO")
    pct_aprovacao = round((aprovadas / total) * 100, 1) if total else 0
    pct_reprovacao = round((reprovadas / total) * 100, 1) if total else 0

    # Volume de peças (não só de inspeções) — pedido do Bruno (02/09/2026):
    # "lote total contém 5 peças, mas dessas 2 unidades ficou com desvio".
    # Sempre em relação ao lote inteiro (item.quantidade), decisão já
    # confirmada com ele.
    total_quantidade_inspecionada = sum((i.item.quantidade or 0) for i in inspecoes if i.item)
    total_quantidade_com_desvio = sum(i.quantidade_com_desvio_total or 0 for i in inspecoes)
    pct_pecas_desvio = (
        round((total_quantidade_com_desvio / total_quantidade_inspecionada) * 100, 1)
        if total_quantidade_inspecionada else 0
    )

    def _quebra_por(chave_fn, opcoes_ordem=None):
        contagem = {}
        for i in inspecoes:
            chave = chave_fn(i) or "—"
            contagem[chave] = contagem.get(chave, 0) + 1
        if opcoes_ordem:
            chaves = list(opcoes_ordem) + sorted(k for k in contagem if k not in opcoes_ordem and k != "—")
            if "—" in contagem:
                chaves.append("—")
            return [{"chave": k, "total": contagem.get(k, 0)} for k in chaves if k in contagem]
        return sorted(({"chave": k, "total": v} for k, v in contagem.items()), key=lambda d: -d["total"])[:10]

    # Produtos/clientes com maior índice de REPROVAÇÃO (não só volume) — só
    # entram no ranking quem já teve ao menos 1 inspeção reprovada.
    def _ranking_reprovacao(chave_fn):
        por_chave = {}
        for i in inspecoes:
            chave = chave_fn(i)
            if not chave:
                continue
            d = por_chave.setdefault(chave, {"total": 0, "reprovadas": 0})
            d["total"] += 1
            if i.resultado == "REPROVADO":
                d["reprovadas"] += 1
        linhas = [
            {"chave": k, "total": v["total"], "reprovadas": v["reprovadas"],
             "pct": round((v["reprovadas"] / v["total"]) * 100, 1)}
            for k, v in por_chave.items() if v["reprovadas"] > 0
        ]
        return sorted(linhas, key=lambda d: (-d["reprovadas"], -d["pct"]))[:10]

    # Evolução mensal de reprovações — últimos 12 meses com pelo menos uma
    # inspeção, em ordem cronológica (pro gráfico de linha).
    por_mes = {}
    for i in inspecoes:
        if not i.data_inspecao:
            continue
        chave = i.data_inspecao.strftime("%m/%Y")
        d = por_mes.setdefault(chave, {"ord": i.data_inspecao.replace(day=1), "total": 0, "reprovadas": 0})
        d["total"] += 1
        if i.resultado == "REPROVADO":
            d["reprovadas"] += 1
    evolucao = [
        {"mes": k, "total": v["total"], "reprovadas": v["reprovadas"]}
        for k, v in sorted(por_mes.items(), key=lambda kv: kv[1]["ord"])
    ][-12:]

    # Componentes do PIG LBD/LUN/SUPERFLEX com mais desvio — pedido do Bruno
    # (11/09/2026, RDIM Fase 5): mesmo espírito de "principais características
    # do desvio" (por_subcategoria_desvio), só que na granularidade de
    # componente do conjunto, contando RdimComponenteDesvio (1 linha por
    # componente apontado, não por inspeção).
    contagem_componentes = {}
    for i in inspecoes:
        for c in i.componentes_desvio:
            contagem_componentes[c.componente] = contagem_componentes.get(c.componente, 0) + 1
    por_componente_desvio = [
        {"chave": k, "total": contagem_componentes.get(k, 0)}
        for k in RDIM_COMPONENTE_LBD_OPCOES if k in contagem_componentes
    ]

    return {
        "total": total,
        "aprovadas": aprovadas,
        "reprovadas": reprovadas,
        "aprovadas_desvio": aprovadas_desvio,
        "pct_aprovacao": pct_aprovacao,
        "pct_reprovacao": pct_reprovacao,
        "por_categoria_desvio": _quebra_por(lambda i: i.categoria_desvio, RDIM_CATEGORIA_DESVIO_OPCOES),
        "por_subcategoria_desvio": _quebra_por(lambda i: i.subcategoria_desvio, RDIM_SUBCATEGORIA_DESVIO_OPCOES),
        "por_estacao": _quebra_por(lambda i: i.estacao, RDIM_ESTACOES_OPCOES),
        "por_componente_desvio": por_componente_desvio,
        "ranking_produtos": _ranking_reprovacao(lambda i: i.produto),
        "ranking_clientes": _ranking_reprovacao(lambda i: i.cliente),
        "evolucao": evolucao,
        "total_quantidade_inspecionada": total_quantidade_inspecionada,
        "total_quantidade_com_desvio": total_quantidade_com_desvio,
        "pct_pecas_desvio": pct_pecas_desvio,
    }


def _validar_quantidade_com_desvio(valor_form, quantidade_item):
    """Lê o campo "Quantidade com desvio" do formulário de inspeção RDIM e
    valida contra a quantidade do lote (item.quantidade) — pedido do Bruno
    (02/09/2026): "lote total contém 5 peças, mas dessas 2 unidades ficou
    com desvio", sempre em relação ao lote inteiro. Retorna (valor, erro) —
    `erro` é None quando válido; quando não, a rota deve mostrar o flash e
    NÃO salvar (mesmo padrão de bloqueio já usado pra resultado/estação).

    `default=None` explícito no _parse_float_form: o default 0.0 da função
    faria todo campo em branco virar "0 peças com desvio" em vez de "não
    informado" — mesma armadilha já corrigida uma vez nas medições."""
    quantidade_com_desvio = _parse_float_form(valor_form, default=None)
    if quantidade_com_desvio is None:
        return None, None
    if quantidade_com_desvio < 0:
        return None, "Quantidade com desvio não pode ser negativa."
    if quantidade_item is not None and quantidade_com_desvio > quantidade_item:
        return None, (
            f"Quantidade com desvio ({quantidade_com_desvio:g}) não pode ser maior que "
            f"a quantidade do lote ({quantidade_item:g})."
        )
    return quantidade_com_desvio, None


def _salvar_medicoes_rdim(inspecao, f, substituir=False):
    """Grava as linhas de medição (grandeza + especificação + faixa medida)
    a partir dos arrays paralelos do formulário — mesmo padrão de arrays
    paralelos (zip) já usado pros itens de um pedido em editar_pedido. Em
    edição, `substituir=True` apaga as medições antigas e recria do zero: é
    mais simples que casar por id e cobre bem o caso comum (poucas linhas,
    reescritas inteiras a cada salvamento) sem precisar de histórico por
    medição — só a InspecaoFinal como um todo entra no histórico de alteração."""
    if substituir:
        for m in list(inspecao.medicoes):
            db.session.delete(m)

    grandezas = f.getlist("grandeza[]")
    esp_mins = f.getlist("especificado_min[]")
    esp_maxs = f.getlist("especificado_max[]")
    med_mins = f.getlist("medido_min[]")
    med_maxs = f.getlist("medido_max[]")

    ordem = 0
    for grandeza, esp_min, esp_max, med_min, med_max in zip(grandezas, esp_mins, esp_maxs, med_mins, med_maxs):
        grandeza = grandeza.strip()
        if not grandeza:
            continue
        db.session.add(
            RdimMedicao(
                inspecao=inspecao,
                grandeza=grandeza,
                especificado_min=_parse_float_form(esp_min, default=None),
                especificado_max=_parse_float_form(esp_max, default=None),
                medido_min=_parse_float_form(med_min, default=None),
                medido_max=_parse_float_form(med_max, default=None),
                ordem=ordem,
            )
        )
        ordem += 1


def _salvar_pecas_desvio_rdim(inspecao, f, substituir=False):
    """Grava o detalhamento peça a peça do desvio (nº da peça + característica
    + valor medido + especificação mín/máx daquela peça) — pedido do Bruno
    (02/09/2026, RDIM Fase 3 e, pros campos especificado_min/max, Fase 4:
    "tolerância era de 0,5mm, peça inspecionada com 0,7mm, peça ficou 0,2mm
    acima da tolerância"), mesmo padrão de arrays paralelos (zip) já usado em
    _salvar_medicoes_rdim. Linhas sem característica selecionada são
    ignoradas (o nº da peça sozinho não basta pra fazer sentido). Não mexe em
    quantidade_com_desvio — os dois campos são independentes, por decisão do
    Bruno."""
    if substituir:
        for p in list(inspecao.pecas_desvio):
            db.session.delete(p)

    pecas_numero = f.getlist("peca_numero[]")
    pecas_caracteristica = f.getlist("peca_caracteristica[]")
    pecas_valor = f.getlist("peca_valor_medido[]")
    pecas_espec_min = f.getlist("peca_especificado_min[]")
    pecas_espec_max = f.getlist("peca_especificado_max[]")

    ordem = 0
    for peca_numero, caracteristica, valor_medido, espec_min, espec_max in zip_longest(
        pecas_numero, pecas_caracteristica, pecas_valor, pecas_espec_min, pecas_espec_max, fillvalue=""
    ):
        caracteristica = (caracteristica or "").strip()
        if not caracteristica:
            continue
        db.session.add(
            RdimPecaDesvio(
                inspecao=inspecao,
                peca_numero=(peca_numero or "").strip() or None,
                caracteristica=caracteristica,
                valor_medido=_parse_float_form(valor_medido, default=None),
                especificado_min=_parse_float_form(espec_min, default=None),
                especificado_max=_parse_float_form(espec_max, default=None),
                ordem=ordem,
            )
        )
        ordem += 1


def _salvar_componentes_desvio_rdim(inspecao, f, quantidade_item, substituir=False):
    """Grava o apontamento de desvio por componente do PIG LBD/LUN/SUPERFLEX
    — pedido do Bruno (11/09/2026, RDIM Fase 5): "o desvio se encontra
    somente no disco selo e não no disco guia... quero a possibilidade de
    inserir para a inspeção e apontamento somente o componente em
    específico". Diferente das medições/peças (arrays paralelos por índice
    de linha), aqui cada um dos 11 componentes de RDIM_COMPONENTE_LBD_OPCOES
    tem um campo próprio indexado pela posição dele na lista (mais simples e
    à prova de desalinhamento do que arrays paralelos, já que a lista de
    componentes é FIXA — não é uma lista dinâmica que o operador monta livre
    como peça a peça). Só cria linha pro componente marcado "Sim"
    (`componente_desvio_{i}` == "SIM"); os demais não geram registro, mesmo
    espírito de _salvar_pecas_desvio_rdim (só quem teve desvio entra).

    Retorna None em sucesso, ou uma mensagem de erro (validação de
    quantidade_com_desvio por componente, mesma regra do campo do lote
    inteiro) — quando há erro, NADA é salvo (a rota deve mostrar o flash e
    não commitar, mesmo padrão de bloqueio já usado no resto do RDIM)."""
    linhas = []
    for i, componente in enumerate(RDIM_COMPONENTE_LBD_OPCOES):
        if (f.get(f"componente_desvio_{i}", "") or "").strip().upper() != "SIM":
            continue
        qtd, erro_qtd = _validar_quantidade_com_desvio(f.get(f"componente_quantidade_com_desvio_{i}"), quantidade_item)
        if erro_qtd:
            return f'Componente "{componente}": {erro_qtd}'
        linhas.append({
            "componente": componente,
            "categoria_desvio": (f.get(f"componente_categoria_desvio_{i}", "") or "").strip() or None,
            "subcategoria_desvio": (f.get(f"componente_subcategoria_desvio_{i}", "") or "").strip() or None,
            "quantidade_com_desvio": qtd,
            "desvio_encontrado": (f.get(f"componente_desvio_encontrado_{i}", "") or "").strip() or None,
        })

    if substituir:
        for c in list(inspecao.componentes_desvio):
            db.session.delete(c)

    for ordem, linha in enumerate(linhas):
        db.session.add(RdimComponenteDesvio(inspecao=inspecao, ordem=ordem, **linha))
    return None


def _inspecoes_rdim_por_item(item_ids):
    """dict item_pedido_id -> InspecaoFinal mais recente (data_inspecao
    desc, depois id desc) — 1 query com IN, sem N+1. Pedido do Bruno
    (02/09/2026): conectar o RDIM com Produção/Operação, mostrando o status
    de qualidade de cada item nas telas que já existem (Consulta Pedido,
    Detalhe do Pedido, Listagem Geral) sem duplicar a lógica em cada uma.
    Se um item tiver mais de uma inspeção, mostra só a mais recente — a
    tabela InspecaoFinal.query já vem ordenada, `setdefault` fica só com a
    primeira ocorrência de cada item_pedido_id."""
    ids = sorted({i for i in item_ids if i})
    if not ids:
        return {}
    inspecoes = (
        InspecaoFinal.query
        .filter(InspecaoFinal.item_pedido_id.in_(ids))
        .order_by(InspecaoFinal.item_pedido_id, InspecaoFinal.data_inspecao.desc().nullslast(), InspecaoFinal.id.desc())
        .all()
    )
    mapa = {}
    for insp in inspecoes:
        mapa.setdefault(insp.item_pedido_id, insp)
    return mapa


def _resumo_rdim_pedido(inspecoes):
    """Agrega as InspecaoFinal de UM pedido (normalmente vindas de
    _inspecoes_rdim_por_item(...).values()) em contagens por resultado + o
    quantitativo de peças com desvio/total — usado no bloco de resumo de
    Qualidade da Consulta Pedido e do Detalhe do Pedido. Retorna None se a
    lista vier vazia (pedido sem nenhuma inspeção RDIM ainda)."""
    inspecoes = list(inspecoes)
    if not inspecoes:
        return None
    return {
        "total": len(inspecoes),
        "aprovadas": sum(1 for i in inspecoes if i.resultado == "APROVADO"),
        "reprovadas": sum(1 for i in inspecoes if i.resultado == "REPROVADO"),
        "aprovadas_desvio": sum(1 for i in inspecoes if i.resultado == "APROVADO_COM_DESVIO"),
        "quantidade_com_desvio": sum(i.quantidade_com_desvio or 0 for i in inspecoes),
        "quantidade_total": sum((i.item.quantidade or 0) for i in inspecoes if i.item),
    }


def _rdim_resumo_por_pedido_venda(pedidos_venda):
    """dict pedido_venda (trim) -> {total, reprovadas, com_desvio,
    quantidade_com_desvio} — pro indicador (mais simples, pedido-level) de
    Qualidade na Listagem Geral de Gestão Operação. Mesmo casamento por
    texto (sem FK, trim() nos dois lados) já usado por
    _itens_producao_por_pedido_venda — só pra exibir uma dica visual, nunca
    aproximado."""
    valores = sorted({_normalizar_pedido_venda(v) for v in pedidos_venda if v and v.strip()})
    if not valores:
        return {}
    inspecoes = (
        InspecaoFinal.query.join(ItemPedido).join(Pedido)
        .options(selectinload(InspecaoFinal.item).selectinload(ItemPedido.pedido))
        .filter(_pedido_venda_normalizado_sql(Pedido.pedido_venda).in_(valores))
        .all()
    )
    mapa = {}
    for insp in inspecoes:
        chave = _normalizar_pedido_venda(insp.pedido_venda)
        if not chave:
            continue
        d = mapa.setdefault(chave, {"total": 0, "reprovadas": 0, "com_desvio": 0, "quantidade_com_desvio": 0})
        d["total"] += 1
        if insp.resultado == "REPROVADO":
            d["reprovadas"] += 1
        if insp.resultado == "APROVADO_COM_DESVIO":
            d["com_desvio"] += 1
        d["quantidade_com_desvio"] += insp.quantidade_com_desvio or 0
    return mapa


# ----------------------------------------------------------------------
# Gestão de Risco / Torre de Controle de OTD (pedido do Bruno, 11/09/2026):
# camada de PREVISÃO sobre a Gestão Operação já existente — pra todo pedido
# CIF ainda não entregue, cruza o prazo comercial prometido (Comercial), o
# lead time de produção (Produção/PCP, "ao vivo" igual _metricas_operacao_
# 360 já faz) com o lead time de transporte JÁ CADASTRADO em Cadastros >
# Lead time Transportadora (fonte única — nunca duplicado aqui, sempre
# consultado ao vivo) e projeta se o pedido vai entregar dentro do prazo.
#
# NADA fica gravado no banco: todo o risco é recalculado a cada carregamento
# da tela, a partir do estado ATUAL de Produção/Operação/Cadastros — por
# isso "recalcula automaticamente sempre que qualquer variável relevante
# mudar" já sai de graça: não existe um campo "risco" salvo que possa ficar
# desatualizado, é sempre a leitura mais recente, igual o resto da Gestão
# Operação (status pedido, leads times, etc. — nenhum é armazenado).
#
# Duas simplificações conscientes, registradas aqui pra ficar rastreável:
# 1) O cadastro de Lead time Transportadora tem 1 prazo por UF/modalidade
#    (não por transportadora) — essa 3ª dimensão não existe cadastrada em
#    lugar nenhum do sistema hoje. O cálculo usa só UF + modalidade; a
#    transportadora do pedido (quando preenchida em Gestão Operação >
#    Logística) aparece como informação de contexto, não como parte da
#    consulta de prazo — evita inventar/duplicar dado que o Bruno não
#    cadastrou.
# 2) Nenhum pedido guarda hoje a modalidade de transporte (rodoviário/
#    aéreo) escolhida — só existe esse conceito no cadastro de lead time.
#    O cálculo BASE sempre assume Rodoviário (modo padrão) e usa o Aéreo só
#    como alternativa de recuperação sugerida quando ajuda a cumprir o
#    prazo — nunca como suposição automática do que já está em curso.
# ----------------------------------------------------------------------

RISCO_OTD_LIMITE_RISCO_DIAS = 2
RISCO_OTD_LIMITE_ATENCAO_DIAS = 5

RISCO_OTD_STATUS_INFO = {
    "INVIAVEL": {"label": "Inviável", "emoji": "🔴", "cor": "danger", "ordem": 0,
                 "descricao": "A previsão de entrega ultrapassa o prazo comercial."},
    "RISCO": {"label": "Risco", "emoji": "🟠", "cor": "risco", "ordem": 1,
              "descricao": "Margem insuficiente — alguma variável operacional ameaça o prazo."},
    "ATENCAO": {"label": "Atenção", "emoji": "🟡", "cor": "warning", "ordem": 2,
                "descricao": "Prazo próximo do limite."},
    "VIAVEL": {"label": "Viável", "emoji": "🟢", "cor": "success", "ordem": 3,
               "descricao": "Existe margem segura para atendimento."},
    "SEM_DADO": {"label": "Sem dado suficiente", "emoji": "⚪", "cor": "secondary", "ordem": 4,
                 "descricao": "Falta prazo comercial, dado de produção ou lead time cadastrado pra UF — "
                              "não dá pra projetar ainda."},
}


def _mapa_lead_time_transportadora():
    """dict (uf, modalidade) -> LeadTimeTransportadora — fonte ÚNICA de lead
    time de transporte pra Gestão de Risco (pedido do Bruno, 11/09/2026:
    "não duplicar... consultar automaticamente os parâmetros existentes em
    Cadastros"). Só linhas ativas. Como hoje só existe 1 origem cadastrada
    (Pindamonhangaba-SP), a chave não inclui origem — se um dia existir mais
    de uma origem, revisitar aqui."""
    linhas = LeadTimeTransportadora.query.filter_by(ativo=True).all()
    mapa = {}
    for linha in linhas:
        mapa.setdefault((linha.uf, linha.modalidade), linha)
    return mapa


def _lead_time_transporte_dias(linha):
    """Converte 1 linha de LeadTimeTransportadora num nº de dias corridos
    pra somar no cálculo de previsão — usa sempre o prazo MÁXIMO (mais
    conservador, "pior caso") quando é uma faixa; quando a unidade é Horas
    (só RJ/SP hoje), converte pra dias arredondando pra cima (48h=2d,
    72h=3d)."""
    if linha is None:
        return None
    valor = linha.prazo_maximo if linha.prazo_maximo else linha.prazo_minimo
    if linha.unidade_prazo == "Horas":
        return math.ceil(valor / 24)
    return math.ceil(valor)


def _calcular_risco_pedido(go, m, pedido_producao, rdim_resumo, gargalos_por_estacao, mapa_lead_time):
    """Projeta o risco de atraso de UM pedido (Gestão Operação, CIF) —
    cruza: prazo comercial prometido (m["solicitada"], mesma fonte "ao
    vivo" da coluna "Data solicitada" da Operação 360); lead time de
    produção (conclusão REAL quando já existe, senão a PREVISÃO do PCP —
    m["conclusao_producao"]/m["expectativa_pcp"], os mesmos campos "ao
    vivo" já usados em toda a Gestão Operação); lead time de transporte
    (consulta o cadastro de Lead time Transportadora pelo Estado do
    pedido, modalidade Rodoviário — ver cabeçalho da seção). Nunca grava
    nada, só calcula em cima do que já existe."""
    hoje = date.today()
    uf = (go.estado or (pedido_producao.estado if pedido_producao else None) or "").strip().upper()

    prazo_comercial_data = m.get("solicitada")
    producao_data = m.get("conclusao_producao") or m.get("expectativa_pcp")
    producao_e_real = m.get("conclusao_producao") is not None

    linha_rodoviario = mapa_lead_time.get((uf, "Rodoviário"))
    linha_aereo = mapa_lead_time.get((uf, "Aéreo"))
    transporte_rodoviario_dias = _lead_time_transporte_dias(linha_rodoviario)
    transporte_aereo_dias = _lead_time_transporte_dias(linha_aereo)

    data_prevista_entrega = None
    if producao_data and transporte_rodoviario_dias is not None:
        data_prevista_entrega = producao_data + timedelta(days=transporte_rodoviario_dias)

    data_maxima_producao = None
    if prazo_comercial_data and transporte_rodoviario_dias is not None:
        data_maxima_producao = prazo_comercial_data - timedelta(days=transporte_rodoviario_dias)

    folga_dias = None
    if prazo_comercial_data and data_prevista_entrega:
        folga_dias = (prazo_comercial_data - data_prevista_entrega).days

    # --- Classificação ---------------------------------------------------
    if prazo_comercial_data is None or producao_data is None or transporte_rodoviario_dias is None:
        status = "SEM_DADO"
    elif folga_dias < 0:
        status = "INVIAVEL"
    elif folga_dias <= RISCO_OTD_LIMITE_RISCO_DIAS:
        status = "RISCO"
    elif folga_dias <= RISCO_OTD_LIMITE_ATENCAO_DIAS:
        status = "ATENCAO"
    else:
        status = "VIAVEL"

    # --- Gargalo principal + ação recomendada -----------------------------
    gargalo = None
    acao_recomendada = None
    alternativas = []
    atraso_projetado = max(0, -folga_dias) if folga_dias is not None else None

    tem_qualidade_pendente = bool(rdim_resumo and (rdim_resumo.get("reprovadas") or rdim_resumo.get("com_desvio")))

    itens_em_gargalo = []
    if pedido_producao:
        for item in pedido_producao.itens:
            if item.status_producao != "FINALIZADO":
                g = gargalos_por_estacao.get(item.estacao)
                if g and (g["fila"] > 0 or g["atraso"] > 0):
                    itens_em_gargalo.append(g)

    if status in ("RISCO", "INVIAVEL"):
        resolve_com_aereo = bool(
            transporte_aereo_dias is not None and prazo_comercial_data and producao_data
            and (prazo_comercial_data - producao_data).days >= transporte_aereo_dias
        )

        if tem_qualidade_pendente:
            gargalo = "Qualidade"
            acao_recomendada = "Verificar reinspeção/reprocesso dos itens com desvio (RDIM) antes de liberar."
        elif m.get("status_pedido_idx") == 1:
            gargalo = "PCP / Fila"
            acao_recomendada = "Priorizar entrada em produção deste pedido no PCP."
        elif itens_em_gargalo:
            pior = max(itens_em_gargalo, key=lambda g: (g["atraso"], g["fila"]))
            gargalo = "Produção"
            acao_recomendada = f'Estação "{pior["estacao"]}" com fila/atraso — priorizar este pedido lá.'
        elif resolve_com_aereo:
            gargalo = "Transporte"
            economia = transporte_rodoviario_dias - transporte_aereo_dias
            acao_recomendada = f"Alterar modalidade para aéreo (economiza ~{economia}d de transporte)."
        elif atraso_projetado:
            gargalo = "Produção"
            acao_recomendada = f"Antecipar produção em {atraso_projetado} dia(s)."
        else:
            gargalo = "Prazo comercial"
            acao_recomendada = "Margem apertada desde a origem — considerar renegociar prazo comercial."

        # Alternativas extras, quando houver mais de uma alavanca possível
        # (pedido do Bruno: "apresentar as opções de recuperação possíveis").
        if resolve_com_aereo and gargalo != "Transporte":
            economia = transporte_rodoviario_dias - transporte_aereo_dias
            alternativas.append(f"Alterar modalidade para aéreo (economiza ~{economia}d).")
        if atraso_projetado and gargalo != "Produção":
            alternativas.append(f"Antecipar produção em {atraso_projetado} dia(s).")
        if gargalo not in ("Prazo comercial", None) and folga_dias is not None and folga_dias < 0:
            alternativas.append("Renegociar prazo comercial com o cliente.")

    return {
        "go": go,
        "pedido_id": go.id,
        "pedido_venda": go.pedido_venda,
        "cliente": go.cliente,
        "uf": uf or None,
        "regiao": REGIAO_POR_UF.get(uf),
        "transportadora": go.go_transportadora.nome if go.go_transportadora else None,
        "prazo_comercial_data": prazo_comercial_data,
        "prazo_comercial_dias": m.get("lead_comercial_dias"),
        "producao_previsao_data": producao_data,
        "producao_e_real": producao_e_real,
        "lead_producao_dias": m.get("lead_producao_dias"),
        "transporte_rodoviario_dias": transporte_rodoviario_dias,
        "transporte_aereo_dias": transporte_aereo_dias,
        "data_prevista_entrega": data_prevista_entrega,
        "data_maxima_producao": data_maxima_producao,
        "folga_dias": folga_dias,
        "atraso_projetado_dias": atraso_projetado,
        "status": status,
        "status_info": RISCO_OTD_STATUS_INFO[status],
        "gargalo": gargalo,
        "acao_recomendada": acao_recomendada,
        "alternativas": alternativas,
        # Itens ainda não finalizados do pedido de PRODUÇÃO (não o de
        # Operação) — pedido do Bruno (11/09/2026, aba Simulação): usado
        # pra achar o lead time de produção PARAMETRIZADO (Cadastros > Lead
        # time Produção) de cada item, sem precisar consultar de novo
        # _pedidos_producao_por_pedido_venda (já foi resolvido pra montar
        # esta linha).
        "itens_producao_abertos": (
            [i for i in pedido_producao.itens if i.status_producao != "FINALIZADO"] if pedido_producao else []
        ),
    }


def _resumo_risco_otd(linhas):
    """KPIs do topo da Gestão de Risco — sobre o MESMO conjunto (já
    filtrado) exibido na tabela abaixo."""
    total = len(linhas)
    por_status = {chave: 0 for chave in RISCO_OTD_STATUS_INFO}
    for l in linhas:
        por_status[l["status"]] += 1

    com_folga = [l["folga_dias"] for l in linhas if l["folga_dias"] is not None]
    projetaveis = [l for l in linhas if l["status"] != "SEM_DADO"]
    no_prazo = [l for l in projetaveis if l["status"] in ("VIAVEL", "ATENCAO")]
    com_atraso = [l for l in linhas if l["atraso_projetado_dias"]]

    return {
        "total": total,
        "por_status": por_status,
        "em_risco_hoje": por_status["RISCO"] + por_status["INVIAVEL"],
        "menor_folga_dias": min(com_folga) if com_folga else None,
        "dias_medios_folga": round(sum(com_folga) / len(com_folga), 1) if com_folga else None,
        "otd_projetado_percentual": round(len(no_prazo) / len(projetaveis) * 100, 1) if projetaveis else None,
        "pedidos_com_atraso_projetado": len(com_atraso),
    }


def _pedidos_risco_otd(args):
    """Monta a lista de risco de OTD pra Gestão de Risco: todo pedido CIF de
    Gestão Operação ainda não entregue, com o cálculo de
    _calcular_risco_pedido pra cada um. Reaproveita TUDO que já existe —
    _metricas_operacao_360, _liberacao_pcp_por_pedido_venda,
    _data_cliente_por_pedido_venda, _pedidos_producao_por_pedido_venda,
    _rdim_resumo_por_pedido_venda, _gargalos_por_estacao — só adiciona a
    camada de lead time de transporte (Cadastros) e a projeção por cima.
    Devolve (linhas, resumo) já ordenado do mais urgente pro menos urgente."""
    busca = (args.get("busca", "") or "").strip()
    status_filtro = [v for v in _getlist_seguro(args, "status") if v in RISCO_OTD_STATUS_INFO]

    query = PedidoOperacao.query.filter(
        PedidoOperacao.frete == "CIF",
        PedidoOperacao.go_data_entregue_cliente.is_(None),
        PedidoOperacao.go_data_real_entrega.is_(None),
    )
    if busca:
        like = f"%{busca}%"
        query = query.filter(or_(PedidoOperacao.pedido_venda.ilike(like), PedidoOperacao.cliente.ilike(like)))

    pedidos_go = query.order_by(PedidoOperacao.data_inclusao_pedido.desc().nullslast()).all()
    if not pedidos_go:
        return [], _resumo_risco_otd([])

    pedidos_venda = [g.pedido_venda for g in pedidos_go]
    liberacao_pcp = _liberacao_pcp_por_pedido_venda(pedidos_venda)
    data_cliente = _data_cliente_por_pedido_venda(pedidos_venda)
    pedidos_producao = _pedidos_producao_por_pedido_venda(pedidos_venda)
    metricas = _metricas_operacao_360(pedidos_go, liberacao_pcp, data_cliente, pedidos_producao)
    rdim_por_pedido = _rdim_resumo_por_pedido_venda(pedidos_venda)
    mapa_lead_time = _mapa_lead_time_transportadora()
    gargalos = {g["estacao"]: g for g in _gargalos_por_estacao()}

    linhas = []
    for go in pedidos_go:
        chave = _normalizar_pedido_venda(go.pedido_venda)
        m = metricas.get(go.id, {})
        pedido_producao = pedidos_producao.get(chave)
        rdim_resumo = rdim_por_pedido.get(chave)
        linhas.append(_calcular_risco_pedido(go, m, pedido_producao, rdim_resumo, gargalos, mapa_lead_time))

    if status_filtro:
        linhas = [l for l in linhas if l["status"] in status_filtro]

    linhas.sort(key=lambda l: (
        RISCO_OTD_STATUS_INFO[l["status"]]["ordem"],
        l["folga_dias"] if l["folga_dias"] is not None else 9999,
    ))
    return linhas, _resumo_risco_otd(linhas)


SIMULACAO_BALANCO_INFO = {
    "concordam": {"label": "Parametrizado bate com a realidade", "emoji": "🟢", "cor": "success"},
    "alerta_operacional": {"label": "Realidade pior que o parametrizado", "emoji": "🔴", "cor": "danger"},
    "revisar_parametro": {"label": "Parâmetro parece desatualizado", "emoji": "🟡", "cor": "warning"},
    "sem_parametro": {"label": "Sem LT de produção parametrizado", "emoji": "⚪", "cor": "secondary"},
}

# Diferença mínima (em dias de folga) pra considerar Simulado A e Simulado B
# realmente divergentes — abaixo disso é só ruído de arredondamento, não
# vale a pena sinalizar como alerta.
SIMULACAO_DIVERGENCIA_DIAS = 2


def _simulacao_otd_linha(linha):
    """Simulado A x Simulado B pra 1 linha já calculada por
    _calcular_risco_pedido (pedido do Bruno, 11/09/2026, aba Simulação):

    Simulado A (parâmetro) = data de inclusão do pedido + LT de produção
    PARAMETRIZADO (pior caso entre os itens ainda abertos, Cadastros > Lead
    time Produção) + LT de transporte parametrizado (o mesmo já calculado
    nesta linha) — funciona mesmo sem o PCP ainda ter planejado nada, então
    já existe pra um pedido recém-criado (ver widget em editar_pedido).

    Simulado B (realidade) = a PRÓPRIA linha da Torre de Controle
    (folga_dias/data_prevista_entrega, que já usam a previsão REAL do PCP +
    o mesmo LT de transporte) — não recalcula nada, só reexibe lado a lado.

    O "balanço" entre os dois é o sinal de decisão pro Bruno: se o parâmetro
    diz que dá e a realidade diz que não dá, o problema é operacional
    (estação atrasada/gargalo); se é o contrário, o parâmetro provavelmente
    está desatualizado (folgado demais) e vale revisar."""
    itens_abertos = linha["itens_producao_abertos"]
    lt_producao, itens_sem_parametro = _lt_producao_parametrizado_pedido(itens_abertos)

    data_inclusao = linha["go"].data_inclusao_pedido
    transporte = linha["transporte_rodoviario_dias"]
    prazo_comercial = linha["prazo_comercial_data"]

    data_prevista_a = None
    if data_inclusao and lt_producao is not None and transporte is not None:
        data_prevista_a = data_inclusao + timedelta(days=lt_producao + transporte)

    folga_a = (prazo_comercial - data_prevista_a).days if (prazo_comercial and data_prevista_a) else None
    folga_b = linha["folga_dias"]

    if folga_a is None or folga_b is None:
        balanco = "sem_parametro"
    elif abs(folga_a - folga_b) < SIMULACAO_DIVERGENCIA_DIAS:
        balanco = "concordam"
    elif folga_a > folga_b:
        # Parâmetro é mais otimista que a realidade -> a operação está
        # performando pior do que o parâmetro promete.
        balanco = "alerta_operacional"
    else:
        # Parâmetro é mais pessimista que a realidade -> a operação está
        # indo melhor do que o parâmetro previa.
        balanco = "revisar_parametro"

    return {
        "data_prevista_a": data_prevista_a,
        "folga_a": folga_a,
        "lt_producao_parametrizado": lt_producao,
        "itens_sem_parametro": itens_sem_parametro,
        "data_prevista_b": linha["data_prevista_entrega"],
        "folga_b": folga_b,
        "balanco": balanco,
        "balanco_info": SIMULACAO_BALANCO_INFO[balanco],
    }


def _simulacao_otd(linhas):
    """Simulado A/B pra TODAS as linhas já calculadas (mesmo conjunto
    filtrado exibido na Torre de Controle) + um resumo por balanço, pra
    alimentar a aba Simulação sem duplicar a consulta de pedidos."""
    simulacoes = {l["pedido_id"]: _simulacao_otd_linha(l) for l in linhas}
    resumo = {chave: 0 for chave in SIMULACAO_BALANCO_INFO}
    for s in simulacoes.values():
        resumo[s["balanco"]] += 1
    return simulacoes, resumo


def _formatar_data_br(d):
    return d.strftime("%d/%m/%Y") if d else ""


def _agora_brt():
    """Hora atual no fuso de Brasília (UTC-3, sem horário de verão) — pedido
    do Bruno (21/09/2026: "atualize a hora e dia para o horario correto do
    relatorio gerado"), achado ao notar que o "Gerado em" do relatório PDF
    de Planejamento Mensal aparecia sempre 3h à frente do horário real: o
    servidor (Render) roda em UTC, então `datetime.now()` direto mostra a
    hora do servidor, não a de Brasília. Mesmo offset fixo já usado (na
    direção oposta, BRT -> UTC) em `_janela_utc_do_dia_brt` — corrigido aqui
    em TODO relatório que carimba "Gerado em" (não só o que o Bruno
    reportou), já que era o mesmo `datetime.now()` direto nos 5 lugares."""
    return datetime.utcnow() - timedelta(hours=3)


def _texto_filtros_risco_otd(filtros):
    """Descrição legível dos filtros aplicados (busca/status) — usada no
    cabeçalho dos relatórios Excel/PDF da Gestão de Risco, pra deixar claro
    pro Bruno (ou quem abrir o arquivo depois) que o relatório reflete só o
    que estava filtrado na tela, não a base inteira."""
    partes = []
    if filtros.get("busca"):
        partes.append(f'Busca: "{filtros["busca"]}"')
    if filtros.get("status"):
        labels = [RISCO_OTD_STATUS_INFO[s]["label"] for s in filtros["status"] if s in RISCO_OTD_STATUS_INFO]
        partes.append("Status: " + ", ".join(labels))
    return " · ".join(partes) if partes else "Nenhum filtro aplicado (todos os pedidos CIF em aberto)"


def _linhas_export_risco_otd(linhas):
    """Detalhamento completo (1 linha por pedido) pro relatório da Gestão de
    Risco — pedido do Bruno (11/09/2026): "relatório completo com todas as
    informações que consta na aba". Reaproveita os MESMOS dicts que já
    alimentam a tabela da tela (_pedidos_risco_otd), sem recalcular nada."""
    cabecalho = [
        "Pedido", "Cliente", "UF", "Região", "Transportadora",
        "Prazo comercial", "Previsão de produção", "Produção real?",
        "Lead transporte rodoviário (dias)", "Lead transporte aéreo (dias)",
        "Previsão de entrega", "Data máxima para produção",
        "Folga (dias)", "Atraso projetado (dias)",
        "Status", "Descrição do status", "Gargalo principal", "Ação recomendada", "Alternativas",
    ]
    linhas_export = [
        [
            l["pedido_venda"], l["cliente"], l["uf"] or "", l["regiao"] or "", l["transportadora"] or "",
            _formatar_data_br(l["prazo_comercial_data"]), _formatar_data_br(l["producao_previsao_data"]),
            "Sim" if l["producao_e_real"] else "Não (previsão PCP)",
            l["transporte_rodoviario_dias"] if l["transporte_rodoviario_dias"] is not None else "",
            l["transporte_aereo_dias"] if l["transporte_aereo_dias"] is not None else "",
            _formatar_data_br(l["data_prevista_entrega"]), _formatar_data_br(l["data_maxima_producao"]),
            l["folga_dias"] if l["folga_dias"] is not None else "",
            l["atraso_projetado_dias"] if l["atraso_projetado_dias"] is not None else "",
            l["status_info"]["label"], l["status_info"]["descricao"],
            l["gargalo"] or "", l["acao_recomendada"] or "", "; ".join(l["alternativas"]),
        ]
        for l in linhas
    ]
    return cabecalho, linhas_export


def _gerar_excel_risco_otd(linhas, resumo, filtros):
    """Relatório Excel completo da Gestão de Risco (pedido do Bruno,
    11/09/2026) — 2 abas: "Resumo" (os mesmos KPIs da tela) e "Pedidos em
    risco" (detalhamento linha a linha, ver _linhas_export_risco_otd).
    Estrutura de múltiplas abas -> monta o Workbook na mão (mesmo padrão de
    _construir_backup_pedidos_wb), não dá pra reaproveitar _responder_xlsx
    (que é sempre 1 aba só)."""
    wb = Workbook()

    ws_resumo = wb.active
    ws_resumo.title = "Resumo"
    linhas_resumo = [
        ("Torre de Controle OTD — Relatório de Risco", ""),
        ("Gerado em", _agora_brt().strftime("%d/%m/%Y %H:%M")),
        ("Filtros aplicados", _texto_filtros_risco_otd(filtros)),
        ("", ""),
        ("Total de pedidos CIF em aberto", resumo["total"]),
        ("Pedidos em risco hoje (Risco + Inviável)", resumo["em_risco_hoje"]),
        ("Menor folga de prazo (dias)", resumo["menor_folga_dias"] if resumo["menor_folga_dias"] is not None else ""),
        ("Dias médios de folga", resumo["dias_medios_folga"] if resumo["dias_medios_folga"] is not None else ""),
        ("OTD projetado (%)", resumo["otd_projetado_percentual"] if resumo["otd_projetado_percentual"] is not None else ""),
        ("Pedidos com atraso projetado", resumo["pedidos_com_atraso_projetado"]),
        ("", ""),
    ]
    for chave in RISCO_OTD_STATUS_INFO:
        info = RISCO_OTD_STATUS_INFO[chave]
        linhas_resumo.append((f'{info["emoji"]} {info["label"]}', resumo["por_status"].get(chave, 0)))
    for linha in linhas_resumo:
        ws_resumo.append(linha)
    ws_resumo["A1"].font = Font(bold=True, size=14)
    for i in (2, 3, 5, 6, 7, 8, 9, 10):
        ws_resumo.cell(row=i, column=1).font = Font(bold=True)
    for i in range(12, 12 + len(RISCO_OTD_STATUS_INFO)):
        ws_resumo.cell(row=i, column=1).font = Font(bold=True)
    ws_resumo.column_dimensions["A"].width = 42
    ws_resumo.column_dimensions["B"].width = 40

    ws_pedidos = wb.create_sheet("Pedidos em risco")
    cabecalho, linhas_export = _linhas_export_risco_otd(linhas)
    ws_pedidos.append(cabecalho)
    for celula in ws_pedidos[1]:
        celula.font = Font(bold=True)
    for linha in linhas_export:
        ws_pedidos.append(linha)
    for coluna in ws_pedidos.columns:
        valores = [len(str(c.value)) for c in coluna if c.value is not None]
        largura = max(valores) if valores else 10
        ws_pedidos.column_dimensions[coluna[0].column_letter].width = min(largura + 2, 45)
    ws_pedidos.freeze_panes = "A2"

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    resposta = Response(
        buffer.getvalue(),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    nome_arquivo = f"torre_controle_otd_{date.today().isoformat()}.xlsx"
    resposta.headers["Content-Disposition"] = f"attachment; filename={nome_arquivo}"
    return resposta


def _gerar_pdf_risco_otd(linhas, resumo, filtros):
    """Relatório PDF completo da Gestão de Risco (pedido do Bruno,
    11/09/2026) — paisagem A4, mesmos KPIs + a mesma tabela por pedido que
    aparecem na tela (inclusive as cores de status, pro relatório impresso
    continuar batendo com o que se vê no navegador). Único PDF do sistema
    até aqui -> usa reportlab (só lib pura Python da lista de dependências
    que gera PDF sem precisar de biblioteca de sistema, ao contrário de
    weasyprint/wkhtmltopdf — importante porque o deploy é via gunicorn no
    Render, sem controle sobre pacotes do SO)."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    # Mesmas cores de fundo por status já usadas na tela (style.css:
    # .linha-risco-inviavel/risco/atencao/sem_dado) — pro PDF bater
    # visualmente com o navegador.
    COR_LINHA_STATUS = {
        "INVIAVEL": colors.HexColor("#f8d7da"),
        "RISCO": colors.HexColor("#ffe5d0"),
        "ATENCAO": colors.HexColor("#fff8e1"),
        "VIAVEL": colors.white,
        "SEM_DADO": colors.HexColor("#f1f3f5"),
    }

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=landscape(A4),
        leftMargin=10 * mm, rightMargin=10 * mm, topMargin=12 * mm, bottomMargin=12 * mm,
        title="Torre de Controle OTD — Relatório de Risco",
    )
    estilos = getSampleStyleSheet()
    # Pedido do Bruno (11/09/2026): letras maiores e cabeçalho em cor clara
    # (fundo escuro consome muita tinta e prejudica a leitura quando
    # impresso) — cabeçalho da tabela usa um azul claro com texto escuro,
    # em vez do preto/branco original.
    estilo_celula = ParagraphStyle("celula", parent=estilos["Normal"], fontSize=8.5, leading=10.5)
    estilo_celula_bold = ParagraphStyle("celula_bold", parent=estilo_celula, fontName="Helvetica-Bold")
    COR_CABECALHO_BG = colors.HexColor("#d3e0f2")
    COR_CABECALHO_TEXTO = colors.HexColor("#1b2a4a")
    estilo_cabecalho_tabela = ParagraphStyle(
        "cabecalho_tabela", parent=estilo_celula_bold, fontSize=9, leading=11, textColor=COR_CABECALHO_TEXTO,
    )

    elementos = [
        Paragraph("Torre de Controle OTD — Relatório de Risco", estilos["Title"]),
        Paragraph(
            f'Gerado em {_agora_brt().strftime("%d/%m/%Y %H:%M")} · {_texto_filtros_risco_otd(filtros)}',
            estilos["Normal"],
        ),
        Spacer(1, 6 * mm),
    ]

    def _kpi(valor, rotulo):
        return [Paragraph(str(valor), ParagraphStyle("kpi_valor", parent=estilos["Normal"], fontSize=17, fontName="Helvetica-Bold", alignment=1)),
                Paragraph(rotulo, ParagraphStyle("kpi_rotulo", parent=estilos["Normal"], fontSize=8.5, alignment=1))]

    kpis_gerais = [
        _kpi(resumo["total"], "Total"),
        _kpi(resumo["em_risco_hoje"], "Em risco hoje"),
        _kpi(resumo["menor_folga_dias"] if resumo["menor_folga_dias"] is not None else "—", "Menor folga (dias)"),
        _kpi(f'{resumo["otd_projetado_percentual"]}%' if resumo["otd_projetado_percentual"] is not None else "—", "OTD projetado"),
        _kpi(resumo["dias_medios_folga"] if resumo["dias_medios_folga"] is not None else "—", "Dias médios de folga"),
        _kpi(resumo["pedidos_com_atraso_projetado"], "Com atraso projetado"),
    ]
    # Fontes padrão (Helvetica) do reportlab não têm os glyphs coloridos de
    # emoji (🔴🟠🟡🟢⚪) — renderizavam como quadrado preto ("tofu"). Em vez
    # do emoji no texto, usa só o rótulo e pinta o fundo da célula do KPI
    # com a MESMA cor da linha da tabela principal (COR_LINHA_STATUS), pra
    # manter o código de cor sem depender de glyph não suportado.
    kpis_status = [(chave, _kpi(resumo["por_status"].get(chave, 0), RISCO_OTD_STATUS_INFO[chave]["label"])) for chave in RISCO_OTD_STATUS_INFO]
    kpis = kpis_gerais + [k for _, k in kpis_status]

    largura_kpi = (landscape(A4)[0] - 20 * mm) / len(kpis)
    tabela_kpis = Table([[k[0] for k in kpis], [k[1] for k in kpis]], colWidths=[largura_kpi] * len(kpis))
    estilo_kpis = [
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#dee2e6")),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#dee2e6")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]
    for i, (chave, _) in enumerate(kpis_status):
        col = len(kpis_gerais) + i
        estilo_kpis.append(("BACKGROUND", (col, 0), (col, -1), COR_LINHA_STATUS.get(chave, colors.white)))
    tabela_kpis.setStyle(TableStyle(estilo_kpis))
    elementos.append(tabela_kpis)
    elementos.append(Spacer(1, 6 * mm))

    cabecalho = [
        "Pedido", "Cliente", "UF/Região", "Transportadora", "Prazo comercial", "Previsão produção",
        "Lead transp.", "Previsão entrega", "Data máx. produção", "Folga/Atraso", "Status", "Gargalo", "Ação recomendada",
    ]
    dados_tabela = [[Paragraph(c, estilo_cabecalho_tabela) for c in cabecalho]]
    cores_linhas = [COR_CABECALHO_BG]

    for l in linhas:
        uf_regiao = (l["uf"] or "—") + (f' ({l["regiao"]})' if l["regiao"] else "")
        producao_txt = _formatar_data_br(l["producao_previsao_data"]) or "—"
        if l["producao_previsao_data"]:
            producao_txt += " (real)" if l["producao_e_real"] else " (previsão)"
        if l["folga_dias"] is None:
            folga_txt = "—"
        elif l["folga_dias"] < 0:
            folga_txt = f'{-l["folga_dias"]}d de atraso'
        else:
            folga_txt = f'{l["folga_dias"]}d de folga'
        acao_txt = l["acao_recomendada"] or "—"
        if l["alternativas"]:
            acao_txt += "<br/>" + "<br/>".join(f"• {a}" for a in l["alternativas"])

        linha_tabela = [
            Paragraph(l["pedido_venda"] or "—", estilo_celula),
            Paragraph(l["cliente"] or "—", estilo_celula),
            Paragraph(uf_regiao, estilo_celula),
            Paragraph(l["transportadora"] or "—", estilo_celula),
            Paragraph(_formatar_data_br(l["prazo_comercial_data"]) or "—", estilo_celula),
            Paragraph(producao_txt, estilo_celula),
            Paragraph(f'{l["transporte_rodoviario_dias"]}d' if l["transporte_rodoviario_dias"] is not None else "—", estilo_celula),
            Paragraph(_formatar_data_br(l["data_prevista_entrega"]) or "—", estilo_celula),
            Paragraph(_formatar_data_br(l["data_maxima_producao"]) or "—", estilo_celula),
            Paragraph(folga_txt, estilo_celula),
            Paragraph(l["status_info"]["label"], estilo_celula_bold),
            Paragraph(l["gargalo"] or "—", estilo_celula),
            Paragraph(acao_txt, estilo_celula),
        ]
        dados_tabela.append(linha_tabela)
        cores_linhas.append(COR_LINHA_STATUS.get(l["status"], colors.white))

    # Pesos relativos de cada coluna — escalados pra ocupar a LARGURA TOTAL
    # disponível da página (paisagem A4 menos as margens laterais), em vez
    # de um valor fixo menor que sobrava espaço em branco na folha e
    # forçava quebra de palavra no meio (ex.: "Prazo come|rcial"). Pedido
    # do Bruno (11/09/2026): distribuir melhor as colunas na página.
    pesos = [10, 15, 11, 10, 13, 14, 9, 12, 12, 11, 13, 14, 24]
    largura_disponivel = landscape(A4)[0] - doc.leftMargin - doc.rightMargin
    soma_pesos = sum(pesos)
    larguras_mm = [p / soma_pesos * largura_disponivel for p in pesos]

    tabela = Table(dados_tabela, colWidths=larguras_mm, repeatRows=1)
    estilo_tabela = [
        ("BACKGROUND", (0, 0), (-1, 0), COR_CABECALHO_BG),
        ("TEXTCOLOR", (0, 0), (-1, 0), COR_CABECALHO_TEXTO),
        ("LINEBELOW", (0, 0), (-1, 0), 1, colors.HexColor("#8fa8cc")),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#ced4da")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 3.5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3.5),
    ]
    for i, cor in enumerate(cores_linhas):
        if i == 0:
            continue
        estilo_tabela.append(("BACKGROUND", (0, i), (-1, i), cor))
    tabela.setStyle(TableStyle(estilo_tabela))
    elementos.append(tabela)

    if not linhas:
        elementos.append(Spacer(1, 6 * mm))
        elementos.append(Paragraph("Nenhum pedido encontrado com o filtro aplicado.", estilos["Normal"]))

    doc.build(elementos)
    buffer.seek(0)
    resposta = Response(buffer.getvalue(), mimetype="application/pdf")
    nome_arquivo = f"torre_controle_otd_{date.today().isoformat()}.pdf"
    resposta.headers["Content-Disposition"] = f"attachment; filename={nome_arquivo}"
    return resposta


# ----------------------------------------------------------------------
# Estações — relatório PDF por estação (pedido do Bruno, 11/09/2026): "dentro
# de cada estação" ele quer gerar um PDF voltado pra Pendente (fila), Em
# produção, ou os dois juntos — visão rápida pra imprimir/levar pro chão de
# fábrica sem precisar abrir o sistema.
# ----------------------------------------------------------------------
RELATORIO_ESTACAO_STATUS_INFO = {
    "pendente": {"label": "Fila (pendente)", "titulo": "Pendente — fila"},
    "em_producao": {"label": "Em produção", "titulo": "Em produção"},
    "ambos": {"label": "Fila + Em produção", "titulo": "Fila + Em produção"},
}


def _itens_relatorio_estacao(nome, status_filtro):
    """Busca os itens de UMA estação pro relatório PDF, já filtrados pelo
    status escolhido — mesmo split PENDENTE (fila) vs demais (em produção,
    exceto FINALIZADO) que _linha() já usa na tela de Estações, pra nunca
    divergir do que o card mostra. Mesma ordenação por prazo (mais urgente
    primeiro) do Kanban (_chave_prazo em estacao_kanban)."""
    query = ItemPedido.query.options(selectinload(ItemPedido.pedido)).filter(ItemPedido.estacao == nome)
    if status_filtro == "pendente":
        query = query.filter(ItemPedido.status_producao == "PENDENTE")
    elif status_filtro == "em_producao":
        query = query.filter(ItemPedido.status_producao.notin_(["FINALIZADO", "PENDENTE"]))
    else:  # "ambos"
        query = query.filter(ItemPedido.status_producao != "FINALIZADO")
    itens = query.all()

    def _chave_prazo(item):
        return (item.liberacao_prevista is None, item.liberacao_prevista or date.max, item.id)

    itens.sort(key=_chave_prazo)
    return itens


def _gerar_pdf_estacao(estacao, itens, status_filtro):
    """Relatório PDF de UMA estação (pedido do Bruno, 11/09/2026, revisado
    17/09/2026 — "deixe mais intuitivo e visual... AGRUPE SEPARADAMENTE o que
    está em produção e o que está pendente, de forma totalmente visual e
    dinâmica"): mesmo estilo claro/print-friendly do relatório da Torre de
    Controle OTD, cores de linha reaproveitando o semáforo de prazo que já
    aparece no Kanban — agora com Em produção/Pendente sempre em blocos
    visualmente separados (faixa colorida própria, MESMAS cores que a tela
    de Estações já usa pro Kanban — STATUS_CHAO_CORES: azul/primary pra "Em
    produção", cinza/secondary pra "Pendente" — pra nunca destoar do que a
    pessoa já reconhece na tela) e um "pingo" colorido de prazo por linha, em
    vez de só o texto.

    Nota técnica: emoji Unicode NÃO é usado aqui de propósito — testei e a
    fonte padrão do PDF (Helvetica/WinAnsi) não tem esses glifos, então cada
    emoji viraria um quadrado preto ao imprimir. O efeito "visual/intuitivo"
    pedido é feito com cor e forma (faixas coloridas, pingo de semáforo),
    que funciona em qualquer impressora sem depender de fonte nenhuma."""
    from reportlab.graphics.shapes import Circle, Drawing
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    COR_SEMAFORO_BG = {
        "vermelho": colors.HexColor("#f8d7da"),
        "amarelo": colors.HexColor("#fff3cd"),
        "verde": colors.white,
        "cinza": colors.HexColor("#f1f3f5"),
    }
    COR_SEMAFORO_PINGO = {
        "vermelho": colors.HexColor("#dc3545"),
        "amarelo": colors.HexColor("#ffc107"),
        "verde": colors.HexColor("#198754"),
        "cinza": colors.HexColor("#adb5bd"),
    }
    # Mesmas cores (hex equivalente aos tokens Bootstrap) já usadas em
    # STATUS_CHAO_CORES pro Kanban de Estações — "Em produção" = primary,
    # "Pendente" = secondary.
    COR_GRUPO_EM_PRODUCAO = colors.HexColor("#0d6efd")
    COR_GRUPO_PENDENTE = colors.HexColor("#6c757d")

    info_filtro = RELATORIO_ESTACAO_STATUS_INFO.get(status_filtro, RELATORIO_ESTACAO_STATUS_INFO["ambos"])
    rotulo = rotulo_estacao(estacao.nome)

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=landscape(A4),
        leftMargin=10 * mm, rightMargin=10 * mm, topMargin=12 * mm, bottomMargin=12 * mm,
        title=f"{rotulo} — Relatório de Produção",
    )
    estilos = getSampleStyleSheet()
    estilo_celula = ParagraphStyle("celula", parent=estilos["Normal"], fontSize=8.5, leading=10.5)
    COR_CABECALHO_BG = colors.HexColor("#d3e0f2")
    COR_CABECALHO_TEXTO = colors.HexColor("#1b2a4a")
    estilo_cabecalho_tabela = ParagraphStyle(
        "cabecalho_tabela", parent=estilo_celula, fontName="Helvetica-Bold", fontSize=9, leading=11,
        textColor=COR_CABECALHO_TEXTO,
    )

    largura_disponivel = landscape(A4)[0] - doc.leftMargin - doc.rightMargin

    elementos = [
        Paragraph(f"{rotulo} — Relatório de Produção", estilos["Title"]),
        Paragraph(
            f'Filtro: {info_filtro["titulo"]} · Gerado em {_agora_brt().strftime("%d/%m/%Y %H:%M")}',
            estilos["Normal"],
        ),
        Spacer(1, 6 * mm),
    ]

    def _kpi(valor, rotulo_kpi):
        return [
            Paragraph(str(valor), ParagraphStyle("kpi_valor", parent=estilos["Normal"], fontSize=17, fontName="Helvetica-Bold", alignment=1)),
            Paragraph(rotulo_kpi, ParagraphStyle("kpi_rotulo", parent=estilos["Normal"], fontSize=8.5, alignment=1)),
        ]

    hoje = date.today()
    total_itens = len(itens)
    total_pecas = sum(item.quantidade or 0 for item in itens)
    total_pecas_txt = int(total_pecas) if total_pecas == int(total_pecas) else total_pecas
    na_fila = sum(1 for item in itens if item.status_producao == "PENDENTE")
    em_producao = total_itens - na_fila
    criticos = sum(
        1 for item in itens
        if item.liberacao_prevista is not None and item.liberacao_prevista < hoje
    )

    kpis = [
        _kpi(total_itens, "OP/produto no relatório"),
        _kpi(total_pecas_txt, "Total de peças"),
        _kpi(na_fila, "Na fila"),
        _kpi(em_producao, "Em produção"),
        _kpi(criticos, "Críticos (atrasados)"),
    ]
    # Faixa de cor no topo de cada KPI (mesma linguagem visual das faixas de
    # grupo abaixo) — pinta rapidamente o que é "fila" (cinza), "em
    # produção" (azul) e "crítico" (vermelho, só quando > 0), sem precisar
    # de emoji.
    cores_topo_kpi = [
        colors.HexColor("#8fa8cc"), colors.HexColor("#8fa8cc"), COR_GRUPO_PENDENTE, COR_GRUPO_EM_PRODUCAO,
        colors.HexColor("#dc3545") if criticos else colors.HexColor("#dee2e6"),
    ]
    largura_kpi = (landscape(A4)[0] - 20 * mm) / len(kpis)
    tabela_kpis = Table([[k[0] for k in kpis], [k[1] for k in kpis]], colWidths=[largura_kpi] * len(kpis))
    estilo_kpis = [
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#dee2e6")),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#dee2e6")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("BACKGROUND", (4, 0), (4, -1), colors.HexColor("#f8d7da") if criticos else colors.white),
    ]
    for i, cor_topo in enumerate(cores_topo_kpi):
        estilo_kpis.append(("LINEABOVE", (i, 0), (i, 0), 2.5, cor_topo))
    tabela_kpis.setStyle(TableStyle(estilo_kpis))
    elementos.append(tabela_kpis)
    elementos.append(Spacer(1, 7 * mm))

    def _pingo(cor_nome):
        diam = 3.2 * mm
        d = Drawing(diam, diam)
        d.add(Circle(diam / 2, diam / 2, diam / 2 - 0.2, fillColor=COR_SEMAFORO_PINGO.get(cor_nome, colors.grey), strokeColor=None))
        return d

    def _faixa_grupo(texto, cor_fundo):
        """Faixa colorida de largura total marcando o início de um bloco
        (Em produção / Pendente) — pedido do Bruno (17/09/2026): "agrupe
        separadamente... de forma totalmente visual e dinâmica"."""
        estilo_faixa = ParagraphStyle(
            "faixa_grupo", parent=estilos["Normal"], fontName="Helvetica-Bold", fontSize=11,
            textColor=colors.white, leading=13,
        )
        t = Table([[Paragraph(texto, estilo_faixa)]], colWidths=[largura_disponivel])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), cor_fundo),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ]))
        return t

    def _agrupar_por_pedido(itens_grupo):
        """Agrupa itens CONSECUTIVOS do mesmo pedido (pedido do Bruno,
        17/09/2026: "quando for do mesmo pedido, agrupe por pedido... em
        TODAS as estações") — preserva a ordem de urgência já aplicada em
        _itens_relatorio_estacao (mais urgente primeiro); só junta os itens
        de um mesmo pedido na posição da 1ª ocorrência dele nessa ordem, pela
        FK direta ItemPedido.pedido_id (sempre presente, mais confiável que
        casar por texto)."""
        grupos = []
        indice_por_pedido = {}
        for item in itens_grupo:
            chave = item.pedido_id
            if chave not in indice_por_pedido:
                indice_por_pedido[chave] = len(grupos)
                grupos.append([])
            grupos[indice_por_pedido[chave]].append(item)
        return grupos

    def _tabela_itens(itens_grupo):
        """Monta a tabela de itens (sem a coluna "Status" — já fica implícita
        na faixa colorida do grupo — e com um "pingo" de semáforo na frente
        da Situação de prazo, no lugar de só texto). Itens do MESMO pedido
        ficam agrupados e as colunas Pedido/Cliente aparecem só 1 vez por
        grupo (célula mesclada verticalmente) — pedido do Bruno (17/09/2026):
        "não repita o mesmo nome do cliente e o mesmo número do pedido"."""
        cabecalho = [
            "", "Pedido", "Cliente", "Produto", "Qtd", "Incluído", "Solicitado", "Previsto", "Início OP", "Situação prazo",
        ]
        dados_tabela = [[Paragraph(c, estilo_cabecalho_tabela) if c else "" for c in cabecalho]]
        cores_linhas = [COR_CABECALHO_BG]
        spans_pedido = []  # (linha_inicio, linha_fim) 1-based (linha 0 = cabeçalho)
        divisores_grupo = []  # linha da ÚLTIMA linha de cada grupo (exceto a última da tabela)

        grupos_pedido = _agrupar_por_pedido(itens_grupo)
        linha_atual = 1
        for grupo in grupos_pedido:
            linha_inicio_grupo = linha_atual
            for indice_no_grupo, item in enumerate(grupo):
                pedido = item.pedido
                cor, dias = item.semaforo
                if dias is None:
                    prazo_txt = "sem prazo"
                elif dias < 0:
                    prazo_txt = f"{-dias}d atrasado"
                else:
                    prazo_txt = f"{dias}d"
                qtd = item.quantidade or 0
                qtd_txt = int(qtd) if qtd == int(qtd) else qtd

                # Pedido/Cliente só na 1ª linha do grupo — as demais ficam em
                # branco e a célula mesclada (SPAN) cobre o grupo inteiro.
                if indice_no_grupo == 0:
                    cel_pedido = Paragraph((pedido.pedido_venda if pedido else None) or "—", estilo_celula)
                    cel_cliente = Paragraph((pedido.cliente if pedido else None) or "—", estilo_celula)
                else:
                    cel_pedido = ""
                    cel_cliente = ""

                linha_tabela = [
                    _pingo(cor),
                    cel_pedido,
                    cel_cliente,
                    Paragraph(item.descricao_produto or "—", estilo_celula),
                    Paragraph(str(qtd_txt), estilo_celula),
                    Paragraph((_formatar_data_br(pedido.data_inclusao_pedido) if pedido else "") or "—", estilo_celula),
                    Paragraph((_formatar_data_br(pedido.data_cliente) if pedido else "") or "—", estilo_celula),
                    Paragraph(_formatar_data_br(item.liberacao_prevista) or "—", estilo_celula),
                    Paragraph(_formatar_data_br(item.inicio_producao) or "—", estilo_celula),
                    Paragraph(prazo_txt, estilo_celula),
                ]
                dados_tabela.append(linha_tabela)
                cores_linhas.append(COR_SEMAFORO_BG.get(cor, colors.white))
                linha_atual += 1

            if len(grupo) > 1:
                spans_pedido.append((linha_inicio_grupo, linha_atual - 1))
            divisores_grupo.append(linha_atual - 1)

        pesos = [4, 11, 16, 22, 6, 10, 10, 10, 10, 11]
        soma_pesos = sum(pesos)
        larguras_mm = [p / soma_pesos * largura_disponivel for p in pesos]

        tabela = Table(dados_tabela, colWidths=larguras_mm, repeatRows=1)
        estilo_tabela = [
            ("BACKGROUND", (0, 0), (-1, 0), COR_CABECALHO_BG),
            ("LINEBELOW", (0, 0), (-1, 0), 1, colors.HexColor("#8fa8cc")),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#ced4da")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (0, -1), "CENTER"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING", (0, 0), (-1, -1), 3.5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3.5),
        ]
        # Mescla Pedido (col 1) e Cliente (col 2) verticalmente pra cada
        # grupo com mais de 1 item — texto aparece só 1 vez, centralizado.
        for linha_inicio, linha_fim in spans_pedido:
            estilo_tabela.append(("SPAN", (1, linha_inicio), (1, linha_fim)))
            estilo_tabela.append(("SPAN", (2, linha_inicio), (2, linha_fim)))
        # Linha divisória um pouco mais forte entre pedidos diferentes, pra
        # reforçar visualmente onde um grupo termina e o outro começa.
        ultima_linha = len(dados_tabela) - 1
        for linha_fim in divisores_grupo:
            if linha_fim != ultima_linha:
                estilo_tabela.append(("LINEBELOW", (0, linha_fim), (-1, linha_fim), 1, colors.HexColor("#8fa8cc")))
        for i, cor_linha in enumerate(cores_linhas):
            if i == 0:
                continue
            estilo_tabela.append(("BACKGROUND", (0, i), (-1, i), cor_linha))
        tabela.setStyle(TableStyle(estilo_tabela))
        return tabela

    if status_filtro == "ambos":
        # "PRINCIPAL: agrupe separadamente o que está em produção e o que
        # está pendente" (pedido do Bruno, 17/09/2026) — 2 blocos sempre
        # separados por uma faixa colorida, cada um com sua própria tabela,
        # em vez da tabela única de antes (onde os dois status ficavam
        # misturados, só distinguíveis pela coluna "Status"/cor da linha).
        # Mantém a ordem por prazo (mais urgente primeiro) DENTRO de cada
        # bloco — mesmo critério de sempre.
        grupos = [
            ("em_producao", "EM PRODUÇÃO", COR_GRUPO_EM_PRODUCAO, [i for i in itens if i.status_producao != "PENDENTE"]),
            ("pendente", "PENDENTE — FILA", COR_GRUPO_PENDENTE, [i for i in itens if i.status_producao == "PENDENTE"]),
        ]
        for _chave, titulo, cor_fundo, itens_grupo in grupos:
            elementos.append(_faixa_grupo(f"{titulo} — {len(itens_grupo)} ITEM(NS)", cor_fundo))
            if itens_grupo:
                elementos.append(_tabela_itens(itens_grupo))
            else:
                elementos.append(Spacer(1, 2 * mm))
                elementos.append(Paragraph("Nenhum item nesta situação no momento.", estilos["Normal"]))
            elementos.append(Spacer(1, 7 * mm))
    else:
        # Filtro já veio de um status só (Pendente OU Em produção) — mantém a
        # MESMA linguagem visual (faixa + tabela sem coluna Status), só com 1
        # bloco em vez de 2, pra nunca destoar do relatório "ambos".
        titulo_unico = "PENDENTE — FILA" if status_filtro == "pendente" else "EM PRODUÇÃO"
        cor_unica = COR_GRUPO_PENDENTE if status_filtro == "pendente" else COR_GRUPO_EM_PRODUCAO
        elementos.append(_faixa_grupo(f"{titulo_unico} — {len(itens)} ITEM(NS)", cor_unica))
        if itens:
            elementos.append(_tabela_itens(itens))
        else:
            elementos.append(Spacer(1, 2 * mm))
            elementos.append(Paragraph("Nenhum item encontrado com o filtro aplicado.", estilos["Normal"]))

    doc.build(elementos)
    buffer.seek(0)
    resposta = Response(buffer.getvalue(), mimetype="application/pdf")
    nome_arquivo = f"estacao_{estacao.nome}_{status_filtro}_{date.today().isoformat()}.pdf"
    resposta.headers["Content-Disposition"] = f"attachment; filename={nome_arquivo}"
    return resposta


# ----------------------------------------------------------------------
# Relatório PDF de VÁRIAS estações de uma vez (pedido do Bruno, 22/09/2026,
# na tela de Estações: "quero que nessa area, tenha a possibilidade geraçãpo
# de relatorio em PDF, com a geração de relatorio do que tem pednnete e em
# andamento de diverdas areas... ex: quero ver em um relaotio em pdf oque
# tem pendente e andamento no PU e Espumagem... crie uma area onde posso
# selecionar diversas areas para gerar o relatorio"). Reaproveita a MESMA
# consulta por estação (_itens_relatorio_estacao) do relatório de uma
# estação só, só que uma vez por estação escolhida, tudo no mesmo PDF — cada
# estação vira sua própria seção, na mesma ordem de agrupamento por processo
# que a tela /estacoes já usa (ver ESTACOES_GRUPOS_MONITORAMENTO).
# ----------------------------------------------------------------------
def _gerar_pdf_estacoes_multiplas(estacoes_com_itens, status_filtro):
    """`estacoes_com_itens` é uma lista de tuplas (Estacao, [ItemPedido...]),
    já filtrada pelo `status_filtro` escolhido (mesma _itens_relatorio_estacao
    do relatório de uma estação só, pra nunca divergir do que a tela mostra).

    Mesma linguagem visual do relatório de uma estação (_gerar_pdf_estacao):
    faixa colorida separando "Em produção"/"Pendente", "pingo" de semáforo
    por linha, agrupamento por pedido — só que repetida por estação, cada
    uma com seu próprio banner de seção (azul-marinho, pra distinguir
    claramente onde uma estação termina e a próxima começa), dentro do MESMO
    documento.

    Página em retrato (mesmo motivo do Planejamento Mensal PCP, pedido do
    Bruno 22/09/2026: "gere na vertical, aproveitando maximo de espaço na
    folha") — aqui a tabela tem só 10 colunas (contra as 13 do relatório de
    Planejamento), então sobra bem mais largura pra Cliente/Produto sem
    precisar espremer tanto as colunas de data quanto foi preciso lá."""
    from reportlab.graphics.shapes import Circle, Drawing
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    COR_SEMAFORO_BG = {
        "vermelho": colors.HexColor("#f8d7da"),
        "amarelo": colors.HexColor("#fff3cd"),
        "verde": colors.white,
        "cinza": colors.HexColor("#f1f3f5"),
    }
    COR_SEMAFORO_PINGO = {
        "vermelho": colors.HexColor("#dc3545"),
        "amarelo": colors.HexColor("#ffc107"),
        "verde": colors.HexColor("#198754"),
        "cinza": colors.HexColor("#adb5bd"),
    }
    # Mesmas cores (equivalente hex dos tokens Bootstrap) já usadas em
    # STATUS_CHAO_CORES/_gerar_pdf_estacao — "Em produção" = primary,
    # "Pendente" = secondary. Banner de estação em azul-marinho (mesmo tom
    # do cabeçalho/rodapé do Planejamento Mensal PCP) pra ficar claramente
    # "um nível acima" das faixas de Em produção/Pendente dentro dela.
    COR_GRUPO_EM_PRODUCAO = colors.HexColor("#0d6efd")
    COR_GRUPO_PENDENTE = colors.HexColor("#6c757d")
    COR_ESTACAO_BANNER = colors.HexColor("#1b2a4a")

    info_filtro = RELATORIO_ESTACAO_STATUS_INFO.get(status_filtro, RELATORIO_ESTACAO_STATUS_INFO["ambos"])

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        leftMargin=8 * mm, rightMargin=8 * mm, topMargin=10 * mm, bottomMargin=10 * mm,
        title="Relatório de Estações Selecionadas",
    )
    estilos = getSampleStyleSheet()
    estilo_celula = ParagraphStyle("celula", parent=estilos["Normal"], fontSize=7.5, leading=8.8)
    COR_CABECALHO_BG = colors.HexColor("#d3e0f2")
    COR_CABECALHO_TEXTO = colors.HexColor("#1b2a4a")
    estilo_cabecalho_tabela = ParagraphStyle(
        "cabecalho_tabela", parent=estilo_celula, fontName="Helvetica-Bold", fontSize=7.8, leading=9.1,
        textColor=COR_CABECALHO_TEXTO,
    )

    largura_disponivel = A4[0] - doc.leftMargin - doc.rightMargin

    rotulos_selecionados = [rotulo_estacao(e.nome) for e, _ in estacoes_com_itens]
    elementos = [
        Paragraph("Relatório de Estações Selecionadas", estilos["Title"]),
        Paragraph(f'Estações: {", ".join(rotulos_selecionados) or "—"}', estilos["Normal"]),
        Paragraph(
            f'Filtro: {info_filtro["titulo"]} · Gerado em {_agora_brt().strftime("%d/%m/%Y %H:%M")}',
            estilos["Normal"],
        ),
        Spacer(1, 5 * mm),
    ]

    def _kpi(valor, rotulo_kpi):
        return [
            Paragraph(str(valor), ParagraphStyle("kpi_valor", parent=estilos["Normal"], fontSize=16, fontName="Helvetica-Bold", alignment=1)),
            Paragraph(rotulo_kpi, ParagraphStyle("kpi_rotulo", parent=estilos["Normal"], fontSize=8, alignment=1)),
        ]

    hoje = date.today()
    todos_itens = [item for _, itens in estacoes_com_itens for item in itens]
    total_itens = len(todos_itens)
    total_pecas = sum(item.quantidade or 0 for item in todos_itens)
    total_pecas_txt = int(total_pecas) if total_pecas == int(total_pecas) else total_pecas
    na_fila = sum(1 for item in todos_itens if item.status_producao == "PENDENTE")
    em_producao = total_itens - na_fila
    criticos = sum(
        1 for item in todos_itens
        if item.liberacao_prevista is not None and item.liberacao_prevista < hoje
    )

    # KPI geral, somando TODAS as estações escolhidas (visão rápida de topo,
    # antes de entrar seção por seção) — mesmo estilo do relatório de uma
    # estação só, com um KPI a mais no início (nº de estações no relatório).
    kpis = [
        _kpi(len(estacoes_com_itens), "Estações no relatório"),
        _kpi(total_itens, "OP/produto no total"),
        _kpi(total_pecas_txt, "Total de peças"),
        _kpi(na_fila, "Na fila"),
        _kpi(em_producao, "Em produção"),
        _kpi(criticos, "Críticos (atrasados)"),
    ]
    cores_topo_kpi = [
        colors.HexColor("#8fa8cc"), colors.HexColor("#8fa8cc"), colors.HexColor("#8fa8cc"),
        COR_GRUPO_PENDENTE, COR_GRUPO_EM_PRODUCAO,
        colors.HexColor("#dc3545") if criticos else colors.HexColor("#dee2e6"),
    ]
    largura_kpi = largura_disponivel / len(kpis)
    tabela_kpis = Table([[k[0] for k in kpis], [k[1] for k in kpis]], colWidths=[largura_kpi] * len(kpis))
    estilo_kpis = [
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#dee2e6")),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#dee2e6")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("BACKGROUND", (5, 0), (5, -1), colors.HexColor("#f8d7da") if criticos else colors.white),
    ]
    for i, cor_topo in enumerate(cores_topo_kpi):
        estilo_kpis.append(("LINEABOVE", (i, 0), (i, 0), 2.5, cor_topo))
    tabela_kpis.setStyle(TableStyle(estilo_kpis))
    elementos.append(tabela_kpis)
    elementos.append(Spacer(1, 6 * mm))

    def _pingo(cor_nome):
        diam = 3 * mm
        d = Drawing(diam, diam)
        d.add(Circle(diam / 2, diam / 2, diam / 2 - 0.2, fillColor=COR_SEMAFORO_PINGO.get(cor_nome, colors.grey), strokeColor=None))
        return d

    def _faixa(texto, cor_fundo, fonte=11):
        estilo_faixa = ParagraphStyle(
            "faixa", parent=estilos["Normal"], fontName="Helvetica-Bold", fontSize=fonte,
            textColor=colors.white, leading=fonte + 2,
        )
        t = Table([[Paragraph(texto, estilo_faixa)]], colWidths=[largura_disponivel])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), cor_fundo),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ]))
        return t

    def _agrupar_por_pedido(itens_grupo):
        grupos = []
        indice_por_pedido = {}
        for item in itens_grupo:
            chave = item.pedido_id
            if chave not in indice_por_pedido:
                indice_por_pedido[chave] = len(grupos)
                grupos.append([])
            grupos[indice_por_pedido[chave]].append(item)
        return grupos

    # Larguras calculadas medindo o texto real que cai em cada coluna nessa
    # fonte (reportlab.stringWidth), mesmo cuidado do relatório de
    # Planejamento Mensal PCP (22/09/2026) — datas ("dd/mm/aaaa") e
    # "Situação prazo" (ex.: "15d atrasado") têm largura mínima garantida
    # pra nunca cortar no meio do texto; Cliente/Produto ficam com o que
    # sobra (aqui, bem mais folga que lá — só 10 colunas, não 13).
    pesos = [18, 40, 104, 120, 30, 46, 46, 46, 46, 54]
    soma_pesos = sum(pesos)
    larguras_colunas = [p / soma_pesos * largura_disponivel for p in pesos]

    def _tabela_itens(itens_grupo):
        cabecalho = [
            "", "Pedido", "Cliente", "Produto", "Qtd", "Incluído", "Solicitado", "Previsto", "Início OP", "Situação prazo",
        ]
        dados_tabela = [[Paragraph(c, estilo_cabecalho_tabela) if c else "" for c in cabecalho]]
        cores_linhas = [COR_CABECALHO_BG]
        spans_pedido = []
        divisores_grupo = []

        grupos_pedido = _agrupar_por_pedido(itens_grupo)
        linha_atual = 1
        for grupo in grupos_pedido:
            linha_inicio_grupo = linha_atual
            for indice_no_grupo, item in enumerate(grupo):
                pedido = item.pedido
                cor, dias = item.semaforo
                if dias is None:
                    prazo_txt = "sem prazo"
                elif dias < 0:
                    prazo_txt = f"{-dias}d atrasado"
                else:
                    prazo_txt = f"{dias}d"
                qtd = item.quantidade or 0
                qtd_txt = int(qtd) if qtd == int(qtd) else qtd

                if indice_no_grupo == 0:
                    cel_pedido = Paragraph((pedido.pedido_venda if pedido else None) or "—", estilo_celula)
                    cel_cliente = Paragraph((pedido.cliente if pedido else None) or "—", estilo_celula)
                else:
                    cel_pedido = ""
                    cel_cliente = ""

                linha_tabela = [
                    _pingo(cor),
                    cel_pedido,
                    cel_cliente,
                    Paragraph(item.descricao_produto or "—", estilo_celula),
                    Paragraph(str(qtd_txt), estilo_celula),
                    Paragraph((_formatar_data_br(pedido.data_inclusao_pedido) if pedido else "") or "—", estilo_celula),
                    Paragraph((_formatar_data_br(pedido.data_cliente) if pedido else "") or "—", estilo_celula),
                    Paragraph(_formatar_data_br(item.liberacao_prevista) or "—", estilo_celula),
                    Paragraph(_formatar_data_br(item.inicio_producao) or "—", estilo_celula),
                    Paragraph(prazo_txt, estilo_celula),
                ]
                dados_tabela.append(linha_tabela)
                cores_linhas.append(COR_SEMAFORO_BG.get(cor, colors.white))
                linha_atual += 1

            if len(grupo) > 1:
                spans_pedido.append((linha_inicio_grupo, linha_atual - 1))
            divisores_grupo.append(linha_atual - 1)

        tabela = Table(dados_tabela, colWidths=larguras_colunas, repeatRows=1)
        estilo_tabela = [
            ("BACKGROUND", (0, 0), (-1, 0), COR_CABECALHO_BG),
            ("LINEBELOW", (0, 0), (-1, 0), 1, colors.HexColor("#8fa8cc")),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#ced4da")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (0, -1), "CENTER"),
            ("TOPPADDING", (0, 0), (-1, -1), 3.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ]
        for linha_inicio, linha_fim in spans_pedido:
            estilo_tabela.append(("SPAN", (1, linha_inicio), (1, linha_fim)))
            estilo_tabela.append(("SPAN", (2, linha_inicio), (2, linha_fim)))
        ultima_linha = len(dados_tabela) - 1
        for linha_fim in divisores_grupo:
            if linha_fim != ultima_linha:
                estilo_tabela.append(("LINEBELOW", (0, linha_fim), (-1, linha_fim), 1, colors.HexColor("#8fa8cc")))
        for i, cor_linha in enumerate(cores_linhas):
            if i == 0:
                continue
            estilo_tabela.append(("BACKGROUND", (0, i), (-1, i), cor_linha))
        tabela.setStyle(TableStyle(estilo_tabela))
        return tabela

    if not estacoes_com_itens:
        elementos.append(Paragraph("Nenhuma estação selecionada.", estilos["Normal"]))

    for estacao, itens in estacoes_com_itens:
        rotulo = rotulo_estacao(estacao.nome)
        elementos.append(_faixa(f"{rotulo} — {len(itens)} ITEM(NS)", COR_ESTACAO_BANNER, fonte=13))
        elementos.append(Spacer(1, 2 * mm))

        if status_filtro == "ambos":
            subgrupos = [
                ("EM PRODUÇÃO", COR_GRUPO_EM_PRODUCAO, [i for i in itens if i.status_producao != "PENDENTE"]),
                ("PENDENTE — FILA", COR_GRUPO_PENDENTE, [i for i in itens if i.status_producao == "PENDENTE"]),
            ]
        else:
            titulo_unico = "PENDENTE — FILA" if status_filtro == "pendente" else "EM PRODUÇÃO"
            cor_unica = COR_GRUPO_PENDENTE if status_filtro == "pendente" else COR_GRUPO_EM_PRODUCAO
            subgrupos = [(titulo_unico, cor_unica, itens)]

        for titulo, cor_fundo, itens_grupo in subgrupos:
            elementos.append(_faixa(f"{titulo} — {len(itens_grupo)} item(ns)", cor_fundo, fonte=10))
            if itens_grupo:
                elementos.append(_tabela_itens(itens_grupo))
            else:
                elementos.append(Spacer(1, 2 * mm))
                elementos.append(Paragraph("Nenhum item nesta situação no momento.", estilos["Normal"]))
            elementos.append(Spacer(1, 5 * mm))

        elementos.append(Spacer(1, 4 * mm))

    doc.build(elementos)
    buffer.seek(0)
    resposta = Response(buffer.getvalue(), mimetype="application/pdf")
    nome_arquivo = f"estacoes_selecionadas_{status_filtro}_{date.today().isoformat()}.pdf"
    resposta.headers["Content-Disposition"] = f"attachment; filename={nome_arquivo}"
    return resposta


# ----------------------------------------------------------------------
# Relatório semanal em PDF da Listagem Geral (pedido do Bruno, 14/09/2026):
# "quero que disponibilize [na Listagem Geral]... vai gerar um relatorio
# semanal de todos os pedidos semanais de setembro, e o total do mes
# projetado pelo pcp... quero de forma agrpado semanalmente e a somatario
# total, tanto de faturamento e tambem do total de numero de pedidos".
# Reaproveita 100% do pipeline que a própria tela já usa (_filtrar_pedidos +
# _linhas_listagem_geral) — o mesmo filtro "Planejamento mensal (PCP)" que já
# existe na tela vira o mês do relatório, então o PDF nunca diverge do que a
# Listagem Geral mostraria filtrando por aquele mês. O agrupamento semanal é
# pelo campo "Planejamento semanal (PCP)" de cada item — a MESMA projeção do
# PCP que já orienta os quadrantes no topo da tela — não por semana de
# calendário da data de inclusão, exatamente o "projetado pelo PCP" pedido.
# ----------------------------------------------------------------------
def _agrupar_linhas_por_semana_pcp(linhas):
    """Agrupa linhas da Listagem Geral (já filtradas por mês) pelo rótulo de
    Planejamento semanal (PCP) de cada item, em ordem cronológica (mesma
    chave — _chave_semana_pcp — que já ordena essa coluna na tela)."""
    grupos = {}
    for l in linhas:
        grupos.setdefault(l.planejamento_semanal, []).append(l)
    rotulos_ordenados = sorted(grupos.keys(), key=_chave_semana_pcp)
    return [(rotulo, grupos[rotulo]) for rotulo in rotulos_ordenados]


def _intervalo_calendario_semana_pcp(mes_ano, rotulo_semana):
    """Data de início/fim (domingo a sábado) da semana de calendário que um
    rótulo "SEMANA NN / MÊS / ANO" representa, só pra exibição no relatório —
    mesma âncora (domingo igual ou anterior ao dia 1 do mês) já usada em
    _quadrantes_planejamento_semanal, pra nunca mostrar um período diferente
    do que os quadrantes da tela mostrariam pra mesma semana."""
    ano, mes = mes_ano
    m = re.search(r"SEMANA\s*(\d+)", (rotulo_semana or "").upper())
    if not m:
        return (None, None)
    n = int(m.group(1))
    primeiro_dia_mes = date(ano, mes, 1)
    domingo_semana_01 = primeiro_dia_mes - timedelta(days=(primeiro_dia_mes.weekday() + 1) % 7)
    inicio = domingo_semana_01 + timedelta(days=7 * (n - 1))
    fim = inicio + timedelta(days=6)
    return (inicio, fim)


def _texto_filtros_listagem_geral_semanal(filtros):
    """Descrição legível dos filtros extras (além do mês, que já vira o
    título do relatório) ativos ao gerar o PDF — mesmo espírito de
    _texto_filtros_risco_otd, pra deixar claro que o relatório reflete só o
    recorte que estava na tela, quando houver algum filtro a mais."""
    partes = []
    if filtros.get("busca"):
        partes.append(f'Busca: "{filtros["busca"]}"')
    if filtros.get("cliente"):
        partes.append(f'Cliente: "{filtros["cliente"]}"')
    if filtros.get("vendedor"):
        partes.append(f'Vendedor: "{filtros["vendedor"]}"')
    if filtros.get("status"):
        partes.append(f'Status: {filtros["status"]}')
    if filtros.get("estacao"):
        partes.append(f'Estação: {filtros["estacao"]}')
    if filtros.get("produto"):
        partes.append(f'Produto: "{filtros["produto"]}"')
    if filtros.get("regiao"):
        partes.append(f'Região: {filtros["regiao"]}')
    if filtros.get("data_inicio") or filtros.get("data_fim"):
        partes.append(f'Incluído {filtros.get("data_inicio") or "—"} a {filtros.get("data_fim") or "—"}')
    if filtros.get("atrasados"):
        partes.append("Só pedidos atrasados")
    return " · ".join(partes)


def _gerar_pdf_planejamento_mensal_pcp(blocos, filtros, modelo="completo"):
    """PDF "Emitir relatório" — Planejamento Mensal PCP/Operação (pedido
    original do Bruno, 14/09/2026, então chamado de "Listagem Geral —
    Relatório Semanal"; aprimorado a pedido dele em 21/09/2026: "quero que
    aprimore a geração de relatorio em pdf... deixe bem mais intuitivo e
    dinamico, tanto visualmente e tambem como gestao... quero gerar
    relatorio tanto do mes atual e tambem juntamente do proximo mes
    (backlog)... no titulo 'listagem geral', quero que mude para algo
    'planejamento mensal pcp/operação'").

    `blocos` é uma lista de 1 ou 2 dicts, cada um com:
      - mes_ano: (ano, mes) do bloco
      - linhas: as _LinhaListagemGeral já filtradas/flatten desse mês
      - rotulo: texto do banner do bloco (só aparece quando há 2 blocos)

    Quando só há 1 bloco (checkbox "Incluir também o mês seguinte como
    Backlog PCP" desmarcado no modal), o layout fica bem próximo do
    relatório de sempre — sem banners nem visão comparativa. Quando há 2
    blocos, cada um ganha uma faixa colorida própria (verde = mês do
    relatório, azul = backlog do mês seguinte — MESMA linguagem de cor já
    usada nos quadrantes "MÊS ATUAL"/"OUTUBRO" do dashboard, pra manter a
    leitura visual consistente "como gestão") e o PDF abre com um
    comparativo rápido entre os dois meses antes de entrar no detalhe.

    Continua agrupando por semana (Planejamento PCP) dentro de cada mês,
    com KPIs, gráfico de barras de faturamento por semana e o rodapé de
    total — a mesma estrutura por mês de antes, só que fatorada em
    _construir_bloco (closure local) pra poder repetir 1x (mês) ou 2x
    (mês + backlog) sem duplicar a montagem inteira da tabela/gráfico.

    `modelo` (pedido do Bruno, 21/09/2026: "quero que tenha dois modelo de
    relatorio em pdf... um modelo que já esta estabelecido... e outro
    modelo mais compacto, onde eu nao quero que detalhe e cite os itens,
    somente cite o pedido... cada linha dentro da semana agrupada
    representara um pedido de venda... porem, totalmente completo (com
    valores, agrupamento, top regioes e clientes, valores, datas etc)")
    controla SÓ a tabela de detalhe dentro de cada semana:
      - "completo" (padrão, o modelo já existente): 1 linha por ITEM,
        agrupado por pedido com uma linha de subtotal quando o pedido tem
        2+ itens (ver _construir_bloco).
      - "compacto": 1 linha por PEDIDO (nunca por item) — soma
        quantidade/valor de todos os itens daquele pedido numa linha só,
        pensado pra reduzir bastante o número de páginas em meses com
        muitos itens por pedido. TUDO mais continua igual nos dois
        modelos: banners de mês/backlog, KPIs, gráfico de faturamento por
        semana, Top 10 clientes, principais regiões e principais clientes
        por região, e o rodapé de total — só a granularidade da tabela de
        pedidos muda."""
    from reportlab.graphics.shapes import Drawing, Rect, String
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    estilos = getSampleStyleSheet()
    # Fonte da tabela de detalhe reduzida de 8/8.5pt pra 7.3/7.6pt (retrato é
    # mais estreito que o landscape anterior — ver comentário na criação do
    # `doc` abaixo) — junto com a realocação de `pesos` mais abaixo, isso
    # evita que datas ("10/04/2026") e status ("FINALIZADO") quebrem no meio
    # do texto por falta de espaço (defeito visual encontrado e corrigido
    # antes de virar padrão — 22/09/2026).
    estilo_celula = ParagraphStyle("celula", parent=estilos["Normal"], fontSize=7.3, leading=8.6)
    COR_CABECALHO_BG = colors.HexColor("#d3e0f2")
    COR_CABECALHO_TEXTO = colors.HexColor("#1b2a4a")
    COR_SEMANA_BG = colors.HexColor("#eaf1fd")
    COR_TOTAL_BG = colors.HexColor("#1b2a4a")
    COR_FINALIZADO_BG = colors.HexColor("#d9f4e0")
    COR_ATUAL_BG = colors.HexColor("#198754")     # verde — mesmo tom do quadrante "success" do dashboard
    COR_BACKLOG_BG = colors.HexColor("#0d6efd")   # azul — mesmo tom do quadrante "primary" (OUTUBRO) do dashboard
    estilo_cabecalho_tabela = ParagraphStyle(
        "cabecalho_tabela", parent=estilo_celula, fontName="Helvetica-Bold", fontSize=7.6, leading=8.9,
        textColor=COR_CABECALHO_TEXTO,
    )

    def _fmt_moeda(v):
        return "R$ " + "{:,.2f}".format(v or 0).replace(",", "X").replace(".", ",").replace("X", ".")

    def _quebravel(texto):
        """Insere um espaço de verdade depois de cada hífen de `texto`, só
        usado nas colunas Cliente/Produto da tabela de detalhe. Sem isso, um
        código de produto tipo "LBD-DG2-DS4-CC2-ELC-MG" (comum no catálogo) é
        um único "token" sem espaço nenhum pro ReportLab quebrar de forma
        natural — na coluna estreita do retrato, ele acaba cortando o texto
        no meio de qualquer jeito (ex.: "LBD-DG2-D" / "S4-CC2..."). Com um
        espaço depois de cada hífen, a quebra (quando precisar) acontece
        sempre logo após um hífen, nunca no meio de um bloco de caracteres.
        (Tentativa inicial usava um espaço de largura zero (U+200B) pra não
        alterar o texto visualmente, mas a fonte base do PDF não tem esse
        glyph e ele aparecia como um quadradinho preto — corrigido antes de
        virar padrão, 22/09/2026.)"""
        return (texto or "").replace("-", "- ")

    titulos_meses = [f"{MESES_PT_EXTENSO[b['mes_ano'][1] - 1].upper()} / {b['mes_ano'][0]}" for b in blocos]
    titulo_periodo = " + ".join(titulos_meses)

    # Página em retrato (pedido do Bruno, 22/09/2026: "quero que gere na
    # vertical, aproveitando maximo de espaço na folha") — A4 retrato tem
    # bem mais altura útil (297mm) que o landscape anterior (210mm), e o
    # conteúdo do relatório é essencialmente empilhado verticalmente (KPIs,
    # gráfico, tabelas), então retrato aproveita a folha melhor e deixa bem
    # menos espaço em branco no rodapé de cada página. Margens laterais
    # reduzidas (8mm) pra compensar a largura menor do retrato e dar o
    # máximo de espaço horizontal possível pra tabela de detalhe (13
    # colunas), que é a parte mais sensível à largura da página.
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        leftMargin=8 * mm, rightMargin=8 * mm, topMargin=10 * mm, bottomMargin=10 * mm,
        title=f"Planejamento Mensal PCP/Operação — {titulo_periodo}",
    )
    largura_disponivel = A4[0] - doc.leftMargin - doc.rightMargin

    sufixo_titulo_modelo = " · Modelo compacto (1 linha por pedido)" if modelo == "compacto" else ""
    elementos = [
        Paragraph("Planejamento Mensal PCP/Operação", estilos["Title"]),
        Paragraph(f"{titulo_periodo} · Gerado em {_agora_brt().strftime('%d/%m/%Y %H:%M')}{sufixo_titulo_modelo}", estilos["Normal"]),
    ]
    texto_filtros = _texto_filtros_listagem_geral_semanal(filtros)
    if texto_filtros:
        elementos.append(Paragraph(f"Filtros ativos: {texto_filtros}", estilos["Normal"]))
    elementos.append(Spacer(1, 5 * mm))

    cores_bloco = (COR_ATUAL_BG, COR_BACKLOG_BG)

    # ---- visão comparativa mês do relatório x backlog PCP (só quando há os
    # 2 blocos) — pedido do Bruno: "possibilidade de visualizar o atual mes
    # e tambem a projeção PCP do proximo mes", "mais intuitivo e
    # dinamico... como gestao" ----
    if len(blocos) > 1:
        linha_rotulo, linha_kpi, linha_valor = [], [], []
        for b in blocos:
            linhas_b = b["linhas"]
            pedidos_b = {l.pedido_id for l in linhas_b}
            faturamento_b = sum(l.venda_total or 0 for l in linhas_b)
            linha_rotulo.append(Paragraph(b["rotulo"], ParagraphStyle("cmp_rotulo", parent=estilos["Normal"], fontSize=13, leading=15, textColor=colors.white, fontName="Helvetica-Bold", alignment=1)))
            linha_kpi.append(Paragraph(f"{len(pedidos_b)} pedido(s) · {len(linhas_b)} item(ns)", ParagraphStyle("cmp_kpi", parent=estilos["Normal"], fontSize=11, leading=13, textColor=colors.white, alignment=1)))
            linha_valor.append(Paragraph(_fmt_moeda(faturamento_b), ParagraphStyle("cmp_valor", parent=estilos["Normal"], fontSize=28, leading=32, textColor=colors.white, fontName="Helvetica-Bold", alignment=1)))
        largura_col = largura_disponivel / len(blocos)
        # Números BEM maiores e chamativos aqui (pedido do Bruno, 21/09/2026:
        # "quero que esses numeros ficam maiores e mais chamativos!!
        # principalmente os numeros da visao geral") — o faturamento de cada
        # bloco (28pt) é de longe o maior número do PDF inteiro, de propósito,
        # já que essa visão comparativa é o primeiro número que salta aos
        # olhos ao abrir o relatório.
        tabela_comparativo = Table([linha_rotulo, linha_kpi, linha_valor], colWidths=[largura_col] * len(blocos))
        estilo_comparativo = [
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, 0), 8),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 2),
            ("TOPPADDING", (0, 1), (-1, 1), 1),
            ("BOTTOMPADDING", (0, 1), (-1, 1), 6),
            ("TOPPADDING", (0, 2), (-1, 2), 2),
            ("BOTTOMPADDING", (0, 2), (-1, 2), 12),
        ]
        for i in range(len(blocos)):
            estilo_comparativo.append(("BACKGROUND", (i, 0), (i, -1), cores_bloco[i % len(cores_bloco)]))
        tabela_comparativo.setStyle(TableStyle(estilo_comparativo))
        elementos.append(Paragraph("Visão geral — mês do relatório x backlog PCP", estilos["Heading4"]))
        elementos.append(tabela_comparativo)
        elementos.append(Spacer(1, 7 * mm))

    # Cabeçalho/larguras da tabela de detalhe por semana — muda de acordo
    # com o `modelo` (completo = 1 linha por item; compacto = 1 linha por
    # pedido, sem citar os itens, pedido do Bruno 21/09/2026).
    # Larguras em pt calculadas medindo o texto real (reportlab.stringWidth)
    # que cai em cada coluna nessa fonte — não são mais só "pesos"
    # proporcionais arbitrários. No retrato (mais estreito que o landscape
    # anterior) isso importa de verdade: as 4 colunas de data (formato fixo
    # "dd/mm/aaaa", sem espaço nenhum pro texto quebrar) e a de Status
    # (palavras como "ANDAMENTO"/"FINALIZADO", também sem espaço) precisam
    # de largura mínima garantida, senão o ReportLab corta o texto no meio
    # de qualquer jeito (ex.: "10/04/2" / "026", "FINALIZ" / "ADO") — defeito
    # visual encontrado e corrigido antes de virar padrão (22/09/2026).
    # Cliente/Produto ficam com o que sobra (ainda a maior fatia) — nomes/
    # códigos muito compridos podem ocasionalmente quebrar num hífen (ver
    # `_quebravel`) ou, no pior caso raro, no meio da palavra mesmo; não tem
    # como evitar 100% disso numa página retrato com 13 colunas.
    if modelo == "compacto":
        cabecalho_tabela = [
            "PV", "Cliente", "Produtos", "Itens", "Frete", "UF / Região", "Cidade",
            "Incluído", "Solicitado", "Liberação prevista", "Liberação real", "Status", "Valor do pedido",
        ]
        pesos = [20, 58, 42, 26, 28, 34, 50, 46, 46, 46, 46, 56, 52]
    else:
        cabecalho_tabela = [
            "PV", "Cliente", "Produto", "Qtd", "Frete", "UF / Região", "Cidade",
            "Incluído", "Solicitado", "Liberação prevista", "Liberação real", "Status", "Venda item",
        ]
        pesos = [20, 58, 48, 20, 28, 34, 50, 46, 46, 46, 46, 56, 52]
    soma_pesos = sum(pesos)
    larguras_colunas = [p / soma_pesos * largura_disponivel for p in pesos]

    def _construir_bloco(b, cor_banner):
        mes_ano_b = b["mes_ano"]
        ano_b, mes_b = mes_ano_b
        titulo_mes_b = f"{MESES_PT_EXTENSO[mes_b - 1].upper()} / {ano_b}"
        linhas_b = b["linhas"]
        elems = []

        if len(blocos) > 1:
            banner = Table(
                [[Paragraph(f"<b>{b['rotulo']}</b> — {titulo_mes_b}", ParagraphStyle("banner", parent=estilos["Normal"], fontSize=12, textColor=colors.white, fontName="Helvetica-Bold"))]],
                colWidths=[largura_disponivel],
            )
            banner.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), cor_banner),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ]))
            elems.append(banner)
            elems.append(Spacer(1, 3 * mm))

        grupos_semana = _agrupar_linhas_por_semana_pcp(linhas_b)
        pedidos_mes = {l.pedido_id for l in linhas_b}
        faturamento_mes = sum(l.venda_total or 0 for l in linhas_b)

        def _kpi(valor, rotulo_kpi):
            return [
                Paragraph(str(valor), ParagraphStyle("kpi_valor", parent=estilos["Normal"], fontSize=20, leading=23, fontName="Helvetica-Bold", alignment=1)),
                Paragraph(rotulo_kpi, ParagraphStyle("kpi_rotulo", parent=estilos["Normal"], fontSize=8.5, alignment=1)),
            ]

        kpis_mes = [
            _kpi(len(pedidos_mes), "Pedidos distintos no mês"),
            _kpi(len(linhas_b), "Itens (produtos) no mês"),
            _kpi(len(grupos_semana), "Semanas com pedido"),
            _kpi(_fmt_moeda(faturamento_mes), "Faturamento total do mês"),
        ]
        largura_kpi = largura_disponivel / len(kpis_mes)
        tabela_kpis = Table([[k[0] for k in kpis_mes], [k[1] for k in kpis_mes]], colWidths=[largura_kpi] * len(kpis_mes))
        tabela_kpis.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#dee2e6")),
            ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#dee2e6")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f8f9fb")),
        ]))
        elems.append(tabela_kpis)
        elems.append(Spacer(1, 6 * mm))

        # ---- barrinha comparando faturamento por semana ("visual e dinâmico") ----
        if len(grupos_semana) > 1:
            resumo_semanas_graf = [
                (rotulo, sum(l.venda_total or 0 for l in linhas_semana))
                for rotulo, linhas_semana in grupos_semana
            ]
            maior_valor = max((v for _, v in resumo_semanas_graf), default=0) or 1
            # Rótulos maiores e mais chamativos (pedido do Bruno, 21/09/2026)
            # — linha mais alta e fontes maiores pro valor de cada semana
            # (o número que ele circulou no print) saltar mais aos olhos.
            altura_linha = 20
            altura_grafico = len(resumo_semanas_graf) * altura_linha + 6
            largura_rotulo = 68
            largura_valor = 92
            largura_barra_max = max(largura_disponivel - largura_rotulo - largura_valor - 6, 10)
            desenho = Drawing(largura_disponivel, altura_grafico)
            for i, (rotulo, valor) in enumerate(resumo_semanas_graf):
                y = altura_grafico - (i + 1) * altura_linha + 4
                m = re.search(r"SEMANA\s*(\d+)", (rotulo or "").upper())
                rotulo_curto = f"Semana {int(m.group(1))}" if m else (rotulo or "—")
                desenho.add(String(0, y, rotulo_curto, fontSize=10, fontName="Helvetica-Bold"))
                largura_barra = (valor / maior_valor) * largura_barra_max if maior_valor else 0
                desenho.add(Rect(largura_rotulo, y - 3, max(largura_barra, 1.5), 13, fillColor=colors.HexColor("#4c8bf5"), strokeColor=None))
                desenho.add(String(largura_rotulo + largura_barra_max + 6, y, _fmt_moeda(valor), fontSize=12, fontName="Helvetica-Bold"))
            elems.append(Paragraph("Faturamento por semana", estilos["Heading4"]))
            elems.append(desenho)
            elems.append(Spacer(1, 6 * mm))

        # ---- top 10 clientes do mês (pedido do Bruno, 21/09/2026: "inclue
        # um top 10 principais clientes") — ranking por faturamento somado
        # de TODOS os itens do cliente no mês (mesma base de dado dos KPIs
        # acima, nunca diverge) ----
        if linhas_b:
            agregados_cliente = {}
            for l in linhas_b:
                chave = l.cliente or "—"
                info = agregados_cliente.setdefault(chave, {"pedidos": set(), "faturamento": 0.0})
                info["pedidos"].add(l.pedido_id)
                info["faturamento"] += l.venda_total or 0
            ranking_clientes = sorted(agregados_cliente.items(), key=lambda kv: kv[1]["faturamento"], reverse=True)[:10]

            dados_top = [[
                Paragraph(c, estilo_cabecalho_tabela) for c in ("#", "Cliente", "Pedidos", "Faturamento no mês")
            ]]
            for i, (cliente_nome, info) in enumerate(ranking_clientes, start=1):
                dados_top.append([
                    Paragraph(str(i), estilo_celula),
                    Paragraph(cliente_nome, estilo_celula),
                    Paragraph(str(len(info["pedidos"])), estilo_celula),
                    Paragraph(_fmt_moeda(info["faturamento"]), estilo_celula),
                ])
            largura_top = [largura_disponivel * 0.06, largura_disponivel * 0.54, largura_disponivel * 0.16, largura_disponivel * 0.24]
            tabela_top = Table(dados_top, colWidths=largura_top, repeatRows=1)
            tabela_top.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), COR_CABECALHO_BG),
                ("TEXTCOLOR", (0, 0), (-1, 0), COR_CABECALHO_TEXTO),
                ("LINEBELOW", (0, 0), (-1, 0), 1, colors.HexColor("#8fa8cc")),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#ced4da")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (2, 1), (-1, -1), "RIGHT"),
                ("TOPPADDING", (0, 0), (-1, -1), 3.5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ]))
            elems.append(KeepTogether([Paragraph("Top 10 clientes do mês", estilos["Heading4"]), tabela_top]))
            elems.append(Spacer(1, 6 * mm))

        # ---- principais regiões de fornecimento + principais clientes por
        # região (pedido do Bruno, 21/09/2026: "principais clientes por
        # regiao e tambem as principais regioes de fornecimento") — mesmo
        # REGIAO_POR_UF já usado na coluna "UF / Região" da tabela de itens
        # (nunca diverge do resto do relatório) ----
        if linhas_b:
            agregados_regiao = {}
            for l in linhas_b:
                regiao_nome = REGIAO_POR_UF.get(l.estado, "Não identificada")
                info = agregados_regiao.setdefault(regiao_nome, {"pedidos": set(), "faturamento": 0.0})
                info["pedidos"].add(l.pedido_id)
                info["faturamento"] += l.venda_total or 0
            ranking_regioes = sorted(agregados_regiao.items(), key=lambda kv: kv[1]["faturamento"], reverse=True)
            faturamento_total_regioes = sum(info["faturamento"] for _, info in ranking_regioes) or 1

            dados_regiao = [[Paragraph(c, estilo_cabecalho_tabela) for c in ("Região", "Pedidos", "Faturamento", "% do mês")]]
            for regiao_nome, info in ranking_regioes[:10]:
                pct = (info["faturamento"] / faturamento_total_regioes) * 100
                dados_regiao.append([
                    Paragraph(regiao_nome, estilo_celula),
                    Paragraph(str(len(info["pedidos"])), estilo_celula),
                    Paragraph(_fmt_moeda(info["faturamento"]), estilo_celula),
                    Paragraph(f"{pct:.1f}%", estilo_celula),
                ])
            largura_regiao = [largura_disponivel * 0.4, largura_disponivel * 0.2, largura_disponivel * 0.24, largura_disponivel * 0.16]
            tabela_regiao = Table(dados_regiao, colWidths=largura_regiao, repeatRows=1)
            tabela_regiao.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), COR_CABECALHO_BG),
                ("TEXTCOLOR", (0, 0), (-1, 0), COR_CABECALHO_TEXTO),
                ("LINEBELOW", (0, 0), (-1, 0), 1, colors.HexColor("#8fa8cc")),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#ced4da")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
                ("TOPPADDING", (0, 0), (-1, -1), 3.5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ]))
            elems.append(KeepTogether([Paragraph("Principais regiões de fornecimento", estilos["Heading4"]), tabela_regiao]))
            elems.append(Spacer(1, 6 * mm))

            # top 3 clientes dentro de cada uma das top 5 regiões (mesmo
            # espírito do Top 10 geral, só que segmentado por região) —
            # tabela única, na ordem do ranking de região acima e, dentro
            # dela, por faturamento do cliente.
            agregados_cliente_regiao = {}
            for l in linhas_b:
                regiao_nome = REGIAO_POR_UF.get(l.estado, "Não identificada")
                cliente_nome = l.cliente or "—"
                chave = (regiao_nome, cliente_nome)
                info = agregados_cliente_regiao.setdefault(chave, {"pedidos": set(), "faturamento": 0.0})
                info["pedidos"].add(l.pedido_id)
                info["faturamento"] += l.venda_total or 0

            dados_cli_regiao = [[Paragraph(c, estilo_cabecalho_tabela) for c in ("Região", "Cliente", "Pedidos", "Faturamento")]]
            for regiao_nome, _ in ranking_regioes[:5]:
                clientes_regiao = sorted(
                    (item for item in agregados_cliente_regiao.items() if item[0][0] == regiao_nome),
                    key=lambda kv: kv[1]["faturamento"], reverse=True,
                )[:3]
                for (_, cliente_nome), info in clientes_regiao:
                    dados_cli_regiao.append([
                        Paragraph(regiao_nome, estilo_celula),
                        Paragraph(cliente_nome, estilo_celula),
                        Paragraph(str(len(info["pedidos"])), estilo_celula),
                        Paragraph(_fmt_moeda(info["faturamento"]), estilo_celula),
                    ])
            if len(dados_cli_regiao) > 1:
                largura_cli_regiao = [largura_disponivel * 0.22, largura_disponivel * 0.38, largura_disponivel * 0.16, largura_disponivel * 0.24]
                tabela_cli_regiao = Table(dados_cli_regiao, colWidths=largura_cli_regiao, repeatRows=1)
                tabela_cli_regiao.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), COR_CABECALHO_BG),
                    ("TEXTCOLOR", (0, 0), (-1, 0), COR_CABECALHO_TEXTO),
                    ("LINEBELOW", (0, 0), (-1, 0), 1, colors.HexColor("#8fa8cc")),
                    ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#ced4da")),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("ALIGN", (2, 1), (-1, -1), "RIGHT"),
                    ("TOPPADDING", (0, 0), (-1, -1), 3.5),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
                    ("LEFTPADDING", (0, 0), (-1, -1), 4),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ]))
                elems.append(KeepTogether([Paragraph("Principais clientes por região", estilos["Heading4"]), tabela_cli_regiao]))
                elems.append(Spacer(1, 6 * mm))

        # ---- 1 seção por semana PCP ----
        for rotulo, linhas_semana in grupos_semana:
            inicio, fim = _intervalo_calendario_semana_pcp(mes_ano_b, rotulo)
            periodo_txt = f" ({inicio.strftime('%d/%m')} a {fim.strftime('%d/%m')})" if inicio and fim else ""
            pedidos_semana = {l.pedido_id for l in linhas_semana}
            faturamento_semana = sum(l.venda_total or 0 for l in linhas_semana)

            cabecalho_semana = Table(
                [[
                    Paragraph(f"<b>{rotulo or 'Sem semana definida'}</b>{periodo_txt}", ParagraphStyle("semana_titulo", parent=estilos["Normal"], fontSize=11, leading=13, textColor=COR_CABECALHO_TEXTO)),
                    Paragraph(f"{len(pedidos_semana)} pedido(s)", ParagraphStyle("semana_kpi", parent=estilos["Normal"], fontSize=10.5, leading=12, alignment=2, fontName="Helvetica-Bold")),
                    Paragraph(_fmt_moeda(faturamento_semana), ParagraphStyle("semana_kpi2", parent=estilos["Normal"], fontSize=12, leading=14, alignment=2, fontName="Helvetica-Bold")),
                ]],
                colWidths=[largura_disponivel * 0.56, largura_disponivel * 0.2, largura_disponivel * 0.24],
            )
            cabecalho_semana.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), COR_SEMANA_BG),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ("LEFTPADDING", (0, 0), (0, 0), 6),
            ]))

            dados_tabela = [[Paragraph(c, estilo_cabecalho_tabela) for c in cabecalho_tabela]]
            cores_linhas = [COR_CABECALHO_BG]
            linhas_subtotal = set()
            linhas_ordenadas = sorted(
                linhas_semana,
                key=lambda l: (l.data_inclusao_pedido or date.max, l.pedido_venda or "", l.item_id),
            )
            # Agrupa por pedido (pedido do Bruno, 21/09/2026: "agrupe os
            # pedidos e totalize o valor total do pedido... tenho um pedido
            # que possui 5 itens, some o valor de todo dele") — os itens já
            # vêm contíguos por pedido graças à ordenação acima, então basta
            # juntar item a item enquanto for o mesmo pedido_id.
            grupos_pedido = []
            for l in linhas_ordenadas:
                if grupos_pedido and grupos_pedido[-1][0] == l.pedido_id:
                    grupos_pedido[-1][1].append(l)
                else:
                    grupos_pedido.append((l.pedido_id, [l]))

            if modelo == "compacto":
                # Modelo compacto (pedido do Bruno, 21/09/2026: "nao quero
                # que detalhe e cite os itens, somente cite o pedido...
                # cada linha dentro da semana agrupada representara um
                # pedido de venda") — 1 linha por PEDIDO, nunca por item.
                # Campos que já são do próprio pedido (frete/UF/cidade/
                # datas de inclusão e do cliente) vêm do primeiro item;
                # campos que variam por item (produto, liberação, status)
                # são resumidos/agregados pra continuar "totalmente
                # completo" sem listar item a item.
                for _, itens_pedido in grupos_pedido:
                    primeiro = itens_pedido[0]
                    total_pedido = sum(l.venda_total or 0 for l in itens_pedido)
                    regiao_txt = REGIAO_POR_UF.get(primeiro.estado, "—")

                    produtos_unicos = []
                    for l in itens_pedido:
                        nome = l.descricao_produto or "—"
                        if nome not in produtos_unicos:
                            produtos_unicos.append(nome)
                    produtos_txt = produtos_unicos[0]
                    if len(produtos_unicos) > 1:
                        produtos_txt += f" (+{len(produtos_unicos) - 1} produto{'s' if len(produtos_unicos) > 2 else ''})"

                    datas_prevista = sorted({l.liberacao_prevista for l in itens_pedido if l.liberacao_prevista})
                    if not datas_prevista:
                        prevista_txt = "—"
                    elif len(datas_prevista) == 1:
                        prevista_txt = _formatar_data_br(datas_prevista[0])
                    else:
                        prevista_txt = f"{_formatar_data_br(datas_prevista[0])} a {_formatar_data_br(datas_prevista[-1])}"

                    qtd_com_real = sum(1 for l in itens_pedido if l.liberacao_real)
                    datas_real = sorted({l.liberacao_real for l in itens_pedido if l.liberacao_real})
                    if qtd_com_real == 0:
                        real_txt = "—"
                    elif qtd_com_real < len(itens_pedido):
                        real_txt = f"{qtd_com_real}/{len(itens_pedido)} liberados"
                    elif len(datas_real) == 1:
                        real_txt = _formatar_data_br(datas_real[0])
                    else:
                        real_txt = f"{_formatar_data_br(datas_real[0])} a {_formatar_data_br(datas_real[-1])}"

                    status_unicos = sorted({l.status_producao or "—" for l in itens_pedido})
                    status_txt = status_unicos[0] if len(status_unicos) == 1 else f"MISTO ({len(status_unicos)} status)"

                    dados_tabela.append([
                        Paragraph(primeiro.pedido_venda or "—", estilo_celula),
                        Paragraph(_quebravel(primeiro.cliente) or "—", estilo_celula),
                        Paragraph(_quebravel(produtos_txt), estilo_celula),
                        Paragraph(str(len(itens_pedido)), estilo_celula),
                        Paragraph(primeiro.frete or "—", estilo_celula),
                        Paragraph(f"{primeiro.estado or '—'} / {regiao_txt}", estilo_celula),
                        Paragraph(primeiro.cidade or "—", estilo_celula),
                        Paragraph(_formatar_data_br(primeiro.data_inclusao_pedido) or "—", estilo_celula),
                        Paragraph(_formatar_data_br(primeiro.data_cliente) or "—", estilo_celula),
                        Paragraph(prevista_txt, estilo_celula),
                        Paragraph(real_txt, estilo_celula),
                        Paragraph(status_txt, estilo_celula),
                        Paragraph(_fmt_moeda(total_pedido), estilo_celula),
                    ])
                    cores_linhas.append(COR_FINALIZADO_BG if qtd_com_real == len(itens_pedido) else colors.white)
            else:
                estilo_subtotal_rotulo = ParagraphStyle("subtotal_rotulo", parent=estilo_celula, fontName="Helvetica-Bold", textColor=COR_CABECALHO_TEXTO)
                estilo_subtotal_valor = ParagraphStyle("subtotal_valor", parent=estilo_celula, fontName="Helvetica-Bold", alignment=2)
                for _, itens_pedido in grupos_pedido:
                    for l in itens_pedido:
                        qtd = l.quantidade or 0
                        qtd_txt = int(qtd) if qtd == int(qtd) else qtd
                        regiao_txt = REGIAO_POR_UF.get(l.estado, "—")
                        dados_tabela.append([
                            Paragraph(l.pedido_venda or "—", estilo_celula),
                            Paragraph(_quebravel(l.cliente) or "—", estilo_celula),
                            Paragraph(_quebravel(l.descricao_produto) or "—", estilo_celula),
                            Paragraph(str(qtd_txt), estilo_celula),
                            Paragraph(l.frete or "—", estilo_celula),
                            Paragraph(f"{l.estado or '—'} / {regiao_txt}", estilo_celula),
                            Paragraph(l.cidade or "—", estilo_celula),
                            Paragraph(_formatar_data_br(l.data_inclusao_pedido) or "—", estilo_celula),
                            Paragraph(_formatar_data_br(l.data_cliente) or "—", estilo_celula),
                            Paragraph(_formatar_data_br(l.liberacao_prevista) or "—", estilo_celula),
                            Paragraph(_formatar_data_br(l.liberacao_real) or "—", estilo_celula),
                            Paragraph(l.status_producao or "—", estilo_celula),
                            Paragraph(_fmt_moeda(l.venda_total), estilo_celula),
                        ])
                        cores_linhas.append(COR_FINALIZADO_BG if l.liberacao_real else colors.white)

                    # Subtotal do pedido — só quando há mais de 1 item (com 1 só
                    # item, a própria linha já mostra o total, repetir seria
                    # redundante); "bem simples" como pedido pelo Bruno.
                    if len(itens_pedido) > 1:
                        primeiro = itens_pedido[0]
                        total_pedido = sum(l.venda_total or 0 for l in itens_pedido)
                        idx_linha = len(dados_tabela)
                        dados_tabela.append([
                            Paragraph(f"Total do pedido {primeiro.pedido_venda or '—'} — {primeiro.cliente or '—'} ({len(itens_pedido)} itens)", estilo_subtotal_rotulo),
                            "", "", "", "", "", "", "", "", "", "", "",
                            Paragraph(_fmt_moeda(total_pedido), estilo_subtotal_valor),
                        ])
                        cores_linhas.append(colors.HexColor("#e9edf5"))
                        linhas_subtotal.add(idx_linha)

            tabela_semana = Table(dados_tabela, colWidths=larguras_colunas, repeatRows=1)
            estilo_tabela = [
                ("BACKGROUND", (0, 0), (-1, 0), COR_CABECALHO_BG),
                ("TEXTCOLOR", (0, 0), (-1, 0), COR_CABECALHO_TEXTO),
                ("LINEBELOW", (0, 0), (-1, 0), 1, colors.HexColor("#8fa8cc")),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#ced4da")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING", (0, 0), (-1, -1), 3.5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ]
            for i, cor in enumerate(cores_linhas):
                if i == 0:
                    continue
                estilo_tabela.append(("BACKGROUND", (0, i), (-1, i), cor))
            for idx_linha in linhas_subtotal:
                estilo_tabela.append(("SPAN", (0, idx_linha), (-2, idx_linha)))
                estilo_tabela.append(("VALIGN", (0, idx_linha), (-1, idx_linha), "MIDDLE"))
                estilo_tabela.append(("TOPPADDING", (0, idx_linha), (-1, idx_linha), 4))
                estilo_tabela.append(("BOTTOMPADDING", (0, idx_linha), (-1, idx_linha), 4))
                estilo_tabela.append(("LINEABOVE", (0, idx_linha), (-1, idx_linha), 0.6, colors.HexColor("#8fa8cc")))
                estilo_tabela.append(("LINEBELOW", (0, idx_linha), (-1, idx_linha), 0.6, colors.HexColor("#8fa8cc")))
            tabela_semana.setStyle(TableStyle(estilo_tabela))

            elems.append(cabecalho_semana)
            elems.append(tabela_semana)
            elems.append(Spacer(1, 6 * mm))

        if not grupos_semana:
            elems.append(Paragraph("Nenhum pedido com Planejamento semanal (PCP) preenchido nesse mês.", estilos["Normal"]))
            elems.append(Spacer(1, 6 * mm))

        # ---- total do mês (rodapé em destaque, pedido explícito do Bruno:
        # "e a somatario total, tanto de faturamento e tambem do total de
        # numero de pedidos") — usa a MESMA cor do banner do bloco quando há
        # 2 meses, pra reforçar visualmente de qual mês é aquele total ----
        tabela_total = Table(
            [[
                Paragraph(f"TOTAL — {titulo_mes_b}", ParagraphStyle("total_titulo", parent=estilos["Normal"], fontSize=12.5, leading=15, textColor=colors.white, fontName="Helvetica-Bold")),
                Paragraph(f"{len(pedidos_mes)} pedido(s)", ParagraphStyle("total_kpi", parent=estilos["Normal"], fontSize=12.5, leading=15, alignment=2, textColor=colors.white, fontName="Helvetica-Bold")),
                Paragraph(_fmt_moeda(faturamento_mes), ParagraphStyle("total_kpi2", parent=estilos["Normal"], fontSize=15, leading=17, alignment=2, textColor=colors.white, fontName="Helvetica-Bold")),
            ]],
            colWidths=[largura_disponivel * 0.56, largura_disponivel * 0.2, largura_disponivel * 0.24],
        )
        tabela_total.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), cor_banner if len(blocos) > 1 else COR_TOTAL_BG),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 7),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ("LEFTPADDING", (0, 0), (0, 0), 6),
        ]))
        elems.append(tabela_total)
        if not linhas_b:
            elems.append(Spacer(1, 4 * mm))
            elems.append(Paragraph("Nenhum pedido encontrado com os filtros aplicados nesse mês.", estilos["Normal"]))

        return elems

    for idx, b in enumerate(blocos):
        elementos.extend(_construir_bloco(b, cores_bloco[idx % len(cores_bloco)]))
        if idx < len(blocos) - 1:
            elementos.append(Spacer(1, 5 * mm))

    doc.build(elementos)
    buffer.seek(0)
    resposta = Response(buffer.getvalue(), mimetype="application/pdf")
    sufixo_nome = "_".join(f"{b['mes_ano'][0]}-{b['mes_ano'][1]:02d}" for b in blocos)
    sufixo_modelo = "_compacto" if modelo == "compacto" else ""
    nome_arquivo = f"planejamento_mensal_pcp{sufixo_modelo}_{sufixo_nome}.pdf"
    resposta.headers["Content-Disposition"] = f"attachment; filename={nome_arquivo}"
    return resposta


# ----------------------------------------------------------------------
# P&D — Pesquisa e Desenvolvimento (Fase 14). Segue o mesmo padrão de
# Qualidade/RNC acima: tabela própria, controle manual, funções auxiliares
# separadas de filtro/listagem/dashboard/form pra não misturar com nenhuma
# outra área do sistema.
# ----------------------------------------------------------------------
def _campos_form_pd(f):
    """Lê e converte todos os campos do formulário de Projeto de P&D (novo/
    editar) — mesma estrutura de _campos_form_rnc."""
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

    resultado_esperado = "; ".join(f.getlist("resultado_esperado")) or None

    return dict(
        codigo=_txt("codigo"),
        nome=_txt("nome"),
        descricao=_txt("descricao"),
        objetivo=_txt("objetivo"),
        justificativa=_txt("justificativa"),
        categoria=_txt("categoria"),
        prioridade=_txt("prioridade"),
        responsavel=_txt("responsavel"),
        participantes=_txt("participantes"),
        cliente=_txt("cliente"),
        produto=_txt("produto"),
        fornecedor=_txt("fornecedor"),
        area_envolvida=_txt("area_envolvida"),
        etapa_atual=_txt("etapa_atual") or "Ideia",
        percentual_conclusao=_num_int("percentual_conclusao") or 0,
        data_inicio=_parse_data_form(f.get("data_inicio")),
        data_prevista_conclusao=_parse_data_form(f.get("data_prevista_conclusao")),
        data_real_conclusao=_parse_data_form(f.get("data_real_conclusao")),
        proxima_entrega=_txt("proxima_entrega"),
        data_proxima_entrega=_parse_data_form(f.get("data_proxima_entrega")),
        responsavel_proxima_entrega=_txt("responsavel_proxima_entrega"),
        custo_previsto=_num_float("custo_previsto"),
        custo_realizado=_num_float("custo_realizado"),
        investimento_previsto=_num_float("investimento_previsto"),
        investimento_realizado=_num_float("investimento_realizado"),
        economia_prevista=_num_float("economia_prevista"),
        economia_realizada=_num_float("economia_realizada"),
        resultado_esperado=resultado_esperado,
        resultado_obtido=_txt("resultado_obtido"),
        problema=_txt("problema"),
        solucao=_txt("solucao"),
        licoes_aprendidas=_txt("licoes_aprendidas"),
        observacoes_gerais=_txt("observacoes_gerais"),
    )


_CAMPOS_DATA_PD = [
    "data_inicio", "data_prevista_conclusao", "data_real_conclusao", "data_proxima_entrega",
]


def _pd_para_form_dict(projeto):
    """Converte um ProjetoPD em dict de strings prontas pra repopular o
    formulário HTML — usado na tela de edição (GET), igual a _rnc_para_form_dict."""
    campos = [
        "codigo", "nome", "descricao", "objetivo", "justificativa", "categoria", "prioridade",
        "responsavel", "participantes", "cliente", "produto", "fornecedor", "area_envolvida",
        "etapa_atual", "percentual_conclusao", "data_inicio", "data_prevista_conclusao",
        "data_real_conclusao", "proxima_entrega", "data_proxima_entrega", "responsavel_proxima_entrega",
        "custo_previsto", "custo_realizado", "investimento_previsto", "investimento_realizado",
        "economia_prevista", "economia_realizada", "resultado_obtido", "problema", "solucao",
        "licoes_aprendidas", "observacoes_gerais",
    ]
    valores = {}
    for campo in campos:
        v = getattr(projeto, campo)
        if v is None:
            valores[campo] = ""
        elif campo in _CAMPOS_DATA_PD:
            valores[campo] = v.isoformat()
        else:
            valores[campo] = str(v)
    valores["resultado_esperado"] = projeto.resultado_esperado_lista
    return valores


def _filtrar_projetos_pd(args):
    busca = args.get("busca", "").strip()
    etapa = args.getlist("etapa")
    categoria = args.getlist("categoria")
    responsavel = args.get("responsavel", "").strip()
    prioridade = args.getlist("prioridade")
    cliente = args.get("cliente", "").strip()
    produto = args.get("produto", "").strip()
    fornecedor = args.get("fornecedor", "").strip()
    area_envolvida = args.get("area_envolvida", "").strip()
    apenas_atrasados = args.get("apenas_atrasados", "").strip()
    apenas_criticos = args.get("apenas_criticos", "").strip()

    query = ProjetoPD.query
    if busca:
        termo = f"%{busca}%"
        query = query.filter(
            or_(
                ProjetoPD.nome.ilike(termo),
                ProjetoPD.codigo.ilike(termo),
                ProjetoPD.cliente.ilike(termo),
                ProjetoPD.produto.ilike(termo),
                ProjetoPD.fornecedor.ilike(termo),
                ProjetoPD.descricao.ilike(termo),
            )
        )
    if etapa:
        query = query.filter(ProjetoPD.etapa_atual.in_(etapa))
    if categoria:
        query = query.filter(ProjetoPD.categoria.in_(categoria))
    if responsavel:
        query = query.filter(ProjetoPD.responsavel.ilike(f"%{responsavel}%"))
    if prioridade:
        query = query.filter(ProjetoPD.prioridade.in_(prioridade))
    if cliente:
        query = query.filter(ProjetoPD.cliente.ilike(f"%{cliente}%"))
    if produto:
        query = query.filter(ProjetoPD.produto.ilike(f"%{produto}%"))
    if fornecedor:
        query = query.filter(ProjetoPD.fornecedor.ilike(f"%{fornecedor}%"))
    if area_envolvida:
        query = query.filter(ProjetoPD.area_envolvida.ilike(f"%{area_envolvida}%"))
    if apenas_atrasados == "1":
        query = query.filter(
            ProjetoPD.data_prevista_conclusao.isnot(None),
            ProjetoPD.data_prevista_conclusao < date.today(),
            ProjetoPD.etapa_atual != "Concluído",
        )
    if apenas_criticos == "1":
        query = query.filter(
            ProjetoPD.prioridade == "ALTA",
            ProjetoPD.data_prevista_conclusao.isnot(None),
            ProjetoPD.data_prevista_conclusao < date.today(),
            ProjetoPD.etapa_atual != "Concluído",
        )

    query = query.order_by(ProjetoPD.data_prevista_conclusao.asc().nullslast(), ProjetoPD.id.desc())

    filtros = dict(
        busca=busca, etapa=etapa, categoria=categoria, responsavel=responsavel,
        prioridade=prioridade, cliente=cliente, produto=produto, fornecedor=fornecedor,
        area_envolvida=area_envolvida, apenas_atrasados=apenas_atrasados, apenas_criticos=apenas_criticos,
    )
    return query, filtros


def _pd_opcoes_filtro(campo, opcoes_curadas):
    """Mesmo espírito de _rnc_opcoes_filtro: combina valores já cadastrados
    (texto livre) com a lista de sugestão curada, pra nenhum valor real ficar
    de fora do filtro."""
    coluna = getattr(ProjetoPD, campo)
    existentes = [v for (v,) in db.session.query(coluna).filter(coluna.isnot(None)).distinct()]
    existentes_lower = {v.lower() for v in existentes}
    extras = [op for op in opcoes_curadas if op.lower() not in existentes_lower]
    return sorted(existentes + extras, key=lambda s: s.lower())


def _linhas_projetos_pd(args):
    page = args.get("page", 1, type=int)
    query, filtros = _filtrar_projetos_pd(args)
    total_filtrado = query.count()
    total_paginas = max(1, (total_filtrado + PAGE_SIZE - 1) // PAGE_SIZE)
    pagina = query.offset((page - 1) * PAGE_SIZE).limit(PAGE_SIZE).all()
    return pagina, page, total_paginas, total_filtrado, filtros


def _dashboard_pd():
    """Recalcula ao vivo os indicadores do Dashboard de P&D — nada fica
    pré-calculado/guardado, sempre em dia com o que estiver cadastrado
    (mesmo espírito de _dashboard_rnc_qualidade)."""
    projetos = ProjetoPD.query.all()
    total = len(projetos)
    ativos = [p for p in projetos if not p.concluido]
    atrasados = [p for p in projetos if p.atrasado]
    criticos = [p for p in projetos if p.critico]
    concluidos = [p for p in projetos if p.concluido]

    no_prazo = [p for p in ativos if not p.atrasado]
    pct_no_prazo = round((len(no_prazo) / len(ativos)) * 100, 1) if ativos else 0
    pct_atrasados = round((len(atrasados) / len(ativos)) * 100, 1) if ativos else 0

    duracoes = [
        (p.data_real_conclusao - p.data_inicio).days
        for p in concluidos if p.data_inicio and p.data_real_conclusao
    ]
    lead_time_medio = round(sum(duracoes) / len(duracoes), 1) if duracoes else 0

    todos_testes = TesteProjetoPD.query.all()
    testes_com_resultado = [t for t in todos_testes if t.resultado in ("Aprovado", "Reprovado")]
    testes_aprovados = [t for t in todos_testes if t.resultado == "Aprovado"]
    taxa_aprovacao_testes = round((len(testes_aprovados) / len(testes_com_resultado)) * 100, 1) if testes_com_resultado else 0

    projetos_homologados_ou_alem = [
        p for p in projetos if PD_ETAPA_OPCOES.index(p.etapa_atual) >= PD_ETAPA_OPCOES.index("Homologação")
    ]
    taxa_homologacao = round((len(projetos_homologados_ou_alem) / total) * 100, 1) if total else 0

    investimento_previsto = sum(p.investimento_previsto or 0 for p in projetos)
    investimento_realizado = sum(p.investimento_realizado or 0 for p in projetos)
    economia_prevista = sum(p.economia_prevista or 0 for p in projetos)
    economia_realizada = sum(p.economia_realizada or 0 for p in projetos)
    roi_geral = (
        round(((economia_realizada - investimento_realizado) / investimento_realizado) * 100, 1)
        if investimento_realizado else None
    )

    def _quebra_por(atributo, opcoes_ordem=None):
        contagem = {}
        for p in projetos:
            chave = getattr(p, atributo) or "—"
            contagem[chave] = contagem.get(chave, 0) + 1
        if opcoes_ordem:
            chaves = list(opcoes_ordem) + sorted(k for k in contagem if k not in opcoes_ordem and k != "—")
            if "—" in contagem:
                chaves.append("—")
            return [{"chave": k, "total": contagem.get(k, 0)} for k in chaves if k in contagem or k in opcoes_ordem]
        return sorted(({"chave": k, "total": v} for k, v in contagem.items()), key=lambda d: -d["total"])

    return {
        "total": total,
        "ativos": len(ativos),
        "em_desenvolvimento": sum(1 for p in projetos if p.etapa_atual == "Desenvolvimento"),
        "em_teste": sum(1 for p in projetos if p.etapa_atual == "Teste"),
        "em_validacao": sum(1 for p in projetos if p.etapa_atual == "Validação"),
        "em_homologacao": sum(1 for p in projetos if p.etapa_atual == "Homologação"),
        "concluidos": len(concluidos),
        "atrasados": len(atrasados),
        "criticos": len(criticos),
        "pct_no_prazo": pct_no_prazo,
        "pct_atrasados": pct_atrasados,
        "lead_time_medio": lead_time_medio,
        "taxa_aprovacao_testes": taxa_aprovacao_testes,
        "taxa_homologacao": taxa_homologacao,
        "investimento_previsto": investimento_previsto,
        "investimento_realizado": investimento_realizado,
        "economia_prevista": economia_prevista,
        "economia_realizada": economia_realizada,
        "roi_geral": roi_geral,
        "por_etapa": _quebra_por("etapa_atual", PD_ETAPA_OPCOES),
        "por_categoria": _quebra_por("categoria", PD_CATEGORIA_OPCOES),
        "por_responsavel": _quebra_por("responsavel"),
        "por_prioridade": _quebra_por("prioridade", PRIORIDADE_OPCOES),
        "projetos_atrasados": sorted(atrasados, key=lambda p: p.data_prevista_conclusao)[:10],
        "projetos_criticos": criticos[:10],
    }


def _cronograma_pd():
    """Monta as linhas do Cronograma Geral (Gantt simplificado, sem
    biblioteca externa — barras posicionadas por porcentagem dentro de uma
    faixa de datas comum). Só entram projetos com data de início preenchida;
    os demais são contados à parte (sem_data) pra não sumirem silenciosamente."""
    projetos = ProjetoPD.query.filter(ProjetoPD.data_inicio.isnot(None)).order_by(ProjetoPD.data_inicio.asc()).all()
    sem_data = ProjetoPD.query.filter(ProjetoPD.data_inicio.is_(None)).count()
    hoje = date.today()

    if not projetos:
        return {"linhas": [], "sem_data": sem_data, "marcadores_mes": [], "hoje_pct": None}

    def _fim_projeto(p):
        fim = p.data_real_conclusao or p.data_prevista_conclusao or p.data_inicio
        return fim if fim >= p.data_inicio else p.data_inicio

    inicio_min = min(p.data_inicio for p in projetos)
    fim_max = max([_fim_projeto(p) for p in projetos] + [hoje])
    total_dias = max((fim_max - inicio_min).days, 1)

    def _pct(d):
        return max(0.0, min(100.0, ((d - inicio_min).days / total_dias) * 100))

    linhas = []
    for p in projetos:
        fim = _fim_projeto(p)
        marcos = []
        for t in p.testes:
            data_evento = t.data_realizada or t.data_planejada
            if data_evento and inicio_min <= data_evento <= fim_max:
                info = PD_TESTE_RESULTADO_INFO.get(t.resultado, {})
                marcos.append({
                    "pos_pct": _pct(data_evento),
                    "label": f"Teste {t.numero}".strip() if t.numero else "Teste",
                    "cor": info.get("cor", "secondary"),
                    "emoji": info.get("emoji", ""),
                })
        linhas.append({
            "projeto": p,
            "offset_pct": _pct(p.data_inicio),
            "largura_pct": max(_pct(fim) - _pct(p.data_inicio), 0.6),
            "marcos": marcos,
        })

    marcadores_mes = []
    cursor = date(inicio_min.year, inicio_min.month, 1)
    while cursor <= fim_max:
        marcadores_mes.append({"pos_pct": _pct(cursor), "label": f"{MESES_PT[cursor.month - 1]}/{cursor.year}"})
        if cursor.month == 12:
            cursor = date(cursor.year + 1, 1, 1)
        else:
            cursor = date(cursor.year, cursor.month + 1, 1)

    return {
        "linhas": linhas, "sem_data": sem_data, "marcadores_mes": marcadores_mes,
        "hoje_pct": _pct(hoje), "inicio_min": inicio_min, "fim_max": fim_max,
    }


def _filtrar_testes_pd(args):
    """Visão global de Testes & Validações — mesmos testes já vistos dentro
    de cada projeto (aba Testes & Validações), aqui juntados de todos os
    projetos numa lista só, com filtro por projeto/resultado/responsável."""
    projeto_id = args.get("projeto_id", type=int)
    resultado = args.getlist("resultado")
    responsavel = args.get("responsavel", "").strip()
    apenas_atrasados = args.get("apenas_atrasados", "").strip()

    query = TesteProjetoPD.query.join(ProjetoPD, TesteProjetoPD.projeto_id == ProjetoPD.id)
    if projeto_id:
        query = query.filter(TesteProjetoPD.projeto_id == projeto_id)
    if resultado:
        query = query.filter(TesteProjetoPD.resultado.in_(resultado))
    if responsavel:
        query = query.filter(TesteProjetoPD.responsavel.ilike(f"%{responsavel}%"))
    if apenas_atrasados == "1":
        query = query.filter(
            TesteProjetoPD.data_planejada.isnot(None),
            TesteProjetoPD.data_planejada < date.today(),
            TesteProjetoPD.data_realizada.is_(None),
        )
    query = query.order_by(TesteProjetoPD.data_planejada.desc().nullslast(), TesteProjetoPD.id.desc())

    filtros = dict(projeto_id=projeto_id, resultado=resultado, responsavel=responsavel, apenas_atrasados=apenas_atrasados)
    return query, filtros


def _custos_pd():
    """Totais previsto x realizado / economia / ROI de todos os projetos —
    página própria de analytics financeiro de P&D (Bruno pediu "Custos &
    Resultados" separado do Dashboard geral)."""
    projetos = ProjetoPD.query.order_by(ProjetoPD.data_prevista_conclusao.asc().nullslast(), ProjetoPD.id.desc()).all()

    totais = {
        "custo_previsto": sum(p.custo_previsto or 0 for p in projetos),
        "custo_realizado": sum(p.custo_realizado or 0 for p in projetos),
        "investimento_previsto": sum(p.investimento_previsto or 0 for p in projetos),
        "investimento_realizado": sum(p.investimento_realizado or 0 for p in projetos),
        "economia_prevista": sum(p.economia_prevista or 0 for p in projetos),
        "economia_realizada": sum(p.economia_realizada or 0 for p in projetos),
    }
    totais["roi_geral"] = (
        round(((totais["economia_realizada"] - totais["investimento_realizado"]) / totais["investimento_realizado"]) * 100, 1)
        if totais["investimento_realizado"] else None
    )

    acima_do_previsto = [
        p for p in projetos
        if (p.custo_realizado is not None and p.custo_previsto is not None and p.custo_realizado > p.custo_previsto)
        or (p.investimento_realizado is not None and p.investimento_previsto is not None and p.investimento_realizado > p.investimento_previsto)
    ]
    return {"projetos": projetos, "totais": totais, "acima_do_previsto": acima_do_previsto}


# Limiares (em dias) usados só pelos alertas de "sem atualização"/"parado" —
# ajustáveis aqui sem mexer no resto da lógica, caso o Bruno peça outro valor.
PD_DIAS_SEM_ATUALIZACAO = 15
PD_DIAS_PARADO = 30


def _alertas_pd():
    """Consolida os alertas de P&D pedidos pelo Bruno (atrasado, prazo
    próximo, teste atrasado, projeto sem atualização, projeto parado, custo
    acima do previsto, homologação em andamento) — só considera projetos
    ainda não concluídos, igual ao resto do sistema já faz pra Pedido."""
    projetos = ProjetoPD.query.filter(ProjetoPD.etapa_atual != "Concluído").all()
    hoje = date.today()
    agora = datetime.utcnow()

    sem_atualizacao, parados = [], []
    for p in projetos:
        referencia = p.atualizado_em or p.criado_em
        if not referencia:
            continue
        dias = (agora - referencia).days
        if dias >= PD_DIAS_PARADO:
            parados.append(p)
        elif dias >= PD_DIAS_SEM_ATUALIZACAO:
            sem_atualizacao.append(p)

    custo_acima = [
        p for p in projetos
        if (p.custo_realizado is not None and p.custo_previsto is not None and p.custo_realizado > p.custo_previsto)
        or (p.investimento_realizado is not None and p.investimento_previsto is not None and p.investimento_realizado > p.investimento_previsto)
    ]

    testes_atrasados = (
        TesteProjetoPD.query.join(ProjetoPD, TesteProjetoPD.projeto_id == ProjetoPD.id)
        .filter(
            TesteProjetoPD.data_planejada.isnot(None),
            TesteProjetoPD.data_planejada < hoje,
            TesteProjetoPD.data_realizada.is_(None),
        )
        .all()
    )

    return {
        "atrasados": [p for p in projetos if p.atrasado],
        "prazo_proximo": [p for p in projetos if p.prazo_proximo],
        "homologacao": [p for p in projetos if p.etapa_atual == "Homologação"],
        "sem_atualizacao": sem_atualizacao,
        "parados": parados,
        "custo_acima": custo_acima,
        "testes_atrasados": testes_atrasados,
    }


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


# ---------------------------------------------------------------------------
# Relatório Excel "gerencial" da Listagem Geral (pedido do Bruno, 21/09/2026:
# "quero que voce melhora a geração de relatorio via excel... quero que
# inclua todo o contexto do pedido, itens, quantidade, produto, valores e
# faturamento, datas... inclua tambem a coluna do planejamento semanal e
# coluna tmb da referencia do mes... totalmente intuitivo e dinamico,
# relatorio padrão gerencial") — substitui o export .xlsx flat (1 linha por
# PEDIDO, poucas colunas) que existia antes. Reaproveita a MESMA lista já
# achatada por ITEM que a tela usa (_linhas_listagem_geral / _LinhaListagemGeral)
# — o relatório nunca diverge do que a Listagem Geral está mostrando com o
# filtro ativo. Formatado como Tabela nativa do Excel (cabeçalho fixo,
# autofiltro, zebra) em vez de uma planilha crua — é isso que fica
# "intuitivo e dinâmico" pra quem abre no Excel: já chega pronta pra
# ordenar/filtrar/explorar sem precisar reformatar nada.
# ---------------------------------------------------------------------------
def _linhas_export_listagem_geral(linhas, inspecoes_rdim):
    """1 linha por item (produto) — mesmo grão da tela — com todo o contexto
    comercial, de produção, faturamento, qualidade e planejamento PCP.
    `inspecoes_rdim` é o dict item_pedido_id -> InspecaoFinal já usado pela
    própria tela (ver _inspecoes_rdim_por_item), reaproveitado aqui pra
    trazer a coluna "Qualidade" sem rodar a mesma query 2x."""
    cabecalho = [
        "Pedido", "Cliente", "CNPJ", "Vendedor", "País", "UF", "Cidade", "Frete", "Prioridade",
        "Produto", "Quantidade", "Venda unitário (R$)", "Venda total item (R$)", "Venda total pedido (R$)",
        "Estação", "Status produção", "Qualidade (RDIM)",
        "Nº nota fiscal", "Valor faturado (R$)",
        "Data de inclusão", "Data do cliente (prazo)",
        "Início produção", "Início inspeção", "Término inspeção",
        "Liberação prevista", "Liberação real", "Liberação faturamento",
        "Planejamento semanal (PCP)", "Mês de referência (Planej. semanal)",
        "Dias (negativo = atrasado)",
    ]
    linhas_export = []
    for l in linhas:
        item = l.item
        insp = inspecoes_rdim.get(l.item_id)
        qualidade = RDIM_RESULTADO_LABELS.get(insp.resultado, insp.resultado) if insp else ""
        mes_ano = _mes_ano_da_semana_pcp(l.planejamento_semanal)
        mes_referencia = f"{MESES_PT[mes_ano[1] - 1]}/{mes_ano[0]}" if mes_ano else ""
        _, dias_atraso = l.semaforo
        linhas_export.append([
            l.pedido_venda or "",
            l.cliente or "",
            l.pedido.cnpj or "",
            l.vendedor or "",
            l.pais or "",
            l.estado or "",
            l.cidade or "",
            l.frete or "",
            l.prioridade or "",
            l.descricao_produto or "",
            l.quantidade or 0,
            round(l.venda_unidade or 0, 2),
            round(l.venda_total or 0, 2),
            round(l.venda_total_pedido or 0, 2),
            l.estacao or "",
            l.status_producao or "",
            qualidade,
            item.numero_nota_fiscal or "",
            round(item.valor_faturamento_realizado or 0, 2),
            l.data_inclusao_pedido,
            l.data_cliente,
            item.inicio_producao,
            item.inicio_inspecao,
            item.termino_inspecao,
            l.liberacao_prevista,
            l.liberacao_real,
            item.liberacao_faturamento,
            l.planejamento_semanal or "",
            mes_referencia,
            dias_atraso if dias_atraso is not None else "",
        ])
    return cabecalho, linhas_export


def _resumo_export_por_semana_pcp(linhas):
    """Agrega as linhas (já achatadas por item, com o filtro da tela já
    aplicado) por Planejamento semanal (PCP) — vira a aba "Resumo por
    Semana PCP" do relatório gerencial, visão rápida sem precisar montar
    tabela dinâmica no Excel. Ordem cronológica (mesma _chave_semana_pcp que
    já ordena essa coluna na tela); itens SEM planejamento ainda (rótulo
    None — exatamente os que o quadrante "Sem planejamento PCP" aponta)
    ficam agrupados num grupo próprio, sempre por último."""
    grupos = {}
    for l in linhas:
        grupos.setdefault(l.planejamento_semanal, []).append(l)

    def chave_ordenacao(rotulo):
        return _chave_semana_pcp(rotulo) or (9999, 99, 0)

    resultado = []
    for rotulo in sorted(grupos.keys(), key=chave_ordenacao):
        linhas_grupo = grupos[rotulo]
        mes_ano = _mes_ano_da_semana_pcp(rotulo)
        mes_referencia = f"{MESES_PT[mes_ano[1] - 1]}/{mes_ano[0]}" if mes_ano else ""
        pedidos_distintos = {l.pedido_id for l in linhas_grupo}
        resultado.append({
            "semana": rotulo or "Sem planejamento",
            "mes_referencia": mes_referencia,
            "pedidos": len(pedidos_distintos),
            "itens": len(linhas_grupo),
            "quantidade_total": sum(l.quantidade or 0 for l in linhas_grupo),
            "venda_total": sum(l.venda_total or 0 for l in linhas_grupo),
            "faturado_total": sum((l.item.valor_faturamento_realizado or 0) for l in linhas_grupo),
        })
    return resultado


def _preencher_aba_relatorio_gerencial(ws, nome_tabela, cabecalho, linhas, colunas_moeda=(), colunas_data=()):
    """Escreve `cabecalho`/`linhas` numa aba já formatada como Tabela nativa
    do Excel — cabeçalho em negrito com fundo, zebra automática (estilo da
    Tabela), autofiltro em toda a extensão, 1ª linha congelada (cabeçalho
    sempre visível ao rolar), moeda/data formatadas e largura de coluna
    automática. `colunas_moeda`/`colunas_data` são conjuntos com o número da
    coluna (1 = primeira) que devem levar cada formato numérico."""
    ws.append(cabecalho)
    fundo_cabecalho = PatternFill(start_color="1B2A4A", end_color="1B2A4A", fill_type="solid")
    for celula in ws[1]:
        celula.font = Font(bold=True, color="FFFFFF")
        celula.fill = fundo_cabecalho
        celula.alignment = Alignment(vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"

    for linha in linhas:
        ws.append(linha)

    ultima_linha = len(linhas) + 1
    for indice in colunas_moeda:
        letra = get_column_letter(indice)
        for celula in ws[f"{letra}2:{letra}{ultima_linha}"]:
            celula[0].number_format = '"R$" #,##0.00'
    for indice in colunas_data:
        letra = get_column_letter(indice)
        for celula in ws[f"{letra}2:{letra}{ultima_linha}"]:
            celula[0].number_format = "DD/MM/YYYY"

    for coluna in ws.columns:
        valores = [len(str(c.value)) for c in coluna if c.value is not None]
        largura = max(valores) if valores else 10
        ws.column_dimensions[coluna[0].column_letter].width = min(max(largura + 2, 10), 42)
    ws.row_dimensions[1].height = 28

    ultima_coluna_letra = get_column_letter(len(cabecalho))
    tabela = Table(displayName=nome_tabela, ref=f"A1:{ultima_coluna_letra}{ultima_linha}")
    tabela.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False,
    )
    ws.add_table(tabela)


def _responder_xlsx_listagem_geral(linhas):
    """Monta o relatório .xlsx "padrão gerencial" da Listagem Geral inteiro
    (2 abas — detalhe por item + resumo por semana PCP, ver funções acima)
    e devolve como download, mesmo padrão de resposta de _responder_xlsx."""
    inspecoes_rdim = _inspecoes_rdim_por_item([l.item_id for l in linhas])
    cabecalho_detalhe, linhas_detalhe = _linhas_export_listagem_geral(linhas, inspecoes_rdim)
    resumo_semanas = _resumo_export_por_semana_pcp(linhas)

    wb = Workbook()
    ws_detalhe = wb.active
    ws_detalhe.title = "Pedidos"
    _preencher_aba_relatorio_gerencial(
        ws_detalhe, "TabelaPedidos", cabecalho_detalhe, linhas_detalhe,
        colunas_moeda={12, 13, 14, 19}, colunas_data={20, 21, 22, 23, 24, 25, 26, 27},
    )

    ws_resumo = wb.create_sheet("Resumo por Semana PCP")
    cabecalho_resumo = [
        "Semana (PCP)", "Mês de referência", "Pedidos distintos", "Itens (produtos)",
        "Quantidade total", "Venda total (R$)", "Faturado total (R$)",
    ]
    linhas_resumo = [
        [r["semana"], r["mes_referencia"], r["pedidos"], r["itens"],
         r["quantidade_total"], round(r["venda_total"], 2), round(r["faturado_total"], 2)]
        for r in resumo_semanas
    ]
    _preencher_aba_relatorio_gerencial(
        ws_resumo, "TabelaResumoSemanal", cabecalho_resumo, linhas_resumo, colunas_moeda={6, 7},
    )

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    resposta = Response(
        buffer.getvalue(),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    resposta.headers["Content-Disposition"] = "attachment; filename=listagem_geral_gerencial.xlsx"
    return resposta


def _construir_backup_pedidos_wb():
    """Monta um Workbook com TODAS as colunas de cada tabela apagada por
    "/admin/zerar-dados" (uma aba cada), lendo as colunas direto do
    mapeamento do SQLAlchemy (não uma lista escrita à mão) — assim não corre
    o risco de esquecer um campo novo que apareça no futuro. Usado como rede
    de segurança antes da exclusão permanente.

    Pedido do Bruno (03/09/2026): "zerar dados" passou a apagar também
    Qualidade (RNC + RDIM) e P&D, então o backup precisa cobrir as mesmas
    tabelas — senão a "rede de segurança" fica incompleta bem na hora que
    mais importa."""
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

    # ---- Produção + Operação (já existia) ----
    _add_sheet("Pedidos", Pedido, Pedido.query.order_by(Pedido.id).all())
    _add_sheet("Itens", ItemPedido, ItemPedido.query.order_by(ItemPedido.id).all())
    _add_sheet("Gestao Operacao", PedidoOperacao, PedidoOperacao.query.order_by(PedidoOperacao.id).all())
    _add_sheet("Programacao", Programacao, Programacao.query.order_by(Programacao.id).all())
    _add_sheet("Historico Alteracoes", HistoricoAlteracao, HistoricoAlteracao.query.order_by(HistoricoAlteracao.id).all())

    # ---- Qualidade: RNC + RDIM (novo, 03/09/2026) ----
    _add_sheet("RNC Qualidade", RncQualidade, RncQualidade.query.order_by(RncQualidade.id).all())
    _add_sheet("Inspecoes RDIM", InspecaoFinal, InspecaoFinal.query.order_by(InspecaoFinal.id).all())
    _add_sheet("RDIM Medicoes", RdimMedicao, RdimMedicao.query.order_by(RdimMedicao.id).all())
    _add_sheet("RDIM Pecas Desvio", RdimPecaDesvio, RdimPecaDesvio.query.order_by(RdimPecaDesvio.id).all())
    _add_sheet("RDIM Componentes Desvio", RdimComponenteDesvio, RdimComponenteDesvio.query.order_by(RdimComponenteDesvio.id).all())

    # ---- P&D (novo, 03/09/2026) ----
    _add_sheet("Projetos PD", ProjetoPD, ProjetoPD.query.order_by(ProjetoPD.id).all())
    _add_sheet("Testes PD", TesteProjetoPD, TesteProjetoPD.query.order_by(TesteProjetoPD.id).all())
    _add_sheet("Visitas Reunioes PD", VisitaReuniaoPD, VisitaReuniaoPD.query.order_by(VisitaReuniaoPD.id).all())

    return wb


# =====================================================================
# GESTÃO DE CUSTOS (pedido do Bruno, 20/09/2026) — módulo novo, Fase 1
# (grupo PIG MANDRIL: LBD, LUN, PU CAST, CORPO MANDRIL, ELC_MG_PC, PIGS EM
# BORRACHA). Plano completo em /root/.claude/plans/joyful-knitting-hoare.md.
#
# Princípio central do módulo inteiro: nenhum custo fica armazenado — tudo
# é calculado ao vivo a partir de MateriaPrima.custo_atual e
# ParametroHoraHomem.valor CORRENTES (_custo_estrutura_produto). É isso que
# garante "mudou o preço, recalcula tudo automaticamente" sem job nenhum.
# =====================================================================

def _hora_homem_atual():
    p = db.session.get(ParametroHoraHomem, 1)
    return p.valor if p else 40.0


def _custo_estrutura_produto(estrutura, _visitados=None, overrides_mp=None, override_hh=None):
    """Calcula o custo de 1 unidade de uma EstruturaProduto (produto numa
    DN), resolvendo itens SUBPRODUTO recursivamente (modela as dependências
    entre abas descobertas na planilha: LUN usa PU CAST, CORPO MANDRIL usa
    LBD). `_visitados` evita loop infinito se alguém cadastrar uma
    referência circular por engano — nunca deveria acontecer num uso normal,
    mas é uma trava de segurança barata.

    `overrides_mp` (dict {materia_prima_id: novo_custo}) e `override_hh`
    (valor R$/h) permitem SIMULAR um cenário hipotético sem alterar nada no
    banco — usados pela tela de Simulação (fase 3). Quando None (uso normal
    de todas as telas de produto/necessidades), o cálculo usa
    `MateriaPrima.custo_atual`/`_hora_homem_atual()` de verdade, exatamente
    como antes. Os overrides se propagam pra baixo em SUBPRODUTO, então
    simular o custo de uma matéria-prima usada bem no fundo de uma cadeia
    (ex. CORPO MANDRIL → LBD → PU CAST) afeta o total corretamente.

    Retorna dict: custo_mp, custo_hh, custo_total, linhas (detalhe de cada
    item, pra tela de composição), incompleto (True se algum SUBPRODUTO não
    tinha estrutura cadastrada na mesma DN — custo fica parcial, mostrado
    com aviso na tela em vez de mentir um total errado)."""
    _visitados = _visitados or set()
    if estrutura.id in _visitados:
        return {"custo_mp": 0.0, "custo_hh": 0.0, "custo_total": 0.0, "linhas": [], "incompleto": True}
    _visitados = _visitados | {estrutura.id}

    custo_mp = 0.0
    linhas = []
    incompleto = False

    for item in estrutura.itens:
        if item.tipo == "SUBPRODUTO":
            sub_estrutura = None
            if item.subproduto_id:
                sub_estrutura = EstruturaProduto.query.filter_by(
                    produto_id=item.subproduto_id, dn=estrutura.dn, ativo=True
                ).first()
            if sub_estrutura is None:
                incompleto = True
                linhas.append({
                    "item": item, "descricao": item.subproduto.codigo if item.subproduto else "?",
                    "quantidade": item.quantidade, "custo_unitario": None, "custo_linha": None,
                    "faltando": True,
                })
                continue
            sub_calc = _custo_estrutura_produto(sub_estrutura, _visitados, overrides_mp=overrides_mp, override_hh=override_hh)
            incompleto = incompleto or sub_calc["incompleto"]
            custo_unit = sub_calc["custo_total"]
            custo_linha = custo_unit * item.quantidade
            custo_mp += custo_linha
            linhas.append({
                "item": item, "descricao": item.subproduto.codigo if item.subproduto else "?",
                "quantidade": item.quantidade, "custo_unitario": custo_unit, "custo_linha": custo_linha,
                "faltando": False,
            })
        else:
            mp = item.materia_prima
            custo_unit = mp.custo_atual if mp else None
            if mp is not None and overrides_mp and mp.id in overrides_mp:
                custo_unit = overrides_mp[mp.id]
            if mp is None:
                incompleto = True
                linhas.append({
                    "item": item, "descricao": "?", "quantidade": item.quantidade,
                    "custo_unitario": None, "custo_linha": None, "faltando": True,
                })
                continue
            custo_linha = custo_unit * item.quantidade
            custo_mp += custo_linha
            linhas.append({
                "item": item, "descricao": f"{mp.codigo} — {mp.descricao}", "quantidade": item.quantidade,
                "unidade": mp.unidade, "custo_unitario": custo_unit, "custo_linha": custo_linha,
                "faltando": False,
            })

    hh_rate = override_hh if override_hh is not None else _hora_homem_atual()
    custo_hh = (estrutura.ciclo_horas or 0) * hh_rate
    return {
        "custo_mp": round(custo_mp, 4),
        "custo_hh": round(custo_hh, 4),
        "custo_total": round(custo_mp + custo_hh, 4),
        "linhas": linhas,
        "incompleto": incompleto,
    }


def _materias_primas_usadas(estrutura, _visitados=None):
    """Coleta (recursivamente, resolvendo SUBPRODUTO na mesma DN) o conjunto de
    MateriaPrima realmente usadas numa EstruturaProduto — pra tela de Simulação
    (fase 3) oferecer um campo de "novo custo" pra cada uma, mesmo as que estão
    escondidas dentro de uma cadeia de sub-produto (ex.: CORPO MANDRIL usa LBD
    usa PU CAST). Retorna lista ordenada por código, sem repetição."""
    _visitados = _visitados or set()
    if estrutura.id in _visitados:
        return []
    _visitados = _visitados | {estrutura.id}

    vistos = {}
    for item in estrutura.itens:
        if item.tipo == "SUBPRODUTO":
            if not item.subproduto_id:
                continue
            sub_estrutura = EstruturaProduto.query.filter_by(
                produto_id=item.subproduto_id, dn=estrutura.dn, ativo=True
            ).first()
            if sub_estrutura is None:
                continue
            for mp in _materias_primas_usadas(sub_estrutura, _visitados):
                vistos[mp.id] = mp
        elif item.materia_prima is not None:
            vistos[item.materia_prima.id] = item.materia_prima
    return sorted(vistos.values(), key=lambda mp: mp.codigo)


def _materiais_consumo_estrutura(estrutura, _visitados=None):
    """Quanto de CADA matéria-prima 1 unidade dessa EstruturaProduto consome
    — mesma recursão de `_custo_estrutura_produto` (resolve SUBPRODUTO pela
    mesma dn), mas soma QUANTIDADE em vez de custo. Pedido do Bruno
    (22/09/2026): "quero enxergar o total de matéria prima" nas Estações —
    a peça que faltava pra virar uma coisa só com o motor de custo já
    existente, sem duplicar a lógica de explosão de estrutura.

    Retorna dict com `por_materia_prima` (lista, 1 linha por MateriaPrima
    usada, já consolidada mesmo se aparecer em mais de um ponto da árvore —
    ex. ELASTOMERO aparece tanto direto quanto dentro do subproduto),
    `kg_total` (soma só das linhas com unidade "kg" — as outras unidades
    ficam listadas mas não entram nesse total, pedido do Bruno: "só o total
    em kg") e `incompleto` (True se algum SUBPRODUTO não tinha estrutura
    cadastrada na mesma dn — mesmo aviso já usado no custo)."""
    _visitados = _visitados or set()
    if estrutura.id in _visitados:
        return {"por_materia_prima": [], "kg_total": 0.0, "tem_kg": False, "incompleto": True}
    _visitados = _visitados | {estrutura.id}

    por_materia_prima = {}  # materia_prima_id -> {"materia_prima": mp, "quantidade": float}
    incompleto = False

    for item in estrutura.itens:
        if item.tipo == "SUBPRODUTO":
            sub_estrutura = None
            if item.subproduto_id:
                sub_estrutura = EstruturaProduto.query.filter_by(
                    produto_id=item.subproduto_id, dn=estrutura.dn, ativo=True
                ).first()
            if sub_estrutura is None:
                incompleto = True
                continue
            sub_calc = _materiais_consumo_estrutura(sub_estrutura, _visitados)
            incompleto = incompleto or sub_calc["incompleto"]
            for linha in sub_calc["por_materia_prima"]:
                mp_id = linha["materia_prima"].id
                bucket = por_materia_prima.setdefault(mp_id, {"materia_prima": linha["materia_prima"], "quantidade": 0.0})
                bucket["quantidade"] += linha["quantidade"] * item.quantidade
        elif item.materia_prima is not None:
            mp = item.materia_prima
            bucket = por_materia_prima.setdefault(mp.id, {"materia_prima": mp, "quantidade": 0.0})
            bucket["quantidade"] += item.quantidade
        else:
            incompleto = True

    linhas = sorted(por_materia_prima.values(), key=lambda l: l["materia_prima"].codigo)
    tem_kg = any(l["materia_prima"].unidade == "kg" for l in linhas)
    kg_total = sum(l["quantidade"] for l in linhas if l["materia_prima"].unidade == "kg")
    return {"por_materia_prima": linhas, "kg_total": round(kg_total, 4), "tem_kg": tem_kg, "incompleto": incompleto}


_FORNECEDORES_MATERIA_PRIMA_PRINCIPAIS = {"AMINO", "COIM", "LANXESS", "TECPUR"}


def _e_materia_prima_principal(mp):
    """Classifica se uma MateriaPrima está entre os grupos "principais" que
    o Bruno pediu pra NUNCA ficar escondidos (23/09/2026, revisão): "isso
    vale pra todas as matérias primas AMINO, COIM, LANXESS, BLOCO DE
    ESPUMA, TECPUR" — e, na mensagem seguinte, pediu pra essas aparecerem
    "ao lado" (sempre visíveis no cabeçalho da coluna), sem precisar
    clicar — diferente do detalhamento completo (todas as matérias-primas,
    inclusive ferragem/acessório), que continua só no popup por clique.

    Usa o campo `fornecedor` quando cadastrado (bate direto com AMINO/COIM/
    LANXESS/TECPUR). "BLOCO DE ESPUMA" não é um fornecedor no cadastro (as
    4 densidades estão como CBP ou DINATEC, inconsistente), então casa pela
    descrição. Um caso como "Bumper PU (MP COIM) — HLCC PC" tem o nome do
    fornecedor dentro da própria descrição mas o campo `fornecedor` ficou
    em branco no cadastro (achado ao implementar isso) — usa a descrição
    como fallback pra não perder esse caso."""
    fornecedor = (mp.fornecedor or "").strip().upper()
    if fornecedor in _FORNECEDORES_MATERIA_PRIMA_PRINCIPAIS:
        return True
    descricao = (mp.descricao or "").upper()
    if "BLOCO ESPUMA" in descricao:
        return True
    return any(grupo in descricao for grupo in _FORNECEDORES_MATERIA_PRIMA_PRINCIPAIS)


def _materiais_item_pedido(item_pedido):
    """Versão "por item do PCP" de `_materiais_consumo_estrutura` — casa o
    item via `_matching_produto_pcp` e multiplica o consumo de 1 unidade
    pela quantidade pedida (`item_pedido.quantidade`). É a função que a tela
    de Estações (Kanban) chama pra cada card — pedido do Bruno (22/09/2026):
    ver o total de matéria-prima ao abrir um item.

    Retorna sempre um dict com `matched` (False quando `_matching_produto_pcp`
    não achou correspondência — nunca inventa um número nesse caso, mesma
    régua de sempre: "não identificado" explícito em vez de estimar errado)
    e `tem_kg` (False quando o produto casou normalmente mas a estrutura
    dele só tem matéria-prima em outra unidade — ex. os subprodutos "DG"/
    "DE"/"DS"/"COPO..." da família PU CAST, importados com o custo já
    consolidado numa única linha "un" em vez de decompostos em química;
    nesse caso 0,00 kg seria um zero ENGANOSO — a tela mostra "sem detalhe
    em kg" em vez de um zero que parece resposta confirmada).

    Quando `estrutura.acessorios_extra` vem preenchido (LBD/LUN com ELC/
    ELP/MG/PC no texto — ver `_acessorios_extras_lbd_lun`, pedido do Bruno
    23/09/2026), o consumo de cada acessório entra somado nas mesmas
    `linhas`/`kg_total` — igual a qualquer outro material da estrutura, sem
    precisar de nenhuma tela nova pra enxergar. `acessorios` na resposta
    lista os códigos detectados, só pra transparência (nunca soma um
    acessório em silêncio sem mostrar qual foi)."""
    estrutura = _matching_produto_pcp(item_pedido)
    if estrutura is None:
        return {"matched": False, "kg_total": 0.0, "tem_kg": False, "linhas": [], "incompleto": False, "acessorios": []}

    consumo = _materiais_consumo_estrutura(estrutura)
    por_materia_prima = {
        l["materia_prima"].id: {"materia_prima": l["materia_prima"], "quantidade": l["quantidade"]}
        for l in consumo["por_materia_prima"]
    }
    incompleto = consumo["incompleto"]

    acessorios_extra = getattr(estrutura, "acessorios_extra", [])
    for acessorio in acessorios_extra:
        consumo_acessorio = _materiais_consumo_estrutura(acessorio)
        incompleto = incompleto or consumo_acessorio["incompleto"]
        for l in consumo_acessorio["por_materia_prima"]:
            mp_id = l["materia_prima"].id
            bucket = por_materia_prima.setdefault(mp_id, {"materia_prima": l["materia_prima"], "quantidade": 0.0})
            bucket["quantidade"] += l["quantidade"]

    quantidade = item_pedido.quantidade or 0
    linhas_unitarias = sorted(por_materia_prima.values(), key=lambda l: l["materia_prima"].codigo)
    linhas = [
        {"materia_prima": l["materia_prima"], "quantidade": round(l["quantidade"] * quantidade, 4)}
        for l in linhas_unitarias
    ]
    kg_total_unitario = sum(l["quantidade"] for l in linhas_unitarias if l["materia_prima"].unidade == "kg")
    tem_kg = any(l["materia_prima"].unidade == "kg" for l in linhas_unitarias)
    return {
        "matched": True,
        "estrutura": estrutura,
        "kg_total": round(kg_total_unitario * quantidade, 3),
        "tem_kg": tem_kg,
        "linhas": linhas,
        "incompleto": incompleto,
        "acessorios": [a.produto.codigo for a in acessorios_extra],
    }


def _produtos_catalogo(familia=None, apenas_ativos=True):
    """Lista Produto + EstruturaProduto com custo calculado — equivalente
    funcional da aba BUSCA DE CUSTO (índice consolidado), item 1/2 do
    pedido. Cada linha = 1 (produto, DN)."""
    q = Produto.query
    if apenas_ativos:
        q = q.filter_by(ativo=True)
    if familia:
        q = q.filter_by(familia=familia)
    produtos = q.order_by(Produto.familia, Produto.codigo).all()

    linhas = []
    for p in produtos:
        estruturas = [e for e in p.estruturas if e.ativo] if apenas_ativos else list(p.estruturas)
        for e in sorted(estruturas, key=lambda e: _chave_ordenacao_dn(e.dn)):
            calc = _custo_estrutura_produto(e)
            linhas.append({"produto": p, "estrutura": e, "calc": calc})
    return linhas


def _chave_ordenacao_dn(dn):
    """Tenta ordenar DN numericamente (2, 3, 4, 6, 8...) em vez de
    alfabeticamente (10 antes de 2) — best-effort, cai pro texto se não
    conseguir extrair número."""
    m = re.search(r"\d+(?:[.,]\d+)?", dn or "")
    return (0, float(m.group(0).replace(",", "."))) if m else (1, dn or "")


_FAMILIAS_PRODUTO_PCP = ("LBD", "LUN", "PU CAST", "CORPO MANDRIL", "ELC_MG_PC", "PIGS EM BORRACHA")

# Famílias da planilha ainda NÃO implementadas — usado só pra mostrar "chega
# numa próxima fase" na tela de Custos dos Produtos, em vez de simplesmente
# omitir sem explicação. Vazio desde a fase 3 (20/09/2026): todas as famílias
# de produto da planilha (PIG MANDRIL, espuma, SUPERFLEX, SILICONE) já foram
# importadas — só a tela de Simulação/Histórico ainda está pendente (ver
# custos_simulacao).
_FAMILIAS_FASE_SEGUINTE = ()


_RE_FRACAO_POL = re.compile(r"\d\s*/\s*\d+\s*(?:''|\"|['’”]|POL(?:EGADAS)?\b)", re.I)


def _dn_extraido_para_matching_custos(descricao):
    """Reaproveita `_classificar_dn_mm` (validada pro item 9 da tela de
    KPIs), mas descarta o resultado quando a medida em polegadas é uma
    fração/número misto (ex. "5 1/8\"", "6 1/2\"" — tamanhos reais de
    tubo/OD que aparecem na base). `_classificar_dn_mm` foi feita pra
    agrupamento em relatório (best-effort, ~95% de cobertura, erro ali só
    põe o item no grupo errado de um relatório) e nesses casos ela captura
    só o denominador da fração (ex. "5 1/8\"" -> "8", "6 1/2\"" -> "2"),
    que aqui viraria uma matéria-prima/custo ERRADO atribuído em silêncio —
    achado ao validar manualmente casamentos reais do PCP nesta sessão.
    Aqui a régua é outra: preferir não identificar a identificar errado."""
    if not descricao or _RE_FRACAO_POL.search(descricao):
        return None
    return _classificar_dn_mm(descricao)


_RE_DENSIDADE_PEDIDO = re.compile(r"\b(ALTA|M[ÉE]DIA|BAIXA)\b", re.I)


def _numero_dn_lider(texto):
    """Extrai só os dígitos LÍDERES de um texto de DN (ex. "10" de
    "10'' MÉDIA", "6" de "6\""). Não usar um `re.sub(r"[^\d.,]", "", ...)`
    ingênuo aqui — ele pegaria TODO dígito solto no texto, inclusive os que
    vêm depois de letras (ex. "10'' BAIXA D26" viraria "1026" em vez de
    "10") — bug real encontrado ao habilitar o casamento da família H
    (dn cadastrado com o sufixo de variante de bloco embutido, ex. "D26")."""
    m = re.match(r"\s*([\d.,]+)", texto or "")
    return m.group(1) if m else ""


def _normalizar_texto_matching_custos(texto):
    """Normalização de texto compartilhada entre a descrição do pedido E o
    `chave_busca` do catálogo, pros dois lados ficarem no mesmo formato
    antes de comparar (pedido do Bruno, 22/09/2026, depois de investigar
    casos reais que não casavam por diferença pura de formatação):
    - espaços duplos (digitação livre, ex. "DISCO  GUIA  DN 18\"");
    - hífen vs espaço (ex. "H-FLEX" no pedido x "HFLEX" no catálogo,
      "HLCC-PC" x "HLCC PC") — trata os dois como equivalentes;
    - "C/SELO" / "C/ SELO" (abreviação comum na digitação do pedido) vira
      "COM SELO", pra casar com o chave_busca escrito por extenso;
    - acento (ex. "CÔNICO" x "CONICO", "ESPAÇADOR" x "ESPACADOR") — comum
      faltar acento na digitação livre do pedido — vira sempre a forma sem
      acento dos dois lados, mesma técnica já usada no filtro de busca
      client-side do Kanban (`estacoes_kanban.html`, NFD + remove marca de
      combinação);
    - palavra de preenchimento "TIPO" (ex. "DISCO TIPO GUIA", "COPO TIPO
      CONICO") — pedido do Bruno (23/09/2026: "DISCO TIPO GUIA OU TIIPO
      SELO ELE NAO ESTA RELACIONANDO... ACREDITO QUE SEJA POR CAUS DA
      PALAVRA 'TIPO'.... CORRIGA ISSO, POIS O TIPO É SO UM TERMO"). O
      `chave_busca` do catálogo é escrito sem essa palavra (ex. "DISCO
      GUIA"), então "TIPO" no meio quebrava o casamento por substring
      contíguo. Confirmado empiricamente contra os 286 itens reais do PCP
      com "TIPO" na descrição: 236 passam a casar (0 regressões) removendo
      só essa palavra — os ~50 restantes continuam sem casar por motivo
      NÃO relacionado a "TIPO" (ex. "COPO TIPO PISTÃO DN 24\"" — lacuna
      real de catálogo, não tem DN 24 cadastrado pra COPO PISTAO)."""
    txt = (texto or "").upper()
    txt = txt.replace("-", " ")
    txt = re.sub(r"C\s*/\s*SELO", "COM SELO", txt)
    txt = "".join(c for c in unicodedata.normalize("NFD", txt) if unicodedata.category(c) != "Mn")
    txt = re.sub(r"\bTIPO\b", " ", txt)
    return re.sub(r"\s+", " ", txt).strip()


# Acessórios da família ELC_MG_PC que podem vir "embutidos" no código de um
# item LBD/LUN do PCP (ex. "LBD-DG2-DS4-ELC-MG-PC DN 20''") — pedido do
# Bruno (23/09/2026): "PARA LBD E LUN, PRECISO INCLUIR O ACESSSORIO... TIPO
# ELC, PC OU MG... PRECISO QUE OS CUSTOS DESSES ACESSORIOS QUE COMPOEM O LBD
# ACOMPNHEM NESSA AREA DE CUSTOS". As 3 siglas que ele citou (ELC/MG/PC) são
# literalmente as que dão nome à família "ELC_MG_PC" já cadastrada (seed de
# 20/09/2026) — ELP entra junto por ser a variante em PP do mesmo ELC (AÇO),
# usando a mesma sigla que já aparece nos pedidos reais (ex. "LBD-DG2-DS4-
# ELP-PC"). Cada sigla mapeia pro `codigo` exato do Produto acessório já
# cadastrado — ver `_acessorios_extras_lbd_lun`.
_ACESSORIOS_LBD_LUN = (
    ("ELC", "ELC (AÇO)"),
    ("ELP", "ELP (PP)"),
    ("MG", "CINTA MAGNÉTICA"),
    ("PC", "PLACA CALIBRADORA"),
)


def _acessorios_extras_lbd_lun(estrutura, descricao_normalizada, numero_dn):
    """Detecta, na descrição já normalizada de um item do PCP, siglas de
    acessório (ELC/ELP/MG/PC — ver `_ACESSORIOS_LBD_LUN`) e retorna a lista
    de EstruturaProduto (mesma DN do item) desses acessórios — só quando o
    produto casado é da família LBD ou LUN (pedido explícito do Bruno,
    23/09/2026, restrito a essas duas famílias pra não arriscar falso
    positivo em outras famílias onde "PC"/"MG" podem não significar
    acessório nenhum).

    Cada sigla usa a mesma fronteira de palavra do casamento de
    `chave_busca` (evita, por exemplo, "MG" casar no meio de outra
    palavra). Quando a sigla aparece mas o acessório não tem estrutura
    cadastrada nessa DN específica, simplesmente não entra na lista — nunca
    quebra o casamento do produto base nem inventa custo."""
    if estrutura is None or numero_dn is None:
        return []
    produto = estrutura.produto
    if produto is None or produto.familia not in ("LBD", "LUN"):
        return []
    extras = []
    codigos_adicionados = set()
    for sigla, codigo_produto in _ACESSORIOS_LBD_LUN:
        if codigo_produto in codigos_adicionados:
            continue
        if not re.search(r"(?<![A-ZÀ-Ü0-9])" + re.escape(sigla) + r"(?![A-ZÀ-Ü0-9])", descricao_normalizada):
            continue
        produto_acessorio = Produto.query.filter_by(codigo=codigo_produto, ativo=True).first()
        if produto_acessorio is None:
            continue
        for e in produto_acessorio.estruturas:
            if e.ativo and _numero_dn_lider(e.dn) == numero_dn:
                extras.append(e)
                codigos_adicionados.add(codigo_produto)
                break
    return extras


def _matching_produto_pcp(item_pedido):
    """Casa um ItemPedido com um (Produto, EstruturaProduto) do catálogo de
    custos, por correspondência de texto — mesmo princípio já usado em
    LeadTimeProducao.produto (ILIKE contido na descrição), decisão
    confirmada com o Bruno em 20/09/2026. DN extraído via
    `_dn_extraido_para_matching_custos` (ver docstring — protege contra
    tamanhos fracionários mal interpretados pela extração genérica de DN).

    O casamento de `chave_busca` usa fronteira de palavra (não substring
    solto) pra um código curto (ex. "H") nunca casar no meio de outra
    palavra, e exige DENSIDADE (ALTA/MÉDIA/BAIXA) quando a EstruturaProduto
    tem esse classificador preenchido (`densidade`, ver
    `_migrar_densidade_estrutura_produto`) — pedido do Bruno (22/09/2026:
    "Modelo produto: HLR, DENSIDADE: ALTA, DN: 4''"), depois de confirmar
    que o casamento antigo aceitava QUALQUER densidade pro mesmo DN e podia
    reportar a matéria-prima da estrutura ERRADA em silêncio. Quando a
    densidade não dá pra identificar no texto do pedido (ou não bate com
    nenhuma estrutura candidata), o item fica "não identificado" — a régua
    aqui é a mesma de sempre: preferir não identificar a identificar
    errado.

    Retorna a EstruturaProduto casada, ou None se não achou correspondência
    (o chamador trata como "não identificado automaticamente" — nunca some
    o item da conta, sempre aparece explicitamente como não-casado). Quando
    o produto casado é LBD/LUN, a EstruturaProduto retornada ganha um
    atributo extra `acessorios_extra` (lista, pode ser vazia) com os
    acessórios ELC/ELP/MG/PC detectados no texto — ver
    `_acessorios_extras_lbd_lun`. É um atributo só em memória (nunca
    persistido), lido pelos chamadores que somam consumo/custo (
    `_materiais_item_pedido`, `_necessidades_pcp_materia_prima`,
    `_visao_rapida_custos`)."""
    descricao = _normalizar_texto_matching_custos(item_pedido.descricao_produto)
    if not descricao:
        return None

    dn_extraido = _dn_extraido_para_matching_custos(item_pedido.descricao_produto)
    numero_extraido = _numero_dn_lider(dn_extraido) if dn_extraido is not None else None

    # Correspondência manual (ferramenta pedida pelo Bruno, 23/09/2026, pra
    # cobrir os casos em que o casamento automático genuinamente não dá
    # conta — ex. "HS DN 5''" sem nenhuma EstruturaProduto cadastrada na DN
    # 5, mas o Bruno quer usar a DN 6'' como base) sempre tem prioridade
    # sobre a heurística automática abaixo — casamento exato pela descrição
    # já normalizada (mesma normalização usada nos dois lados).
    correspondencia = CorrespondenciaManualCusto.query.filter_by(descricao_normalizada=descricao).first()
    if correspondencia is not None:
        estrutura_manual = EstruturaProduto.query.filter_by(
            produto_id=correspondencia.produto_id, dn=correspondencia.dn, ativo=True
        ).first()
        if estrutura_manual is not None:
            estrutura_manual.acessorios_extra = _acessorios_extras_lbd_lun(estrutura_manual, descricao, numero_extraido)
            return estrutura_manual
        # a correspondência aponta pra um produto/DN que foi desativado
        # desde então — cai pro casamento automático abaixo em vez de
        # travar o item em silêncio.

    if dn_extraido is None:
        return None

    m_densidade = _RE_DENSIDADE_PEDIDO.search(descricao)
    densidade_pedido = m_densidade.group(1).upper().replace("MEDIA", "MÉDIA") if m_densidade else None

    candidatos = Produto.query.filter(Produto.ativo == True, Produto.chave_busca.isnot(None)).all()  # noqa: E712
    melhor = None
    melhor_especificidade = -1
    for produto in candidatos:
        chave = _normalizar_texto_matching_custos(produto.chave_busca)
        if not chave:
            continue
        # fronteira de palavra: o caractere logo antes/depois da chave (se
        # existir) não pode ser letra/número/acento — impede que uma chave
        # curta (ex. "H") case no meio de outra palavra.
        if not re.search(r"(?<![A-ZÀ-Ü0-9])" + re.escape(chave) + r"(?![A-ZÀ-Ü0-9])", descricao):
            continue
        for estrutura in produto.estruturas:
            if not estrutura.ativo:
                continue
            # compara o número puro do DN (ignora aspas/"mm"/espaços) —
            # dn_extraido vem no formato 6" ou 150MM; estrutura.dn é só o
            # número (ex. "6") como na planilha original.
            numero_estrutura = _numero_dn_lider(estrutura.dn)
            if not (numero_extraido and numero_estrutura and numero_extraido == numero_estrutura):
                continue
            if estrutura.densidade:
                # primeiro "token" de densidade.densidade (ex. "BAIXA" de
                # "BAIXA D26") — a variante de bloco (D26/D45/D60/D80) nunca
                # aparece no texto do pedido, então não dá pra (nem precisa,
                # ver docstring de _migrar_densidade_estrutura_produto: as
                # matérias-primas em kg são idênticas entre variantes de
                # bloco) desambiguar além do 1º token.
                densidade_estrutura = estrutura.densidade.split(" ")[0]
                if densidade_pedido != densidade_estrutura:
                    continue
            # produto mais específico (chave de busca mais longa) ganha em
            # caso de mais de um bater (ex. "LBD" e "LBD-DG2-DS4")
            if len(chave) > melhor_especificidade:
                melhor = estrutura
                melhor_especificidade = len(chave)
    if melhor is not None:
        melhor.acessorios_extra = _acessorios_extras_lbd_lun(melhor, descricao, numero_extraido)
    return melhor


def _necessidades_pcp_materia_prima():
    """Item 5/6 do pedido — o objetivo principal do módulo: a partir da fila
    do PCP (itens ainda não finalizados), identifica automaticamente
    produto -> estrutura -> matérias-primas necessárias, consolidando por
    matéria-prima entre todos os produtos que a usam.

    NUNCA grava nada (item 7: é sempre "necessidade prevista" calculada na
    hora, nunca consumo real, nunca mexe em estoque/produção) — recalcula do
    zero a cada chamada, então acompanha qualquer mudança do PCP
    automaticamente (item 8), sem duplicidade e sem precisar de
    sincronização manual."""
    itens_pendentes = (
        ItemPedido.query.join(Pedido)
        .filter(ItemPedido.status_producao != "FINALIZADO")
        .order_by(Pedido.data_inclusao_pedido.desc().nullslast())
        .all()
    )

    necessidades = {}  # materia_prima_id -> {materia_prima, necessidade, origens: [...]}
    nao_identificados = []

    def _explodir(estrutura, quantidade_produto, item_pedido, _visitados=None):
        _visitados = _visitados or set()
        if estrutura.id in _visitados:
            return
        _visitados = _visitados | {estrutura.id}
        for comp in estrutura.itens:
            if comp.tipo == "SUBPRODUTO":
                if not comp.subproduto_id:
                    continue
                sub_estrutura = EstruturaProduto.query.filter_by(
                    produto_id=comp.subproduto_id, dn=estrutura.dn, ativo=True
                ).first()
                if sub_estrutura:
                    _explodir(sub_estrutura, quantidade_produto * comp.quantidade, item_pedido, _visitados)
            else:
                if not comp.materia_prima_id:
                    continue
                mp = comp.materia_prima
                qtd_necessaria = quantidade_produto * comp.quantidade
                bucket = necessidades.setdefault(mp.id, {"materia_prima": mp, "necessidade_prevista": 0.0, "origens": []})
                bucket["necessidade_prevista"] += qtd_necessaria
                bucket["origens"].append({
                    "pedido": item_pedido.pedido, "item": item_pedido, "quantidade_mp": qtd_necessaria,
                })

    for item in itens_pendentes:
        estrutura = _matching_produto_pcp(item)
        if estrutura is None:
            nao_identificados.append(item)
            continue
        _explodir(estrutura, item.quantidade or 0, item)
        # Acessórios LBD/LUN detectados no texto (ELC/ELP/MG/PC — pedido do
        # Bruno 23/09/2026) entram na mesma necessidade prevista, cada um
        # explodido do zero (_visitados novo por chamada, sem interferir na
        # árvore do produto base).
        for acessorio in getattr(estrutura, "acessorios_extra", []):
            _explodir(acessorio, item.quantidade or 0, item)

    linhas = sorted(necessidades.values(), key=lambda b: b["materia_prima"].descricao)
    for linha in linhas:
        linha["necessidade_prevista"] = round(linha["necessidade_prevista"], 3)

    return {"linhas": linhas, "nao_identificados": nao_identificados}


def _visao_rapida_custos():
    """Item 9 do pedido — indicadores do painel de Gestão de Custos. Os 2
    indicadores que o texto original liga a estoque ("MP em risco",
    "produtos impactados por falta de MP") não têm como significar risco de
    ruptura sem estoque real cadastrado no app (nenhum existe hoje — decisão
    já combinada com o Bruno) — nesta fase eles viram indicador de
    QUALIDADE DE DADO (matéria-prima sem custo cadastrado / item do PCP não
    identificado automaticamente), rotulados de forma explícita na tela pra
    não prometer um risco de estoque que o dado não sustenta."""
    necessidades = _necessidades_pcp_materia_prima()

    produtos_cadastrados = Produto.query.filter_by(ativo=True).count()
    mps_cadastradas = MateriaPrima.query.filter_by(ativo=True).count()

    custo_total_previsto = 0.0
    itens_pendentes = ItemPedido.query.filter(ItemPedido.status_producao != "FINALIZADO").all()
    for item in itens_pendentes:
        estrutura = _matching_produto_pcp(item)
        if estrutura is None:
            continue
        custo_unitario = _custo_estrutura_produto(estrutura)["custo_total"]
        # Acessórios LBD/LUN (ELC/ELP/MG/PC) somam custo no mesmo item —
        # pedido do Bruno (23/09/2026): "precisO que os custos desses
        # acessórios que compõem o LBD acompanhem nessa área de custos".
        for acessorio in getattr(estrutura, "acessorios_extra", []):
            custo_unitario += _custo_estrutura_produto(acessorio)["custo_total"]
        custo_total_previsto += custo_unitario * (item.quantidade or 0)

    mps_sem_custo = [
        linha["materia_prima"] for linha in necessidades["linhas"]
        if not linha["materia_prima"].custo_atual
    ]

    return {
        "produtos_cadastrados": produtos_cadastrados,
        "mps_cadastradas": mps_cadastradas,
        "custo_total_previsto": round(custo_total_previsto, 2),
        "materias_primas_com_necessidade": len(necessidades["linhas"]),
        "mps_sem_custo": mps_sem_custo,
        "itens_nao_identificados": necessidades["nao_identificados"],
        "hora_homem_atual": _hora_homem_atual(),
    }


def _validar_materia_prima_form(f, ignorar_id=None):
    codigo = (f.get("codigo") or "").strip().upper()
    descricao = (f.get("descricao") or "").strip()
    unidade = (f.get("unidade") or "").strip() or "kg"
    categoria = (f.get("categoria") or "").strip() or None
    fornecedor = (f.get("fornecedor") or "").strip() or None
    custo_atual = _parse_float_form(f.get("custo_atual"), default=None)

    if not codigo:
        return None, "Informe o código da matéria-prima."
    if not descricao:
        return None, "Informe a descrição da matéria-prima."
    if custo_atual is None or custo_atual < 0:
        return None, "Informe um custo atual válido (maior ou igual a zero)."

    conflito = MateriaPrima.query.filter_by(codigo=codigo)
    if ignorar_id is not None:
        conflito = conflito.filter(MateriaPrima.id != ignorar_id)
    if conflito.first() is not None:
        return None, f'Já existe uma matéria-prima cadastrada com o código "{codigo}".'

    return {
        "codigo": codigo, "descricao": descricao, "unidade": unidade, "categoria": categoria,
        "fornecedor": fornecedor, "custo_atual": custo_atual,
    }, None


def _registrar_revisao_materia_prima(mp, valor_anterior, valor_novo, motivo):
    if valor_anterior == valor_novo:
        return
    db.session.add(
        MateriaPrimaHistorico(
            materia_prima_id=mp.id, custo_anterior=valor_anterior, custo_novo=valor_novo,
            motivo=(motivo or "").strip() or None,
            usuario_nome=current_user.nome if current_user.is_authenticated else None,
        )
    )


def register_routes(app):
    @app.before_request
    def _restringir_acesso_por_papel():
        """Papel PD (Líder de P&D): só pode acessar P&D + visualização de
        Gestão Produção > Estações — ver ENDPOINTS_PERMITIDOS_PD em
        permissoes.py. Todo o resto do sistema continua sem restrição de
        visualização por papel (só edição é restrita, via requer_role)."""
        if not current_user.is_authenticated:
            return
        endpoint = request.endpoint
        if endpoint is None:
            return
        if not pode_acessar_endpoint(current_user, endpoint):
            abort(403)

    @app.route("/login", methods=["GET", "POST"])
    def login():
        if current_user.is_authenticated:
            return redirect(_pagina_inicial(current_user))

        if request.method == "POST":
            username = request.form.get("username", "").strip()
            senha = request.form.get("senha", "")
            usuario = Usuario.query.filter_by(username=username).first()
            if usuario and usuario.checar_senha(senha):
                login_user(usuario)
                destino = request.args.get("next") or _pagina_inicial(usuario)
                return redirect(destino)
            flash("Usuário ou senha inválidos.", "danger")

        return render_template("login.html")

    @app.route("/logout")
    @login_required
    def logout():
        logout_user()
        return redirect(url_for("login"))

    @app.route("/consulta-pedido/busca-sugestoes")
    @login_required
    def consulta_pedido_busca_sugestoes():
        """Pedido do Bruno (01/09/2026, movido pra aba própria em 02/09/2026):
        sugestões dinâmicas (digitando) pro canal de busca de status da aba
        Consulta Pedido — devolve um fragmento HTML pronto (mesmo padrão do
        resto do sistema, que é 100% server-rendered) pra ser injetado direto
        na página via fetch()."""
        termo = request.args.get("q", "")
        sugestoes = _buscar_pedidos_para_status(termo)
        return render_template("_consulta_pedido_sugestoes.html", sugestoes=sugestoes, termo=termo.strip())

    @app.route("/consulta-pedido/busca-detalhe")
    @login_required
    def consulta_pedido_busca_detalhe():
        """Status completo de UM pedido (Produção + Operação, cruzados por
        pedido_venda) — devolve o fragmento HTML do painel de detalhe,
        aberto ao clicar numa sugestão/atalho da aba Consulta Pedido."""
        chave = _normalizar_pedido_venda(request.args.get("pedido_venda", ""))
        if not chave:
            return "", 204

        pedido = (
            Pedido.query.options(selectinload(Pedido.itens))
            .filter(_pedido_venda_normalizado_sql(Pedido.pedido_venda) == chave)
            .first()
        )
        go = PedidoOperacao.query.filter(_pedido_venda_normalizado_sql(PedidoOperacao.pedido_venda) == chave).first()

        if pedido is None and go is None:
            return render_template("_consulta_pedido_detalhe.html", nao_encontrado=True, pedido_venda=chave)

        liberacao_pcp = _liberacao_pcp_por_pedido_venda([chave]).get(chave, {})
        data_cliente_producao = _data_cliente_por_pedido_venda([chave]).get(chave)
        situacao_entrega = _situacao_entrega_go(go, pedido)
        otd_pedido = _otd_do_pedido(go)
        etapas = _etapas_acompanhamento_pedido(pedido, go)
        # Pedido do Bruno (09/09/2026): lead time em dias corridos, prazo
        # comercial (inclusão -> solicitado cliente) e prazo de operação
        # completa (inclusão -> liberação efetiva PCP) — ver _prazos_pedido.
        prazos = _prazos_pedido(pedido, go, liberacao_pcp, data_cliente_producao)

        # Qualidade (RDIM) — pedido do Bruno (02/09/2026): ver o processo de
        # qualidade item a item dentro do pedido, na mesma tela que já une
        # Produção + Operação.
        inspecoes_rdim = _inspecoes_rdim_por_item([i.id for i in pedido.itens]) if pedido else {}
        resumo_rdim = _resumo_rdim_pedido(inspecoes_rdim.values())

        return render_template(
            "_consulta_pedido_detalhe.html",
            pedido=pedido, go=go, pedido_venda=chave,
            liberacao_pcp=liberacao_pcp, data_cliente_producao=data_cliente_producao,
            situacao_entrega=situacao_entrega, otd_pedido=otd_pedido,
            etapas=etapas, nao_encontrado=False,
            inspecoes_rdim=inspecoes_rdim, resumo_rdim=resumo_rdim,
            prazos=prazos,
        )

    @app.route("/consulta-pedido")
    @login_required
    def consulta_pedido():
        """Aba própria "🔍 Consulta Pedido" (pedido do Bruno, 02/09/2026): o
        canal de busca de status que antes vivia dentro do Painel virou uma
        tela dedicada — só a busca por nº do pedido/cliente. Sem atalhos de
        Atrasados/Vencendo aqui de propósito (pedido do Bruno, 02/09/2026):
        esta área é de acesso comercial, sem esse tipo de informação interna."""
        pedido_venda_inicial = (request.args.get("pedido_venda", "") or "").strip()
        return render_template("consulta_pedido.html", pedido_venda_inicial=pedido_venda_inicial)

    @app.route("/painel")
    @login_required
    def painel():
        resumo = _calcular_resumo()
        atrasados = Pedido.query.filter(_predicado_atrasado()).count()
        vencendo = Pedido.query.filter(_predicado_vencendo()).count()
        backlog = resumo["total"] - resumo["finalizado"]

        hoje = date.today()
        # Faturamento previsto (mês atual) + Backlog mês seguinte — pedido do
        # Bruno (16/09/2026): sempre a partir do planejamento PCP da
        # Listagem Geral (ItemPedido.planejamento_semanal), não mais da
        # liberação prevista (_faturamento_mes, que continua existindo pros
        # outros usos que já tinha: gráfico anual, tendência etc.). Os dois
        # cards usam o mesmo _resumo_mes_pcp — inclui contagem de pedidos e
        # Top 5 por valor (pedido do Bruno, 16/09/2026).
        resumo_mes_atual_pcp = _resumo_mes_pcp(hoje.year, hoje.month)
        previsto_mes = resumo_mes_atual_pcp["valor_total"]
        resumo_mes_seguinte_pcp = _resumo_mes_seguinte_pcp(hoje)

        # Mini gestão de risco de prazos, FOB x CIF — pedido do Bruno
        # (16/09/2026), logo abaixo de "Últimos apontamentos de Qualidade".
        mini_risco_prazos = _mini_risco_prazos_painel()

        # Detalhamento de Lead Time (chão de fábrica, total, fila de espera,
        # prazo comercial) — pedido do Bruno (16/09/2026), quadrante próprio.
        lead_time_detalhado = _lead_time_detalhado_painel()

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

        # OTD/Lead Time da Operação, recorte "mês atual" — pedido do Bruno
        # (03/09/2026): "os principais KPIs da operação já sejam visuais logo
        # no painel", trazendo pro Painel os mesmos indicadores da tela
        # Resultados/OTD (mesmas funções, mesmo critério/meta).
        query_operacao_mes = _pedidos_operacao_do_periodo("mes", hoje.year, hoje.month)
        otd_operacao_mes = _resumo_otd(query_operacao_mes)
        lead_times_operacao_mes = _resumo_lead_times(query_operacao_mes)
        mes_atual_label = f"{MESES_PT[hoje.month - 1]}/{hoje.year}"

        # Gráficos anuais de faturamento realizado e OTD (pedido do Bruno,
        # 03/09/2026, no lugar do antigo "previsto × realizado (6 meses)") —
        # 1 filtro de ano simples, compartilhado pelos dois gráficos.
        anos_disponiveis_graficos = list(range(hoje.year - 2, hoje.year + 1))
        ano_graficos = request.args.get("ano_graficos", hoje.year, type=int)
        if ano_graficos not in anos_disponiveis_graficos:
            ano_graficos = hoje.year

        # Apontamentos recentes de Qualidade (RDIM + RNC) — pedido do Bruno
        # (03/09/2026): "notificação breve" de todo apontamento novo, pra ele
        # ter ciência sem precisar entrar na área de Qualidade. "Novo desde a
        # última visita" é guardado na sessão do navegador (sem tabela nova),
        # e é sempre atualizado DEPOIS de calcular a lista, pra este mesmo
        # carregamento ainda mostrar o que entrou desde a visita anterior.
        ultima_visita_str = session.get("ultima_visita_painel")
        ultima_visita = datetime.fromisoformat(ultima_visita_str) if ultima_visita_str else None
        apontamentos_qualidade = _apontamentos_recentes_qualidade(desde=ultima_visita)
        novos_apontamentos = sum(1 for a in apontamentos_qualidade if a["novo"])

        # Últimos pedidos incluídos em Gestão Produção + últimas atualizações
        # de P&D — pedido do Bruno (03/09/2026), mesmo formato/mesma "última
        # visita" do feed de Qualidade acima.
        pedidos_recentes_producao = _pedidos_recentes_producao(desde=ultima_visita)
        novos_pedidos_producao = sum(1 for e in pedidos_recentes_producao if e["novo"])
        atualizacoes_pd = _atualizacoes_recentes_pd(desde=ultima_visita)
        novas_atualizacoes_pd = sum(1 for e in atualizacoes_pd if e["novo"])

        # Prévia horizontal do Planejamento Semanal PCP — pedido do Bruno
        # (09/09/2026), logo abaixo de "Últimos pedidos incluídos". Desde
        # 16/09/2026 mostra sempre TODAS as semanas do mês atual (não mais
        # uma janela rolante) — ver docstring de _preview_semanal_pcp_painel.
        preview_pcp_semanal = _preview_semanal_pcp_painel(hoje)

        session["ultima_visita_painel"] = datetime.utcnow().isoformat()

        return render_template(
            "painel.html",
            resumo=resumo,
            atrasados=atrasados,
            vencendo=vencendo,
            backlog=backlog,
            previsto_mes=previsto_mes,
            resumo_mes_atual_pcp=resumo_mes_atual_pcp,
            resumo_mes_seguinte_pcp=resumo_mes_seguinte_pcp,
            mini_risco_prazos=mini_risco_prazos,
            lead_time_detalhado=lead_time_detalhado,
            lead_time_medio=_lead_time_medio_dias(),
            otd=_otd_percentual(),
            backlog_estacao=_backlog_por_estacao(),
            pedidos_atrasados=pedidos_atrasados,
            projecao_pcp=_projecao_pcp(),
            projecao_pcp_mensal=_projecao_pcp_mensal(pcp_de, pcp_ate),
            pcp_de_str=f"{pcp_de[0]:04d}-{pcp_de[1]:02d}",
            pcp_ate_str=f"{pcp_ate[0]:04d}-{pcp_ate[1]:02d}",
            otd_operacao_mes=otd_operacao_mes,
            lead_times_operacao_mes=lead_times_operacao_mes,
            mes_atual_label=mes_atual_label,
            faturamento_mensal_ano=_faturamento_mensal_ano(ano_graficos),
            otd_mensal_ano=_otd_mensal_ano(ano_graficos),
            ano_graficos=ano_graficos,
            anos_disponiveis_graficos=anos_disponiveis_graficos,
            apontamentos_qualidade=apontamentos_qualidade,
            novos_apontamentos=novos_apontamentos,
            pedidos_recentes_producao=pedidos_recentes_producao,
            novos_pedidos_producao=novos_pedidos_producao,
            atualizacoes_pd=atualizacoes_pd,
            novas_atualizacoes_pd=novas_atualizacoes_pd,
            preview_pcp_semanal=preview_pcp_semanal,
        )

    @app.route("/kpis")
    @login_required
    def kpis():
        # Reconstrução completa (pedido do Bruno, 17-18/09/2026) — 2
        # controles de período independentes na mesma tela: `meses`
        # (3/6/12, pro Bloco A — lead time/fila, dinâmico por período, igual
        # a tela antiga já tinha) e `ano`/`mes` (mês navegável, mesmo padrão
        # de faturamento(), pro Bloco D — controle mensal itens 6/7/8/9).
        meses = request.args.get("meses", 6, type=int)
        if meses not in (3, 6, 12):
            meses = 6
        desde = date.today() - timedelta(days=30 * meses)

        hoje = date.today()
        ano = request.args.get("ano", hoje.year, type=int)
        mes = request.args.get("mes", hoje.month, type=int)
        if not (1 <= mes <= 12):
            mes = hoje.month
        mes_anterior_ano, mes_anterior_mes = (ano, mes - 1) if mes > 1 else (ano - 1, 12)
        mes_seguinte_ano, mes_seguinte_mes = (ano, mes + 1) if mes < 12 else (ano + 1, 1)

        tendencia_fila = _tendencia_fila_lead_time(meses=meses)

        return render_template(
            "kpis.html",
            meses=meses,
            ano=ano,
            mes=mes,
            mes_label=f"{MESES_PT[mes - 1]}/{ano}",
            mes_anterior=dict(ano=mes_anterior_ano, mes=mes_anterior_mes),
            mes_seguinte=dict(ano=mes_seguinte_ano, mes=mes_seguinte_mes),
            lead_time_fila_estacao=_lead_time_fila_por_estacao(desde=desde),
            tendencia=tendencia_fila["tendencia"],
            variacoes=tendencia_fila["variacoes"],
            gargalos=_gargalos_por_estacao(),
            produtividade=_produtividade_por_setor(ano, mes),
            kpi_mensal=_kpi_gerencial_mensal(ano, mes),
        )

    @app.route("/kpis/relatorio.pdf")
    @login_required
    def kpis_relatorio_pdf():
        """Relatório PDF da tela de KPIs (pedido do Bruno, 18/09/2026) —
        aceita os MESMOS parâmetros `meses`/`ano`/`mes` da tela (o botão
        "Emitir relatório" do kpis.html já manda o período/mês que estava
        selecionado na hora), e recalcula os dados com as MESMAS funções da
        rota kpis() acima — nunca corre o risco de o PDF divergir do que a
        tela mostra."""
        meses = request.args.get("meses", 6, type=int)
        if meses not in (3, 6, 12):
            meses = 6
        desde = date.today() - timedelta(days=30 * meses)

        hoje = date.today()
        ano = request.args.get("ano", hoje.year, type=int)
        mes = request.args.get("mes", hoje.month, type=int)
        if not (1 <= mes <= 12):
            mes = hoje.month

        tendencia_fila = _tendencia_fila_lead_time(meses=meses)
        return _gerar_pdf_kpis(
            meses=meses,
            ano=ano,
            mes=mes,
            lead_time_fila_estacao=_lead_time_fila_por_estacao(desde=desde),
            tendencia=tendencia_fila["tendencia"],
            variacoes=tendencia_fila["variacoes"],
            gargalos=_gargalos_por_estacao(),
            produtividade=_produtividade_por_setor(ano, mes),
            kpi_mensal=_kpi_gerencial_mensal(ano, mes),
        )

    @app.route("/kpis/mensal/<int:ano>/<int:mes>/salvar", methods=["POST"])
    @requer_role("ADMIN", "PCP")
    def kpis_mensal_salvar(ano, mes):
        """Salva os KPIs manuais do mês (itens 6/7/8 + observações de 4/5,
        pedido do Bruno, 17-18/09/2026) — 1 linha por ano+mes
        (KpiGerencialMensal), criada na hora se ainda não existir."""
        if not (1 <= mes <= 12):
            flash("Mês inválido.", "danger")
            return redirect(url_for("kpis"))

        linha = KpiGerencialMensal.query.filter_by(ano=ano, mes=mes).first()
        novo = linha is None
        if novo:
            linha = KpiGerencialMensal(ano=ano, mes=mes)

        def _float_ou_none(nome):
            valor = request.form.get(nome, "").strip().replace(",", ".")
            if not valor:
                return None
            try:
                return float(valor)
            except ValueError:
                return None

        campos = [
            "aderencia_planejado",
            "aderencia_realizado",
            "consumo_materia_prima",
            "indice_perdas",
            "indice_refugos",
            "indice_descartes",
            "obs_causas_fila_lead_time",
            "obs_gargalos_plano_acao",
        ]
        antes = {} if novo else {campo: getattr(linha, campo) for campo in campos}

        linha.aderencia_planejado = _float_ou_none("aderencia_planejado")
        linha.aderencia_realizado = _float_ou_none("aderencia_realizado")
        linha.consumo_materia_prima = _float_ou_none("consumo_materia_prima")
        linha.indice_perdas = _float_ou_none("indice_perdas")
        linha.indice_refugos = _float_ou_none("indice_refugos")
        linha.indice_descartes = _float_ou_none("indice_descartes")
        linha.obs_causas_fila_lead_time = request.form.get("obs_causas_fila_lead_time", "").strip() or None
        linha.obs_gargalos_plano_acao = request.form.get("obs_gargalos_plano_acao", "").strip() or None
        linha.atualizado_por = current_user.nome if current_user.is_authenticated else None

        if novo:
            db.session.add(linha)
            db.session.flush()  # garante linha.id preenchido antes do histórico

        depois = {campo: getattr(linha, campo) for campo in campos}
        _registrar_alteracoes("kpi_gerencial_mensal", linha.id, None, antes, depois, campos)

        db.session.commit()
        flash(f"KPIs de {MESES_PT[mes - 1]}/{ano} salvos.", "success")
        return redirect(url_for("kpis", ano=ano, mes=mes))

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

    # ------------------------------------------------------------------
    # Cadastro "Lead time Transportadora" (pedido do Bruno, 11/09/2026):
    # prazo de entrega por UF/modalidade, simulando frete saindo de
    # Pindamonhangaba-SP — ver _seed_lead_time_transportadora (dados
    # iniciais) e LeadTimeTransportadora (models.py).
    # ------------------------------------------------------------------
    _LEAD_TIME_ORDEM_REGIAO = ["Sudeste", "Sul", "Centro-Oeste", "Norte", "Nordeste"]

    def _ordenar_lead_time(linhas):
        def chave(linha):
            posicao_regiao = (
                _LEAD_TIME_ORDEM_REGIAO.index(linha.regiao)
                if linha.regiao in _LEAD_TIME_ORDEM_REGIAO
                else len(_LEAD_TIME_ORDEM_REGIAO)
            )
            posicao_modalidade = (
                LEAD_TIME_MODALIDADE_OPCOES.index(linha.modalidade)
                if linha.modalidade in LEAD_TIME_MODALIDADE_OPCOES
                else len(LEAD_TIME_MODALIDADE_OPCOES)
            )
            return (posicao_regiao, linha.uf, posicao_modalidade)

        return sorted(linhas, key=chave)

    def _validar_lead_time_form(f, ignorar_id=None):
        """Valida/normaliza o formulário de Lead time Transportadora — devolve
        (dados, None) se ok, ou (None, mensagem_erro) se algo não bater.
        Compartilhado entre novo/editar pra nunca divergir a validação."""
        origem = f.get("origem", "").strip() or _LEAD_TIME_ORIGEM_PADRAO
        uf = f.get("uf", "").strip().upper()
        modalidade = f.get("modalidade", "").strip()
        unidade_prazo = f.get("unidade_prazo", "").strip()
        observacao = f.get("observacao", "").strip() or None
        prazo_minimo = _parse_float_form(f.get("prazo_minimo"), default=None)
        prazo_maximo = _parse_float_form(f.get("prazo_maximo"), default=None)

        if uf not in UFS_BRASIL:
            return None, "Selecione um estado (UF) válido."
        if modalidade not in LEAD_TIME_MODALIDADE_OPCOES:
            return None, "Selecione a modalidade (Rodoviário ou Aéreo)."
        if unidade_prazo not in LEAD_TIME_UNIDADE_OPCOES:
            return None, "Selecione a unidade do prazo (Dias úteis ou Horas)."
        if prazo_minimo is None or prazo_minimo <= 0:
            return None, "Informe um prazo mínimo válido (maior que zero)."
        if prazo_maximo is not None and prazo_maximo < prazo_minimo:
            return None, "O prazo máximo não pode ser menor que o prazo mínimo."

        conflito = LeadTimeTransportadora.query.filter_by(origem=origem, uf=uf, modalidade=modalidade)
        if ignorar_id is not None:
            conflito = conflito.filter(LeadTimeTransportadora.id != ignorar_id)
        if conflito.first() is not None:
            return None, f"Já existe um lead time cadastrado para {uf} / {modalidade} (origem {origem})."

        return {
            "origem": origem,
            "uf": uf,
            "regiao": REGIAO_POR_UF.get(uf, ""),
            "modalidade": modalidade,
            "prazo_minimo": prazo_minimo,
            "prazo_maximo": prazo_maximo,
            "unidade_prazo": unidade_prazo,
            "observacao": observacao,
        }, None

    @app.route("/cadastros/lead-time-transportadora")
    @requer_role("ADMIN", "PCP")
    def cadastros_lead_time_transportadora():
        linhas = _ordenar_lead_time(LeadTimeTransportadora.query.all())
        return render_template(
            "cadastros_lead_time_transportadora.html",
            linhas=linhas,
            origem_padrao=_LEAD_TIME_ORIGEM_PADRAO,
        )

    @app.route("/cadastros/lead-time-transportadora/novo", methods=["GET", "POST"])
    @requer_role("ADMIN", "PCP")
    def cadastros_lead_time_transportadora_novo():
        if request.method == "POST":
            f = request.form
            dados, erro = _validar_lead_time_form(f)
            if erro:
                flash(erro, "danger")
                return render_template("cadastros_lead_time_transportadora_form.html", linha=None, form=f)

            nova = LeadTimeTransportadora(ativo=True, **dados)
            db.session.add(nova)
            db.session.commit()
            flash(f"Lead time {nova.uf} / {nova.modalidade} cadastrado com sucesso.", "success")
            return redirect(url_for("cadastros_lead_time_transportadora"))

        return render_template(
            "cadastros_lead_time_transportadora_form.html", linha=None,
            form={"origem": _LEAD_TIME_ORIGEM_PADRAO, "unidade_prazo": "Dias úteis"},
        )

    @app.route("/cadastros/lead-time-transportadora/<int:linha_id>/editar", methods=["GET", "POST"])
    @requer_role("ADMIN", "PCP")
    def cadastros_lead_time_transportadora_editar(linha_id):
        linha = db.session.get(LeadTimeTransportadora, linha_id)
        if linha is None:
            flash("Lead time não encontrado.", "danger")
            return redirect(url_for("cadastros_lead_time_transportadora"))

        if request.method == "POST":
            f = request.form
            dados, erro = _validar_lead_time_form(f, ignorar_id=linha.id)
            if erro:
                flash(erro, "danger")
                return render_template("cadastros_lead_time_transportadora_form.html", linha=linha, form=f)

            for campo, valor in dados.items():
                setattr(linha, campo, valor)
            linha.ativo = bool(f.get("ativo"))
            db.session.commit()
            flash(f"Lead time {linha.uf} / {linha.modalidade} atualizado com sucesso.", "success")
            return redirect(url_for("cadastros_lead_time_transportadora"))

        return render_template("cadastros_lead_time_transportadora_form.html", linha=linha, form={})

    # ------------------------------------------------------------------
    # Cadastros > Lead time Produção (pedido do Bruno, 11/09/2026) — mesmo
    # padrão CRUD do Lead time Transportadora acima, agora por Produto +
    # Estação. O histórico real (média/melhor/pior/tendência) NUNCA é
    # gravado aqui — é sempre recalculado ao vivo por
    # _estatisticas_lead_time_producao (ver seção antes de
    # _gargalos_por_estacao).
    # ------------------------------------------------------------------
    def _validar_lead_time_producao_form(f, ignorar_id=None):
        """Valida/normaliza o formulário de Lead time Produção — mesmo
        espírito de _validar_lead_time_form (devolve (dados, None) ou
        (None, erro), compartilhado entre novo/editar)."""
        produto = f.get("produto", "").strip()
        familia = f.get("familia", "").strip() or None
        estacao_id = f.get("estacao_id", type=int)
        lt_padrao_dias = _parse_float_form(f.get("lt_padrao_dias"), default=None)

        if not produto:
            return None, "Informe o produto (ou trecho do nome do produto)."
        if not estacao_id or db.session.get(Estacao, estacao_id) is None:
            return None, "Selecione uma estação válida."
        if lt_padrao_dias is None or lt_padrao_dias <= 0:
            return None, "Informe um LT padrão válido (em dias corridos, maior que zero)."

        conflito = LeadTimeProducao.query.filter_by(produto=produto, estacao_id=estacao_id)
        if ignorar_id is not None:
            conflito = conflito.filter(LeadTimeProducao.id != ignorar_id)
        if conflito.first() is not None:
            return None, f'Já existe um Lead time cadastrado para "{produto}" nessa estação.'

        return {
            "produto": produto,
            "familia": familia,
            "estacao_id": estacao_id,
            "lt_padrao_dias": lt_padrao_dias,
        }, None

    def _registrar_revisao_lt_producao(entrada, valor_anterior, valor_novo, motivo):
        """Grava 1 linha em LeadTimeProducaoHistorico quando o LT padrão
        muda de valor (pedido do Bruno, item 4: valor anterior/novo/data/
        responsável/motivo). Só grava quando o valor realmente mudou —
        editar outros campos (família, ativo) não gera histórico de
        revisão de LT."""
        if valor_anterior == valor_novo:
            return
        db.session.add(
            LeadTimeProducaoHistorico(
                lead_time_producao_id=entrada.id,
                valor_anterior=valor_anterior,
                valor_novo=valor_novo,
                motivo=(motivo or "").strip() or None,
                usuario_nome=current_user.nome if current_user.is_authenticated else None,
            )
        )

    @app.route("/cadastros/lead-time-producao")
    @requer_role("ADMIN", "PCP")
    def cadastros_lead_time_producao():
        entradas = (
            LeadTimeProducao.query.join(Estacao, LeadTimeProducao.estacao_id == Estacao.id)
            .order_by(Estacao.ordem_exibicao, LeadTimeProducao.produto)
            .all()
        )
        linhas = [
            {"entrada": e, "estatisticas": _estatisticas_lead_time_producao(e.produto, e.estacao.nome, e.lt_padrao_dias)}
            for e in entradas
        ]
        return render_template(
            "cadastros_lead_time_producao.html",
            linhas=linhas,
            estacoes=Estacao.query.filter_by(ativo=True).order_by(Estacao.ordem_exibicao).all(),
        )

    @app.route("/cadastros/lead-time-producao/novo", methods=["GET", "POST"])
    @requer_role("ADMIN", "PCP")
    def cadastros_lead_time_producao_novo():
        if request.method == "POST":
            f = request.form
            dados, erro = _validar_lead_time_producao_form(f)
            if erro:
                flash(erro, "danger")
                return render_template(
                    "cadastros_lead_time_producao_form.html", entrada=None, form=f,
                    estacoes=Estacao.query.filter_by(ativo=True).order_by(Estacao.ordem_exibicao).all(),
                    historico=[],
                )

            nova = LeadTimeProducao(
                ativo=True,
                responsavel_revisao=current_user.nome if current_user.is_authenticated else None,
                data_ultima_revisao=date.today(),
                **dados,
            )
            db.session.add(nova)
            db.session.commit()
            flash(f'Lead time de produção "{nova.produto}" cadastrado com sucesso.', "success")
            return redirect(url_for("cadastros_lead_time_producao"))

        return render_template(
            "cadastros_lead_time_producao_form.html", entrada=None, form={},
            estacoes=Estacao.query.filter_by(ativo=True).order_by(Estacao.ordem_exibicao).all(),
            historico=[],
        )

    @app.route("/cadastros/lead-time-producao/<int:entrada_id>/editar", methods=["GET", "POST"])
    @requer_role("ADMIN", "PCP")
    def cadastros_lead_time_producao_editar(entrada_id):
        entrada = db.session.get(LeadTimeProducao, entrada_id)
        if entrada is None:
            flash("Lead time de produção não encontrado.", "danger")
            return redirect(url_for("cadastros_lead_time_producao"))

        historico = (
            LeadTimeProducaoHistorico.query.filter_by(lead_time_producao_id=entrada.id)
            .order_by(LeadTimeProducaoHistorico.criado_em.desc())
            .all()
        )
        estacoes_ativas = Estacao.query.filter_by(ativo=True).order_by(Estacao.ordem_exibicao).all()

        if request.method == "POST":
            f = request.form
            dados, erro = _validar_lead_time_producao_form(f, ignorar_id=entrada.id)
            if erro:
                flash(erro, "danger")
                return render_template(
                    "cadastros_lead_time_producao_form.html", entrada=entrada, form=f,
                    estacoes=estacoes_ativas, historico=historico,
                )

            lt_anterior = entrada.lt_padrao_dias
            for campo, valor in dados.items():
                setattr(entrada, campo, valor)
            entrada.ativo = bool(f.get("ativo"))

            if dados["lt_padrao_dias"] != lt_anterior:
                _registrar_revisao_lt_producao(entrada, lt_anterior, dados["lt_padrao_dias"], f.get("motivo_alteracao"))
                entrada.data_ultima_revisao = date.today()
                entrada.responsavel_revisao = current_user.nome if current_user.is_authenticated else None

            db.session.commit()
            flash(f'Lead time de produção "{entrada.produto}" atualizado com sucesso.', "success")
            return redirect(url_for("cadastros_lead_time_producao"))

        return render_template(
            "cadastros_lead_time_producao_form.html", entrada=entrada, form={},
            estacoes=estacoes_ativas, historico=historico,
        )

    # ------------------------------------------------------------------
    # GESTÃO DE CUSTOS (pedido do Bruno, 20/09/2026) — Fase 1: grupo PIG
    # MANDRIL. Visualização liberada pra ADMIN/PCP/GESTAO (mesmo grupo que
    # já vê KPIs/Gargalos/Faturamento); edição (matéria-prima, hora-homem,
    # estrutura) só ADMIN/PCP, via @requer_role — mesmo padrão do resto do
    # app. Plano completo em /root/.claude/plans/joyful-knitting-hoare.md.
    # ------------------------------------------------------------------
    @app.route("/custos")
    @requer_role("ADMIN", "PCP", "GESTAO")
    def custos_painel():
        return render_template("custos_painel.html", visao=_visao_rapida_custos())

    @app.route("/custos/produtos")
    @requer_role("ADMIN", "PCP", "GESTAO")
    def custos_produtos():
        familia = request.args.get("familia") or None
        linhas = _produtos_catalogo(familia=familia)
        familias_com_dados = sorted({p.familia for p in Produto.query.filter_by(ativo=True).all()})
        return render_template(
            "custos_produtos.html", linhas=linhas, familia_selecionada=familia,
            familias_com_dados=familias_com_dados, familias_futuras=_FAMILIAS_FASE_SEGUINTE,
        )

    @app.route("/custos/produtos/<int:produto_id>")
    @requer_role("ADMIN", "PCP", "GESTAO")
    def custos_produto_detalhe(produto_id):
        produto = db.session.get(Produto, produto_id)
        if produto is None:
            flash("Produto não encontrado.", "danger")
            return redirect(url_for("custos_produtos"))
        estruturas = sorted([e for e in produto.estruturas if e.ativo], key=lambda e: _chave_ordenacao_dn(e.dn))
        composicoes = [{"estrutura": e, "calc": _custo_estrutura_produto(e)} for e in estruturas]
        return render_template("custos_produto_detalhe.html", produto=produto, composicoes=composicoes)

    @app.route("/custos/configurador-pig")
    @requer_role("ADMIN", "PCP", "GESTAO")
    def custos_configurador_pig():
        """Configurador de acessórios pro LBD/LUN (pedido do Bruno, 21/09/2026): escolhe
        o PIG (LBD ou LUN) + DN, marca quais acessórios entram (ELC aço, ELP PP, Cinta
        Magnética, Placa Calibradora — hoje já cadastrados como Produto próprio da
        família ELC_MG_PC, com custo por DN — e Alojamento) e vê o custo total do
        conjunto ajustar na hora. Tela própria e simples (sem os campos de "novo
        custo"/histórico da tela de Simulação), só GET, nada é salvo.

        Réplica fiel do "CONFIGURADOR DE CUSTOS ADICIONAIS" que já existe dentro da
        própria aba LBD/LUN (linhas 31+/32+, colunas A-S — achado só depois do Bruno
        perguntar sobre o disco espaçador; na fase 1 eu tinha checado só as colunas
        BH/BI/BJ daquele intervalo, que ficam vazias, e concluído errado que era um
        bloco sem dado). 2 regras que não são óbvias e vêm direto das fórmulas de lá
        (D32/F32/H32/J32/O32/P32/Q32 da aba LBD, idêntico na LUN):
        1) ELC (AÇO) e ELP (PP) somam, além do próprio custo, um valor de "CALANDRA"
           por DN (tabela própria, ver `_seed_custos_pig_calandra`) — CINTA MAGNÉTICA e
           PLACA CALIBRADORA não somam calandra.
        2) Cada acessório marcado entre ELC/ELP/CINTA/PLACA (Alojamento NÃO conta)
           soma mais 1 Disco Espaçador (matéria-prima "PUCAST-MP-DE-DN{dn}", a mesma
           já usada na estrutura padrão do LBD/LUN) ao custo — 2 acessórios marcados
           = 2 espaçadores extras, e assim por diante."""
        produtos_base = Produto.query.filter(Produto.familia.in_(("LBD", "LUN")), Produto.ativo == True).order_by(Produto.familia, Produto.codigo).all()  # noqa: E712
        produto_id = request.args.get("produto_id", type=int)
        produto_selecionado = db.session.get(Produto, produto_id) if produto_id else None
        if produto_selecionado is None or produto_selecionado.familia not in ("LBD", "LUN"):
            produto_selecionado = produtos_base[0] if produtos_base else None

        dns_disponiveis = []
        if produto_selecionado is not None:
            dns_disponiveis = sorted([e.dn for e in produto_selecionado.estruturas if e.ativo], key=_chave_ordenacao_dn)
        dn_selecionado = request.args.get("dn") or (dns_disponiveis[0] if dns_disponiveis else None)

        estrutura_base = None
        calc_base = None
        if produto_selecionado is not None and dn_selecionado:
            estrutura_base = EstruturaProduto.query.filter_by(produto_id=produto_selecionado.id, dn=dn_selecionado, ativo=True).first()
        if estrutura_base is not None:
            calc_base = _custo_estrutura_produto(estrutura_base)

        acessorios = []
        qtd_conta_espacador = 0
        if dn_selecionado:
            for chave, nome_produto, tem_calandra in (
                ("elc", "ELC (AÇO)", True), ("elp", "ELP (PP)", True),
                ("cinta", "CINTA MAGNÉTICA", False), ("placa", "PLACA CALIBRADORA", False),
            ):
                produto_acc = Produto.query.filter_by(familia="ELC_MG_PC", codigo=nome_produto, ativo=True).first()
                calc_acc = None
                if produto_acc is not None:
                    estrutura_acc = EstruturaProduto.query.filter_by(produto_id=produto_acc.id, dn=dn_selecionado, ativo=True).first()
                    if estrutura_acc is not None:
                        calc_acc = _custo_estrutura_produto(estrutura_acc)
                custo_calandra = None
                if tem_calandra:
                    mp_calandra = MateriaPrima.query.filter_by(codigo=f"CALANDRA-DN{dn_selecionado}", ativo=True).first()
                    custo_calandra = mp_calandra.custo_atual if mp_calandra else 0.0
                custo_item = None
                if calc_acc is not None:
                    custo_item = calc_acc["custo_total"] + (custo_calandra or 0.0)
                marcado = request.args.get(f"acc_{chave}") == "1"
                acessorios.append({
                    "chave": chave, "nome": nome_produto, "disponivel": custo_item is not None,
                    "custo": custo_item, "custo_calandra": custo_calandra if tem_calandra else None,
                    "marcado": marcado, "conta_espacador": True,
                })
                if marcado and custo_item is not None:
                    qtd_conta_espacador += 1
            mp_alojamento = MateriaPrima.query.filter_by(codigo=f"ALOJAMENTO-DN{dn_selecionado}", ativo=True).first()
            acessorios.append({
                "chave": "alojamento", "nome": "Alojamento (embalagem)", "disponivel": mp_alojamento is not None,
                "custo": mp_alojamento.custo_atual if mp_alojamento else None, "custo_calandra": None,
                "marcado": request.args.get("acc_alojamento") == "1", "conta_espacador": False,
            })

        mp_espacador = MateriaPrima.query.filter_by(codigo=f"PUCAST-MP-DE-DN{dn_selecionado}", ativo=True).first() if dn_selecionado else None
        custo_unit_espacador = mp_espacador.custo_atual if mp_espacador else 0.0
        custo_espacadores = qtd_conta_espacador * custo_unit_espacador

        custo_acessorios = sum((a["custo"] or 0) for a in acessorios if a["marcado"] and a["disponivel"])
        custo_base = calc_base["custo_total"] if calc_base else 0.0
        custo_total_combinado = custo_base + custo_acessorios + custo_espacadores

        return render_template(
            "custos_configurador_pig.html", produtos_base=produtos_base, produto_selecionado=produto_selecionado,
            dns_disponiveis=dns_disponiveis, dn_selecionado=dn_selecionado, calc_base=calc_base,
            acessorios=acessorios, custo_acessorios=custo_acessorios, qtd_conta_espacador=qtd_conta_espacador,
            custo_unit_espacador=custo_unit_espacador, custo_espacadores=custo_espacadores,
            custo_total_combinado=custo_total_combinado,
        )

    @app.route("/custos/estrutura/<int:produto_id>/<path:dn>/editar", methods=["GET", "POST"])
    @requer_role("ADMIN", "PCP")
    def custos_estrutura_editar(produto_id, dn):
        produto = db.session.get(Produto, produto_id)
        if produto is None:
            flash("Produto não encontrado.", "danger")
            return redirect(url_for("custos_produtos"))
        estrutura = EstruturaProduto.query.filter_by(produto_id=produto_id, dn=dn).first()
        if estrutura is None:
            flash(f'Estrutura para DN "{dn}" não encontrada.', "danger")
            return redirect(url_for("custos_produto_detalhe", produto_id=produto_id))

        if request.method == "POST":
            acao = request.form.get("acao")
            if acao == "adicionar_item":
                tipo = request.form.get("tipo") or "MATERIA_PRIMA"
                quantidade = _parse_float_form(request.form.get("quantidade"), default=None)
                if quantidade is None or quantidade <= 0:
                    flash("Informe uma quantidade válida (maior que zero).", "danger")
                elif tipo == "SUBPRODUTO":
                    subproduto_id = request.form.get("subproduto_id", type=int)
                    if not subproduto_id:
                        flash("Selecione o subproduto.", "danger")
                    else:
                        db.session.add(EstruturaProdutoItem(
                            estrutura_id=estrutura.id, tipo="SUBPRODUTO", subproduto_id=subproduto_id,
                            quantidade=quantidade, observacao=(request.form.get("observacao") or "").strip() or None,
                            ordem=len(estrutura.itens),
                        ))
                        db.session.commit()
                        flash("Item adicionado à estrutura.", "success")
                else:
                    materia_prima_id = request.form.get("materia_prima_id", type=int)
                    if not materia_prima_id:
                        flash("Selecione a matéria-prima.", "danger")
                    else:
                        db.session.add(EstruturaProdutoItem(
                            estrutura_id=estrutura.id, tipo="MATERIA_PRIMA", materia_prima_id=materia_prima_id,
                            quantidade=quantidade, observacao=(request.form.get("observacao") or "").strip() or None,
                            ordem=len(estrutura.itens),
                        ))
                        db.session.commit()
                        flash("Item adicionado à estrutura.", "success")
            elif acao == "remover_item":
                item_id = request.form.get("item_id", type=int)
                item = db.session.get(EstruturaProdutoItem, item_id)
                if item and item.estrutura_id == estrutura.id:
                    db.session.delete(item)
                    db.session.commit()
                    flash("Item removido da estrutura.", "success")
            elif acao == "salvar_ciclo":
                ciclo = _parse_float_form(request.form.get("ciclo_horas"), default=None)
                if ciclo is None or ciclo < 0:
                    flash("Informe um ciclo de horas válido.", "danger")
                else:
                    estrutura.ciclo_horas = ciclo
                    estrutura.observacao = (request.form.get("observacao") or "").strip() or None
                    estrutura.atualizado_por = current_user.nome if current_user.is_authenticated else None
                    db.session.commit()
                    flash("Ciclo de horas atualizado.", "success")
            return redirect(url_for("custos_estrutura_editar", produto_id=produto_id, dn=dn))

        materias_primas = MateriaPrima.query.filter_by(ativo=True).order_by(MateriaPrima.codigo).all()
        subprodutos = Produto.query.filter(Produto.ativo == True, Produto.id != produto_id).order_by(Produto.familia, Produto.codigo).all()  # noqa: E712
        calc = _custo_estrutura_produto(estrutura)
        return render_template(
            "custos_estrutura_editar.html", produto=produto, estrutura=estrutura, calc=calc,
            materias_primas=materias_primas, subprodutos=subprodutos,
        )

    @app.route("/custos/materias-primas")
    @requer_role("ADMIN", "PCP", "GESTAO")
    def custos_materias_primas():
        """Pedido do Bruno (21/09/2026): a tela não pode misturar matéria-prima
        de verdade (validada na aba PARÂMETROS da planilha) com produto
        acabado/custo consolidado (Placa Calibradora, Cinta Magnética, ELC,
        ELP, Alojamento, Calandra, discos "PU CAST consolidado"...).

        Mostra primeiro, agrupado exatamente como a aba PARÂMETROS (mesmas
        seções, mesma ordem), só o que é matéria-prima validada. Os itens
        "derivados" continuam existindo (o motor de custo depende deles)
        mas ficam numa seção separada e recolhida no fim da página — dá pra
        editar o custo deles do mesmo jeito, só não aparecem misturados."""

        def _chave_ordenacao_mp(mp):
            m = re.search(r"DN(\d+)", mp.codigo)
            if m:
                return (0, int(m.group(1)))
            return (1, mp.descricao or mp.codigo)

        todas = MateriaPrima.query.all()
        parametros = [mp for mp in todas if mp.origem_planilha == "PARAMETROS"]
        derivadas = [mp for mp in todas if mp.origem_planilha != "PARAMETROS"]

        secoes = []
        for nome, _prefixos in _ORIGEM_MP_SECOES:
            itens = sorted((mp for mp in parametros if _secao_materia_prima(mp.codigo) == nome), key=_chave_ordenacao_mp)
            if itens:
                secoes.append({"nome": nome, "itens": itens})

        derivadas_por_categoria = {}
        for mp in sorted(derivadas, key=lambda mp: (mp.categoria or "", mp.codigo)):
            derivadas_por_categoria.setdefault(mp.categoria or "Sem categoria", []).append(mp)

        return render_template(
            "custos_materias_primas.html",
            secoes=secoes, derivadas_por_categoria=derivadas_por_categoria,
            total_parametros=len(parametros), total_derivadas=len(derivadas),
        )

    @app.route("/custos/materias-primas/novo", methods=["GET", "POST"])
    @requer_role("ADMIN", "PCP")
    def custos_materias_primas_novo():
        if request.method == "POST":
            f = request.form
            dados, erro = _validar_materia_prima_form(f)
            if erro:
                flash(erro, "danger")
                return render_template("custos_materia_prima_form.html", mp=None, form=f, historico=[])
            nova = MateriaPrima(ativo=True, **dados)
            db.session.add(nova)
            db.session.commit()
            flash(f'Matéria-prima "{nova.codigo}" cadastrada com sucesso.', "success")
            return redirect(url_for("custos_materias_primas"))
        return render_template("custos_materia_prima_form.html", mp=None, form={}, historico=[])

    @app.route("/custos/materias-primas/<int:mp_id>/editar", methods=["GET", "POST"])
    @requer_role("ADMIN", "PCP")
    def custos_materias_primas_editar(mp_id):
        mp = db.session.get(MateriaPrima, mp_id)
        if mp is None:
            flash("Matéria-prima não encontrada.", "danger")
            return redirect(url_for("custos_materias_primas"))
        historico = MateriaPrimaHistorico.query.filter_by(materia_prima_id=mp.id).order_by(MateriaPrimaHistorico.criado_em.desc()).all()

        if request.method == "POST":
            f = request.form
            dados, erro = _validar_materia_prima_form(f, ignorar_id=mp.id)
            if erro:
                flash(erro, "danger")
                return render_template("custos_materia_prima_form.html", mp=mp, form=f, historico=historico)

            custo_anterior = mp.custo_atual
            for campo, valor in dados.items():
                setattr(mp, campo, valor)
            mp.ativo = bool(f.get("ativo"))

            if dados["custo_atual"] != custo_anterior:
                _registrar_revisao_materia_prima(mp, custo_anterior, dados["custo_atual"], f.get("motivo_alteracao"))
                mp.data_atualizacao_fornecedor = date.today()

            db.session.commit()
            flash(f'Matéria-prima "{mp.codigo}" atualizada com sucesso.', "success")
            return redirect(url_for("custos_materias_primas"))

        return render_template("custos_materia_prima_form.html", mp=mp, form={}, historico=historico)

    @app.route("/custos/hora-homem", methods=["GET", "POST"])
    @requer_role("ADMIN", "PCP", "GESTAO")
    def custos_hora_homem():
        param = db.session.get(ParametroHoraHomem, 1)
        if param is None:
            param = ParametroHoraHomem(id=1, valor=40.0)
            db.session.add(param)
            db.session.commit()
        historico = ParametroHoraHomemHistorico.query.order_by(ParametroHoraHomemHistorico.criado_em.desc()).all()

        if request.method == "POST":
            if current_user.role not in ("ADMIN", "PCP"):
                abort(403)
            novo_valor = _parse_float_form(request.form.get("valor"), default=None)
            if novo_valor is None or novo_valor <= 0:
                flash("Informe um valor de hora-homem válido (maior que zero).", "danger")
                return redirect(url_for("custos_hora_homem"))
            valor_anterior = param.valor
            if novo_valor != valor_anterior:
                db.session.add(ParametroHoraHomemHistorico(
                    valor_anterior=valor_anterior, valor_novo=novo_valor,
                    motivo=(request.form.get("motivo_alteracao") or "").strip() or None,
                    usuario_nome=current_user.nome if current_user.is_authenticated else None,
                ))
                param.valor = novo_valor
                param.atualizado_por = current_user.nome if current_user.is_authenticated else None
                db.session.commit()
                flash(f"Valor de hora-homem atualizado para R$ {novo_valor:.2f}/h — todos os produtos foram recalculados automaticamente.", "success")
            return redirect(url_for("custos_hora_homem"))

        return render_template("custos_hora_homem.html", param=param, historico=historico)

    @app.route("/custos/correspondencias-manuais")
    @requer_role("ADMIN", "PCP", "GESTAO")
    def custos_correspondencias_manuais():
        """Ferramenta visível de correlação manual pedida pelo Bruno
        (23/09/2026) — lista/gerencia as correspondências já cadastradas.
        Ver docstring de CorrespondenciaManualCusto (models.py) e da checagem
        em _matching_produto_pcp pro racional completo."""
        correspondencias = CorrespondenciaManualCusto.query.order_by(CorrespondenciaManualCusto.criado_em.desc()).all()
        return render_template("custos_correspondencias_manuais.html", correspondencias=correspondencias)

    @app.route("/custos/correspondencias-manuais/nova", methods=["GET", "POST"])
    @requer_role("ADMIN", "PCP")
    def custos_correspondencias_manuais_nova():
        item_id = request.values.get("item_id", type=int)
        item_pedido = db.session.get(ItemPedido, item_id) if item_id else None

        produtos = Produto.query.filter_by(ativo=True).order_by(Produto.familia, Produto.codigo).all()
        produtos_dns = {
            p.id: sorted([e.dn for e in p.estruturas if e.ativo], key=_chave_ordenacao_dn)
            for p in produtos
        }

        if request.method == "POST":
            descricao_original = (request.form.get("descricao_original") or "").strip()
            produto_id = request.form.get("produto_id", type=int)
            dn = (request.form.get("dn") or "").strip()
            observacao = (request.form.get("observacao") or "").strip() or None
            produto_selecionado = db.session.get(Produto, produto_id) if produto_id else None

            erro = None
            estrutura_base = None
            if not descricao_original:
                erro = "Informe a descrição do item do PCP que deve usar esta correspondência."
            elif produto_selecionado is None:
                erro = "Selecione um produto do catálogo de custos."
            elif not dn:
                erro = "Selecione a DN (estrutura) do produto que servirá de base."
            else:
                estrutura_base = EstruturaProduto.query.filter_by(produto_id=produto_selecionado.id, dn=dn, ativo=True).first()
                if estrutura_base is None:
                    erro = "Essa DN não tem estrutura de custo cadastrada para o produto selecionado."

            if erro:
                flash(erro, "danger")
                return render_template(
                    "custos_correspondencia_manual_form.html", item_pedido=item_pedido,
                    produtos=produtos, produtos_dns=produtos_dns, produto_selecionado=produto_selecionado,
                    dn_selecionado=dn, descricao_original=descricao_original, observacao=observacao,
                )

            descricao_normalizada = _normalizar_texto_matching_custos(descricao_original)
            existente = CorrespondenciaManualCusto.query.filter_by(descricao_normalizada=descricao_normalizada).first()
            nome_usuario = current_user.nome if current_user.is_authenticated else None
            if existente is not None:
                existente.descricao_original = descricao_original
                existente.produto_id = produto_selecionado.id
                existente.dn = dn
                existente.observacao = observacao
                existente.criado_por = nome_usuario
                flash(f'Correspondência manual de "{descricao_original}" atualizada com sucesso.', "success")
            else:
                db.session.add(CorrespondenciaManualCusto(
                    descricao_normalizada=descricao_normalizada, descricao_original=descricao_original,
                    produto_id=produto_selecionado.id, dn=dn, observacao=observacao, criado_por=nome_usuario,
                ))
                flash(f'Correspondência manual de "{descricao_original}" cadastrada com sucesso — a partir de agora, qualquer item do PCP com essa descrição usa {produto_selecionado.codigo} DN {dn} automaticamente.', "success")
            db.session.commit()
            return redirect(url_for("custos_correspondencias_manuais"))

        descricao_inicial = item_pedido.descricao_produto if item_pedido else (request.args.get("descricao_original") or "")
        return render_template(
            "custos_correspondencia_manual_form.html", item_pedido=item_pedido,
            produtos=produtos, produtos_dns=produtos_dns, produto_selecionado=None,
            dn_selecionado=None, descricao_original=descricao_inicial, observacao="",
        )

    @app.route("/custos/correspondencias-manuais/<int:corr_id>/excluir", methods=["POST"])
    @requer_role("ADMIN", "PCP")
    def custos_correspondencias_manuais_excluir(corr_id):
        corr = db.session.get(CorrespondenciaManualCusto, corr_id)
        if corr is not None:
            descricao = corr.descricao_original
            db.session.delete(corr)
            db.session.commit()
            flash(f'Correspondência manual de "{descricao}" removida.', "success")
        return redirect(url_for("custos_correspondencias_manuais"))

    @app.route("/custos/necessidades-pcp")
    @requer_role("ADMIN", "PCP", "GESTAO")
    def custos_necessidades_pcp():
        dados = _necessidades_pcp_materia_prima()
        return render_template("custos_necessidades_pcp.html", linhas=dados["linhas"], nao_identificados=dados["nao_identificados"])

    @app.route("/custos/simulacao")
    @requer_role("ADMIN", "PCP", "GESTAO")
    def custos_simulacao():
        # --- Simulação: "e se" um custo de matéria-prima / hora-homem / quantidade
        # mudasse — sem alterar nada de verdade (GET puro, nada é salvo). ---
        produtos_pai = (
            Produto.query.filter_by(ativo=True)
            .order_by(Produto.familia, Produto.codigo)
            .all()
        )
        produto_id = request.args.get("produto_id", type=int)
        dn_selecionado = request.args.get("dn") or None

        produto_selecionado = None
        dns_disponiveis = []
        estrutura = None
        materias_usadas = []
        calc_atual = None
        calc_simulado = None
        hh_atual = _hora_homem_atual()
        hh_override_valor = _parse_float_form(request.args.get("hh_override"), default=None)
        quantidade = _parse_float_form(request.args.get("quantidade"), default=None) or 1.0
        simulando = False

        if produto_id:
            produto_selecionado = db.session.get(Produto, produto_id)
        if produto_selecionado is not None:
            dns_disponiveis = sorted(
                [e.dn for e in produto_selecionado.estruturas if e.ativo],
                key=_chave_ordenacao_dn,
            )
            if dn_selecionado is None and dns_disponiveis:
                dn_selecionado = dns_disponiveis[0]
            if dn_selecionado:
                estrutura = EstruturaProduto.query.filter_by(
                    produto_id=produto_selecionado.id, dn=dn_selecionado, ativo=True
                ).first()
            if estrutura is not None:
                materias_usadas = _materias_primas_usadas(estrutura)
                calc_atual = _custo_estrutura_produto(estrutura)

                # overrides vêm de "mp_<id>" na query string — só monta o dict de
                # simulação se pelo menos 1 valor foi realmente alterado do padrão
                # (senão a "simulação" seria idêntica ao atual e só confundiria).
                overrides_mp = {}
                algo_alterado = False
                for mp in materias_usadas:
                    valor = _parse_float_form(request.args.get(f"mp_{mp.id}"), default=None)
                    if valor is not None:
                        overrides_mp[mp.id] = valor
                        if round(valor, 6) != round(mp.custo_atual, 6):
                            algo_alterado = True
                if hh_override_valor is not None and round(hh_override_valor, 6) != round(hh_atual, 6):
                    algo_alterado = True
                if quantidade != 1.0:
                    algo_alterado = True

                if request.args.get("simular") and algo_alterado:
                    simulando = True
                    calc_simulado = _custo_estrutura_produto(
                        estrutura,
                        overrides_mp=overrides_mp or None,
                        override_hh=hh_override_valor,
                    )

        # --- Histórico: MateriaPrimaHistorico + ParametroHoraHomemHistorico, num
        # único timeline (filtrável por matéria-prima). ---
        mp_historico_id = request.args.get("mp_historico_id", type=int)
        mps_por_id = {mp.id: mp for mp in MateriaPrima.query.all()}
        q_historico = MateriaPrimaHistorico.query
        if mp_historico_id:
            q_historico = q_historico.filter_by(materia_prima_id=mp_historico_id)
        eventos = []
        for h in q_historico.order_by(MateriaPrimaHistorico.criado_em.desc()).limit(200).all():
            mp_ref = mps_por_id.get(h.materia_prima_id)
            eventos.append({
                "data": h.criado_em, "tipo": "Matéria-prima",
                "item": f"{mp_ref.codigo} — {mp_ref.descricao}" if mp_ref else "?",
                "valor_anterior": h.custo_anterior, "valor_novo": h.custo_novo,
                "motivo": h.motivo, "usuario_nome": h.usuario_nome,
            })
        if not mp_historico_id:
            for h in ParametroHoraHomemHistorico.query.order_by(ParametroHoraHomemHistorico.criado_em.desc()).limit(50).all():
                eventos.append({
                    "data": h.criado_em, "tipo": "Hora-homem",
                    "item": "Valor de hora-homem (R$/h)",
                    "valor_anterior": h.valor_anterior, "valor_novo": h.valor_novo,
                    "motivo": h.motivo, "usuario_nome": h.usuario_nome,
                })
            eventos.sort(key=lambda e: e["data"] or datetime.min, reverse=True)
        eventos = eventos[:200]

        materias_para_filtro = MateriaPrima.query.order_by(MateriaPrima.codigo).all()

        return render_template(
            "custos_simulacao.html",
            produtos_pai=produtos_pai, produto_selecionado=produto_selecionado,
            dns_disponiveis=dns_disponiveis, dn_selecionado=dn_selecionado,
            estrutura=estrutura, materias_usadas=materias_usadas,
            calc_atual=calc_atual, calc_simulado=calc_simulado, simulando=simulando,
            hh_atual=hh_atual, hh_override_valor=hh_override_valor, quantidade=quantidade,
            eventos=eventos, materias_para_filtro=materias_para_filtro, mp_historico_id=mp_historico_id,
        )

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
        pd_alertas = _alertas_pd()

        return render_template(
            "alertas.html",
            pedidos_atrasados=pedidos_atrasados,
            pedidos_vencendo=pedidos_vencendo,
            gargalos_criticos=gargalos_criticos,
            faturamento_pendente=faturamento_pendente,
            pd_alertas=pd_alertas,
            PD_DIAS_SEM_ATUALIZACAO=PD_DIAS_SEM_ATUALIZACAO,
            PD_DIAS_PARADO=PD_DIAS_PARADO,
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

    @app.route("/api/relatorio-diario")
    def api_relatorio_diario():
        """Relatório Diário (Qualidade + Produção) pra automação externa —
        pedido do Bruno (09/09/2026): notificação automática todo dia de
        manhã, "sem eu precisar colocar a mão no site". NÃO é uma tela do
        site (sem link em lugar nenhum da interface, sem @login_required) —
        quem chama é uma tarefa agendada fora do navegador, então a
        autenticação é por token compartilhado contra a variável de ambiente
        RELATORIO_DIARIO_TOKEN no Render, em vez de sessão de usuário. Sem
        essa variável configurada no ambiente, a rota fica sempre bloqueada
        — nunca fica aberta por acidente em produção.

        Token aceito de 2 formas: header X-Report-Key (preferível — não fica
        em log nenhum) OU querystring ?chave=... (10/09/2026: a automação
        precisou trocar de um Bash/curl com header pra um fetch de página
        sem suporte a header customizado, por causa de uma restrição de
        rede de saída do ambiente onde a automação roda — sem acesso ao
        plano Team/Enterprise do Bruno pra liberar isso nas configurações
        da organização. Ele confirmou ciente do trade-off: o token pode
        aparecer em logs de acesso do Render dessa forma).

        `?dia=AAAA-MM-DD` (opcional) força um dia específico, pra testar;
        sem o parâmetro, usa "ontem" no fuso de Brasília (uso normal, pela
        tarefa agendada)."""
        token_esperado = os.environ.get("RELATORIO_DIARIO_TOKEN")
        token_recebido = request.headers.get("X-Report-Key") or request.args.get("chave")
        if not token_esperado or token_recebido != token_esperado:
            abort(403)
        dia_brt = None
        dia_str = request.args.get("dia", "").strip()
        if dia_str:
            try:
                dia_brt = datetime.strptime(dia_str, "%Y-%m-%d").date()
            except ValueError:
                dia_brt = None
        return jsonify(_relatorio_diario_dados(dia_brt))

    @app.route("/relatorios/listagem.csv")
    @login_required
    def relatorio_listagem_csv():
        query, _ = _filtrar_pedidos(request.args)
        cabecalho, linhas = _linhas_export_listagem(query.all())
        return _responder_csv("listagem_pedidos.csv", cabecalho, linhas)

    @app.route("/relatorios/listagem.xlsx")
    @login_required
    def relatorio_listagem_xlsx():
        """Relatório Excel "padrão gerencial" (pedido do Bruno, 21/09/2026) —
        reaproveita a MESMA lista achatada por item que a Listagem Geral usa
        na tela (_linhas_listagem_geral), então o Excel sempre reflete
        exatamente o filtro que estava ativo quando o botão foi clicado.
        Ver _responder_xlsx_listagem_geral pro detalhe do que entra em cada
        aba."""
        query, _ = _filtrar_pedidos(request.args)
        linhas = _linhas_listagem_geral(query.all(), request.args)
        return _responder_xlsx_listagem_geral(linhas)

    @app.route("/relatorios/listagem-geral-semanal.pdf")
    @login_required
    def relatorio_listagem_geral_semanal_pdf():
        """PDF "Emitir relatório" — Planejamento Mensal PCP/Operação (pedido
        original do Bruno, 14/09/2026, então "Listagem Geral"; aprimorado a
        pedido dele em 21/09/2026 com a opção de incluir também o mês
        seguinte como "backlog"/projeção PCP no mesmo PDF). Agrupado por
        semana (Planejamento PCP) dentro de cada mês, com somatório de
        faturamento e de nº de pedidos por semana e do mês inteiro. Ver
        _gerar_pdf_planejamento_mensal_pcp.

        O mês é OBRIGATÓRIO pro relatório fazer sentido (agrupar por semana
        só cabe dentro de 1 mês por vez) — se a tela não tinha nenhum
        "Planejamento mensal (PCP)" selecionado, cai no mês atual por
        padrão, em vez de dar erro ou mostrar tudo misturado. Os OUTROS
        filtros que já estiverem ativos na tela (cliente, vendedor, status,
        estação, busca, produto, região, datas de inclusão, atrasados)
        continuam valendo — reaproveita _filtrar_pedidos, então o relatório
        nunca diverge do que a própria tela mostraria com esse recorte.

        `incluir_backlog=1` (checkbox do modal "Emitir relatório") soma um
        segundo bloco com o mês SEGUINTE ao escolhido (via _somar_meses,
        nunca mês fixo) — mesmos outros filtros, sua própria consulta
        (_filtrar_pedidos de novo com planejamento_mensal trocado), sem
        misturar as linhas de um mês com o outro em nenhum momento.

        `modelo_relatorio` (select do modal, pedido do Bruno, 21/09/2026:
        "quero que tenha dois modelo de relatorio em pdf... um... que já
        esta estabelecido... e outro... mais compacto... cada linha dentro
        da semana agrupada representara um pedido de venda") — "completo"
        (padrão) ou "compacto"; vale pros dois blocos (mês + backlog)
        igual, só passa direto pra _gerar_pdf_planejamento_mensal_pcp."""
        args = request.args.to_dict()
        if not args.get("planejamento_mensal", "").strip():
            args["planejamento_mensal"] = date.today().strftime("%Y-%m")
        query, filtros = _filtrar_pedidos(args)
        pedidos = query.all()
        linhas = _linhas_listagem_geral(pedidos, args)
        mes_ano = _parse_mes_ano_form(filtros.get("planejamento_mensal"), None) or (date.today().year, date.today().month)

        incluir_backlog = args.get("incluir_backlog", "").strip() == "1"
        if incluir_backlog:
            blocos = [
                {"mes_ano": mes_ano, "linhas": linhas, "rotulo": "MÊS DO RELATÓRIO"},
            ]
            ano_bl, mes_bl = _somar_meses(mes_ano[0], mes_ano[1], 1)
            args_backlog = dict(args)
            args_backlog["planejamento_mensal"] = f"{ano_bl:04d}-{mes_bl:02d}"
            query_bl, _ = _filtrar_pedidos(args_backlog)
            linhas_bl = _linhas_listagem_geral(query_bl.all(), args_backlog)
            blocos.append({"mes_ano": (ano_bl, mes_bl), "linhas": linhas_bl, "rotulo": "BACKLOG — MÊS SEGUINTE (PROJEÇÃO PCP)"})
        else:
            blocos = [{"mes_ano": mes_ano, "linhas": linhas, "rotulo": "MÊS DO RELATÓRIO"}]

        modelo_relatorio = args.get("modelo_relatorio", "completo").strip() or "completo"
        if modelo_relatorio not in ("completo", "compacto"):
            modelo_relatorio = "completo"
        return _gerar_pdf_planejamento_mensal_pcp(blocos, filtros, modelo_relatorio)

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

        # Pedido do Bruno (03/09/2026): os cards do topo (Total/Pendentes/Em
        # tratativa/Em andamento/Finalizados/Valor total) agora recalculam em
        # cima do resultado já filtrado, em vez do banco inteiro sempre —
        # ver _calcular_resumo_filtrado.
        resumo = _calcular_resumo_filtrado(linhas)

        filtros_paginacao = dict(filtros, sort=sort, dir=dir_ordenacao)

        # Link de cada card de status: mantém os OUTROS filtros ativos
        # (mês, região, estação, busca...) e só troca o status — clicar em
        # "Pendentes" com um filtro de região já aplicado continua só
        # naquela região, em vez de resetar tudo (mesmo pedido do Bruno, já
        # que os números do card agora refletem esse recorte).
        filtros_status_cards = {
            "total": dict(filtros, status=""),
            "pendente": dict(filtros, status="PENDENTE"),
            "em_tratativa": dict(filtros, status="EM TRATATIVA"),
            "andamento": dict(filtros, status="ANDAMENTO"),
            "finalizado": dict(filtros, status="FINALIZADO"),
        }

        # Qualidade (RDIM) — pedido do Bruno (02/09/2026): coluna de status
        # de qualidade por item, direto na Listagem Geral.
        inspecoes_rdim = _inspecoes_rdim_por_item([l.item_id for l in linhas_pagina])

        # Quadrantes de Planejamento Semanal/Mensal PCP (pedido do Bruno,
        # 10/09/2026) — ver _quadrantes_planejamento_semanal.
        quadrantes_pcp = _quadrantes_planejamento_semanal(filtros)

        return render_template(
            "dashboard.html",
            linhas=linhas_pagina,
            resumo=resumo,
            page=page,
            total_paginas=total_paginas,
            total_filtrado=total_filtrado,
            filtros=filtros,
            filtros_paginacao=filtros_paginacao,
            filtros_status_cards=filtros_status_cards,
            sort=sort,
            dir_ordenacao=dir_ordenacao,
            inspecoes_rdim=inspecoes_rdim,
            quadrantes_pcp=quadrantes_pcp,
            # Valor padrão do seletor de mês no modal "Emitir relatório"
            # (pedido do Bruno, 14/09/2026) quando nenhum "Planejamento
            # mensal (PCP)" já estiver filtrado na tela.
            mes_atual_input=date.today().strftime("%Y-%m"),
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
            # Alimenta automaticamente a Listagem Geral de Gestão Operação
            # (pedido do Bruno, 01/09/2026) — 1 PedidoOperacao criado junto,
            # cópia inicial dos dados; dali em diante cada um é editado
            # independente na sua própria tela (ver
            # _criar_pedido_operacao_a_partir_de_producao).
            _criar_pedido_operacao_a_partir_de_producao(pedido, f)
            db.session.commit()
            flash(
                f"Pedido de {pedido.cliente} incluído com sucesso ({len(itens)} item(ns)) "
                "— também já apareceu na Listagem Geral de Gestão Operação.",
                "success",
            )
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
            inicios_producao = f.getlist("item_inicio_producao[]")
            # Pedido do Bruno (01/09/2026): a tela de edição não mostra mais
            # início/término de inspeção-embalagem e liberação de faturamento
            # como campos separados — só "Conclusão produção", um único campo
            # que grava nos dois (termino_inspecao E liberacao_faturamento),
            # exatamente como o botão "Avançar" do Kanban já fazia (ver
            # estacao_kanban_mover) — os dois sempre andam juntos na prática,
            # então isso não muda em nada o cálculo de FINALIZADO nem os
            # relatórios de lead time/gargalos, que continuam lendo
            # termino_inspecao normalmente.
            conclusoes_producao = f.getlist("item_conclusao_producao[]")

            itens_originais = {item.id: item for item in pedido.itens}
            # snapshot ANTES de qualquer alteração, só dos itens que já existiam
            # (itens novos não têm "antes" pra comparar — são uma inclusão, não uma mudança)
            itens_antes = {iid: {c: getattr(item, c) for c in CAMPOS_HISTORICO_ITEM} for iid, item in itens_originais.items()}
            ids_mantidos = set()

            linhas = zip(
                item_ids, descricoes, quantidades, custos, estacoes, status_itens,
                inicios_producao, conclusoes_producao,
            )
            for (item_id, desc, qtd, custo, estacao_item, status_item,
                 ini_prod, concl_prod) in linhas:
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
                item.inicio_producao = _parse_data_form(ini_prod)
                data_conclusao = _parse_data_form(concl_prod)
                item.termino_inspecao = data_conclusao
                item.liberacao_faturamento = data_conclusao
                item.liberacao_prevista = lib_prevista_pedido
                item.liberacao_real = lib_real_pedido
                item.planejamento_semanal = planejamento_semanal_pedido

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

        # Simulado A (pedido do Bruno, 11/09/2026): prazo comercial x LT de
        # produção PARAMETRIZADO + LT de transporte parametrizado — é pra
        # esta tela que o sistema já redireciona ao salvar um pedido novo,
        # então já mostra a projeção na hora, sem depender do PCP ainda ter
        # planejado nada (Simulado B, na Gestão de Risco, já é a versão com
        # a previsão REAL do PCP). Só considera itens ainda não finalizados.
        simulado_a = None
        itens_abertos = [i for i in pedido.itens if i.status_producao != "FINALIZADO"]
        if itens_abertos:
            lt_producao, itens_sem_parametro = _lt_producao_parametrizado_pedido(itens_abertos)

            transporte_dias = None
            transporte_aplicavel = pedido.frete == "CIF"
            if transporte_aplicavel and pedido.estado:
                uf = pedido.estado.strip().upper()
                linha_transporte = _mapa_lead_time_transportadora().get((uf, "Rodoviário"))
                transporte_dias = _lead_time_transporte_dias(linha_transporte)

            data_prevista = None
            if pedido.data_inclusao_pedido and lt_producao is not None and (transporte_dias is not None or not transporte_aplicavel):
                data_prevista = pedido.data_inclusao_pedido + timedelta(days=lt_producao + (transporte_dias or 0))

            folga = (pedido.data_cliente - data_prevista).days if (pedido.data_cliente and data_prevista) else None

            if folga is None:
                indicador = "sem_dado"
            elif folga < 0 or folga <= RISCO_OTD_LIMITE_RISCO_DIAS:
                indicador = "vermelho"
            elif folga <= RISCO_OTD_LIMITE_ATENCAO_DIAS:
                indicador = "amarelo"
            else:
                indicador = "verde"

            simulado_a = {
                "lt_producao": lt_producao,
                "transporte_dias": transporte_dias,
                "transporte_aplicavel": transporte_aplicavel,
                "data_prevista": data_prevista,
                "folga": folga,
                "indicador": indicador,
                "itens_sem_parametro": itens_sem_parametro,
            }

        return render_template(
            "editar_pedido.html", pedido=pedido, transportadoras=transportadoras, simulado_a=simulado_a,
        )

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
        inspecoes_rdim = _inspecoes_rdim_por_item([i.id for i in pedido.itens])
        resumo_rdim = _resumo_rdim_pedido(inspecoes_rdim.values())
        return render_template(
            "detalhe_pedido.html", pedido=pedido, historico=historico, timeline=timeline,
            inspecoes_rdim=inspecoes_rdim, resumo_rdim=resumo_rdim,
        )

    @app.route("/pedidos/<int:pedido_id>/excluir", methods=["POST"])
    @requer_role("ADMIN", "PCP")
    def excluir_pedido(pedido_id):
        pedido = db.session.get(Pedido, pedido_id)
        if pedido is not None:
            db.session.delete(pedido)
            try:
                db.session.commit()
            except IntegrityError:
                # Defesa extra (bug real corrigido 03/09/2026, pedido do
                # Bruno "não consigo apagar nenhum pedido"): faltavam
                # cascades de exclusão pra Programacao/InspecaoFinal/
                # HistoricoAlteracao, o que fazia qualquer pedido já editado
                # ou inspecionado estourar um 500 sem aviso nenhum. Os
                # relacionamentos já foram corrigidos com cascade="all,
                # delete-orphan" em models.py — isto aqui é só uma rede de
                # segurança pra nunca mais devolver um 500 cru se algum
                # cadastro novo no futuro esquecer o mesmo cuidado.
                db.session.rollback()
                flash(
                    "Não foi possível excluir este pedido: ainda há registros vinculados a ele "
                    "(histórico, inspeção ou programação) que não puderam ser removidos. Avise o suporte.",
                    "danger",
                )
                return redirect(url_for("dashboard"))
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
    # A aba PCP (rota "/gestao-operacao/pcp") foi apagada a pedido do Bruno
    # (16/09/2026): com o grupo Gestão Produção já completo (Planejamento
    # Semanal PCP na Listagem Geral), a listagem/quadrantes que existiam
    # aqui ficaram redundantes — Status produção/Previsão/Data efetiva/
    # Solicitada cliente/Término semanal já aparecem, ao vivo, na Operação
    # 360 (ver _metricas_operacao_360); "Custo produção real" ganhou coluna
    # própria lá (única coisa que só existia nesta aba). O FORMULÁRIO de
    # edição da seção PCP por pedido continua existindo normalmente em
    # gestao_operacao_editar (fallback manual pra pedido ainda não lançado
    # em Produção).

    @app.route("/gestao-operacao/logistica")
    @login_required
    def gestao_operacao_logistica():
        pedidos, page, total_paginas, total_filtrado, filtros, _query_operacao = _linhas_gestao_operacao(request.args)
        # Kanban Expedição (pedido do Bruno, 16/09/2026) — ver
        # _pedidos_kanban_expedicao. Independente dos filtros/paginação da
        # tabela abaixo: mostra SEMPRE todos os pedidos parados na
        # expedição, no site inteiro.
        kanban_expedicao = _pedidos_kanban_expedicao()
        return render_template(
            "gestao_operacao_logistica.html",
            pedidos=pedidos, page=page, total_paginas=total_paginas,
            total_filtrado=total_filtrado, filtros=filtros,
            kanban_expedicao=kanban_expedicao,
        )

    @app.route("/gestao-operacao/<int:pedido_id>/marcar-expedido", methods=["POST"])
    @requer_role("ADMIN", "PCP")
    def gestao_operacao_marcar_expedido(pedido_id):
        """Ação rápida do Kanban Expedição (pedido do Bruno, 16/09/2026): marca
        a Data de expedição de hoje sem precisar abrir o formulário de edição
        completo — o pedido sai do kanban sozinho na próxima carga (ver
        _pedidos_kanban_expedicao, que já filtra por go_data_pedido_expedido
        vazio). Mesmo padrão de auditoria (_registrar_alteracoes) e de
        permissão (@requer_role) já usados em gestao_operacao_editar."""
        pedido = db.session.get(PedidoOperacao, pedido_id)
        if pedido is None:
            flash("Pedido não encontrado.", "danger")
            return redirect(url_for("gestao_operacao_logistica"))

        antes = {"go_data_pedido_expedido": pedido.go_data_pedido_expedido}
        pedido.go_data_pedido_expedido = date.today()
        depois = {"go_data_pedido_expedido": pedido.go_data_pedido_expedido}
        _registrar_alteracoes("pedido_operacao", pedido.id, None, antes, depois, ["go_data_pedido_expedido"])

        db.session.commit()
        flash(f"Pedido {pedido.pedido_venda or pedido.id} marcado como expedido.", "success")
        return redirect(request.referrer or url_for("gestao_operacao_logistica"))

    @app.route("/gestao-operacao/resultados")
    @login_required
    def gestao_operacao_resultados():
        # Período (pedido do Bruno, 28/08/2026: mês selecionável; ampliado
        # 03/09/2026 — "inclua em formato de lista... além do mês, inclua
        # também trimestre, semestre... e o filtro geral" — ver
        # _parse_periodo/_opcoes_periodo). Sem `periodo` na URL, cai no mês
        # atual (mesmo default de sempre). Calculado ANTES de
        # _linhas_gestao_operacao pra poder usar `periodo_str` (já com o
        # default aplicado) no campo oculto do formulário de segmento (ver
        # abaixo) — assim o filtro de segmento sempre navega com um período
        # explícito, mesmo na primeira visita à página.
        tipo_periodo, ano_periodo, valor_periodo, periodo_label = _parse_periodo(request.args.get("periodo", ""))
        periodo_str = _periodo_para_str(tipo_periodo, ano_periodo, valor_periodo)
        faturamento_semanal = _faturamento_por_periodo(tipo_periodo, ano_periodo, valor_periodo)
        periodo_anterior = _periodo_vizinho(tipo_periodo, ano_periodo, valor_periodo, -1)
        periodo_seguinte = _periodo_vizinho(tipo_periodo, ano_periodo, valor_periodo, 1)

        pedidos, page, total_paginas, total_filtrado, filtros, query_operacao = _linhas_gestao_operacao(request.args)
        # Pedido do Bruno (03/09/2026): "quero ver todos os pedidos de julho
        # faturados ou dentro do planejamento semanal do pcp, com isso ver o
        # otd" — filtros["segmento"] (junto com o período) já veio aplicado
        # em query_operacao (ver _filtrar_pedidos_operacao); o resumo de OTD
        # abaixo usa o MESMO recorte, em vez de sempre olhar pra todos os
        # pedidos.
        otd = _resumo_otd(query_operacao)

        # Resumo fixo do período (pedido do Bruno, 03/09/2026: "preciso ver
        # os resultados detalhados de cada mês, como otd, lead time
        # operação, lead time chão de fábrica (produção), lead time operação
        # cif, lead time operação fob... de preferência no topo da página,
        # do lado do faturamento") — SEMPRE o período selecionado em cima,
        # no mesmo recorte "planejamento" (Término Semanal PCP no período)
        # usado em "Qtd/Valor liberado", independente do dropdown de
        # segmento mais abaixo (que continua controlando só os cards de OTD
        # "geral" e a lista de pedidos).
        query_periodo = _pedidos_operacao_do_periodo(tipo_periodo, ano_periodo, valor_periodo)
        otd_mes = _resumo_otd(query_periodo)
        lead_times_mes = _resumo_lead_times(query_periodo)

        return render_template(
            "gestao_operacao_resultados.html",
            pedidos=pedidos, page=page, total_paginas=total_paginas,
            total_filtrado=total_filtrado, filtros=filtros, otd=otd,
            faturamento_semanal=faturamento_semanal,
            otd_mes=otd_mes, lead_times_mes=lead_times_mes,
            periodo=periodo_str, tipo_periodo=tipo_periodo, mes_label=periodo_label,
            periodo_anterior=periodo_anterior, periodo_seguinte=periodo_seguinte,
            opcoes_periodo=_opcoes_periodo(),
        )

    @app.route("/gestao-operacao/listagem-geral")
    @login_required
    def gestao_operacao_listagem_geral():
        """"Operação 360" (pedido do Bruno, 10/09/2026 — antes "Listagem
        Geral"; nome da aba mudou, endpoint/URL continuam os mesmos pra não
        quebrar link nenhum) — 1 linha por PEDIDO (não por produto, diferente
        da Listagem Geral de Gestão Produção), com as principais informações
        do pedido de ponta a ponta: inclusão -> solicitação -> produção ->
        expedição -> entrega, os 3 lead times (comercial/produção/operação),
        OTD e qualidade. Passar o mouse (ou clicar, no touch) sobre "Itens"
        mostra os produtos/quantidades já preenchidos em Gestão Produção pelo
        PCP, casando pelo nº de pedido de venda — sem criar nenhum vínculo
        real entre as duas tabelas."""
        pedidos, page, total_paginas, total_filtrado, filtros, query_operacao = _linhas_gestao_operacao(request.args)
        itens_por_pedido_venda = _itens_producao_por_pedido_venda([p.pedido_venda for p in pedidos])
        # Qualidade (RDIM) — pedido do Bruno (02/09/2026): indicador simples
        # (pedido-level, sem granularidade de item/estação) de "contém
        # desvio" nesta listagem.
        rdim_por_pedido_venda = _rdim_resumo_por_pedido_venda([p.pedido_venda for p in pedidos])
        # "Data solic. cliente" também acompanha ao vivo a "Data do cliente"
        # de Gestão Produção — pedido do Bruno (03/09/2026).
        data_cliente_por_pedido_venda = _data_cliente_por_pedido_venda([p.pedido_venda for p in pedidos])
        # "Data conclusão produção"/"Semanal planejamento PCP" também
        # acompanham ao vivo a Liberação real/Planejamento semanal de Gestão
        # Produção — pedido do Bruno (10/09/2026, "Operação 360").
        liberacao_pcp_por_pedido_venda = _liberacao_pcp_por_pedido_venda([p.pedido_venda for p in pedidos])
        # "Status pedido" (coluna com emoji, pedido do Bruno 10/09/2026)
        # reaproveita _indice_etapa_pedido, que precisa do Pedido inteiro
        # (não só os itens já resumidos acima) — ver
        # _pedidos_producao_por_pedido_venda.
        pedidos_producao_por_pedido_venda = _pedidos_producao_por_pedido_venda([p.pedido_venda for p in pedidos])
        # Datas/lead times/OTD/status da "Operação 360" — ver _metricas_operacao_360.
        # (calculado só pra página atual — o que aparece na tabela).
        metricas_operacao_360 = _metricas_operacao_360(
            pedidos, liberacao_pcp_por_pedido_venda, data_cliente_por_pedido_venda, pedidos_producao_por_pedido_venda,
        )

        # Quadrante "NFs emitidas no mês" + painel dinâmico de valores/
        # faturamento/lead time (pedido do Bruno, 10/09/2026) — no lugar dos
        # quadrantes semanais/mensais de PCP, que saíram desta tela e
        # continuam só na Listagem Geral de Produção, como já estabelecido.
        # Precisa do conjunto TOTAL filtrado (não só a página atual), por
        # isso recalcula liberação PCP/data cliente/Pedido de Produção pra
        # TODOS os pedidos filtrados — ver _painel_operacao_360.
        pedidos_filtrados_completo = query_operacao.all()
        pedidos_venda_completo = [p.pedido_venda for p in pedidos_filtrados_completo]
        liberacao_pcp_completo = _liberacao_pcp_por_pedido_venda(pedidos_venda_completo)
        data_cliente_completo = _data_cliente_por_pedido_venda(pedidos_venda_completo)
        pedidos_producao_completo = _pedidos_producao_por_pedido_venda(pedidos_venda_completo)
        metricas_completo = _metricas_operacao_360(
            pedidos_filtrados_completo, liberacao_pcp_completo, data_cliente_completo, pedidos_producao_completo,
        )
        painel_operacao_360 = _painel_operacao_360(filtros, pedidos_filtrados_completo, metricas_completo)

        # Opções do filtro "Status produção" (lista multi-seleção, pedido do
        # Bruno 10/09/2026) — reaproveita _ETAPAS_ACOMPANHAMENTO_PEDIDO/
        # _ETAPA_EMOJI (privados a este módulo, por isso montados aqui em vez
        # de expostos via inject_globals) pra nunca divergir dos rótulos já
        # usados na coluna "Status pedido" e em Consulta Pedido.
        status_pedido_opcoes = [
            {"valor": str(idx), "label": etapa["label"], "emoji": _ETAPA_EMOJI[idx - 1]}
            for idx, etapa in enumerate(_ETAPAS_ACOMPANHAMENTO_PEDIDO, start=1)
        ]

        return render_template(
            "gestao_operacao_listagem_geral.html",
            pedidos=pedidos, page=page, total_paginas=total_paginas,
            total_filtrado=total_filtrado, filtros=filtros,
            itens_por_pedido_venda=itens_por_pedido_venda,
            rdim_por_pedido_venda=rdim_por_pedido_venda,
            data_cliente_por_pedido_venda=data_cliente_por_pedido_venda,
            metricas_operacao_360=metricas_operacao_360,
            painel_operacao_360=painel_operacao_360,
            status_pedido_opcoes=status_pedido_opcoes,
        )

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
            if secao == "pcp":
                # Pedido do Bruno (01/09/2026, ampliado 03/09/2026): quando já
                # existe o dado automático vindo de Gestão Produção pra
                # "Previsão liberação PCP"/"Data efetiva liberação"/
                # "Solicitada cliente/retira"/"Término semanal", essas colunas
                # viram só leitura aqui — não sobrescreve (nem zera) o campo
                # antigo de PedidoOperacao ao salvar o resto da aba PCP, senão
                # o valor manual histórico se perderia à toa mesmo sem o
                # usuário ter mexido nele (o formulário nem mostra mais um
                # <input> pra esses campos nesse caso — ver
                # gestao_operacao_editar.html).
                chave_pv = _normalizar_pedido_venda(pedido.pedido_venda)
                liberacao_real_pcp = _liberacao_pcp_por_pedido_venda([pedido.pedido_venda]).get(chave_pv) or {}
                data_cliente_real = _data_cliente_por_pedido_venda([pedido.pedido_venda]).get(chave_pv)
                campos_secao = [
                    c for c in campos_secao
                    if not (c == "go_previsao_liberacao_pcp" and liberacao_real_pcp.get("previsao"))
                    and not (c == "go_data_efetiva_liberacao_pcp" and liberacao_real_pcp.get("efetiva"))
                    and not (c == "go_data_solicitada_cliente_retira" and data_cliente_real)
                    and not (c == "go_termino_semanal_pcp" and liberacao_real_pcp.get("termino_semanal"))
                ]
            elif secao == "comercial":
                # Mesmo espírito acima, pro campo equivalente da aba Comercial
                # (pedido do Bruno, 03/09/2026: "quero que todos os dados
                # dentro da gestão operação seja extraída automaticamente da
                # gestão produção").
                chave_pv = _normalizar_pedido_venda(pedido.pedido_venda)
                data_cliente_real = _data_cliente_por_pedido_venda([pedido.pedido_venda]).get(chave_pv)
                campos_secao = [
                    c for c in campos_secao
                    if not (c == "go_data_solicitada_entrega" and data_cliente_real)
                ]
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

        chave_pv = _normalizar_pedido_venda(pedido.pedido_venda)
        status_real = _status_producao_por_pedido_venda([pedido.pedido_venda]).get(chave_pv)
        liberacao_real = _liberacao_pcp_por_pedido_venda([pedido.pedido_venda]).get(chave_pv) or {}
        data_cliente_real = _data_cliente_por_pedido_venda([pedido.pedido_venda]).get(chave_pv)

        return render_template(
            "gestao_operacao_editar.html", pedido=pedido, transportadoras=transportadoras,
            secao=secao, GO_SECOES=GO_SECOES, GO_SECAO_ENDPOINT=GO_SECAO_ENDPOINT, GO_SECAO_LABEL=GO_SECAO_LABEL,
            status_real=status_real, liberacao_real=liberacao_real, data_cliente_real=data_cliente_real,
        )

    # ------------------------------------------------------------------
    # Gestão de Risco / Torre de Controle de OTD (pedido do Bruno,
    # 11/09/2026) — ver cabeçalho de _pedidos_risco_otd (app.py) pro
    # desenho completo. Só pedidos CIF, sempre calculado ao vivo.
    # ------------------------------------------------------------------
    @app.route("/gestao-risco")
    @login_required
    def gestao_risco():
        linhas, resumo = _pedidos_risco_otd(request.args)
        filtros = {
            "busca": (request.args.get("busca", "") or "").strip(),
            "status": [v for v in _getlist_seguro(request.args, "status") if v in RISCO_OTD_STATUS_INFO],
        }
        # Aba Simulação (pedido do Bruno, 11/09/2026): Simulado A (parâmetro)
        # x Simulado B (realidade, = a própria linha acima) — mesmo conjunto
        # já filtrado, sem duplicar a consulta de pedidos.
        simulacoes, resumo_simulacao = _simulacao_otd(linhas)
        return render_template(
            "gestao_risco.html",
            linhas=linhas, resumo=resumo, filtros=filtros,
            RISCO_OTD_STATUS_INFO=RISCO_OTD_STATUS_INFO,
            simulacoes=simulacoes, resumo_simulacao=resumo_simulacao,
            SIMULACAO_BALANCO_INFO=SIMULACAO_BALANCO_INFO,
        )

    @app.route("/gestao-risco/relatorio.xlsx")
    @login_required
    def gestao_risco_xlsx():
        linhas, resumo = _pedidos_risco_otd(request.args)
        filtros = {
            "busca": (request.args.get("busca", "") or "").strip(),
            "status": [v for v in _getlist_seguro(request.args, "status") if v in RISCO_OTD_STATUS_INFO],
        }
        return _gerar_excel_risco_otd(linhas, resumo, filtros)

    @app.route("/gestao-risco/relatorio.pdf")
    @login_required
    def gestao_risco_pdf():
        linhas, resumo = _pedidos_risco_otd(request.args)
        filtros = {
            "busca": (request.args.get("busca", "") or "").strip(),
            "status": [v for v in _getlist_seguro(request.args, "status") if v in RISCO_OTD_STATUS_INFO],
        }
        return _gerar_pdf_risco_otd(linhas, resumo, filtros)

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
    # Qualidade — Inspeção Final / RDIM (pedido do Bruno, 02/09/2026).
    # ------------------------------------------------------------------
    @app.route("/qualidade/rdim/buscar-itens")
    @login_required
    def rdim_buscar_itens():
        """Sugestões (JSON) pro campo de busca de item/OP da tela de Nova
        Inspeção — só itens das estações MANDRIL/PU/SILICONE."""
        termo = request.args.get("q", "")
        return jsonify(_itens_rdim_disponiveis(termo))

    @app.route("/qualidade/rdim/dashboard")
    @login_required
    def rdim_dashboard():
        dados = _dashboard_rdim()
        return render_template("qualidade_rdim_dashboard.html", dados=dados)

    @app.route("/qualidade/rdim")
    @login_required
    def rdim_lista():
        inspecoes, page, total_paginas, total_filtrado, filtros = _linhas_inspecoes_finais(request.args)
        responsaveis = (
            Usuario.query.join(InspecaoFinal, InspecaoFinal.responsavel_id == Usuario.id)
            .distinct().order_by(Usuario.nome).all()
        )
        return render_template(
            "qualidade_rdim_lista.html",
            inspecoes=inspecoes, page=page, total_paginas=total_paginas,
            total_filtrado=total_filtrado, filtros=filtros, responsaveis=responsaveis,
        )

    @app.route("/qualidade/rdim/novo", methods=["GET", "POST"])
    @login_required
    def rdim_novo():
        item_id_inicial = request.args.get("item_id", "").strip()

        if request.method == "POST":
            f = request.form
            item_id = f.get("item_pedido_id", "").strip()
            resultado = f.get("resultado", "").strip()
            tipo_produto = f.get("tipo_produto_inspecionado", "").strip()
            item = db.session.get(ItemPedido, int(item_id)) if item_id.isdigit() else None

            if item is None:
                flash("Selecione uma OP (item de pedido) válida antes de salvar.", "danger")
                return render_template("qualidade_rdim_novo.html", valores=f, item_selecionado=None)
            if resultado not in RDIM_RESULTADO_OPCOES:
                flash("Selecione o resultado da inspeção (Aprovado / Reprovado / Aprovado com desvio).", "danger")
                return render_template("qualidade_rdim_novo.html", valores=f, item_selecionado=item)
            if tipo_produto not in RDIM_TIPO_PRODUTO_OPCOES:
                flash("Selecione o tipo de produto a ser inspecionado.", "danger")
                return render_template("qualidade_rdim_novo.html", valores=f, item_selecionado=item)

            quantidade_com_desvio, erro_qtd = _validar_quantidade_com_desvio(f.get("quantidade_com_desvio"), item.quantidade)
            if erro_qtd:
                flash(erro_qtd, "danger")
                return render_template("qualidade_rdim_novo.html", valores=f, item_selecionado=item)

            nova = InspecaoFinal(
                item_pedido_id=item.id,
                estacao=item.estacao,
                data_inspecao=_parse_data_form(f.get("data_inspecao")) or date.today(),
                responsavel_id=current_user.id,
                numero_rif=f.get("numero_rif", "").strip() or None,
                procedimento=f.get("procedimento", "").strip() or None,
                norma=f.get("norma", "").strip() or None,
                instrucao_trabalho=f.get("instrucao_trabalho", "").strip() or None,
                inspecao_visual=f.get("inspecao_visual", "").strip() or None,
                desvio_encontrado=f.get("desvio_encontrado", "").strip() or None,
                categoria_desvio=f.get("categoria_desvio", "").strip() or None,
                subcategoria_desvio=f.get("subcategoria_desvio", "").strip() or None,
                observacao=f.get("observacao", "").strip() or None,
                resultado=resultado,
                quantidade_com_desvio=quantidade_com_desvio,
                tipo_produto_inspecionado=tipo_produto,
                criado_por_id=current_user.id,
            )
            db.session.add(nova)
            _salvar_medicoes_rdim(nova, f)
            _salvar_pecas_desvio_rdim(nova, f)
            erro_componentes = _salvar_componentes_desvio_rdim(nova, f, item.quantidade)
            if erro_componentes:
                db.session.rollback()
                flash(erro_componentes, "danger")
                return render_template("qualidade_rdim_novo.html", valores=f, item_selecionado=item)
            db.session.commit()
            flash("Inspeção final registrada com sucesso.", "success")
            return redirect(url_for("rdim_editar", inspecao_id=nova.id))

        item_selecionado = None
        if item_id_inicial.isdigit():
            item_selecionado = db.session.get(ItemPedido, int(item_id_inicial))
        return render_template("qualidade_rdim_novo.html", valores={}, item_selecionado=item_selecionado)

    @app.route("/qualidade/rdim/<int:inspecao_id>/editar", methods=["GET", "POST"])
    @login_required
    def rdim_editar(inspecao_id):
        inspecao = db.session.get(InspecaoFinal, inspecao_id)
        if inspecao is None:
            flash("Inspeção não encontrada.", "danger")
            return redirect(url_for("rdim_lista"))

        if request.method == "POST":
            f = request.form
            resultado = f.get("resultado", "").strip()
            tipo_produto = f.get("tipo_produto_inspecionado", "").strip()
            if resultado not in RDIM_RESULTADO_OPCOES:
                flash("Selecione o resultado da inspeção (Aprovado / Reprovado / Aprovado com desvio).", "danger")
                historico = (
                    HistoricoAlteracao.query.filter_by(entidade_tipo="inspecao_final", entidade_id=inspecao.id)
                    .order_by(HistoricoAlteracao.criado_em.desc()).all()
                )
                return render_template("qualidade_rdim_editar.html", inspecao=inspecao, historico=historico)
            if tipo_produto not in RDIM_TIPO_PRODUTO_OPCOES:
                flash("Selecione o tipo de produto a ser inspecionado.", "danger")
                historico = (
                    HistoricoAlteracao.query.filter_by(entidade_tipo="inspecao_final", entidade_id=inspecao.id)
                    .order_by(HistoricoAlteracao.criado_em.desc()).all()
                )
                return render_template("qualidade_rdim_editar.html", inspecao=inspecao, historico=historico)

            quantidade_item = inspecao.item.quantidade if inspecao.item else None
            quantidade_com_desvio, erro_qtd = _validar_quantidade_com_desvio(f.get("quantidade_com_desvio"), quantidade_item)
            if erro_qtd:
                flash(erro_qtd, "danger")
                historico = (
                    HistoricoAlteracao.query.filter_by(entidade_tipo="inspecao_final", entidade_id=inspecao.id)
                    .order_by(HistoricoAlteracao.criado_em.desc()).all()
                )
                return render_template("qualidade_rdim_editar.html", inspecao=inspecao, historico=historico)

            erro_componentes = _salvar_componentes_desvio_rdim(inspecao, f, quantidade_item, substituir=True)
            if erro_componentes:
                flash(erro_componentes, "danger")
                historico = (
                    HistoricoAlteracao.query.filter_by(entidade_tipo="inspecao_final", entidade_id=inspecao.id)
                    .order_by(HistoricoAlteracao.criado_em.desc()).all()
                )
                return render_template("qualidade_rdim_editar.html", inspecao=inspecao, historico=historico)

            antes = {c: getattr(inspecao, c) for c in CAMPOS_HISTORICO_INSPECAO_FINAL}
            inspecao.data_inspecao = _parse_data_form(f.get("data_inspecao")) or inspecao.data_inspecao
            inspecao.numero_rif = f.get("numero_rif", "").strip() or None
            inspecao.procedimento = f.get("procedimento", "").strip() or None
            inspecao.norma = f.get("norma", "").strip() or None
            inspecao.instrucao_trabalho = f.get("instrucao_trabalho", "").strip() or None
            inspecao.inspecao_visual = f.get("inspecao_visual", "").strip() or None
            inspecao.desvio_encontrado = f.get("desvio_encontrado", "").strip() or None
            inspecao.categoria_desvio = f.get("categoria_desvio", "").strip() or None
            inspecao.subcategoria_desvio = f.get("subcategoria_desvio", "").strip() or None
            inspecao.observacao = f.get("observacao", "").strip() or None
            inspecao.resultado = resultado
            inspecao.quantidade_com_desvio = quantidade_com_desvio
            inspecao.tipo_produto_inspecionado = tipo_produto
            depois = {c: getattr(inspecao, c) for c in CAMPOS_HISTORICO_INSPECAO_FINAL}
            _registrar_alteracoes("inspecao_final", inspecao.id, inspecao.item.pedido_id if inspecao.item else None,
                                   antes, depois, CAMPOS_HISTORICO_INSPECAO_FINAL)

            _salvar_medicoes_rdim(inspecao, f, substituir=True)
            _salvar_pecas_desvio_rdim(inspecao, f, substituir=True)
            db.session.commit()
            flash("Inspeção atualizada com sucesso.", "success")
            return redirect(url_for("rdim_editar", inspecao_id=inspecao.id))

        historico = (
            HistoricoAlteracao.query.filter_by(entidade_tipo="inspecao_final", entidade_id=inspecao.id)
            .order_by(HistoricoAlteracao.criado_em.desc()).all()
        )
        return render_template("qualidade_rdim_editar.html", inspecao=inspecao, historico=historico)

    @app.route("/qualidade/rdim/<int:inspecao_id>/excluir", methods=["POST"])
    @requer_role("ADMIN", "PCP")
    def rdim_excluir(inspecao_id):
        """Apagar uma inspeção RDIM definitivamente — pedido do Bruno
        (02/09/2026): só tinha editar, faltava excluir. Restrito a ADMIN/PCP,
        mesmo critério já usado em excluir_pedido — é um registro formal de
        inspeção (RIF/norma/procedimento), então a exclusão fica mais
        controlada que criar/editar (aberto a todo usuário autenticado).
        RdimMedicao e RdimPecaDesvio somem juntos (cascade="all,
        delete-orphan" já configurado no relacionamento)."""
        inspecao = db.session.get(InspecaoFinal, inspecao_id)
        if inspecao is not None:
            db.session.delete(inspecao)
            db.session.commit()
            flash("Inspeção RDIM excluída.", "info")
        return redirect(url_for("rdim_lista"))

    # ------------------------------------------------------------------
    # P&D — Pesquisa e Desenvolvimento (Fase 14, 01/09/2026). Área nova,
    # independente de PCP/Produção/Operação/Qualidade. Mesmo critério de
    # acesso já adotado em Qualidade (aberto a todo usuário autenticado,
    # sem role dedicado) — usada principalmente por Bruno e Gustavo Fugita,
    # mas sem restringir os outros papéis já existentes no sistema.
    # ------------------------------------------------------------------
    @app.route("/pd/dashboard")
    @login_required
    def pd_dashboard():
        dados = _dashboard_pd()
        return render_template("pd_dashboard.html", dados=dados)

    @app.route("/pd")
    @login_required
    def pd_lista():
        projetos, page, total_paginas, total_filtrado, filtros = _linhas_projetos_pd(request.args)
        opcoes_filtro = dict(
            etapa=PD_ETAPA_OPCOES,
            categoria=_pd_opcoes_filtro("categoria", PD_CATEGORIA_OPCOES),
            prioridade=PRIORIDADE_OPCOES,
        )
        return render_template(
            "pd_lista.html",
            projetos=projetos, page=page, total_paginas=total_paginas,
            total_filtrado=total_filtrado, filtros=filtros, opcoes_filtro=opcoes_filtro,
        )

    @app.route("/pd/novo", methods=["GET", "POST"])
    @login_required
    def pd_novo():
        if request.method == "POST":
            f = request.form
            nome = f.get("nome", "").strip()
            if not nome:
                flash("Nome do projeto é obrigatório.", "danger")
                valores_repopular = {k: f.get(k, "") for k in f.keys()}
                valores_repopular["resultado_esperado"] = f.getlist("resultado_esperado")
                return render_template("pd_novo.html", valores=valores_repopular)

            valores = _campos_form_pd(f)
            novo = ProjetoPD(criado_por_id=current_user.id, **valores)
            db.session.add(novo)
            db.session.commit()
            flash(f"Projeto {novo.codigo or ('#' + str(novo.id))} cadastrado com sucesso.", "success")
            return redirect(url_for("pd_editar", projeto_id=novo.id))

        return render_template("pd_novo.html", valores={})

    @app.route("/pd/<int:projeto_id>/editar", methods=["GET", "POST"])
    @login_required
    def pd_editar(projeto_id):
        projeto = db.session.get(ProjetoPD, projeto_id)
        if projeto is None:
            flash("Projeto de P&D não encontrado.", "danger")
            return redirect(url_for("pd_lista"))

        if request.method == "POST":
            f = request.form
            nome = f.get("nome", "").strip()
            if not nome:
                flash("Nome do projeto é obrigatório.", "danger")
                return redirect(url_for("pd_editar", projeto_id=projeto.id))

            antes = {c: getattr(projeto, c) for c in CAMPOS_HISTORICO_PD}
            valores = _campos_form_pd(f)
            for campo, valor in valores.items():
                setattr(projeto, campo, valor)
            depois = {c: getattr(projeto, c) for c in CAMPOS_HISTORICO_PD}
            _registrar_alteracoes("projeto_pd", projeto.id, None, antes, depois, CAMPOS_HISTORICO_PD)

            db.session.commit()
            flash("Projeto atualizado com sucesso.", "success")
            return redirect(url_for("pd_editar", projeto_id=projeto.id))

        historico = (
            HistoricoAlteracao.query
            .filter_by(entidade_tipo="projeto_pd", entidade_id=projeto.id)
            .order_by(HistoricoAlteracao.criado_em.desc())
            .all()
        )
        return render_template(
            "pd_editar.html", projeto=projeto, valores=_pd_para_form_dict(projeto), historico=historico,
        )

    @app.route("/pd/<int:projeto_id>/testes/novo", methods=["POST"])
    @login_required
    def pd_teste_novo(projeto_id):
        projeto = db.session.get(ProjetoPD, projeto_id)
        if projeto is None:
            flash("Projeto de P&D não encontrado.", "danger")
            return redirect(url_for("pd_lista"))

        f = request.form
        teste = TesteProjetoPD(
            projeto_id=projeto.id,
            numero=f.get("numero", "").strip() or None,
            data_planejada=_parse_data_form(f.get("data_planejada")),
            data_realizada=_parse_data_form(f.get("data_realizada")),
            responsavel=f.get("responsavel", "").strip() or None,
            material_utilizado=f.get("material_utilizado", "").strip() or None,
            lote=f.get("lote", "").strip() or None,
            fornecedor=f.get("fornecedor", "").strip() or None,
            condicoes=f.get("condicoes", "").strip() or None,
            resultado=f.get("resultado", "").strip() or "Planejado",
            observacoes=f.get("observacoes", "").strip() or None,
            anexos=f.get("anexos", "").strip() or None,
        )
        db.session.add(teste)
        db.session.commit()
        flash("Teste registrado com sucesso.", "success")
        return redirect(url_for("pd_editar", projeto_id=projeto.id) + "#testes")

    @app.route("/pd/<int:projeto_id>/eventos/novo", methods=["POST"])
    @login_required
    def pd_evento_novo(projeto_id):
        projeto = db.session.get(ProjetoPD, projeto_id)
        if projeto is None:
            flash("Projeto de P&D não encontrado.", "danger")
            return redirect(url_for("pd_lista"))

        f = request.form
        evento = VisitaReuniaoPD(
            projeto_id=projeto.id,
            data=_parse_data_form(f.get("data")),
            tipo=f.get("tipo", "").strip() or None,
            participantes=f.get("participantes", "").strip() or None,
            local=f.get("local", "").strip() or None,
            objetivo=f.get("objetivo", "").strip() or None,
            resultado=f.get("resultado", "").strip() or None,
            proximas_acoes=f.get("proximas_acoes", "").strip() or None,
            responsavel=f.get("responsavel", "").strip() or None,
            anexos=f.get("anexos", "").strip() or None,
        )
        db.session.add(evento)
        db.session.commit()
        flash("Visita/reunião registrada com sucesso.", "success")
        return redirect(url_for("pd_editar", projeto_id=projeto.id) + "#eventos")

    # Kanban de P&D — pedido do Bruno: "visão alternada" dentro de Projetos,
    # colunas = as 8 etapas do ciclo de vida, arrastável. Mover uma coluna
    # pra outra usa a mesma rota que o <select> de fallback (sem JS/toque) —
    # os dois caminhos (drag-and-drop e o <select>) chamam pd_mover_etapa.
    @app.route("/pd/kanban")
    @login_required
    def pd_kanban():
        projetos = ProjetoPD.query.order_by(ProjetoPD.data_prevista_conclusao.asc().nullslast(), ProjetoPD.id.desc()).all()
        colunas = {etapa: [] for etapa in PD_ETAPA_OPCOES}
        for p in projetos:
            colunas.setdefault(p.etapa_atual, []).append(p)
        return render_template("pd_kanban.html", colunas=colunas)

    @app.route("/pd/<int:projeto_id>/mover-etapa", methods=["POST"])
    @login_required
    def pd_mover_etapa(projeto_id):
        projeto = db.session.get(ProjetoPD, projeto_id)
        if projeto is None:
            flash("Projeto de P&D não encontrado.", "danger")
            return redirect(url_for("pd_kanban"))

        nova_etapa = request.form.get("etapa", "").strip()
        if nova_etapa not in PD_ETAPA_OPCOES:
            flash("Etapa inválida.", "danger")
            return redirect(url_for("pd_kanban"))

        if nova_etapa != projeto.etapa_atual:
            antes = {c: getattr(projeto, c) for c in CAMPOS_HISTORICO_PD}
            projeto.etapa_atual = nova_etapa
            # Conveniência: ao mover pra "Concluído" sem data real de
            # conclusão ainda preenchida, carimba hoje sozinho (o Bruno pode
            # sempre corrigir depois na aba "Dados do projeto").
            if nova_etapa == "Concluído" and not projeto.data_real_conclusao:
                projeto.data_real_conclusao = date.today()
            depois = {c: getattr(projeto, c) for c in CAMPOS_HISTORICO_PD}
            _registrar_alteracoes("projeto_pd", projeto.id, None, antes, depois, CAMPOS_HISTORICO_PD)
            db.session.commit()
            flash(f"\"{projeto.nome}\" movido para {nova_etapa}.", "success")

        return redirect(url_for("pd_kanban"))

    @app.route("/pd/cronograma")
    @login_required
    def pd_cronograma():
        dados = _cronograma_pd()
        return render_template("pd_cronograma.html", **dados)

    @app.route("/pd/testes")
    @login_required
    def pd_testes_lista():
        query, filtros = _filtrar_testes_pd(request.args)
        testes = query.all()
        projetos_opcoes = ProjetoPD.query.order_by(ProjetoPD.nome).all()
        return render_template("pd_testes.html", testes=testes, projetos_opcoes=projetos_opcoes, filtros=filtros)

    @app.route("/pd/custos")
    @login_required
    def pd_custos():
        dados = _custos_pd()
        return render_template("pd_custos.html", **dados)

    @app.route("/pd/conhecimento")
    @login_required
    def pd_conhecimento():
        busca = request.args.get("busca", "").strip()
        query = ProjetoPD.query.filter(
            or_(ProjetoPD.problema.isnot(None), ProjetoPD.solucao.isnot(None), ProjetoPD.licoes_aprendidas.isnot(None))
        )
        if busca:
            termo = f"%{busca}%"
            query = query.filter(
                or_(
                    ProjetoPD.nome.ilike(termo), ProjetoPD.problema.ilike(termo), ProjetoPD.solucao.ilike(termo),
                    ProjetoPD.licoes_aprendidas.ilike(termo), ProjetoPD.categoria.ilike(termo),
                    ProjetoPD.cliente.ilike(termo), ProjetoPD.produto.ilike(termo), ProjetoPD.fornecedor.ilike(termo),
                )
            )
        projetos = query.order_by(ProjetoPD.data_real_conclusao.desc().nullslast(), ProjetoPD.id.desc()).all()
        return render_template("pd_conhecimento.html", projetos=projetos, busca=busca)

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
    # Zerar dados de teste (pedido do Bruno, 28/08/2026; ampliado pra
    # Qualidade + P&D em 03/09/2026 — "quero que exclua não só produção e
    # operação, mas também todos os dados de testes (P&D e Qualidade)") —
    # apaga TODOS os Pedido/ItemPedido (Gestão Produção), PedidoOperacao
    # (Gestão Operação), RncQualidade + InspecaoFinal/RdimMedicao/
    # RdimPecaDesvio (Qualidade) e ProjetoPD/TesteProjetoPD/VisitaReuniaoPD
    # (P&D) de uma vez, pra ele testar manualmente do zero. Só ADMIN, com
    # confirmação por texto digitado — e sempre com um backup (.xlsx)
    # disponível antes, já que é uma exclusão permanente. Cadastros
    # (usuários, estações, transportadoras) nunca são tocados.
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
            total_rnc=RncQualidade.query.count(),
            total_inspecoes_rdim=InspecaoFinal.query.count(),
            total_projetos_pd=ProjetoPD.query.count(),
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
        total_rnc = RncQualidade.query.count()
        total_inspecoes_rdim = InspecaoFinal.query.count()
        total_projetos_pd = ProjetoPD.query.count()

        # Ordem importa: filhos antes dos pais, por causa das foreign keys —
        # bulk delete (.query.delete()) não aciona cascade de ORM, só as
        # normais do banco, então cada FK precisa ser removida "na mão" na
        # ordem certa (RdimPecaDesvio/RdimMedicao/RdimComponenteDesvio ->
        # InspecaoFinal -> ItemPedido; Programacao -> ItemPedido;
        # HistoricoAlteracao -> Pedido; TesteProjetoPD/VisitaReuniaoPD ->
        # ProjetoPD). RncQualidade e ProjetoPD/PedidoOperacao são tabelas
        # independentes (sem FK com o resto), podem vir em qualquer ordem.
        RdimPecaDesvio.query.delete(synchronize_session=False)
        RdimMedicao.query.delete(synchronize_session=False)
        RdimComponenteDesvio.query.delete(synchronize_session=False)
        InspecaoFinal.query.delete(synchronize_session=False)
        Programacao.query.delete(synchronize_session=False)
        HistoricoAlteracao.query.delete(synchronize_session=False)
        ItemPedido.query.delete(synchronize_session=False)
        Pedido.query.delete(synchronize_session=False)
        PedidoOperacao.query.delete(synchronize_session=False)
        RncQualidade.query.delete(synchronize_session=False)
        TesteProjetoPD.query.delete(synchronize_session=False)
        VisitaReuniaoPD.query.delete(synchronize_session=False)
        ProjetoPD.query.delete(synchronize_session=False)

        # Marca que foi de propósito — sem isso, _seed_inicial e
        # _importar_gestao_operacao reimportariam a planilha antiga sozinhos
        # no próximo deploy, assim que virem as tabelas vazias. (RNC e P&D
        # não têm reimportação automática recorrente — _seed_rnc_qualidade já
        # roda uma única vez, guardado por ControleSistema próprio, e nunca
        # reimporta de novo mesmo com a tabela vazia; P&D não tem seed nenhum.)
        if ControleSistema.query.filter_by(chave=_CHAVE_DADOS_ZERADOS_MANUALMENTE).first() is None:
            db.session.add(ControleSistema(chave=_CHAVE_DADOS_ZERADOS_MANUALMENTE))

        db.session.commit()
        app.logger.warning(
            "ZERAR DADOS: %s executou a limpeza manual — %d pedidos, %d itens, %d pedidos de Gestão "
            "Operação, %d RNCs, %d inspeções RDIM e %d projetos de P&D apagados.",
            current_user.nome, total_pedidos, total_itens, total_pedidos_operacao,
            total_rnc, total_inspecoes_rdim, total_projetos_pd,
        )
        flash(
            f"Pronto: {total_pedidos} pedido(s), {total_itens} item(ns), {total_pedidos_operacao} "
            f"pedido(s) de Gestão Operação, {total_rnc} RNC(s), {total_inspecoes_rdim} inspeção(ões) "
            f"RDIM e {total_projetos_pd} projeto(s) de P&D foram apagados. Pode começar a testar do zero.",
            "success",
        )
        return redirect(url_for("admin_zerar_dados"))

    # ------------------------------------------------------------------
    # Estações — visão geral por setor + Kanban de produção
    # ------------------------------------------------------------------
    @app.route("/estacoes")
    @login_required
    def estacoes_lista():
        hoje = date.today()
        estacoes_por_nome = {e.nome: e for e in Estacao.query.all()}

        def _op_e_itens(*filtros):
            """(nº de OP/produto, soma de itens) — pedido do Bruno (11/09/2026):
            cada ItemPedido é 1 linha de OP/produto no pedido, e o campo
            quantidade é o total de peças físicas daquela linha. Uma query só
            (count + sum) pra não duplicar ida ao banco por estação/bucket."""
            op, itens = db.session.query(
                func.count(ItemPedido.id), func.coalesce(func.sum(ItemPedido.quantidade), 0)
            ).filter(*filtros).one()
            return op, int(round(itens))

        def _linha(e):
            # Correção (pedido do Bruno, 11/09/2026): "na fila" contava TODO
            # item não finalizado (inclusive os que já estavam em produção),
            # então uma estação sem nada esperando pra começar ainda assim
            # aparecia com itens "na fila". Agora "fila" é só PENDENTE (ainda
            # não começou) e "em produção" é ANDAMENTO/EM TRATATIVA (já
            # começou, mesma régua do status_chao usado no Kanban da
            # estação, pra nunca mais divergir). Cada bucket mostra 2
            # números: quantas OPs/produtos (linhas de ItemPedido) e o
            # total de itens (soma de quantidade) dentro delas.
            fila_op, fila_itens = _op_e_itens(ItemPedido.estacao == e.nome, ItemPedido.status_producao == "PENDENTE")
            producao_op, producao_itens = _op_e_itens(
                ItemPedido.estacao == e.nome, ItemPedido.status_producao.notin_(["FINALIZADO", "PENDENTE"])
            )
            criticos = ItemPedido.query.filter(
                ItemPedido.estacao == e.nome,
                ItemPedido.status_producao != "FINALIZADO",
                ItemPedido.liberacao_prevista.isnot(None),
                ItemPedido.liberacao_prevista < hoje,
            ).count()
            lead_times = _lead_times_estacao(e.nome)
            return {
                "estacao": e, "rotulo": rotulo_estacao(e.nome),
                "fila": fila_op, "fila_itens": fila_itens,
                "em_producao": producao_op, "em_producao_itens": producao_itens,
                "criticos": criticos, "lt_medio": lead_times["chao"], "lead_times": lead_times,
            }

        # 3 colunas fixas (pedido do Bruno, 03/09/2026) — ver
        # ESTACOES_GRUPOS_MONITORAMENTO em models.py. Só entram estações
        # ativas; uma estação ativa que não conste em nenhum grupo cai numa
        # coluna extra "Outras estações" no final, pra nunca sumir em
        # silêncio do monitoramento (ex.: uma estação nova cadastrada depois
        # sem eu saber encaixar num dos 3 grupos).
        grupos = []
        nomes_agrupados = set()
        for grupo in ESTACOES_GRUPOS_MONITORAMENTO:
            linhas_grupo = []
            for nome in grupo["estacoes"]:
                nomes_agrupados.add(nome)
                e = estacoes_por_nome.get(nome)
                if e is not None and e.ativo:
                    linhas_grupo.append(_linha(e))
            grupos.append({"titulo": grupo["titulo"], "linhas": linhas_grupo})

        extras = [
            _linha(e)
            for nome, e in sorted(estacoes_por_nome.items())
            if nome not in nomes_agrupados and e.ativo
        ]
        if extras:
            grupos.append({"titulo": "Outras estações", "linhas": extras})

        return render_template("estacoes_lista.html", grupos=grupos)

    @app.route("/estacoes/<nome>")
    @login_required
    def estacao_kanban(nome):
        estacao = Estacao.query.filter_by(nome=nome).first()
        if estacao is None:
            flash("Estação não encontrada.", "danger")
            return redirect(url_for("estacoes_lista"))

        itens = ItemPedido.query.options(selectinload(ItemPedido.pedido)).filter(ItemPedido.estacao == nome).all()

        colunas = {chave: [] for chave in STATUS_CHAO_OPCOES}
        for item in itens:
            colunas[item.status_chao].append(item)

        # Pedido do Bruno (03/09/2026, em TODAS as estações): em Pendente e Em
        # produção, prazo mais curto (liberação prevista) sempre primeiro —
        # quem "aperta" mais fica visível sem precisar rolar a coluna. Item
        # sem liberação prevista cadastrada vai pro final da coluna (não tem
        # como saber se é urgente ou não).
        def _chave_prazo(item):
            return (item.liberacao_prevista is None, item.liberacao_prevista or date.max, item.id)

        for chave in ("PENDENTE", "EM_PRODUCAO"):
            colunas[chave].sort(key=_chave_prazo)

        # Finalizado continua no critério antigo (25/08/2026): item mais
        # recém concluído/movimentado no topo — depois de entregue, prazo
        # previsto deixa de ser o que importa pra essa coluna.
        colunas["FINALIZADO"].sort(key=lambda i: (i.atualizado_em or datetime.min, i.id), reverse=True)

        def _agrupar_por_pedido(itens_ordenados):
            """Agrupa os itens (já ordenados pelo critério da coluna acima) por
            PEDIDO — pedido do Bruno (10/09/2026): "agrupe os itens da
            novelis... agrupe também os do pedido 873" — cards soltos
            repetindo cliente/frete/UF/cidade em cada item viravam ruído
            visual quando o mesmo pedido tinha vários produtos na mesma
            coluna.

            A posição de cada GRUPO na coluna é a do seu item mais urgente —
            o primeiro que aparece na lista já ordenada — então a ordem de
            prioridade da coluna continua valendo, só que por pedido; dentro
            do grupo, os itens mantêm a ordem relativa que já vinham
            (também por urgência). Um pedido com produtos em estações
            diferentes só agrupa os itens QUE ESTÃO nesta estação — não
            mistura com itens de outras estações."""
            grupos = {}
            ordem_pedidos = []
            for item in itens_ordenados:
                pid = item.pedido_id
                if pid not in grupos:
                    grupos[pid] = []
                    ordem_pedidos.append(pid)
                grupos[pid].append(item)
            return [(grupos[pid][0].pedido, grupos[pid]) for pid in ordem_pedidos]

        colunas_agrupadas = {chave: _agrupar_por_pedido(colunas[chave]) for chave in STATUS_CHAO_OPCOES}

        # Total de matéria-prima (kg) por item e por coluna — pedido do
        # Bruno (22/09/2026): "quero enxergar o total de materia prima" nas
        # Estações, "quando eu abrir um item, ou colunas do kanban".
        # `_materiais_item_pedido` casa cada item com o catálogo de Gestão
        # de Custos (Produto/EstruturaProduto/MateriaPrima) — item sem
        # correspondência automática fica com matched=False, mostrado como
        # "não identificado" no card em vez de forjar um número.
        materiais_por_item = {item.id: _materiais_item_pedido(item) for item in itens}
        kg_por_coluna = {}
        nao_identificados_por_coluna = {}
        # Pedido do Bruno (22/09/2026, revisão): o total borrado da coluna
        # ("≈ 75,1 kg") não é o que ele quer ver de cara — quer o consumo
        # discriminado por matéria-prima (ex. "2471 200 kg, 2475 200 kg, 122
        # 200 kg"), a mesma lista nomeada do pedido original. Agrega as
        # `linhas` de todo item IDENTIFICADO da coluna, consolidando por
        # matéria-prima — abre num modal por coluna (mesmo padrão do modal
        # por item).
        #
        # Pedido do Bruno (22/09/2026, à noite): "eu não quero que você deixe
        # escondido, quero que deixe visível... isso vale pra todas as
        # matérias-primas AMINO, COIM, LANXESS, BLOCO DE ESPUMA, TECPUR" —
        # ou seja, TODAS as matérias-primas da estrutura, não só as em kg.
        # A primeira versão deste agregado filtrava `unidade != "kg"` (pra
        # bater com a decisão anterior de só SOMAR em kg no total do topo),
        # mas isso escondia por completo do popup da coluna matérias como
        # BLOCO ESPUMA D26/D45/D60/D80 (m³) — que aparecem certinho no popup
        # de cada ITEM, só não estavam sendo agregadas aqui na coluna. Agora
        # agrega TODAS, cada uma na sua própria unidade — o total em kg no
        # topo do popup continua só em kg (decisão de 22/09/2026 cedo, "só o
        # total em kg"), mas nenhuma matéria-prima fica de fora da lista.
        materiais_agrupados_por_coluna = {}
        for chave in STATUS_CHAO_OPCOES:
            info_coluna = [materiais_por_item[item.id] for item in colunas[chave]]
            kg_por_coluna[chave] = round(sum(i["kg_total"] for i in info_coluna), 1)
            nao_identificados_por_coluna[chave] = sum(1 for i in info_coluna if not i["matched"])

            agregados = {}
            for info in info_coluna:
                for linha in info["linhas"]:
                    mp = linha["materia_prima"]
                    bucket = agregados.setdefault(mp.id, {"materia_prima": mp, "quantidade": 0.0})
                    bucket["quantidade"] += linha["quantidade"]
            # kg primeiro (bate com o total do topo do popup), depois as
            # demais unidades — dentro de cada grupo, maior quantidade primeiro.
            materiais_agrupados_por_coluna[chave] = sorted(
                agregados.values(),
                key=lambda l: (l["materia_prima"].unidade != "kg", -l["quantidade"]),
            )

        # Pedido do Bruno (23/09/2026): "eu quero que detalhe AO LADO a
        # somatória de cada material... focado nos principais como citei
        # acima" (AMINO, COIM, LANXESS, BLOCO DE ESPUMA, TECPUR) — ele não
        # quer precisar clicar pra ver isso (o popup completo continua
        # existindo, mas só pra ferragem/acessório e pro resto que não é
        # um desses 5 grupos). Subconjunto de `materiais_agrupados_por_coluna`
        # já calculado acima, filtrado por `_e_materia_prima_principal` —
        # normalmente só 3 a 8 linhas por coluna (bem menor que o total de
        # matérias-primas de uma coluna cheia), então cabe direto no
        # cabeçalho sem repetir o estouro de layout que já corrigimos antes.
        materiais_principais_por_coluna = {
            chave: [l for l in linhas if _e_materia_prima_principal(l["materia_prima"])]
            for chave, linhas in materiais_agrupados_por_coluna.items()
        }

        return render_template(
            "estacoes_kanban.html",
            estacao=estacao,
            rotulo=rotulo_estacao(estacao.nome),
            colunas=colunas,
            colunas_agrupadas=colunas_agrupadas,
            pode_editar=pode_editar_estacao(current_user, nome),
            RELATORIO_ESTACAO_STATUS_INFO=RELATORIO_ESTACAO_STATUS_INFO,
            materiais_por_item=materiais_por_item,
            kg_por_coluna=kg_por_coluna,
            nao_identificados_por_coluna=nao_identificados_por_coluna,
            materiais_agrupados_por_coluna=materiais_agrupados_por_coluna,
            materiais_principais_por_coluna=materiais_principais_por_coluna,
        )

    @app.route("/estacoes/<nome>/relatorio.pdf")
    @login_required
    def estacao_relatorio_pdf(nome):
        estacao = Estacao.query.filter_by(nome=nome).first()
        if estacao is None:
            flash("Estação não encontrada.", "danger")
            return redirect(url_for("estacoes_lista"))

        status_filtro = request.args.get("status", "ambos")
        if status_filtro not in RELATORIO_ESTACAO_STATUS_INFO:
            status_filtro = "ambos"

        itens = _itens_relatorio_estacao(nome, status_filtro)
        return _gerar_pdf_estacao(estacao, itens, status_filtro)

    @app.route("/estacoes/relatorio-multiplo.pdf")
    @login_required
    def estacoes_relatorio_multiplo_pdf():
        """Relatório PDF de VÁRIAS estações escolhidas de uma vez (pedido do
        Bruno, 22/09/2026: "crie uma area onde posso selecionar diversas
        areas para gerar o relatorio... ex: quero ver em um relatorio em pdf
        oque tem pendente e andamento no PU e Espumagem"). Rota SEM
        "/<nome>" — registrada como caminho fixo, então tem prioridade sobre
        "/estacoes/<nome>" (Kanban) pra esse path exato, sem risco de
        conflito de rota."""
        nomes = [v for v in request.args.getlist("estacao") if v]
        status_filtro = request.args.get("status", "ambos")
        if status_filtro not in RELATORIO_ESTACAO_STATUS_INFO:
            status_filtro = "ambos"

        if not nomes:
            flash("Selecione ao menos uma estação para gerar o relatório.", "warning")
            return redirect(url_for("estacoes_lista"))

        estacoes_por_nome = {e.nome: e for e in Estacao.query.filter(Estacao.nome.in_(nomes)).all()}

        # Mantém a MESMA ordem de agrupamento por processo da tela /estacoes
        # (ESTACOES_GRUPOS_MONITORAMENTO) em vez da ordem em que os
        # checkboxes chegaram no formulário — o PDF sempre lê na mesma ordem
        # visual que a pessoa já reconhece na tela, não importa a ordem que
        # ela marcou.
        ordem = []
        vistos = set()
        for grupo in ESTACOES_GRUPOS_MONITORAMENTO:
            for nome_grupo in grupo["estacoes"]:
                if nome_grupo in estacoes_por_nome and nome_grupo not in vistos:
                    ordem.append(nome_grupo)
                    vistos.add(nome_grupo)
        for nome_extra in sorted(estacoes_por_nome):
            if nome_extra not in vistos:
                ordem.append(nome_extra)
                vistos.add(nome_extra)

        estacoes_com_itens = [
            (estacoes_por_nome[nome_ord], _itens_relatorio_estacao(nome_ord, status_filtro))
            for nome_ord in ordem
        ]
        return _gerar_pdf_estacoes_multiplas(estacoes_com_itens, status_filtro)

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
