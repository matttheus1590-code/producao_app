from calendar import monthrange
from datetime import date, datetime

from flask_login import UserMixin
from werkzeug.security import check_password_hash, generate_password_hash

from extensions import db

# ---------------------------------------------------------------------------
# Opções fixas usadas nos formulários (dropdowns). Ajuste livremente aqui —
# o resto do sistema lê sempre destas listas.
# ---------------------------------------------------------------------------
ESTACOES = [
    "CORTE",
    "MANDRIL",
    "ESPUMAGEM",
    "SILICONE",
    "PU",
    "REFORMA",
    "REVENDA",
    "PROJETO ESPECIAL",
    "OUTROS",
]

STATUS_OPCOES = ["PENDENTE", "EM TRATATIVA", "ANDAMENTO", "FINALIZADO"]

# Status de CHÃO DE FÁBRICA — mais granular que o STATUS_OPCOES "comercial" acima
# (PENDENTE/EM TRATATIVA/ANDAMENTO/FINALIZADO). É só uma leitura mais detalhada
# das MESMAS datas que já existem em cada item (não é um campo novo, não duplica
# nada) — usado no Kanban da tela Estações.
STATUS_CHAO_OPCOES = ["NAO_INICIADO", "PROGRAMADO", "EM_PRODUCAO", "INSPECAO", "EMBALAGEM", "FINALIZADO"]
STATUS_CHAO_LABELS = {
    "NAO_INICIADO": "Pendente",
    "PROGRAMADO": "Programado",
    "EM_PRODUCAO": "Em produção",
    "INSPECAO": "Inspeção",
    "EMBALAGEM": "Embalagem",
    "FINALIZADO": "Finalizado",
}
STATUS_CHAO_CORES = {
    "NAO_INICIADO": "secondary",
    "PROGRAMADO": "info",
    "EM_PRODUCAO": "primary",
    "INSPECAO": "warning",
    "EMBALAGEM": "warning",
    "FINALIZADO": "success",
}

# Região de cada UF, usada só para o filtro "Região" da listagem — a planilha
# original não tem uma coluna de região separada, então agrupamos a partir do
# estado (Pedido.estado) já existente.
REGIAO_POR_UF = {
    "AC": "Norte", "AP": "Norte", "AM": "Norte", "PA": "Norte", "RO": "Norte", "RR": "Norte", "TO": "Norte",
    "AL": "Nordeste", "BA": "Nordeste", "CE": "Nordeste", "MA": "Nordeste", "PB": "Nordeste",
    "PE": "Nordeste", "PI": "Nordeste", "RN": "Nordeste", "SE": "Nordeste",
    "DF": "Centro-Oeste", "GO": "Centro-Oeste", "MT": "Centro-Oeste", "MS": "Centro-Oeste",
    "ES": "Sudeste", "MG": "Sudeste", "RJ": "Sudeste", "SP": "Sudeste",
    "PR": "Sul", "RS": "Sul", "SC": "Sul",
}
REGIOES_OPCOES = ["Norte", "Nordeste", "Centro-Oeste", "Sudeste", "Sul"]

PRIORIDADE_OPCOES = ["BAIXA", "MÉDIA", "ALTA"]

FRETE_OPCOES = ["CIF", "FOB", "SEM FRETE", "EXPORTAÇÃO", "OUTRO"]

STATUS_CORES = {
    "PENDENTE": "secondary",
    "EM TRATATIVA": "warning",
    "ANDAMENTO": "info",
    "FINALIZADO": "success",
}

PRIORIDADE_CORES = {
    "BAIXA": "secondary",
    "MÉDIA": "info",
    "ALTA": "danger",
}

# ---------------------------------------------------------------------------
# Semáforo de prazo: compara a liberação prevista com hoje.
#   verde    -> ainda falta bastante tempo
#   amarelo  -> vencendo (dentro de PRAZO_ALERTA_DIAS dias)
#   vermelho -> já passou do prazo
#   cinza    -> não dá pra saber (sem data prevista, ou item já finalizado)
# ---------------------------------------------------------------------------
PRAZO_ALERTA_DIAS = 3  # ajuste este número se quiser um limite diferente de "vencendo"

SEMAFORO_CORES = {"verde": "success", "amarelo": "warning", "vermelho": "danger", "cinza": "secondary"}
SEMAFORO_LABELS = {"verde": "No prazo", "amarelo": "Vencendo", "vermelho": "Atrasado", "cinza": "Sem prazo"}
_SEMAFORO_PRIORIDADE = {"vermelho": 0, "amarelo": 1, "verde": 2, "cinza": 3}  # menor = mais urgente

# ---------------------------------------------------------------------------
# Gestão Operação (Fase 13) — meta mínima de OTD (On-Time Delivery), pedida
# pelo Bruno. Ajuste este número se a meta mudar — todo o resto lê daqui.
# ---------------------------------------------------------------------------
GO_OTD_META_PERCENTUAL = 78

_MESES_ABREV_PCP = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]


def gerar_semanas_pcp(meses_atras=1, meses_frente=6, hoje=None):
    """Gera as opções de "semana" pro campo Término Semanal PCP, no mesmo
    formato usado na planilha importada: "SEMANA NN / MÊS / ANO" — onde a
    semana do mês é ceil(dia/7) (semana 1 = dias 1-7, semana 2 = 8-14, ...,
    semana 5 = 29-31). Cobre de `meses_atras` meses atrás até `meses_frente`
    meses à frente do mês atual, sempre em ordem cronológica (útil tanto pro
    dropdown do formulário quanto pra ordenar o gráfico de projeção do Painel).
    """
    hoje = hoje or date.today()
    ano, mes = hoje.year, hoje.month
    mes -= meses_atras
    while mes <= 0:
        mes += 12
        ano -= 1

    opcoes = []
    for _ in range(meses_atras + meses_frente + 1):
        dias_no_mes = monthrange(ano, mes)[1]
        n_semanas = -(-dias_no_mes // 7)  # ceil(dias_no_mes / 7)
        for semana in range(1, n_semanas + 1):
            opcoes.append(f"SEMANA {semana:02d} / {_MESES_ABREV_PCP[mes - 1]} / {ano}")
        mes += 1
        if mes > 12:
            mes = 1
            ano += 1
    return opcoes


def semaforo_prazo(liberacao_prevista, finalizado=False, hoje=None):
    """Retorna (cor, dias) comparando a liberação prevista com hoje.

    dias negativo = atrasado; dias positivo = quanto falta.
    """
    hoje = hoje or date.today()
    if finalizado or not liberacao_prevista:
        return ("cinza", None)
    dias = (liberacao_prevista - hoje).days
    if dias < 0:
        return ("vermelho", dias)
    if dias <= PRAZO_ALERTA_DIAS:
        return ("amarelo", dias)
    return ("verde", dias)


class Usuario(UserMixin, db.Model):
    __tablename__ = "usuarios"

    id = db.Column(db.Integer, primary_key=True)
    nome = db.Column(db.String(120), nullable=False)
    username = db.Column(db.String(80), unique=True, nullable=False)
    senha_hash = db.Column(db.String(255), nullable=False)

    # ---- Papéis de acesso (ADMIN / PCP / LIDER / GESTAO) ----
    # ADMIN e PCP podem cadastrar/editar pedidos e programação.
    # LIDER só edita a produção da própria estação (campo "setor").
    # GESTAO acompanha tudo, mas não edita pedidos/produção.
    role = db.Column(db.String(20), nullable=False, default="PCP")
    setor = db.Column(db.String(40), nullable=True)  # usado só quando role == "LIDER"
    ativo = db.Column(db.Boolean, nullable=False, default=True)

    def set_senha(self, senha):
        self.senha_hash = generate_password_hash(senha)

    def checar_senha(self, senha):
        return check_password_hash(self.senha_hash, senha)

    @property
    def is_active(self):
        # Sobrescreve o padrão do Flask-Login (que assume sempre True) para que
        # usuários desativados não consigam mais entrar, sem precisar apagá-los.
        return self.ativo


class Estacao(db.Model):
    """Estação de produção — antes era só uma lista fixa (ESTACOES) em Python,
    agora é uma tabela de verdade, o que permite ordenar, desativar e (no
    futuro) guardar meta de lead time por estação sem mexer em código.

    Os itens continuam guardando a estação pelo NOME (ItemPedido.estacao é
    texto, não uma chave estrangeira) — assim nenhum pedido existente perde a
    referência quando esta tabela é criada."""

    __tablename__ = "estacoes"

    id = db.Column(db.Integer, primary_key=True)
    nome = db.Column(db.String(40), unique=True, nullable=False)
    ordem_exibicao = db.Column(db.Integer, nullable=False, default=0)
    meta_lead_time_dias = db.Column(db.Integer, nullable=True)
    ativo = db.Column(db.Boolean, nullable=False, default=True)


class Transportadora(db.Model):
    """Cadastro de transportadoras — dado 100% novo, não existe nada parecido
    na planilha original (só o "Modelo Frete"/incoterm, que é outra coisa)."""

    __tablename__ = "transportadoras"

    id = db.Column(db.Integer, primary_key=True)
    nome = db.Column(db.String(120), unique=True, nullable=False)
    ativo = db.Column(db.Boolean, nullable=False, default=True)


class Pedido(db.Model):
    __tablename__ = "pedidos"

    id = db.Column(db.Integer, primary_key=True)

    # ---------------- Dados BASE (inclusão manual do pedido) ----------------
    data_cliente = db.Column(db.Date, nullable=True)
    data_inclusao_pedido = db.Column(db.Date, nullable=True)
    cliente = db.Column(db.String(200), nullable=False)
    cnpj = db.Column(db.String(30), nullable=True)
    cidade = db.Column(db.String(120), nullable=True)
    estado = db.Column(db.String(60), nullable=True)
    pais = db.Column(db.String(60), nullable=True, default="Brasil")
    frete = db.Column(db.String(30), nullable=True)
    vendedor = db.Column(db.String(120), nullable=True)
    pedido_venda = db.Column(db.String(40), nullable=True)

    # Um pedido pode ter vários produtos diferentes (itens)
    itens = db.relationship(
        "ItemPedido",
        backref="pedido",
        cascade="all, delete-orphan",
        order_by="ItemPedido.id",
    )

    # ------------- Dados do pedido como um todo -------------
    prioridade = db.Column(db.String(20), default="MÉDIA")
    obs = db.Column(db.Text, nullable=True)

    # Estação, status, datas de produção e RNC ficam em cada ITEM (ItemPedido),
    # já que produtos diferentes do mesmo pedido podem estar em etapas diferentes.

    # ---------------------------------------------------------------------
    # Gestão Operação (Fase 13) — fluxo comercial completo do PEDIDO (não do
    # item), em 4 blocos que espelham a planilha "Gestão de Fluxo Produtivo":
    # Comercial (azul) -> PCP (cinza) -> Logística/NF (amarelo) -> Resultados
    # e OTD (verde). Campos 100% novos e opcionais — não substituem nada que
    # já existe; onde já havia um campo equivalente em Pedido (cliente,
    # vendedor, pedido_venda, data_inclusao_pedido, prioridade, frete, pais,
    # estado, cidade) ele é reaproveitado, não duplicado aqui.
    # ---------------------------------------------------------------------

    # -- Comercial (azul) --
    go_tipo_pedido = db.Column(db.String(60), nullable=True)
    go_contrato = db.Column(db.String(60), nullable=True)
    go_pedido_compra_cliente = db.Column(db.String(60), nullable=True)
    go_proposta = db.Column(db.String(60), nullable=True)
    go_data_solicitada_entrega = db.Column(db.Date, nullable=True)
    go_status_pedido_info = db.Column(db.String(120), nullable=True)
    go_valor_pedido_operacao = db.Column(db.Float, nullable=True)

    # -- PCP (cinza) --
    go_previsao_liberacao_pcp = db.Column(db.Date, nullable=True)
    go_data_efetiva_liberacao_pcp = db.Column(db.Date, nullable=True)
    go_data_solicitada_cliente_retira = db.Column(db.Date, nullable=True)
    go_custo_producao_real = db.Column(db.Float, nullable=True)
    go_termino_semanal_pcp = db.Column(db.String(40), nullable=True)

    # -- Logística / NF (amarelo) --
    go_data_emissao_nf = db.Column(db.Date, nullable=True)
    go_valor_nf_emitida = db.Column(db.Float, nullable=True)
    go_numero_nf = db.Column(db.String(30), nullable=True)
    go_status_logistica = db.Column(db.String(60), nullable=True)
    go_data_pedido_expedido = db.Column(db.Date, nullable=True)
    # Mesmo padrão do ItemPedido.transportadora_id (sem FK de banco de verdade
    # — validado na aplicação, não no banco).
    go_transportadora_id = db.Column(db.Integer, nullable=True)
    go_custo_frete_previsto = db.Column(db.Float, nullable=True)
    go_custo_frete_final = db.Column(db.Float, nullable=True)
    go_custo_frete_sobre_nota = db.Column(db.Float, nullable=True)
    go_data_prevista_entrega = db.Column(db.Date, nullable=True)
    go_data_real_entrega = db.Column(db.Date, nullable=True)

    # -- Resultados / OTD (verde) --
    go_otd_realizado = db.Column(db.String(10), nullable=True)  # "SIM" / "NÃO"
    go_data_solicitada_cliente_final = db.Column(db.Date, nullable=True)
    go_data_entregue_cliente = db.Column(db.Date, nullable=True)
    go_obs_operacao = db.Column(db.Text, nullable=True)
    go_status_final_alinhamento = db.Column(db.String(60), nullable=True)

    criado_em = db.Column(db.DateTime, default=datetime.utcnow)
    atualizado_em = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    # ---------------------------------------------------------------------
    # Campos 100% calculados — nunca editados diretamente pelo usuário
    # ---------------------------------------------------------------------
    @property
    def valor_total(self):
        return round(sum((item.valor_total for item in self.itens), 0.0), 2)

    @property
    def quantidade_total(self):
        return sum((item.quantidade or 0) for item in self.itens)

    @property
    def descricao_resumo(self):
        """Texto curto para listagens: primeiro produto + contagem dos demais."""
        if not self.itens:
            return "—"
        primeiro = self.itens[0].descricao_produto
        if len(self.itens) > 1:
            return f"{primeiro} (+{len(self.itens) - 1} item(ns))"
        return primeiro

    @property
    def estacao_resumo(self):
        """Texto curto para listagens: estação, ou aviso se os itens estão em estações diferentes."""
        estacoes = {item.estacao for item in self.itens if item.estacao}
        if not estacoes:
            return "—"
        if len(estacoes) == 1:
            return next(iter(estacoes))
        return f"{len(estacoes)} estações"

    @property
    def status_producao(self):
        """Status do pedido como um todo, calculado a partir do status de cada item.

        Regras:
          - Algum item "EM TRATATIVA"          -> pedido "EM TRATATIVA" (precisa de atenção)
          - Todos os itens "FINALIZADO"        -> pedido "FINALIZADO"
          - Todos os itens "PENDENTE"          -> pedido "PENDENTE"
          - Qualquer outra mistura de status   -> pedido "ANDAMENTO"
        """
        if not self.itens:
            return "PENDENTE"
        status_itens = {item.status_producao for item in self.itens}
        if "EM TRATATIVA" in status_itens:
            return "EM TRATATIVA"
        if status_itens == {"FINALIZADO"}:
            return "FINALIZADO"
        if status_itens == {"PENDENTE"}:
            return "PENDENTE"
        return "ANDAMENTO"

    @property
    def lt_comercial_dias(self):
        if self.data_cliente and self.data_inclusao_pedido:
            return (self.data_inclusao_pedido - self.data_cliente).days
        return None

    @property
    def prazo_total_dias(self):
        """Prazo do pedido inteiro: da data do cliente até o item que terminar por último
        (ou até hoje, se ainda tiver item em aberto)."""
        if not self.data_cliente:
            return None
        if self.itens and all(item.termino_inspecao for item in self.itens):
            fim = max(item.termino_inspecao for item in self.itens)
        else:
            fim = date.today()
        return (fim - self.data_cliente).days

    @property
    def semaforo(self):
        """Pior semáforo de prazo entre os itens do pedido (o mais urgente manda)."""
        if not self.itens:
            return ("cinza", None)
        piores = sorted((item.semaforo for item in self.itens), key=lambda s: _SEMAFORO_PRIORIDADE[s[0]])
        return piores[0]

    # ---------------------------------------------------------------------
    # Gestão Operação — calculados, nunca editados diretamente (mesmo padrão
    # de lt_comercial_dias/prazo_total_dias acima).
    # ---------------------------------------------------------------------
    @property
    def go_transportadora(self):
        if not self.go_transportadora_id:
            return None
        return db.session.get(Transportadora, self.go_transportadora_id)

    @property
    def go_lead_time_frete_dias(self):
        """Saída (expedição) até chegada (entrega/coleta real)."""
        if self.go_data_pedido_expedido and self.go_data_real_entrega:
            return (self.go_data_real_entrega - self.go_data_pedido_expedido).days
        return None

    @property
    def go_lead_time_operacao_dias(self):
        """Inclusão do pedido até a entrega efetiva no cliente."""
        if self.data_inclusao_pedido and self.go_data_entregue_cliente:
            return (self.go_data_entregue_cliente - self.data_inclusao_pedido).days
        return None

    @property
    def go_dias_atraso_antecipacao(self):
        """Positivo = entregue depois do solicitado (atraso); negativo = antes (antecipação)."""
        if self.go_data_solicitada_cliente_final and self.go_data_entregue_cliente:
            return (self.go_data_entregue_cliente - self.go_data_solicitada_cliente_final).days
        return None


class ItemPedido(db.Model):
    """Um produto dentro de um pedido. Um pedido pode ter vários itens, e cada
    item avança pela produção de forma independente (estação, status, datas)."""

    __tablename__ = "itens_pedido"

    id = db.Column(db.Integer, primary_key=True)
    pedido_id = db.Column(db.Integer, db.ForeignKey("pedidos.id"), nullable=False)

    descricao_produto = db.Column(db.String(300), nullable=False)
    quantidade = db.Column(db.Float, nullable=False, default=0)
    custo_unitario = db.Column(db.Float, nullable=False, default=0)

    # ------------- Dados PARAMETRIZADOS (evoluem durante a produção deste item) -------------
    estacao = db.Column(db.String(40), nullable=True)
    status_producao = db.Column(db.String(20), default="PENDENTE")
    # quando True, o status foi definido manualmente como "EM TRATATIVA" e o
    # cálculo automático não deve sobrescrevê-lo até que as datas de término existam
    status_manual = db.Column(db.Boolean, default=False)

    inicio_producao = db.Column(db.Date, nullable=True)
    inicio_inspecao = db.Column(db.Date, nullable=True)
    termino_inspecao = db.Column(db.Date, nullable=True)
    liberacao_faturamento = db.Column(db.Date, nullable=True)
    liberacao_prevista = db.Column(db.Date, nullable=True)
    # Preenchida manualmente (não é calculada a partir da liberação prevista
    # nem de nenhuma outra data) — a data em que o item foi liberado de
    # verdade, pra comparar com a liberação prevista.
    liberacao_real = db.Column(db.Date, nullable=True)
    # Preenchido manualmente pelo PCP, junto com a liberação prevista — não é
    # calculado a partir dela. Guarda o rótulo de semana no mesmo formato de
    # gerar_semanas_pcp() (ex.: "SEMANA 03 / AGO / 2026").
    planejamento_semanal = db.Column(db.String(40), nullable=True)

    rnc = db.Column(db.String(120), nullable=True)

    # ------------- Faturamento (dados novos — a planilha original não tinha isso) -------------
    numero_nota_fiscal = db.Column(db.String(30), nullable=True)
    valor_faturado = db.Column(db.Float, nullable=True)

    # ------------- Logística (dados novos) -------------
    # transportadora_id NÃO usa db.ForeignKey de propósito: essa coluna é
    # adicionada numa tabela que já existe (via ALTER TABLE), e SQLite não
    # aplica chave estrangeira de verdade nesse caso — então validamos na
    # aplicação (rota) em vez de depender do banco pra isso.
    transportadora_id = db.Column(db.Integer, nullable=True)
    data_envio = db.Column(db.Date, nullable=True)

    @property
    def transportadora(self):
        if not self.transportadora_id:
            return None
        return db.session.get(Transportadora, self.transportadora_id)

    @property
    def valor_faturamento_realizado(self):
        """Valor a considerar como "realizado" no previsto × realizado: usa o
        valor faturado de verdade quando preenchido, senão cai pro valor do
        item (comportamento antigo, antes deste campo existir)."""
        return self.valor_faturado if self.valor_faturado is not None else self.valor_total

    @property
    def valor_total(self):
        return round((self.quantidade or 0) * (self.custo_unitario or 0), 2)

    @property
    def tempo_espera_dias(self):
        if self.pedido and self.pedido.data_inclusao_pedido and self.inicio_producao:
            return (self.inicio_producao - self.pedido.data_inclusao_pedido).days
        return None

    @property
    def lt_producao_dias(self):
        if self.inicio_producao and self.termino_inspecao:
            return (self.termino_inspecao - self.inicio_producao).days
        return None

    @property
    def prazo_total_dias(self):
        if not (self.pedido and self.pedido.data_cliente):
            return None
        fim = self.termino_inspecao or date.today()
        return (fim - self.pedido.data_cliente).days

    @property
    def semaforo(self):
        return semaforo_prazo(self.liberacao_prevista, finalizado=(self.status_producao == "FINALIZADO"))

    @property
    def status_chao(self):
        """Status de chão de fábrica (mais granular), calculado a partir das
        MESMAS datas que status_producao já usa — nenhum campo novo — mais a
        existência de uma Programação ativa para decidir entre "Pendente" e
        "Programado"."""
        if self.termino_inspecao and self.liberacao_faturamento:
            return "FINALIZADO"
        if self.termino_inspecao:
            return "EMBALAGEM"
        if self.inicio_inspecao:
            return "INSPECAO"
        if self.inicio_producao:
            return "EM_PRODUCAO"
        if any(p.status == "ATIVA" for p in self.programacoes):
            return "PROGRAMADO"
        return "NAO_INICIADO"

    def atualizar_status_automatico(self):
        """Recalcula status_producao deste item a partir das datas preenchidas.

        Regras:
          - Término de inspeção + liberação de faturamento preenchidos -> FINALIZADO
          - Início de produção preenchido (mas ainda sem término)      -> ANDAMENTO
          - Nada preenchido                                            -> PENDENTE
          - "EM TRATATIVA" é um estado manual: uma vez selecionado pelo
            usuário ele fica travado até o item ser finalizado.
        """
        if self.termino_inspecao and self.liberacao_faturamento:
            self.status_producao = "FINALIZADO"
            self.status_manual = False
            return

        if self.status_manual and self.status_producao == "EM TRATATIVA":
            return

        if self.inicio_producao:
            self.status_producao = "ANDAMENTO"
        else:
            self.status_producao = "PENDENTE"


class Programacao(db.Model):
    """Um agendamento de produção: "este item vai ser trabalhado na estação X,
    no dia Y". Reprogramar NUNCA sobrescreve a data antiga — marca o registro
    velho como REPROGRAMADA e cria um novo ATIVO, então o histórico de
    reagendamentos fica registrado automaticamente, sem precisar duplicar isso
    no histórico de alterações."""

    __tablename__ = "programacoes"

    id = db.Column(db.Integer, primary_key=True)
    item_pedido_id = db.Column(db.Integer, db.ForeignKey("itens_pedido.id"), nullable=False)
    item = db.relationship("ItemPedido", backref="programacoes")

    data_programada = db.Column(db.Date, nullable=False)
    estacao = db.Column(db.String(40), nullable=False)  # snapshot — pode divergir do item se ele for reatribuído depois
    prioridade_producao = db.Column(db.String(20), nullable=True)  # independente da prioridade comercial do pedido
    observacao = db.Column(db.Text, nullable=True)

    responsavel_id = db.Column(db.Integer, db.ForeignKey("usuarios.id"), nullable=True)
    responsavel = db.relationship("Usuario", foreign_keys=[responsavel_id])

    status = db.Column(db.String(20), nullable=False, default="ATIVA")  # ATIVA / REPROGRAMADA / CANCELADA

    criado_por_id = db.Column(db.Integer, db.ForeignKey("usuarios.id"), nullable=True)
    criado_por = db.relationship("Usuario", foreign_keys=[criado_por_id])
    criado_em = db.Column(db.DateTime, default=datetime.utcnow)


class PedidoOperacao(db.Model):
    """Gestão Operação — fluxo comercial completo de um pedido (Comercial -> PCP
    -> Logística/NF -> Resultados/OTD). TOTALMENTE INDEPENDENTE de Pedido/ItemPedido
    (Gestão Produção) — tabela própria, id próprio, sem FK nem relação nenhuma com
    a tabela `pedidos`. Cada linha aqui é um pedido comercial, sem a duplicação por
    linha-de-planilha que o histórico legado de Produção tinha.

    Antes (Fase 13) esses mesmos campos viviam como colunas `go_*` dentro de
    `Pedido` — a pedido do Bruno, os dois grupos do site passaram a ser
    independentes um do outro. Essa tabela nasce vazia e é populada uma única vez
    por um backfill a partir dos dados antigos (ver
    _migrar_dados_go_para_pedidos_operacao em app.py); dali em diante, pelo
    formulário "Novo pedido" da Gestão Operação e pelo importador de planilha
    (importar_gestao_operacao.py). Mantém o prefixo "go_" dos campos (mesmo nome de
    antes) só pra minimizar mudança nos templates — não tem mais nenhum significado
    de "coluna extra dentro de Pedido".
    """

    __tablename__ = "pedidos_operacao"

    id = db.Column(db.Integer, primary_key=True)

    # ---------------- Identidade do pedido (antes era reaproveitada de Pedido) ----------------
    pedido_venda = db.Column(db.String(40), nullable=True)
    cliente = db.Column(db.String(200), nullable=False)
    vendedor = db.Column(db.String(120), nullable=True)
    data_inclusao_pedido = db.Column(db.Date, nullable=True)
    prioridade = db.Column(db.String(20), default="MÉDIA")  # Criticidade
    frete = db.Column(db.String(30), nullable=True)
    pais = db.Column(db.String(60), nullable=True, default="Brasil")
    estado = db.Column(db.String(60), nullable=True)
    cidade = db.Column(db.String(120), nullable=True)

    # -- Comercial (azul) --
    go_tipo_pedido = db.Column(db.String(60), nullable=True)
    go_contrato = db.Column(db.String(60), nullable=True)
    go_pedido_compra_cliente = db.Column(db.String(60), nullable=True)
    go_proposta = db.Column(db.String(60), nullable=True)
    go_data_solicitada_entrega = db.Column(db.Date, nullable=True)
    go_status_pedido_info = db.Column(db.String(120), nullable=True)
    go_valor_pedido_operacao = db.Column(db.Float, nullable=True)

    # -- PCP (cinza) --
    go_previsao_liberacao_pcp = db.Column(db.Date, nullable=True)
    go_data_efetiva_liberacao_pcp = db.Column(db.Date, nullable=True)
    go_data_solicitada_cliente_retira = db.Column(db.Date, nullable=True)
    go_custo_producao_real = db.Column(db.Float, nullable=True)
    go_termino_semanal_pcp = db.Column(db.String(40), nullable=True)

    # -- Logística / NF (amarelo) --
    go_data_emissao_nf = db.Column(db.Date, nullable=True)
    go_valor_nf_emitida = db.Column(db.Float, nullable=True)
    go_numero_nf = db.Column(db.String(30), nullable=True)
    go_status_logistica = db.Column(db.String(60), nullable=True)
    go_data_pedido_expedido = db.Column(db.Date, nullable=True)
    # Mesmo padrão do ItemPedido.transportadora_id — sem FK de banco de verdade,
    # validado na aplicação.
    go_transportadora_id = db.Column(db.Integer, nullable=True)
    go_custo_frete_previsto = db.Column(db.Float, nullable=True)
    go_custo_frete_final = db.Column(db.Float, nullable=True)
    go_custo_frete_sobre_nota = db.Column(db.Float, nullable=True)
    go_data_prevista_entrega = db.Column(db.Date, nullable=True)
    go_data_real_entrega = db.Column(db.Date, nullable=True)

    # -- Resultados / OTD (verde) --
    go_otd_realizado = db.Column(db.String(10), nullable=True)  # "SIM" / "NÃO"
    go_data_solicitada_cliente_final = db.Column(db.Date, nullable=True)
    go_data_entregue_cliente = db.Column(db.Date, nullable=True)
    go_obs_operacao = db.Column(db.Text, nullable=True)
    go_status_final_alinhamento = db.Column(db.String(60), nullable=True)

    criado_em = db.Column(db.DateTime, default=datetime.utcnow)
    atualizado_em = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    # ---------------------------------------------------------------------
    # Calculados — mesmo padrão de Pedido, nunca editados diretamente.
    # ---------------------------------------------------------------------
    @property
    def go_transportadora(self):
        if not self.go_transportadora_id:
            return None
        return db.session.get(Transportadora, self.go_transportadora_id)

    @property
    def go_lead_time_frete_dias(self):
        """Saída (expedição) até chegada (entrega/coleta real)."""
        if self.go_data_pedido_expedido and self.go_data_real_entrega:
            return (self.go_data_real_entrega - self.go_data_pedido_expedido).days
        return None

    @property
    def go_lead_time_operacao_dias(self):
        """Inclusão do pedido até a entrega efetiva no cliente."""
        if self.data_inclusao_pedido and self.go_data_entregue_cliente:
            return (self.go_data_entregue_cliente - self.data_inclusao_pedido).days
        return None

    @property
    def go_dias_atraso_antecipacao(self):
        """Positivo = entregue depois do solicitado (atraso); negativo = antes (antecipação)."""
        if self.go_data_solicitada_cliente_final and self.go_data_entregue_cliente:
            return (self.go_data_entregue_cliente - self.go_data_solicitada_cliente_final).days
        return None

    @property
    def status_producao(self):
        """Status calculado só a partir de dados da própria Gestão Operação — sem
        depender de ItemPedido/estação (isso é Gestão Produção, mundo à parte).

        Regras:
          - OTD já registrado (SIM ou NÃO)             -> "FINALIZADO"
          - Já tem alguma data de PCP ou logística      -> "ANDAMENTO"
          - Só tem dado comercial preenchido ainda      -> "PENDENTE"
        """
        if self.go_otd_realizado:
            return "FINALIZADO"
        if self.go_data_efetiva_liberacao_pcp or self.go_data_pedido_expedido or self.go_data_real_entrega:
            return "ANDAMENTO"
        return "PENDENTE"


class HistoricoAlteracao(db.Model):
    """Registro de auditoria: quem mudou o quê, quando, e qual era o valor antes.

    Gravado manualmente nos pontos onde o pedido/item é salvo (não usa eventos
    automáticos do SQLAlchemy) — assim cada rota escolhe exatamente quais
    campos importam registrar, sem "mágica" escondida.
    """

    __tablename__ = "historico_alteracoes"

    id = db.Column(db.Integer, primary_key=True)
    entidade_tipo = db.Column(db.String(20), nullable=False)  # "pedido" ou "item_pedido"
    entidade_id = db.Column(db.Integer, nullable=False)
    # pedido_id fica preenchido mesmo quando a mudança é de um item, para que o
    # histórico completo de um pedido (itens inclusos) seja uma única consulta.
    pedido_id = db.Column(db.Integer, db.ForeignKey("pedidos.id"), nullable=True)
    usuario_id = db.Column(db.Integer, db.ForeignKey("usuarios.id"), nullable=True)
    usuario_nome = db.Column(db.String(120), nullable=True)  # snapshot, sobrevive se o usuário for desativado
    campo = db.Column(db.String(60), nullable=False)
    valor_anterior = db.Column(db.Text, nullable=True)
    valor_novo = db.Column(db.Text, nullable=True)
    criado_em = db.Column(db.DateTime, default=datetime.utcnow)
