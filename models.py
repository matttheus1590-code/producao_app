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

# Status de CHÃO DE FÁBRICA — usado só no Kanban da tela Estações. Pedido do
# Bruno (25/08/2026): só 3 colunas, lidas DIRETO de status_producao (o mesmo
# campo "confiável" que a Listagem Geral usa) — antes esse status era
# calculado a partir de inicio_inspecao/termino_inspecao/liberacao_faturamento,
# datas que na prática quase nunca são preenchidas (a planilha de referência
# nem tem essas colunas), então itens já FINALIZADO na Listagem apareciam como
# "Pendente" no Kanban. Ver ItemPedido.status_chao (abaixo) — não duplica
# nenhum campo novo, só lê status_producao com um rótulo diferente.
STATUS_CHAO_OPCOES = ["PENDENTE", "EM_PRODUCAO", "FINALIZADO"]
STATUS_CHAO_LABELS = {
    "PENDENTE": "Pendente",
    "EM_PRODUCAO": "Em produção",
    "FINALIZADO": "Finalizado",
}
STATUS_CHAO_CORES = {
    "PENDENTE": "secondary",
    "EM_PRODUCAO": "primary",
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

# Pedido do Bruno (01/09/2026): na tela de Novo Pedido (Gestão Produção),
# "Tipo de pedido" e "Status do pedido (informações)" — dois campos
# comerciais de Gestão Operação preenchidos já na inclusão — deixam de ser
# texto livre e viram listas fechadas, pra padronizar o que é digitado (e,
# no caso do status, colorir igual às outras badges de status do sistema).
GO_TIPO_PEDIDO_OPCOES = ["Pedido Padrão", "Reforma", "Serviço", "Emergencial"]

GO_STATUS_PEDIDO_INFO_OPCOES = ["OK", "EM TRATATIVA", "PENDENTE"]

GO_STATUS_PEDIDO_INFO_CORES = {
    "OK": "success",
    "EM TRATATIVA": "warning",
    "PENDENTE": "danger",
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

# ---------------------------------------------------------------------------
# Qualidade — RNC (Relatório de Não Conformidade). Vieram da planilha
# "Controle RNC" que o Bruno enviou (aba "Listas"), servem só como sugestão
# nos <select> dos formulários — o campo no banco é texto livre (String), não
# um Enum, porque a própria planilha original do Bruno já tinha valores fora
# dessas listas em alguns RNCs reais (ex.: "Documental (OP)" no Tipo de NC).
# Controle 100% manual, então o sistema nunca deve travar/rejeitar um valor
# digitado que não esteja aqui — só oferecer como atalho.
# ---------------------------------------------------------------------------
RNC_EMITENTE_OPCOES = ["Qualidade", "Engenharia", "Produção"]
RNC_SETOR_OPCOES = ["Qualidade", "Engenharia", "Produção", "Metrologia", "Almoxarifado"]
RNC_ORIGEM_OPCOES = [
    "Inspeção Interna",
    "Reclamação de Cliente",
    "Auditoria Interna",
    "Auditoria Externa",
    "Reunião Análise Crítica da Direção",
]
RNC_LOCAL_SETOR_OPCOES = [
    "PU", "Usinagem", "Montagem", "Solda", "Pintura", "Almoxarifado",
    "Expedição", "Recebimento", "Metrologia", "Injeção", "PCP", "Mecânica", "Espumagem",
]
RNC_TIPO_NC_OPCOES = [
    "Dimensional",
    "Dureza / Material",
    "Visual / Estética",
    "Funcional",
    "Documentação",
    "Embalagem",
    "Identificação / Rastreabilidade",
    "Processo",
    "Fornecedor",
    "Outro",
]
RNC_SEVERIDADE_OPCOES = ["Crítica", "Maior", "Menor"]
RNC_FERRAMENTA_ANALISE_OPCOES = [
    "5 Porquês", "Ishikawa (Espinha de Peixe)", "Pareto", "FMEA", "Diagrama de Dispersão", "Outro",
]
RNC_DISPOSICAO_OPCOES = [
    "Retrabalhar", "Reparar", "Uso sob Concessão", "Devolução ao Fornecedor", "Sucatear", "Reciclar",
]
RNC_STATUS_ACAO_OPCOES = ["Não Iniciada", "Em Andamento", "Concluída", "Atrasada"]
RNC_EFICACIA_OPCOES = ["Eficaz", "Não Eficaz", "Pendente"]
RNC_SIM_NAO_OPCOES = ["Sim", "Não"]
RNC_STATUS_GERAL_OPCOES = [
    "Aberto", "Em Análise", "Em Ação Corretiva", "Aguardando Verificação", "Fechado", "Cancelado",
]
RNC_STATUS_GERAL_ABERTOS = ["Aberto", "Em Análise", "Em Ação Corretiva", "Aguardando Verificação"]

RNC_SEVERIDADE_CORES = {"Crítica": "danger", "Maior": "warning", "Menor": "secondary"}
RNC_STATUS_GERAL_CORES = {
    "Aberto": "danger",
    "Em Análise": "warning",
    "Em Ação Corretiva": "info",
    "Aguardando Verificação": "primary",
    "Fechado": "success",
    "Cancelado": "secondary",
}
RNC_EFICACIA_CORES = {"Eficaz": "success", "Não Eficaz": "danger", "Pendente": "secondary"}

# ---------------------------------------------------------------------------
# Qualidade — Inspeção Final / RDIM (Relatório de Inspeção Dimensional e
# Dureza), pedido do Bruno em 02/09/2026. Digitaliza a inspeção final hoje
# feita em planilha/papel nas estações MANDRIL, PU e SILICONE. Diferente da
# RNC (que não tem FK com nada), a InspecaoFinal usa FK real pra ItemPedido —
# "OP", nesta base, É o próprio ItemPedido (o sistema nunca teve um número de
# Ordem de Produção separado; confirmado por busca completa no código antes
# de desenhar esta área).
# ---------------------------------------------------------------------------
RDIM_ESTACOES_OPCOES = ["MANDRIL", "PU", "SILICONE"]

RDIM_RESULTADO_OPCOES = ["APROVADO", "REPROVADO", "APROVADO_COM_DESVIO"]
RDIM_RESULTADO_LABELS = {
    "APROVADO": "Aprovado",
    "REPROVADO": "Reprovado",
    "APROVADO_COM_DESVIO": "Aprovado com desvio",
}
RDIM_RESULTADO_CORES = {"APROVADO": "success", "REPROVADO": "danger", "APROVADO_COM_DESVIO": "warning"}

RDIM_INSPECAO_VISUAL_OPCOES = ["OK", "Não OK"]

# Nasce da própria distinção que o modelo oficial já faz na aprovação com
# desvio ("Estético" vs. "Dimensional/Dureza") — é o que alimenta "principais
# tipos de desvio" e "ranking de causas" no dashboard.
RDIM_CATEGORIA_DESVIO_OPCOES = ["Dimensional", "Dureza", "Estético/Visual", "Outro"]

# Subcategoria — pedido do Bruno (02/09/2026, depois de já usar a área):
# mais granular que a categoria acima, aponta exatamente QUAL característica
# desviou (alimenta "principais características do desvio" no dashboard).
RDIM_SUBCATEGORIA_DESVIO_OPCOES = [
    "Espessura", "Diâmetro Externo", "Diâmetro Interno", "Comprimento", "Deformação/Ranhura",
]

# Sugestões pré-preenchidas na tela de Nova Inspeção — o operador pode editar,
# remover ou adicionar outras (grandeza é texto livre em RdimMedicao), então
# esta lista não trava nada, só acelera o preenchimento.
RDIM_GRANDEZAS_PADRAO = [
    "Diâmetro Externo (mm)",
    "Diâmetro Interno (mm)",
    "Espessura/Altura (mm)",
    "Dureza Shore A",
]

# ---------------------------------------------------------------------------
# P&D — Pesquisa e Desenvolvimento (Fase 14, 01/09/2026). Nova área pedida
# pelo Bruno: "Central de Gestão de Projetos de Desenvolvimento, Inovação e
# Melhoria", usada principalmente por ele (Coordenador de Operações
# Industriais) e por Gustavo Fugita (Líder de P&D/Desenvolvimento).
# ---------------------------------------------------------------------------
PD_CATEGORIA_OPCOES = [
    "Inovação",
    "Redução de custos",
    "Desenvolvimento de matéria-prima",
    "Desenvolvimento de produto",
    "Desenvolvimento tecnológico",
    "Desenvolvimento de fornecedor",
    "Melhoria de processo",
    "Melhoria de qualidade",
    "Novas soluções",
    "Projetos específicos para clientes",
]

# Ciclo de vida padrão pedido pelo Bruno — pode "andar pra trás" (ex.: Teste
# reprovado -> volta pra Desenvolvimento -> novo Teste), o que é permitido
# livremente (etapa_atual é só um texto de lista, não uma máquina de estados
# travada) — cada mudança de etapa gera uma linha em HistoricoAlteracao.
PD_ETAPA_OPCOES = [
    "Ideia", "Planejamento", "Desenvolvimento", "Teste",
    "Validação", "Homologação", "Implementação", "Concluído",
]
PD_ETAPA_CORES = {
    "Ideia": "secondary",
    "Planejamento": "info",
    "Desenvolvimento": "primary",
    "Teste": "warning",
    "Validação": "warning",
    "Homologação": "info",
    "Implementação": "primary",
    "Concluído": "success",
}

# "Resultado esperado" é multi-seleção (o Bruno listou várias opções que um
# projeto pode combinar ao mesmo tempo) — guardado como texto único (lista
# separada por "; "), no mesmo espírito dos campos de texto livre do resto
# do app: simples de somar/buscar, sem precisar de tabela auxiliar nova.
PD_RESULTADO_ESPERADO_OPCOES = [
    "Redução de custo",
    "Aumento de produtividade",
    "Melhoria de qualidade",
    "Redução de desperdício",
    "Redução de lead time",
    "Ganho tecnológico",
    "Novo produto",
    "Substituição de matéria-prima",
    "Desenvolvimento de fornecedor",
    "Outros",
]

PD_TIPO_EVENTO_OPCOES = [
    "Visita a fornecedor", "Visita a cliente", "Reunião técnica",
    "Reunião interna", "Validação", "Homologação",
]

# Status de cada Teste — cores e emoji seguem exatamente o que o Bruno pediu
# (🟡🔵🟣🟢🔴🟠), usados juntos (emoji + badge) nos templates.
PD_TESTE_RESULTADO_OPCOES = [
    "Planejado", "Agendado", "Realizado", "Aprovado", "Reprovado", "Novo teste necessário",
]
PD_TESTE_RESULTADO_INFO = {
    "Planejado": {"emoji": "🟡", "cor": "warning"},
    "Agendado": {"emoji": "🔵", "cor": "info"},
    "Realizado": {"emoji": "🟣", "cor": "primary"},
    "Aprovado": {"emoji": "🟢", "cor": "success"},
    "Reprovado": {"emoji": "🔴", "cor": "danger"},
    "Novo teste necessário": {"emoji": "🟠", "cor": "warning"},
}

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

    # Carimbo de "última alteração" — o SQLAlchemy atualiza sozinho
    # (onupdate) toda vez que o item é salvo, seja editando o pedido ou
    # clicando em "Avançar" no Kanban das Estações. Usado pra ordenar o
    # Kanban sempre com os itens mais novos/recém movimentados no topo de
    # cada coluna, sem precisar de nenhuma lógica extra no código de rota.
    atualizado_em = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

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
        """Status de chão de fábrica pro Kanban da tela Estações — só 3
        colunas, lidas direto de status_producao (mesmo campo que a Listagem
        Geral usa, então os dois nunca mais divergem). "EM TRATATIVA" entra
        junto com "ANDAMENTO" na coluna "Em produção" — não tem coluna própria
        pra ele no Kanban."""
        if self.status_producao == "FINALIZADO":
            return "FINALIZADO"
        if self.status_producao == "PENDENTE":
            return "PENDENTE"
        return "EM_PRODUCAO"

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
    def prazo_total_dias_corridos(self):
        """Prazo total combinado com o cliente, em dias corridos: da inclusão
        do pedido até a data solicitada de entrega (bloco Comercial) — usado
        na tela "Listagem Geral" de Gestão Operação. Não é o mesmo que
        go_dias_atraso_antecipacao (que compara solicitado x entregue de
        verdade, lá em Resultados) — este é só o prazo combinado em si."""
        if self.data_inclusao_pedido and self.go_data_solicitada_entrega:
            return (self.go_data_solicitada_entrega - self.data_inclusao_pedido).days
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


class RncQualidade(db.Model):
    """Qualidade — RNC (Relatório de Não Conformidade). Réplica da planilha
    "Controle RNC" do Bruno (aba "Controle RNC - Agosto" + aba "Dashboard",
    calculada aqui a partir destas linhas). TOTALMENTE INDEPENDENTE de
    Pedido/ItemPedido e de PedidoOperacao — tabela própria, id próprio, sem
    FK com nenhuma outra tabela, mesmo padrão de PedidoOperacao. Controle
    100% manual (Bruno: "controle manual com entrada de números e textos") —
    nenhum campo aqui é calculado a partir de Gestão Produção/Operação.

    `numero_rnc` é texto (não sequencial automático) porque a planilha usa um
    formato "NN/AA" (ex.: "07/26") escolhido por quem abre o RNC, e o Bruno
    pode querer manter essa numeração ao migrar RNCs futuros.
    """

    __tablename__ = "rnc_qualidade"

    id = db.Column(db.Integer, primary_key=True)

    # ---------------- Identificação ----------------
    numero_rnc = db.Column(db.String(20), nullable=True)
    revisao = db.Column(db.Integer, nullable=True, default=0)
    data_emissao = db.Column(db.Date, nullable=True)
    emitente = db.Column(db.String(60), nullable=True)
    setor = db.Column(db.String(60), nullable=True)
    origem = db.Column(db.String(80), nullable=True)

    # ---------------- Origem da não conformidade ----------------
    cliente_projeto = db.Column(db.String(200), nullable=True)
    numero_pedido_contrato = db.Column(db.String(60), nullable=True)
    produto_equipamento = db.Column(db.String(200), nullable=True)
    numero_op = db.Column(db.String(30), nullable=True)
    local_setor = db.Column(db.String(60), nullable=True)
    data_identificacao = db.Column(db.Date, nullable=True)
    responsavel_identificacao = db.Column(db.String(120), nullable=True)

    # ---------------- Descrição da não conformidade ----------------
    descricao_nc = db.Column(db.Text, nullable=True)
    qtd_nao_conforme = db.Column(db.Integer, nullable=True)
    requisito_nao_atendido = db.Column(db.String(200), nullable=True)
    tipo_nc = db.Column(db.String(60), nullable=True)
    severidade = db.Column(db.String(20), nullable=True)
    acao_contencao_imediata = db.Column(db.Text, nullable=True)

    # ---------------- Análise de causa raiz (5 Porquês) ----------------
    porque_1 = db.Column(db.Text, nullable=True)
    porque_2 = db.Column(db.Text, nullable=True)
    porque_3 = db.Column(db.Text, nullable=True)
    porque_4 = db.Column(db.Text, nullable=True)
    porque_5 = db.Column(db.Text, nullable=True)
    causa_raiz = db.Column(db.Text, nullable=True)
    ferramenta_analise = db.Column(db.String(60), nullable=True)
    disposicao_produto = db.Column(db.String(60), nullable=True)

    # ---------------- Ação corretiva ----------------
    acao_corretiva_descricao = db.Column(db.Text, nullable=True)
    responsavel_acao_corretiva = db.Column(db.String(120), nullable=True)
    prazo_acao_corretiva = db.Column(db.Date, nullable=True)
    data_realizacao = db.Column(db.Date, nullable=True)
    status_acao_corretiva = db.Column(db.String(30), nullable=True)

    # ---------------- Verificação de eficácia ----------------
    data_verificacao_eficacia = db.Column(db.Date, nullable=True)
    eficacia_acao = db.Column(db.String(20), nullable=True)
    obs_verificacao = db.Column(db.Text, nullable=True)
    reincidencia = db.Column(db.String(10), nullable=True)  # "Sim" / "Não"
    numero_rnc_relacionada = db.Column(db.String(20), nullable=True)

    # ---------------- Encerramento ----------------
    custo_estimado = db.Column(db.Float, nullable=True)
    status_geral = db.Column(db.String(30), nullable=True, default="Aberto")
    responsavel_qualidade = db.Column(db.String(120), nullable=True)
    data_fechamento = db.Column(db.Date, nullable=True)
    evidencias_anexos = db.Column(db.Text, nullable=True)
    observacoes_gerais = db.Column(db.Text, nullable=True)

    criado_por_id = db.Column(db.Integer, db.ForeignKey("usuarios.id"), nullable=True)
    criado_por = db.relationship("Usuario", foreign_keys=[criado_por_id])
    criado_em = db.Column(db.DateTime, default=datetime.utcnow)
    atualizado_em = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    @property
    def dias_em_aberto(self):
        """Dias desde a identificação até o fechamento — ou até hoje, se ainda
        aberto. Calculado (não guardado) para nunca ficar desatualizado
        enquanto o RNC segue em aberto, ao contrário da planilha original
        (coluna "Dias em Aberto" fixa, congelada no dia da exportação)."""
        if not self.data_identificacao:
            return None
        fim = self.data_fechamento or date.today()
        return (fim - self.data_identificacao).days

    @property
    def esta_aberto(self):
        return (self.status_geral or "Aberto") in RNC_STATUS_GERAL_ABERTOS


class InspecaoFinal(db.Model):
    """Qualidade — Inspeção Final / RDIM (pedido do Bruno, 02/09/2026).

    Diferente da RncQualidade (que não tem FK com nada), esta tabela usa FK
    real pra ItemPedido: cada registro É a inspeção final de um item de
    pedido — "OP", nesta base, é o próprio ItemPedido, já que o sistema nunca
    teve um número de Ordem de Produção separado. Cliente/Nº Pedido/Produto
    nunca são duplicados aqui — sempre lidos através de `item` (properties
    abaixo)."""

    __tablename__ = "inspecoes_finais"

    id = db.Column(db.Integer, primary_key=True)

    item_pedido_id = db.Column(db.Integer, db.ForeignKey("itens_pedido.id"), nullable=False)
    item = db.relationship("ItemPedido", foreign_keys=[item_pedido_id])

    # Copiada do item no momento da criação — não é recalculada depois (se o
    # item mudar de estação mais tarde, esta inspeção continua contando a
    # estação de quando foi realmente inspecionada).
    estacao = db.Column(db.String(40), nullable=True)

    data_inspecao = db.Column(db.Date, nullable=True)
    responsavel_id = db.Column(db.Integer, db.ForeignKey("usuarios.id"), nullable=True)
    responsavel = db.relationship("Usuario", foreign_keys=[responsavel_id])

    numero_rif = db.Column(db.String(30), nullable=True)
    procedimento = db.Column(db.String(60), nullable=True)
    norma = db.Column(db.String(60), nullable=True)
    instrucao_trabalho = db.Column(db.String(60), nullable=True)

    inspecao_visual = db.Column(db.String(20), nullable=True)
    desvio_encontrado = db.Column(db.Text, nullable=True)
    categoria_desvio = db.Column(db.String(30), nullable=True)
    # Subcategoria (característica específica que desviou) — pedido do Bruno
    # (02/09/2026), mais granular que categoria_desvio. Coluna nova em tabela
    # que já existe em produção -> precisa de _migrar_rdim_inspecao_final em
    # app.py (ALTER TABLE), não é coberta só por db.create_all().
    subcategoria_desvio = db.Column(db.String(40), nullable=True)
    observacao = db.Column(db.Text, nullable=True)
    resultado = db.Column(db.String(24), nullable=True)

    # Quantitativo de peças do lote (item.quantidade) que apresentaram
    # desvio — pedido do Bruno (02/09/2026): "lote total contém 5 peças, mas
    # dessas 2 unidades ficou com desvio". Sempre em relação ao lote inteiro
    # (decisão confirmada com Bruno) — não existe campo separado de
    # "quantidade inspecionada"; a % de desvio é sempre quantidade_com_desvio
    # / item.quantidade. Float (não Integer) pra bater com o tipo de
    # ItemPedido.quantidade.
    quantidade_com_desvio = db.Column(db.Float, nullable=True)

    criado_por_id = db.Column(db.Integer, db.ForeignKey("usuarios.id"), nullable=True)
    criado_por = db.relationship("Usuario", foreign_keys=[criado_por_id])
    criado_em = db.Column(db.DateTime, default=datetime.utcnow)
    atualizado_em = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    medicoes = db.relationship(
        "RdimMedicao",
        backref="inspecao",
        cascade="all, delete-orphan",
        order_by="RdimMedicao.ordem",
    )

    # Detalhamento peça a peça do desvio — pedido do Bruno (02/09/2026, RDIM
    # Fase 3): "quantidade_com_desvio" acima é só o total do lote; aqui fica
    # o apontamento individual (nº da peça + característica + valor medido
    # daquela peça específica). Tabela nova (não coluna em tabela
    # existente) -> criada sozinha pelo db.create_all(), sem precisar de
    # migração em app.py, mesmo caso de RdimMedicao na Fase 1.
    pecas_desvio = db.relationship(
        "RdimPecaDesvio",
        backref="inspecao",
        cascade="all, delete-orphan",
        order_by="RdimPecaDesvio.ordem",
    )

    @property
    def pedido(self):
        return self.item.pedido if self.item else None

    @property
    def cliente(self):
        return self.pedido.cliente if self.pedido else None

    @property
    def pedido_venda(self):
        return self.pedido.pedido_venda if self.pedido else None

    @property
    def produto(self):
        return self.item.descricao_produto if self.item else None

    @property
    def tem_desvio_fora_tolerancia(self):
        """True se alguma medição registrada estourou a especificação — usado
        pra sinalização visual e pros KPIs do dashboard, sempre recalculado a
        partir das medições reais (nunca guardado como campo à parte)."""
        return any(m.dentro_da_tolerancia is False for m in self.medicoes)

    @property
    def resumo_pecas_desvio(self):
        """Texto curto tipo "1: Espessura; 2: Espessura; 3: Diâmetro
        Externo" a partir de pecas_desvio — usado só como tooltip na
        listagem RDIM, pra dar uma prévia do detalhamento peça a peça sem
        precisar abrir a inspeção."""
        partes = []
        for p in self.pecas_desvio:
            rotulo = p.peca_numero or "—"
            if p.caracteristica:
                rotulo += ": " + p.caracteristica
            partes.append(rotulo)
        return "; ".join(partes)

    @property
    def contexto_pecas_desvio(self):
        """Texto multi-linha com o "contexto do desvio" de cada peça
        apontada (espec. x medido x variação) — pedido do Bruno (02/09/2026,
        RDIM Fase 4): "medida solicitada X junto com a medida inspecionada X
        junto com a variação que teve da peça". Usado como tooltip completo
        na listagem RDIM (complementa o resumo compacto de
        pior_apontamento_peca, que mostra só a pior linha na própria
        célula)."""
        linhas = []
        for p in self.pecas_desvio:
            rotulo = p.peca_numero or "—"
            if p.caracteristica:
                rotulo += ": " + p.caracteristica
            if p.especificado_min is not None or p.especificado_max is not None:
                espec = f"{p.especificado_min if p.especificado_min is not None else '—'} a {p.especificado_max if p.especificado_max is not None else '—'}"
            else:
                espec = "—"
            medido = p.valor_medido if p.valor_medido is not None else "—"
            if p.variacao is None:
                variacao_txt = ""
            elif p.variacao == 0:
                variacao_txt = " (dentro da tolerância)"
            elif p.acima_da_tolerancia:
                variacao_txt = f" (+{p.variacao} acima da tolerância)"
            else:
                variacao_txt = f" (-{p.variacao} abaixo da tolerância)"
            linhas.append(f"{rotulo} — espec. {espec} x medido {medido}{variacao_txt}")
        return "; ".join(linhas)

    @property
    def pior_apontamento_peca(self):
        """A linha de pecas_desvio com a MAIOR variação em relação à
        tolerância (RdimPecaDesvio.variacao) — pedido do Bruno (02/09/2026,
        RDIM Fase 4): na listagem de inspeções, mostrar o "contexto do
        desvio" (medida solicitada x medida inspecionada x variação),
        exemplo dele: "tolerância era de 0,5mm, peça inspecionada com
        0,7mm, peça ficou 0,2mm acima da tolerância". Com várias peças
        apontadas numa mesma inspeção, mostra a pior (maior variacao) como
        resumo rápido na linha da tabela; o detalhamento completo continua
        disponível ao abrir a inspeção. None se não há nenhum apontamento
        com variação calculável (falta espec. ou valor medido)."""
        pior = None
        for p in self.pecas_desvio:
            v = p.variacao
            if v is None:
                continue
            if pior is None or v > pior.variacao:
                pior = p
        return pior


class RdimMedicao(db.Model):
    """Uma linha = uma grandeza medida dentro de uma InspecaoFinal (diâmetro
    externo, diâmetro interno, espessura/altura, dureza Shore A, ou qualquer
    outra que o inspetor adicionar). Grandezas flexíveis por inspeção — não 4
    colunas fixas — porque Mandril, PU e Silicone medem coisas diferentes."""

    __tablename__ = "rdim_medicoes"

    id = db.Column(db.Integer, primary_key=True)
    inspecao_final_id = db.Column(db.Integer, db.ForeignKey("inspecoes_finais.id"), nullable=False)

    grandeza = db.Column(db.String(60), nullable=False)
    especificado_min = db.Column(db.Float, nullable=True)
    especificado_max = db.Column(db.Float, nullable=True)
    # Faixa (mín/máx) encontrada no lote inspecionado — igual ao modelo
    # oficial, não um valor único.
    medido_min = db.Column(db.Float, nullable=True)
    medido_max = db.Column(db.Float, nullable=True)
    ordem = db.Column(db.Integer, nullable=False, default=0)

    @property
    def dentro_da_tolerancia(self):
        """True/False, ou None quando falta dado suficiente pra comparar (sem
        especificação, ou sem medição ainda) — usado pra sinalização visual
        na tela e pros KPIs do dashboard. Nunca guardado como coluna: sempre
        recalculado a partir de especificado_*/medido_* para não poder ficar
        desatualizado."""
        if self.especificado_min is None and self.especificado_max is None:
            return None
        if self.medido_min is None and self.medido_max is None:
            return None
        if self.especificado_min is not None:
            if self.medido_min is not None and self.medido_min < self.especificado_min:
                return False
            if self.medido_max is not None and self.medido_max < self.especificado_min:
                return False
        if self.especificado_max is not None:
            if self.medido_max is not None and self.medido_max > self.especificado_max:
                return False
            if self.medido_min is not None and self.medido_min > self.especificado_max:
                return False
        return True


class RdimPecaDesvio(db.Model):
    """Uma linha = uma peça com desvio numa característica específica,
    dentro de uma InspecaoFinal — pedido do Bruno (02/09/2026, RDIM Fase 3):
    "lote total contém 10 peças, mas 5 tiveram desvio na espessura" não era
    detalhado o suficiente; ele quer registrar a variação peça a peça. Uma
    peça com desvio em 2 características vira 2 linhas (decisão confirmada
    com o Bruno) — mesmo espírito de "grandezas flexíveis" de RdimMedicao,
    só que por peça em vez de por lote. NÃO recalcula nem substitui
    quantidade_com_desvio (InspecaoFinal): esse continua sendo digitado à
    mão, separado — por decisão dele, os dois números podem não bater 1:1,
    esta lista é só um detalhamento complementar."""

    __tablename__ = "rdim_pecas_desvio"

    id = db.Column(db.Integer, primary_key=True)
    inspecao_final_id = db.Column(db.Integer, db.ForeignKey("inspecoes_finais.id"), nullable=False)

    # Número/identificação da peça dentro do lote (ex.: 1, 2, 3...) — texto
    # livre (não Integer) pra também caber uma identificação tipo "peça 3B"
    # ou uma tag/serial, caso o Bruno numere as peças de outro jeito.
    peca_numero = db.Column(db.String(20), nullable=True)
    caracteristica = db.Column(db.String(40), nullable=True)
    valor_medido = db.Column(db.Float, nullable=True)

    # Especificação (tolerância) daquela característica, na própria peça —
    # pedido do Bruno (02/09/2026): "tolerância era de 0,5mm, peça
    # inspecionada com 0,7mm, peça ficou 0,2mm acima da tolerância". Campos
    # novos numa tabela que já existe em produção (rdim_pecas_desvio criada
    # na Fase 3) -> precisa de _migrar_rdim_pecas_desvio em app.py (ALTER
    # TABLE), mesmo caso de subcategoria_desvio/quantidade_com_desvio antes.
    especificado_min = db.Column(db.Float, nullable=True)
    especificado_max = db.Column(db.Float, nullable=True)

    ordem = db.Column(db.Integer, nullable=False, default=0)

    @property
    def variacao(self):
        """Quanto o valor medido ficou fora da tolerância — sempre positivo
        (não importa se estourou pra cima ou pra baixo), None quando não dá
        pra calcular (falta valor medido, ou nenhuma especificação
        informada). Ex.: espec. máx 0,5mm, medido 0,7mm -> 0,2 (peça ficou
        0,2mm ACIMA da tolerância). Dentro da tolerância -> 0.0. Nunca
        guardado como coluna: sempre recalculado a partir de
        especificado_*/valor_medido, mesmo espírito de
        RdimMedicao.dentro_da_tolerancia."""
        if self.valor_medido is None:
            return None
        if self.especificado_min is None and self.especificado_max is None:
            return None
        if self.especificado_max is not None and self.valor_medido > self.especificado_max:
            return round(self.valor_medido - self.especificado_max, 4)
        if self.especificado_min is not None and self.valor_medido < self.especificado_min:
            return round(self.especificado_min - self.valor_medido, 4)
        return 0.0

    @property
    def acima_da_tolerancia(self):
        """True se valor_medido > especificado_max, False se < especificado_min
        (abaixo), None se dentro da tolerância ou sem dado suficiente —
        usado só pra escolher a seta/sinal no texto ("+0,2mm acima" vs
        "-0,2mm abaixo")."""
        if self.valor_medido is None:
            return None
        if self.especificado_max is not None and self.valor_medido > self.especificado_max:
            return True
        if self.especificado_min is not None and self.valor_medido < self.especificado_min:
            return False
        return None


class ProjetoPD(db.Model):
    """P&D — Central de Gestão de Projetos de Desenvolvimento, Inovação e
    Melhoria (Fase 14, pedido do Bruno em 01/09/2026). TOTALMENTE
    INDEPENDENTE de Pedido/ItemPedido/PedidoOperacao/RncQualidade — tabela
    própria, sem FK com nenhuma outra área do sistema, mesmo espírito de
    RncQualidade (controle manual, campos de texto livre onde a planilha do
    Bruno usava texto livre, sem inventar cadastros de Cliente/Produto/
    Fornecedor que não existem em nenhum outro lugar do app).

    "código" não é sequencial automático (mesma lógica do numero_rnc do
    RNC): texto livre, pra deixar o Bruno/Gustavo escolherem o formato deles.
    """

    __tablename__ = "projetos_pd"

    id = db.Column(db.Integer, primary_key=True)

    # ---------------- Informações gerais ----------------
    codigo = db.Column(db.String(30), nullable=True)
    nome = db.Column(db.String(200), nullable=False)
    descricao = db.Column(db.Text, nullable=True)
    objetivo = db.Column(db.Text, nullable=True)
    justificativa = db.Column(db.Text, nullable=True)
    categoria = db.Column(db.String(60), nullable=True)
    prioridade = db.Column(db.String(10), nullable=True)  # BAIXA/MÉDIA/ALTA — mesma lista de Pedido
    responsavel = db.Column(db.String(120), nullable=True)
    participantes = db.Column(db.Text, nullable=True)
    cliente = db.Column(db.String(200), nullable=True)
    produto = db.Column(db.String(200), nullable=True)
    fornecedor = db.Column(db.String(200), nullable=True)
    area_envolvida = db.Column(db.String(120), nullable=True)

    # ---------------- Cronograma ----------------
    etapa_atual = db.Column(db.String(30), nullable=False, default="Ideia")
    percentual_conclusao = db.Column(db.Integer, nullable=True, default=0)
    data_inicio = db.Column(db.Date, nullable=True)
    data_prevista_conclusao = db.Column(db.Date, nullable=True)
    data_real_conclusao = db.Column(db.Date, nullable=True)
    proxima_entrega = db.Column(db.String(200), nullable=True)
    data_proxima_entrega = db.Column(db.Date, nullable=True)
    responsavel_proxima_entrega = db.Column(db.String(120), nullable=True)

    # ---------------- Financeiro ----------------
    custo_previsto = db.Column(db.Float, nullable=True)
    custo_realizado = db.Column(db.Float, nullable=True)
    investimento_previsto = db.Column(db.Float, nullable=True)
    investimento_realizado = db.Column(db.Float, nullable=True)
    economia_prevista = db.Column(db.Float, nullable=True)
    economia_realizada = db.Column(db.Float, nullable=True)

    # ---------------- Resultado esperado / obtido ----------------
    resultado_esperado = db.Column(db.Text, nullable=True)  # lista separada por "; "
    resultado_obtido = db.Column(db.Text, nullable=True)

    # ---------------- Base de conhecimento (preenchido ao concluir) ----------------
    problema = db.Column(db.Text, nullable=True)
    solucao = db.Column(db.Text, nullable=True)
    licoes_aprendidas = db.Column(db.Text, nullable=True)

    observacoes_gerais = db.Column(db.Text, nullable=True)

    criado_por_id = db.Column(db.Integer, db.ForeignKey("usuarios.id"), nullable=True)
    criado_por = db.relationship("Usuario", foreign_keys=[criado_por_id])
    criado_em = db.Column(db.DateTime, default=datetime.utcnow)
    atualizado_em = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    testes = db.relationship(
        "TesteProjetoPD", backref="projeto", order_by="TesteProjetoPD.id", cascade="all, delete-orphan",
    )
    eventos = db.relationship(
        "VisitaReuniaoPD", backref="projeto", order_by="VisitaReuniaoPD.data.desc()", cascade="all, delete-orphan",
    )

    @property
    def concluido(self):
        return self.etapa_atual == "Concluído"

    @property
    def dias_restantes(self):
        if not self.data_prevista_conclusao or self.concluido:
            return None
        return (self.data_prevista_conclusao - date.today()).days

    @property
    def atrasado(self):
        """Mesma regra usada em todo o resto do sistema: só "atrasado" se
        tiver data prevista, ainda não tiver concluído, e a data já tiver
        passado."""
        dias = self.dias_restantes
        return dias is not None and dias < 0

    @property
    def prazo_proximo(self):
        dias = self.dias_restantes
        return dias is not None and 0 <= dias <= PRAZO_ALERTA_DIAS

    @property
    def critico(self):
        return self.atrasado and self.prioridade == "ALTA"

    @property
    def roi_percentual(self):
        """(economia - investimento) / investimento, em %. None se não der
        pra calcular (sem investimento realizado lançado ainda)."""
        if not self.investimento_realizado:
            return None
        economia = self.economia_realizada or 0
        return round(((economia - self.investimento_realizado) / self.investimento_realizado) * 100, 1)

    @property
    def resultado_esperado_lista(self):
        return [r.strip() for r in (self.resultado_esperado or "").split(";") if r.strip()]


class TesteProjetoPD(db.Model):
    """Um projeto de P&D pode ter vários testes/validações (ex.: "Teste
    reprovado -> novo teste"). Cada linha é um teste independente — o
    histórico de tentativas fica visível listando todas as linhas do
    projeto, sem sobrescrever nada."""

    __tablename__ = "testes_projeto_pd"

    id = db.Column(db.Integer, primary_key=True)
    projeto_id = db.Column(db.Integer, db.ForeignKey("projetos_pd.id"), nullable=False)

    numero = db.Column(db.String(30), nullable=True)
    data_planejada = db.Column(db.Date, nullable=True)
    data_realizada = db.Column(db.Date, nullable=True)
    responsavel = db.Column(db.String(120), nullable=True)
    material_utilizado = db.Column(db.String(200), nullable=True)
    lote = db.Column(db.String(60), nullable=True)
    fornecedor = db.Column(db.String(200), nullable=True)
    condicoes = db.Column(db.Text, nullable=True)
    resultado = db.Column(db.String(30), nullable=True, default="Planejado")
    observacoes = db.Column(db.Text, nullable=True)
    # Sem infraestrutura de upload de arquivo no sistema ainda (nenhuma área
    # do app tem hoje) — igual a RncQualidade.evidencias_anexos, fica como
    # referência/link em texto por enquanto; upload real de foto/documento
    # fica pra um próximo incremento.
    anexos = db.Column(db.Text, nullable=True)

    criado_em = db.Column(db.DateTime, default=datetime.utcnow)


class VisitaReuniaoPD(db.Model):
    """Visitas a fornecedores/clientes e reuniões técnicas/internas/
    validações/homologações ligadas a um projeto de P&D."""

    __tablename__ = "visitas_reunioes_pd"

    id = db.Column(db.Integer, primary_key=True)
    projeto_id = db.Column(db.Integer, db.ForeignKey("projetos_pd.id"), nullable=False)

    data = db.Column(db.Date, nullable=True)
    tipo = db.Column(db.String(40), nullable=True)
    participantes = db.Column(db.Text, nullable=True)
    local = db.Column(db.String(200), nullable=True)
    objetivo = db.Column(db.Text, nullable=True)
    resultado = db.Column(db.Text, nullable=True)
    proximas_acoes = db.Column(db.Text, nullable=True)
    responsavel = db.Column(db.String(120), nullable=True)
    anexos = db.Column(db.Text, nullable=True)

    criado_em = db.Column(db.DateTime, default=datetime.utcnow)


class ControleSistema(db.Model):
    """Tabela bem pequena, chave-valor, usada só como marcador de "isso já
    rodou uma vez" para migrações/importações que não podem usar o padrão
    "só roda se a tabela estiver vazia" (ex.: sincronizar uma planilha
    específica por cima de dados que já existem em produção — precisa rodar
    exatamente uma vez, mesmo com o banco já povoado)."""

    __tablename__ = "controle_sistema"

    id = db.Column(db.Integer, primary_key=True)
    chave = db.Column(db.String(80), unique=True, nullable=False)
    aplicado_em = db.Column(db.DateTime, default=datetime.utcnow)


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
