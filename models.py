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


class Usuario(UserMixin, db.Model):
    __tablename__ = "usuarios"

    id = db.Column(db.Integer, primary_key=True)
    nome = db.Column(db.String(120), nullable=False)
    username = db.Column(db.String(80), unique=True, nullable=False)
    senha_hash = db.Column(db.String(255), nullable=False)

    def set_senha(self, senha):
        self.senha_hash = generate_password_hash(senha)

    def checar_senha(self, senha):
        return check_password_hash(self.senha_hash, senha)


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

    # ------------- Dados PARAMETRIZADOS (evoluem durante a produção) -------------
    prioridade = db.Column(db.String(20), default="MÉDIA")
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

    rnc = db.Column(db.String(120), nullable=True)
    obs = db.Column(db.Text, nullable=True)

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
    def lt_comercial_dias(self):
        if self.data_cliente and self.data_inclusao_pedido:
            return (self.data_inclusao_pedido - self.data_cliente).days
        return None

    @property
    def tempo_espera_dias(self):
        if self.data_inclusao_pedido and self.inicio_producao:
            return (self.inicio_producao - self.data_inclusao_pedido).days
        return None

    @property
    def lt_producao_dias(self):
        if self.inicio_producao and self.termino_inspecao:
            return (self.termino_inspecao - self.inicio_producao).days
        return None

    @property
    def prazo_total_dias(self):
        if not self.data_cliente:
            return None
        fim = self.termino_inspecao or date.today()
        return (fim - self.data_cliente).days

    def atualizar_status_automatico(self):
        """Recalcula status_producao a partir das datas preenchidas.

        Regras:
          - Término de inspeção + liberação de faturamento preenchidos -> FINALIZADO
          - Início de produção preenchido (mas ainda sem término)      -> ANDAMENTO
          - Nada preenchido                                            -> PENDENTE
          - "EM TRATATIVA" é um estado manual: uma vez selecionado pelo
            usuário ele fica travado até o pedido ser finalizado.
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


class ItemPedido(db.Model):
    """Um produto dentro de um pedido. Um pedido pode ter vários itens."""

    __tablename__ = "itens_pedido"

    id = db.Column(db.Integer, primary_key=True)
    pedido_id = db.Column(db.Integer, db.ForeignKey("pedidos.id"), nullable=False)

    descricao_produto = db.Column(db.String(300), nullable=False)
    quantidade = db.Column(db.Float, nullable=False, default=0)
    custo_unitario = db.Column(db.Float, nullable=False, default=0)

    @property
    def valor_total(self):
        return round((self.quantidade or 0) * (self.custo_unitario or 0), 2)
