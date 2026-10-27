"""Relatório executivo (PDF) da Mini gestão de risco de prazos (pedido do
Bruno, 27/10/2026): "relatório de gestão de risco totalmente completo, para
a alta direção, focado em antecipar qualquer risco já com a tomada de decisão
orientada — CIF e FOB, 5, 10 e 20 dias, o mais compacto possível, com datas
e horários, logo, intuitivo e completo com todas as datas e regiões".

Módulo independente (não importa o app): recebe os dados já calculados por
_dados_gestao_risco_completa() e só apresenta. Nada de glyph fora do
Latin-1 (setas, ✓...) — a fonte padrão do reportlab os desenha como
quadrado preto; ícones vêm de PNGs (função `icone`)."""

import io
from collections import defaultdict
from datetime import timedelta
from xml.sax.saxutils import escape as esc

from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.shapes import Drawing
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    CondPageBreak, HRFlowable, Image, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)

NAVY = "#1b2a4a"
VERMELHO = "#c0392b"
AMARELO = "#d99a00"
CINZA = "#6c757d"
VERDE = "#198754"
AZUL = "#0d6efd"
BORDA = colors.HexColor("#dee2e6")

_DIAS_SEMANA = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]

_REGIAO_POR_UF = {
    **{u: "Sudeste" for u in ("SP", "RJ", "MG", "ES")},
    **{u: "Sul" for u in ("PR", "SC", "RS")},
    **{u: "Centro-Oeste" for u in ("DF", "GO", "MT", "MS")},
    **{u: "Nordeste" for u in ("BA", "SE", "AL", "PE", "PB", "RN", "CE", "PI", "MA")},
    **{u: "Norte" for u in ("AM", "PA", "AC", "RO", "RR", "AP", "TO")},
    "EX": "Exterior",
}

_ORDEM_CLASSE = {"atraso": 0, "sem_plano": 1, "risco": 2}
_ROTULO_CLASSE = {"atraso": "Atraso projetado", "sem_plano": "Sem plano PCP", "risco": "Em risco"}
_COR_CLASSE = {"atraso": VERMELHO, "sem_plano": CINZA, "risco": AMARELO}
_FUNDO_PRIO = {"atraso": "#f8d7da", "sem_plano": "#e2e3e5", "risco": "#fff3cd"}
_PRIO = {"atraso": "P1", "sem_plano": "P2", "risco": "P3"}


def _regiao(uf):
    return _REGIAO_POR_UF.get((uf or "").upper(), "—")


def _dm(d, semana=True):
    if not d:
        return "—"
    return f"{d.strftime('%d/%m/%Y')}" + (f" {_DIAS_SEMANA[d.weekday()]}" if semana else "")


def _dm_curto(d):
    return f"{d.strftime('%d/%m')} {_DIAS_SEMANA[d.weekday()]}" if d else "—"


def _faixa(dias_ate_prazo, horizontes):
    for h in horizontes:
        if dias_ate_prazo <= h:
            return h
    return horizontes[-1]


def _acao(l):
    """Ação orientada, curta, por pedido."""
    rev = l.get("revisao")
    lim = _dm_curto(l.get("limite_liberacao"))
    if l["classe"] == "atraso":
        if l["dias_ate_prazo"] < 0:
            base = "VENCIDO: acelerar a entrega e alinhar nova data com o cliente"
        else:
            base = f"Antecipar {-l['folga_dias']}d (liberar até {lim}) ou renegociar prazo"
        if rev:
            base = "Cobrar retorno do comercial (revisão já pedida)"
            if rev.novo_prazo_sugerido:
                base += f"; novo prazo sugerido {_dm_curto(rev.novo_prazo_sugerido)}"
        return base
    if l["classe"] == "sem_plano":
        parcial = l.get("itens_sem_previsao") and l.get("data_prevista_pcp")
        return ("Programar itens restantes no PCP" if parcial else "Programar no PCP") + f" — liberar até {lim}"
    return f"Blindar a fila: sem folga — liberar até {lim}"


# Só ícones com arte real (alguns PNGs de emoji_pdf são placeholders 24x24
# cinza que viram "quadradinho" no PDF).
_ICONES_OK = {"warning", "target", "calendar", "check", "red_circle", "yellow_circle", "green_circle",
              "hourglass", "magnifier", "bar_chart", "factory", "funnel", "clipboard"}


def gerar(dados, agora, logo_path, icone_base, moeda):
    def icone(nome, tamanho=9):
        return icone_base(nome, tamanho) if nome in _ICONES_OK else ""

    hoje = dados["hoje"]
    horizontes = tuple(dados["horizontes"])
    linhas = sorted(
        dados["linhas"],
        key=lambda l: (_ORDEM_CLASSE[l["classe"]], l["folga_dias"] if l["folga_dias"] is not None else 9999, l["dias_ate_prazo"]),
    )

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=landscape(A4), leftMargin=8 * mm, rightMargin=8 * mm, topMargin=8 * mm, bottomMargin=12 * mm,
        title="Gestão de risco de prazos — visão executiva",
    )
    largura = landscape(A4)[0] - doc.leftMargin - doc.rightMargin
    est = getSampleStyleSheet()
    base = ParagraphStyle("b", parent=est["Normal"], fontSize=7, leading=8.4)
    base_b = ParagraphStyle("bb", parent=base, fontName="Helvetica-Bold")
    peq = ParagraphStyle("peq", parent=base, fontSize=6.3, leading=7.6)
    nota = ParagraphStyle("nota", parent=base, fontSize=6.8, leading=8.6, textColor=colors.HexColor("#555"))
    centro = ParagraphStyle("c", parent=base, alignment=1)
    centro_b = ParagraphStyle("cb", parent=base_b, alignment=1)
    direita = ParagraphStyle("d", parent=base, alignment=2)
    cab_st = ParagraphStyle("cab", parent=base_b, fontSize=6.6, leading=7.8, textColor=colors.HexColor("#1b2a4a"))
    cab_c = ParagraphStyle("cabc", parent=cab_st, alignment=1)
    titulo = ParagraphStyle("t", parent=est["Title"], fontSize=19, leading=22, alignment=0, spaceAfter=0, textColor=colors.HexColor(NAVY))
    sub = ParagraphStyle("sub", parent=base, fontSize=8, leading=10.5, textColor=colors.HexColor("#444"))

    def P(t, st=None):
        return Paragraph(t if t is not None else "—", st or base)

    def T(t, st=None):
        return Paragraph(esc(str(t)) if t not in (None, "") else "—", st or base)

    def cor(t, hexc, b=True):
        return f'<font color="{hexc}">{"<b>" if b else ""}{t}{"</b>" if b else ""}</font>'

    def banner(txt, cor_bg, sub_txt=None, ic=None):
        st = ParagraphStyle("bn", parent=base_b, fontSize=10, leading=12, textColor=colors.white)
        t = Table([[P((icone(ic, 11) + " " if ic else "") + esc(txt), st)]], colWidths=[largura])
        t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(cor_bg)), ("TOPPADDING", (0, 0), (-1, -1), 3.2),
                               ("BOTTOMPADDING", (0, 0), (-1, -1), 3.2), ("LEFTPADDING", (0, 0), (-1, -1), 6)]))
        out = [CondPageBreak(38 * mm), t]
        if sub_txt:
            out += [Spacer(1, 0.8 * mm), P(sub_txt, nota)]
        out.append(Spacer(1, 1.2 * mm))
        return out

    def tabela(cab, rows, pesos, alin=None, extra=None, zebra=True, fonte_cab=None):
        alin = alin or "L" * len(cab)
        cabs = [P(c, cab_c if alin[i] == "C" else cab_st) for i, c in enumerate(cab)]
        t = Table([cabs] + rows, colWidths=[p / float(sum(pesos)) * largura for p in pesos], repeatRows=1)
        stl = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#dbe5f4")),
            ("LINEBELOW", (0, 0), (-1, 0), 0.8, colors.HexColor("#8fa8cc")),
            ("GRID", (0, 0), (-1, -1), 0.25, BORDA),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 1.7), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.7),
            ("LEFTPADDING", (0, 0), (-1, -1), 2.6), ("RIGHTPADDING", (0, 0), (-1, -1), 2.6),
        ]
        if zebra:
            for i in range(2, len(rows) + 1, 2):
                stl.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#f8f9fa")))
        stl += extra or []
        t.setStyle(TableStyle(stl))
        return t

    def agrega(lista):
        return {
            "n": len(lista),
            "atraso": sum(1 for l in lista if l["classe"] == "atraso"),
            "risco": sum(1 for l in lista if l["classe"] == "risco"),
            "sem_plano": sum(1 for l in lista if l["classe"] == "sem_plano"),
            "valor": sum(l["valor_aberto"] for l in lista),
            "rev": sum(1 for l in lista if l["revisao"] is not None),
            "a_solicitar": sum(1 for l in lista if l["revisao"] is None and l["classe"] in ("atraso", "sem_plano")),
        }

    el = []

    # ============================ CABEÇALHO ============================
    import os
    logo = Image(logo_path, width=50 * mm, height=50 * mm * (63 / 261)) if os.path.exists(logo_path) else P("")
    bloco = [
        P(f'{icone("warning", 17)} GESTÃO DE RISCO DE PRAZOS — VISÃO EXECUTIVA', titulo),
        Spacer(1, 1 * mm),
        P(f'{icone("clipboard", 10)} Emitido em <b>{agora.strftime("%d/%m/%Y")}</b> às <b>{agora.strftime("%H:%M")}</b> (horário de Brasília) · '
          f'{icone("calendar", 10)} Posição de <b>{_dm(hoje)}</b> · FOB e CIF · horizontes de <b>{" / ".join(str(h) for h in horizontes)} dias</b> '
          f'(pelo prazo do cliente; os já vencidos entram no 1º horizonte)', sub),
    ]
    cab = Table([[logo, bloco]], colWidths=[56 * mm, largura - 56 * mm])
    cab.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    el += [cab, Spacer(1, 1.2 * mm), HRFlowable(width="100%", thickness=1.5, color=colors.HexColor(NAVY)), Spacer(1, 2 * mm)]

    # legenda
    el.append(P(
        f'{icone("red_circle", 7)} <b>P1 Atraso projetado</b>: entrega projetada (liberação PCP + frete CIF da UF) depois do prazo do cliente · '
        f'{icone("hourglass", 7)} <b>P2 Sem plano PCP</b>: sem previsão de liberação em todos/parte dos itens abertos · '
        f'{icone("yellow_circle", 7)} <b>P3 Em risco</b>: folga de até {dados["folga_alerta"]} dias · FOB: entrega = liberação (frete do cliente).', nota))
    el.append(Spacer(1, 2 * mm))

    # ============================ 1. PAINEL EXECUTIVO ============================
    tot = agrega(linhas)
    kp = [
        (str(tot["n"]), "Pedidos com risco (até 20d)", NAVY),
        (str(tot["atraso"]), "Atraso projetado (P1)", VERMELHO),
        (str(tot["sem_plano"]), "Sem plano PCP (P2)", CINZA),
        (str(tot["risco"]), "Em risco (P3)", AMARELO),
        (moeda(tot["valor"]), "Valor em risco (itens abertos)", NAVY),
        (str(tot["a_solicitar"]), "A solicitar revisão ao comercial", VERMELHO if tot["a_solicitar"] else VERDE),
        (str(tot["rev"]), "Revisão já solicitada", VERDE),
    ]
    topo = [P(v, ParagraphStyle("kv", parent=base_b, fontSize=13 if len(v) < 12 else 10.5, leading=15, alignment=1, textColor=colors.HexColor(c))) for v, _, c in kp]
    rot = [P(esc(r), ParagraphStyle("kr", parent=base, fontSize=6.6, leading=7.8, alignment=1)) for _, r, _ in kp]
    k = Table([topo, rot], colWidths=[largura / len(kp)] * len(kp))
    k.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.5, BORDA), ("INNERGRID", (0, 0), (-1, -1), 0.5, BORDA), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                           ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3), ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f4f7fc"))]))
    el += [k, Spacer(1, 2.2 * mm)]

    # --- matriz horizonte x frete (acumulada) ---
    linhas_m, est_m = [], []
    for h in horizontes:
        data_lim = hoje + timedelta(days=h)
        grupo = [l for l in linhas if l["dias_ate_prazo"] <= h]
        for frete in ("FOB", "CIF", "TOTAL"):
            sel = grupo if frete == "TOTAL" else [l for l in grupo if l["frete"] == frete]
            a = agrega(sel)
            negrito = frete == "TOTAL"
            stb = base_b if negrito else base
            linhas_m.append([
                (T(f"Até {h} dias", base_b) if frete == "FOB" else P("")), (T(f"prazo até {_dm_curto(data_lim)}", peq) if frete == "FOB" else P("")),
                T(frete, stb), T(a["n"], ParagraphStyle("x", parent=stb, alignment=1)),
                P(cor(a["atraso"], VERMELHO) if a["atraso"] else "0", centro), P(cor(a["sem_plano"], CINZA) if a["sem_plano"] else "0", centro),
                P(cor(a["risco"], AMARELO) if a["risco"] else "0", centro), T(moeda(a["valor"]), ParagraphStyle("v", parent=stb, alignment=2)),
                T(a["rev"], centro), P(cor(a["a_solicitar"], VERMELHO) if a["a_solicitar"] else "0", centro),
            ])
            if negrito:
                est_m.append(("BACKGROUND", (0, len(linhas_m)), (-1, len(linhas_m)), colors.HexColor("#e8eef8")))
        est_m.append(("LINEBELOW", (0, len(linhas_m)), (-1, len(linhas_m)), 0.8, colors.HexColor("#8fa8cc")))
    el.append(tabela(["Horizonte", "Data limite", "Frete", "Pedidos", "P1 Atraso", "P2 Sem plano", "P3 Em risco", "Valor em risco", "Revisão solicitada", "A solicitar"],
                     linhas_m, [11, 14, 8, 8, 9, 10, 9, 15, 11, 9], alin="LLLCCCCRCC", extra=est_m, zebra=False))
    el.append(Spacer(1, 2.4 * mm))

    # --- decisões recomendadas ---
    dec = []

    def top_txt(lista, n=3, campo_dias=True):
        pedaco = []
        for l in lista[:n]:
            d = (f"{-l['folga_dias']}d de atraso" if l["classe"] == "atraso" and l["folga_dias"] is not None else
                 (f"folga {l['folga_dias']}d" if l["folga_dias"] is not None else "sem plano"))
            pedaco.append(f"PV {esc(str(l['pedido_venda'] or l['id']))} {esc((l['cliente'] or '')[:22])} ({esc(l['uf'] or '—')}, {d})")
        return "; ".join(pedaco)

    p1_sem_rev = [l for l in linhas if l["classe"] == "atraso" and not l["revisao"]]
    p1_com_rev = [l for l in linhas if l["classe"] == "atraso" and l["revisao"]]
    vencidos = [l for l in linhas if l["dias_ate_prazo"] < 0]
    sem_plano = sorted([l for l in linhas if l["classe"] == "sem_plano"], key=lambda l: l["dias_ate_prazo"])
    risco_l = [l for l in linhas if l["classe"] == "risco"]
    if vencidos:
        mx = max(vencidos, key=lambda l: -l["dias_ate_prazo"])
        dec.append((VERMELHO, "AGIR HOJE — prazos do cliente JÁ VENCIDOS",
                    f'{len(vencidos)} pedido(s), {moeda(sum(l["valor_aberto"] for l in vencidos))} em aberto; maior vencimento: {-mx["dias_ate_prazo"]} dias (PV {esc(str(mx["pedido_venda"] or mx["id"]))}). '
                    "Decidir entre priorizar a entrega e formalizar nova data com o cliente."))
    if p1_sem_rev:
        dec.append((VERMELHO, "PEDIR REVISÃO DE PRAZO AO COMERCIAL (P1 sem solicitação)",
                    f'{len(p1_sem_rev)} pedido(s) com atraso projetado, {moeda(sum(l["valor_aberto"] for l in p1_sem_rev))}. Piores: {top_txt(p1_sem_rev)}. '
                    "Alternativa: antecipar a liberação (ver coluna Ação)."))
    if p1_com_rev:
        mais_antigo = min(p1_com_rev, key=lambda l: l["revisao"].solicitado_em)
        dec.append((AZUL, "COBRAR RETORNO DO COMERCIAL",
                    f'{len(p1_com_rev)} pedido(s) com atraso projetado e revisão já solicitada; a solicitação mais antiga é de '
                    f'{(mais_antigo["revisao"].solicitado_em - timedelta(hours=3)).strftime("%d/%m/%Y %H:%M")} (PV {esc(str(mais_antigo["pedido_venda"] or mais_antigo["id"]))}).'))
    if sem_plano:
        prim = sem_plano[0]
        dec.append((CINZA, "PROGRAMAR NO PCP (P2 sem plano)",
                    f'{len(sem_plano)} pedido(s) sem previsão de liberação ({moeda(sum(l["valor_aberto"] for l in sem_plano))}). '
                    f'Primeiro limite de liberação: {_dm(prim["limite_liberacao"])} (PV {esc(str(prim["pedido_venda"] or prim["id"]))}).'))
    if risco_l:
        dec.append((AMARELO, "BLINDAR A FILA (P3 em risco)",
                    f'{len(risco_l)} pedido(s) com folga de até {dados["folga_alerta"]} dias ({moeda(sum(l["valor_aberto"] for l in risco_l))}): qualquer deslize vira atraso. '
                    f'Menores folgas: {top_txt(risco_l)}.'))
    por_uf = defaultdict(list)
    for l in linhas:
        por_uf[(l["uf"] or "—")].append(l)
    if por_uf:
        ranking = sorted(por_uf.items(), key=lambda kv: (-sum(1 for l in kv[1] if l["classe"] != "risco"), -len(kv[1])))[:3]
        dec.append((NAVY, "CONCENTRAÇÃO REGIONAL",
                    "; ".join(f'{esc(uf)} ({_regiao(uf)}): {len(v)} pedido(s), {moeda(sum(l["valor_aberto"] for l in v))}' for uf, v in ranking)
                    + ". Ver quadro de regiões na próxima página."))
    if not dec:
        dec.append((VERDE, "SEM RISCO IDENTIFICADO", "Nenhum pedido com risco de prazo nos próximos 20 dias."))
    linhas_dec = []
    est_dec = []
    for i, (c, t, txt) in enumerate(dec, 1):
        linhas_dec.append([P(f'<font color="white"><b>{i}</b></font>', centro), P(cor(esc(t), c)), P(txt)])
        est_dec.append(("BACKGROUND", (0, i - 1), (0, i - 1), colors.HexColor(c)))
    td = Table(linhas_dec, colWidths=[7 * mm, 62 * mm, largura - 69 * mm])
    td.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.25, BORDA), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("TOPPADDING", (0, 0), (-1, -1), 2.2),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 2.2), ("LEFTPADDING", (0, 0), (-1, -1), 3)] + est_dec))
    el += banner("DECISÕES RECOMENDADAS À DIREÇÃO", NAVY, "Geradas automaticamente a partir dos pedidos abaixo, em ordem de urgência.", "target")
    el += [td, Spacer(1, 2.4 * mm)]

    # --- gráfico: prazos do cliente por dia ---
    cats = ["Venc."] + [(hoje + timedelta(days=i)).strftime("%d/%m") for i in range(0, 21)]
    ser = {"atraso": [0] * len(cats), "sem_plano": [0] * len(cats), "risco": [0] * len(cats)}
    for l in linhas:
        idx = 0 if l["dias_ate_prazo"] < 0 else min(l["dias_ate_prazo"], 20) + 1
        ser[l["classe"]][idx] += 1
    alt = 40 * mm
    dw = Drawing(largura, alt)
    ch = VerticalBarChart()
    ch.x, ch.y, ch.width, ch.height = 9 * mm, 8 * mm, largura - 12 * mm, alt - 11 * mm
    ch.data = [ser["atraso"], ser["sem_plano"], ser["risco"]]
    ch.categoryAxis.categoryNames = cats
    ch.categoryAxis.style = "stacked"
    ch.categoryAxis.labels.fontSize = 6
    ch.categoryAxis.labels.angle = 0
    ch.categoryAxis.labels.dy = -2
    ch.valueAxis.valueMin = 0
    ch.valueAxis.labels.fontSize = 6.5
    ch.valueAxis.visibleGrid = True
    ch.valueAxis.gridStrokeColor = BORDA
    ch.groupSpacing = 2
    for i, c in enumerate((VERMELHO, CINZA, AMARELO)):
        ch.bars[i].fillColor = colors.HexColor(c)
    ch.valueAxis.valueStep = 1 if max([sum(x) for x in zip(*ch.data)] or [0]) <= 8 else None
    dw.add(ch)
    el += banner("CALENDÁRIO DE PRAZOS EM RISCO — QUANTOS PEDIDOS VENCEM POR DIA", NAVY,
                 "Cada barra = pedidos em risco cujo prazo do cliente cai no dia (vermelho P1, cinza P2, amarelo P3). “Venc.” = prazos já vencidos.", "calendar")
    el += [dw]

    # ============================ 2. REGIÕES ============================
    el += banner("RISCO POR REGIÃO E UF", "#0a6e8c",
                 "Frete rodoviário = lead time cadastrado por UF (soma na entrega projetada do CIF). Ordenado por exposição (P1 + P2, depois nº de pedidos).", "bar_chart")
    ufs = defaultdict(lambda: {"FOB": [], "CIF": []})
    lt_uf = {}
    for l in linhas:
        ufs[l["uf"] or "—"][l["frete"]].append(l)
        if l["frete"] == "CIF" and l["frete_dias"] is not None:
            lt_uf[l["uf"]] = l["frete_dias"]
    def exposicao(item):
        todos = item[1]["FOB"] + item[1]["CIF"]
        return (-sum(1 for l in todos if l["classe"] != "risco"), -len(todos))
    rows_r, est_r = [], []
    for uf, d in sorted(ufs.items(), key=exposicao):
        todos = d["FOB"] + d["CIF"]
        a, f, c = agrega(todos), agrega(d["FOB"]), agrega(d["CIF"])
        pior = min((l["folga_dias"] for l in todos if l["folga_dias"] is not None), default=None)
        proximo = min(todos, key=lambda l: l["dias_ate_prazo"])
        def cel(ag):
            if not ag["n"]:
                return P("—", centro)
            return P(f'{cor(ag["atraso"], VERMELHO) if ag["atraso"] else "0"} / {cor(ag["sem_plano"], CINZA) if ag["sem_plano"] else "0"} / {cor(ag["risco"], AMARELO) if ag["risco"] else "0"}', centro)
        rows_r.append([
            T(uf, base_b), T(_regiao(uf)), T(f"{lt_uf[uf]}d" if uf in lt_uf else ("—" if not d["CIF"] else "s/ cadastro"), centro),
            cel(f), cel(c), T(a["n"], centro_b), T(moeda(a["valor"]), direita),
            P(cor(f"{-pior}d de atraso", VERMELHO) if pior is not None and pior < 0 else (f"folga {pior}d" if pior is not None else "—"), centro),
            T(f'{_dm_curto(proximo["data_solicitada_cliente"])} (PV {proximo["pedido_venda"] or proximo["id"]})'),
        ])
    tot_f, tot_c = agrega([l for l in linhas if l["frete"] == "FOB"]), agrega([l for l in linhas if l["frete"] == "CIF"])
    el.append(tabela(["UF", "Região", "Frete CIF (lead time)", "FOB  P1 / P2 / P3", "CIF  P1 / P2 / P3", "Pedidos", "Valor em risco", "Pior situação", "Prazo mais próximo"],
                     rows_r, [5, 11, 10, 12, 12, 7, 13, 12, 18], alin="LLCCCCRCL"))
    el.append(Spacer(1, 2.4 * mm))

    # por macro-região
    mr = defaultdict(list)
    for l in linhas:
        mr[_regiao(l["uf"])].append(l)
    rows_mr = []
    for nome, lst in sorted(mr.items(), key=lambda kv: -len(kv[1])):
        a = agrega(lst)
        ufs_lista = ", ".join(sorted({l["uf"] or "—" for l in lst}))
        rows_mr.append([T(nome, base_b), T(ufs_lista), T(a["n"], centro_b), P(cor(a["atraso"], VERMELHO) if a["atraso"] else "0", centro),
                        P(cor(a["sem_plano"], CINZA) if a["sem_plano"] else "0", centro), P(cor(a["risco"], AMARELO) if a["risco"] else "0", centro),
                        T(moeda(a["valor"]), direita), T(f'{round(100 * a["valor"] / tot["valor"])}%' if tot["valor"] else "—", centro)])
    el += banner("RESUMO POR MACRO-REGIÃO", "#0a6e8c", None, "bar_chart")
    el.append(tabela(["Região", "UFs com risco", "Pedidos", "P1", "P2", "P3", "Valor em risco", "% do valor"], rows_mr,
                     [14, 30, 8, 7, 7, 7, 14, 9], alin="LLCCCCRC"))

    # ============================ 3. DETALHE POR FRETE x HORIZONTE ============================
    cab_det = ["Prio", "Pedido", "Cliente / Vendedor", "Cidade / UF", "Prazo cliente", "Previsto PCP", "Frete", "Entrega proj.", "Folga", "Valor aberto", "Itens", "Revisão comercial", "Ação orientada"]
    pesos_det = [4, 10, 19, 13, 13, 11, 5, 11, 6, 11, 5, 19, 28]
    alin_det = "CLLLLLCLCRCLL"
    faixas = []
    ant = -10**6
    for h in horizontes:
        faixas.append((h, ant, h))
        ant = h

    el.append(Spacer(1, 3 * mm))
    for frete, cor_f in (("FOB", "#b8860b"), ("CIF", "#0d6efd")):
        selecao_f = [l for l in linhas if l["frete"] == frete]
        af = agrega(selecao_f)
        el += banner(f'{frete} — DETALHE POR HORIZONTE ({af["n"]} pedido(s) · {moeda(af["valor"])})', cor_f,
                     "FOB: frete por conta do cliente, entrega = liberação do PCP." if frete == "FOB"
                     else "CIF: entrega projetada = liberação do PCP + frete rodoviário da UF (lead time cadastrado; sem cadastro = frete não somado).",
                     "funnel")
        for h, _, ate in faixas:
            ini = [x for x in horizontes if x < h]
            lim_inf = (ini[-1] + 1) if ini else None
            sel = [l for l in selecao_f if (lim_inf is None or l["dias_ate_prazo"] >= lim_inf) and l["dias_ate_prazo"] <= h]
            a = agrega(sel)
            titulo_f = (f'Horizonte {h} dias — prazos do cliente até {_dm(hoje + timedelta(days=h))} (inclui vencidos)' if lim_inf is None
                        else f'Horizonte {h} dias — prazos de {_dm(hoje + timedelta(days=lim_inf))} a {_dm(hoje + timedelta(days=h))}')
            cabf = Table([[P(f'<b>{esc(titulo_f)}</b>', ParagraphStyle("hz", parent=base_b, fontSize=8, leading=10, textColor=colors.HexColor(NAVY))),
                           P(f'{a["n"]} pedido(s) · {cor(a["atraso"], VERMELHO)} P1 · {cor(a["sem_plano"], CINZA)} P2 · {cor(a["risco"], AMARELO)} P3 · {moeda(a["valor"])} · '
                             f'{a["rev"]} c/ revisão solicitada · {cor(a["a_solicitar"], VERMELHO)} a solicitar', ParagraphStyle("hz2", parent=base, alignment=2))]],
                         colWidths=[largura * 0.5, largura * 0.5])
            cabf.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#eef2f8")), ("TOPPADDING", (0, 0), (-1, -1), 2.2),
                                      ("BOTTOMPADDING", (0, 0), (-1, -1), 2.2), ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                                      ("LINEBELOW", (0, 0), (-1, -1), 0.6, colors.HexColor("#8fa8cc"))]))
            bloco_h = [CondPageBreak(30 * mm), cabf]
            if not sel:
                bloco_h += [Spacer(1, 0.8 * mm), P(f'{icone("check", 8)} Nenhum pedido {frete} com risco neste horizonte.', nota), Spacer(1, 2 * mm)]
                el += bloco_h
                continue
            rows, extra = [], []
            for i, l in enumerate(sel, 1):
                rev = l["revisao"]
                if rev:
                    rev_txt = (f'<font color="{VERDE}"><b>Solicitada</b></font> {(rev.solicitado_em - timedelta(hours=3)).strftime("%d/%m %H:%M")} · {esc(rev.solicitado_por or "—")}'
                               + (f'<br/>novo prazo sug.: <b>{_dm_curto(rev.novo_prazo_sugerido)}</b>' if rev.novo_prazo_sugerido else "")
                               + (f'<br/><i>{esc((rev.motivo or "")[:60])}</i>' if rev.motivo else ""))
                elif l["classe"] == "risco":
                    rev_txt = "—"
                else:
                    rev_txt = cor("A solicitar", VERMELHO)
                dias = l["dias_ate_prazo"]
                prazo_txt = f'<b>{_dm(l["data_solicitada_cliente"])}</b><br/>' + (cor(f"venceu há {-dias}d", VERMELHO, False) if dias < 0 else ("vence hoje" if dias == 0 else f"em {dias}d"))
                prev = l["data_prevista_pcp"]
                prev_txt = _dm(prev) if prev else cor("sem previsão", CINZA, False)
                if l["itens_sem_previsao"] and prev:
                    prev_txt += f'<br/>{l["itens_sem_previsao"]} item(ns) s/ prev.'
                folga = l["folga_dias"]
                folga_txt = (cor(f"{folga:+d}d", VERMELHO if folga < 0 else AMARELO) if folga is not None else "—")
                frete_txt = f'+{l["frete_dias"]}d' if l["frete"] == "CIF" and l["frete_dias"] is not None else ("s/ cad." if l["frete"] == "CIF" else "FOB")
                rows.append([
                    P(f'<font color="{_COR_CLASSE[l["classe"]]}"><b>{_PRIO[l["classe"]]}</b></font>', centro),
                    T(l["pedido_venda"] or l["id"], base_b),
                    P(f'<b>{esc((l["cliente"] or "—")[:34])}</b><br/>{esc((l["vendedor"] or "—")[:22])}', base),
                    P(f'{esc((l["cidade"] or "—")[:22])}<br/><b>{esc(l["uf"] or "—")}</b> · {_regiao(l["uf"])}', base),
                    P(prazo_txt, base), P(prev_txt, base), T(frete_txt, centro),
                    P(f'<b>{_dm(l["entrega_projetada"])}</b>' if l["entrega_projetada"] else "—", base),
                    P(folga_txt, centro), T(moeda(l["valor_aberto"]), direita), T(f'{l["itens_abertos"]}/{l["itens_total"]}', centro),
                    P(rev_txt, peq), P(esc(_acao(l)), peq),
                ])
                extra.append(("BACKGROUND", (0, i), (0, i), colors.HexColor(_FUNDO_PRIO[l["classe"]])))
            bloco_h += [Spacer(1, 0.6 * mm), tabela(cab_det, rows, pesos_det, alin=alin_det, extra=extra), Spacer(1, 2.4 * mm)]
            el += bloco_h
        sp = dados["sem_prazo"].get(frete, 0)
        if sp:
            el.append(P(f'{icone("pin", 8)} {sp} pedido(s) {frete} em aberto <b>sem prazo do cliente cadastrado</b> não entram na projeção — cadastrar o prazo para avaliar o risco.', nota))

    # ============================ 4. OCULTOS + METODOLOGIA ============================
    el.append(CondPageBreak(60 * mm))
    ocultos = dados.get("ocultos") or []
    if ocultos:
        el += banner(f"PEDIDOS OCULTOS DA GESTÃO DE RISCO ({len(ocultos)})", CINZA,
                     "Pedidos fora da curva da carteira, retirados dos totais pelo PCP. Reexibir no Painel > Mini gestão de risco.", "check")
        rows_o = [[T(o["pedido_venda"] or o["pedido_id"], base_b), T(o["cliente"]), T(o["frete"]), T(_dm(o["prazo_cliente"])), T(o["motivo"] or "—"),
                   T(f'{o["quando"] or "—"} · {o["oculto_por"] or "—"}')] for o in ocultos]
        el += [tabela(["Pedido", "Cliente", "Frete", "Prazo cliente", "Motivo", "Ocultado em / por"], rows_o, [8, 26, 5, 12, 30, 22]), Spacer(1, 2.4 * mm)]
    el += banner("COMO LER ESTE RELATÓRIO", NAVY, None, "check")
    metodo = [
        f'<b>Entrega projetada</b> = liberação prevista pelo PCP do ÚLTIMO item ainda aberto (se a previsão já passou, vale hoje, {_dm(hoje)}) + frete rodoviário da UF no CIF. '
        'No FOB o frete é do cliente, então a entrega projetada é a própria liberação.',
        f'<b>Folga</b> = prazo do cliente − entrega projetada. Negativa = <b>P1 atraso projetado</b>; de 0 a {dados["folga_alerta"]} dias = <b>P3 em risco</b>; pedidos sem previsão de liberação do PCP = <b>P2 sem plano</b>.',
        '<b>Horizontes</b> contam os dias a partir de hoje até o prazo do cliente; os 5 dias incluem os prazos já vencidos. Os totais da página 1 são acumulados (10d inclui 5d; 20d inclui 10d); '
        'o detalhe por frete lista cada pedido uma única vez, na sua faixa.',
        '<b>Revisão comercial</b>: apontamento feito pelo PCP no Painel (data/hora, responsável, novo prazo sugerido e motivo). “A solicitar” = P1/P2 sem apontamento.',
        'Dados recalculados no momento da emissão; datas e horários no fuso de Brasília.',
    ]
    for m in metodo:
        el.append(P("• " + m, nota))

    def rodape(c, d):
        c.saveState()
        c.setFont("Helvetica", 7)
        c.setFillColor(colors.HexColor("#888"))
        c.drawString(doc.leftMargin, 6 * mm, "4PIPE Solutions · Gestão da Produção e Operação · Gestão de risco de prazos — visão executiva")
        c.drawRightString(landscape(A4)[0] - doc.rightMargin, 6 * mm, f'Emitido em {agora.strftime("%d/%m/%Y %H:%M")} · Página {d.page}')
        c.restoreState()

    doc.build(el, onFirstPage=rodape, onLaterPages=rodape)
    return buf.getvalue()
