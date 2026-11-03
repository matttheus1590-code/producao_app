"""PDFs do Centro de Comando (pedido do Bruno, 03/11/2026): "gerar relatório em
PDF, com logo, horário etc." em todas as abas gerenciais de Gestão Produção /
Operação.

Módulo independente (não importa o app): recebe os dados JÁ calculados pelas
funções _dados_comando/_dados_alertas/_gaps_operacao/_dados_faturamento_projecao
do app.py e só apresenta — por isso tela e PDF nunca divergem. Nada de glyph
fora do Latin-1 (setas, ✓, emoji...): a fonte padrão do reportlab os desenha
como quadrado preto; ícones coloridos vêm de PNGs (função `ctx["icone"]`).

Todos os geradores montam o MESMO cabeçalho (logo 4PIPE, título, data/hora de
emissão, usuário) e o MESMO rodapé (página n) — `_documento`.
"""

import io
from xml.sax.saxutils import escape as esc

from reportlab.graphics.charts.barcharts import HorizontalBarChart, VerticalBarChart
from reportlab.graphics.shapes import Drawing, String
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    CondPageBreak, HRFlowable, Image, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)

NAVY = "#1b2a4a"
COR = {"vermelho": "#c0392b", "amarelo": "#b8860b", "verde": "#198754", "azul": "#0d6efd", "cinza": "#6c757d"}
FUNDO = {"vermelho": "#f8d7da", "amarelo": "#fff3cd", "verde": "#d1e7dd", "azul": "#cfe2ff", "cinza": "#e9ecef"}
BORDA = colors.HexColor("#dee2e6")
ZEBRA = colors.HexColor("#f8f9fa")
CAB_BG = colors.HexColor("#d3e0f2")

# emoji da tela -> PNG pré-renderado em static/img/emoji_pdf
ICONE = {
    "📦": "package", "🔴": "red_circle", "🟡": "yellow_circle", "🟢": "green_circle", "🛡️": "warning", "🎯": "target",
    "⏱️": "stopwatch", "🏭": "factory", "💵": "moneybag", "🚚": "truck", "🤝": "handshake", "🗓️": "calendar",
    "🧾": "receipt", "🔁": "gear", "⏳": "hourglass", "📃": "clipboard", "📈": "bar_chart", "🚨": "red_circle",
    "🧪": "test_tube", "🔍": "magnifier", "⚠️": "warning", "🌐": "globe",
}

_DIAS = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sab", "Dom"]


def _d(valor):
    return valor.strftime("%d/%m/%Y") if valor else "-"


def _dm(valor):
    return valor.strftime("%d/%m") if valor else "-"


def _t(texto):
    """Texto seguro pro Paragraph (escapa e troca o que a fonte padrão não desenha)."""
    s = esc(str(texto if texto is not None else "-"))
    for velho, novo in (("→", "-&gt;"), ("—", "-"), ("–", "-"), ("·", "-"), ("’", "'"), ("“", '"'), ("”", '"')):
        s = s.replace(velho, novo)
    return s


def _estilos():
    base = getSampleStyleSheet()
    normal = ParagraphStyle("c_normal", parent=base["Normal"], fontSize=8.5, leading=10.5)
    return {
        "normal": normal,
        "negrito": ParagraphStyle("c_neg", parent=normal, fontName="Helvetica-Bold"),
        "pequeno": ParagraphStyle("c_peq", parent=normal, fontSize=7.5, leading=9, textColor=colors.HexColor("#495057")),
        "cab": ParagraphStyle("c_cab", parent=normal, fontName="Helvetica-Bold", fontSize=8.5, textColor=colors.HexColor(NAVY)),
        "secao": ParagraphStyle("c_sec", parent=base["Heading2"], fontSize=13, leading=16, textColor=colors.HexColor(NAVY), spaceBefore=6, spaceAfter=4),
        "sub": ParagraphStyle("c_sub", parent=base["Heading3"], fontSize=10, leading=12, textColor=colors.HexColor(NAVY), spaceBefore=4, spaceAfter=3),
        "direita": ParagraphStyle("c_dir", parent=normal, alignment=2),
        "centro": ParagraphStyle("c_cen", parent=normal, alignment=1),
        "obs": ParagraphStyle("c_obs", parent=normal, fontSize=8, textColor=colors.HexColor("#495057")),
    }


def _ic(ctx, emoji, tamanho=9):
    nome = ICONE.get(emoji)
    return (ctx["icone"](nome, tamanho) + " ") if nome else ""


def _tabela(dados, pesos, largura, cab_linhas=1, zebra=True, extra=None):
    larguras = [p / sum(pesos) * largura for p in pesos]
    t = Table(dados, colWidths=larguras, repeatRows=cab_linhas)
    estilo = [
        ("BACKGROUND", (0, 0), (-1, cab_linhas - 1), CAB_BG),
        ("GRID", (0, 0), (-1, -1), 0.4, BORDA),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]
    if zebra:
        for i in range(cab_linhas, len(dados)):
            if (i - cab_linhas) % 2:
                estilo.append(("BACKGROUND", (0, i), (-1, i), ZEBRA))
    t.setStyle(TableStyle(estilo + (extra or [])))
    return t


def _kpi_grid(ctx, tiles, largura, colunas=4):
    """Cartões coloridos (semáforo) com ícone, título, valor grande e legenda."""
    est = _estilos()
    celulas = []
    for ti in tiles:
        cor = ti.get("cor", "cinza")
        celulas.append([
            Paragraph(_ic(ctx, ti.get("emoji", ""), 9) + f"<b>{_t(ti['titulo'])}</b>", ParagraphStyle("kt", parent=est["pequeno"], textColor=colors.HexColor(NAVY))),
            Paragraph(_t(ti["valor"]), ParagraphStyle("kv", parent=est["normal"], fontName="Helvetica-Bold", fontSize=18, leading=21, textColor=colors.HexColor(COR[cor]))),
            Paragraph(_t(ti.get("sub", "")), est["pequeno"]),
        ])
    while len(celulas) % colunas:
        celulas.append("")
    linhas = [celulas[i:i + colunas] for i in range(0, len(celulas), colunas)]
    t = Table(linhas, colWidths=[largura / colunas] * colunas)
    estilo = [("VALIGN", (0, 0), (-1, -1), "TOP"), ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
              ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6)]
    for r, linha in enumerate(linhas):
        for c, cel in enumerate(linha):
            if cel == "":
                continue
            cor = tiles[r * colunas + c].get("cor", "cinza")
            estilo += [("BOX", (c, r), (c, r), 1.0, colors.HexColor(COR[cor])), ("BACKGROUND", (c, r), (c, r), colors.HexColor(FUNDO[cor]))]
    estilo += [("LINEBEFORE", (1, 0), (-1, -1), 4, colors.white), ("LINEABOVE", (0, 1), (-1, -1), 4, colors.white)]
    t.setStyle(TableStyle(estilo))
    return t


def _celula_cor(texto, cor, estilo):
    return Paragraph(f'<font color="{COR[cor]}"><b>{_t(texto)}</b></font>', estilo)


def _documento(ctx, titulo, subtitulo, construir, paisagem=True):
    """Monta cabeçalho (logo + título + emissão), corpo e rodapé de página."""
    pagina = landscape(A4) if paisagem else A4
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=pagina, leftMargin=10 * mm, rightMargin=10 * mm, topMargin=10 * mm, bottomMargin=13 * mm,
                            title=titulo, author="4PIPE Solutions")
    largura = pagina[0] - doc.leftMargin - doc.rightMargin
    est = _estilos()
    agora = ctx["agora"]
    carimbo = f'{agora.strftime("%d/%m/%Y")} às {agora.strftime("%H:%M")}'
    try:
        logo = Image(ctx["logo"], width=44 * mm, height=44 * mm * (63 / 261))
    except Exception:
        logo = Paragraph("", est["normal"])
    cab = Table([[
        logo,
        [Paragraph(f'<font size="15" color="{NAVY}"><b>{_t(titulo)}</b></font>', ParagraphStyle("ct", parent=est["normal"], fontSize=15, leading=19)),
         Paragraph(_t(subtitulo), est["pequeno"])],
        [Paragraph(f'<b>Emitido em</b> {carimbo}', ParagraphStyle("ce", parent=est["pequeno"], alignment=2)),
         Paragraph(f'por {_t(ctx.get("usuario") or "sistema")}', ParagraphStyle("ce2", parent=est["pequeno"], alignment=2))],
    ]], colWidths=[largura * 0.24, largura * 0.52, largura * 0.24])
    cab.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    elementos = [cab, HRFlowable(width="100%", thickness=1.2, color=colors.HexColor(NAVY), spaceBefore=3, spaceAfter=5)]
    elementos += construir(ctx, largura)

    def rodape(canvas_obj, doc_obj):
        canvas_obj.saveState()
        canvas_obj.setFont("Helvetica", 7.5)
        canvas_obj.setFillColor(colors.HexColor("#888888"))
        canvas_obj.drawString(doc_obj.leftMargin, 6 * mm, f"4PIPE Solutions - Gestão da Produção e Operação - {titulo}")
        canvas_obj.drawRightString(pagina[0] - doc_obj.rightMargin, 6 * mm, f"Emitido em {carimbo} - Página {doc_obj.page}")
        canvas_obj.restoreState()

    doc.build(elementos, onFirstPage=rodape, onLaterPages=rodape)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# Seções (reaproveitadas pelo "Pacote gerencial" e pelo PDF de KPIs)
# ---------------------------------------------------------------------------

def sec_comando(ctx, d, largura):
    est = _estilos()
    moeda = ctx["moeda"]
    out = [Paragraph(_ic(ctx, "🎯", 12) + "Centro de Comando - resumo executivo", est["secao"]),
           _kpi_grid(ctx, d["tiles"], largura), Spacer(1, 4 * mm)]

    # Gargalos | Pedidos críticos lado a lado
    cab_g = [Paragraph(c, est["cab"]) for c in ("Estação", "Abertos", "Atras.", "Em risco", "Risco", "Cobertura")]
    linhas_g = [cab_g]
    extra_g = []
    for i, g in enumerate(d["gargalos"], start=1):
        linhas_g.append([
            Paragraph(_ic(ctx, g.get("emoji", "")) + _t(g["curto"]), est["normal"]),
            Paragraph(str(g["abertos"]), est["centro"]), Paragraph(str(g["atrasados"]), est["centro"]),
            Paragraph(str(g["em_risco"]), est["centro"]),
            _celula_cor(f'{g["risco_pct"]}%', g["cor"], est["centro"]),
            _celula_cor(f'{g["cobertura"]:.1f} sem'.replace(".", ",") if g.get("cobertura") is not None else "-", g.get("cap_cor", "cinza"), est["centro"]),
        ])
        extra_g.append(("BACKGROUND", (4, i), (4, i), colors.HexColor(FUNDO[g["cor"]])))
    if len(linhas_g) == 1:
        linhas_g.append([Paragraph("Nenhuma estação com itens abertos.", est["normal"]), "", "", "", "", ""])
    larg_g = largura * 0.46
    tab_g = _tabela(linhas_g, [26, 11, 10, 11, 11, 15], larg_g, extra=extra_g)

    cab_c = [Paragraph(c, est["cab"]) for c in ("Pedido", "Cliente", "Atraso", "Valor", "Causa")]
    linhas_c = [cab_c]
    for p in d["criticos"]:
        linhas_c.append([
            Paragraph(f'<b>{_t(p["pedido_venda"])}</b>', est["normal"]), Paragraph(_t((p["cliente"] or "")[:22]), est["normal"]),
            _celula_cor(f'{p["dias_atraso"]}d', "vermelho", est["centro"]), Paragraph(_t(moeda(p["valor"])), est["direita"]),
            Paragraph(_ic(ctx, p.get("motivo_emoji", "")) + _t(p["motivo_titulo"]), est["pequeno"]),
        ])
    if len(linhas_c) == 1:
        linhas_c.append([Paragraph("Nenhum pedido atrasado.", est["normal"]), "", "", "", ""])
    larg_c = largura * 0.52
    tab_c = _tabela(linhas_c, [13, 22, 9, 17, 30], larg_c)

    lado = Table([[[Paragraph(_ic(ctx, "🏭") + "Gargalos por estação", est["sub"]), tab_g],
                   [Paragraph(_ic(ctx, "🔴") + "Pedidos críticos (mais atrasados)", est["sub"]), tab_c]]],
                 colWidths=[largura * 0.48, largura * 0.52])
    lado.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 2)]))
    out += [lado, Spacer(1, 3 * mm)]

    # O que fazer agora + carga projetada
    itens = []
    for i, dec in enumerate(d["decisoes"], start=1):
        itens.append([Paragraph(f'<b>{i}</b>', est["centro"]), Paragraph(_ic(ctx, dec["emoji"]) + _t(dec["texto"]), est["normal"])])
    if not itens:
        itens.append(["", Paragraph("Nenhuma ação urgente identificada. Operação dentro do esperado.", est["normal"])])
    t_dec = Table(itens, colWidths=[8 * mm, largura * 0.56 - 8 * mm])
    estilo_dec = [("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("GRID", (0, 0), (-1, -1), 0.4, BORDA),
                  ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]
    for i, dec in enumerate(d["decisoes"]):
        estilo_dec.append(("BACKGROUND", (0, i), (0, i), colors.HexColor(FUNDO[dec["nivel"]])))
    t_dec.setStyle(TableStyle(estilo_dec))

    carga = d["carga"]
    graf = Drawing(largura * 0.38, 46 * mm)
    if carga["cargas"]:
        bc = VerticalBarChart()
        bc.bars.strokeColor = None
        bc.x, bc.y, bc.width, bc.height = 10 * mm, 9 * mm, largura * 0.38 - 16 * mm, 30 * mm
        bc.data = [list(carga["cargas"])]
        bc.bars[0].fillColor = colors.HexColor("#0d6efd")
        bc.categoryAxis.categoryNames = [r["titulo"] for r in carga["rotulos"]][: len(carga["cargas"])]
        bc.categoryAxis.labels.fontSize = 7
        bc.categoryAxis.labels.fontName = "Helvetica"
        bc.valueAxis.labels.fontSize = 7
        bc.valueAxis.labels.fontName = "Helvetica"
        bc.valueAxis.valueMin = 0
        bc.valueAxis.valueMax = max(carga["cargas"] + [1]) * 1.3
        bc.barLabelFormat = "%d"
        bc.barLabels.fontSize = 8
        bc.barLabels.fontName = "Helvetica"
        bc.barLabels.nudge = 6
        graf.add(bc)
        graf.add(String(2 * mm, 42 * mm, "Itens abertos por semana (carga projetada pelo PCP)", fontName="Helvetica-Bold", fontSize=8, fillColor=colors.HexColor(NAVY)))
    rodape_carga = Paragraph(
        f'Ritmo médio: <b>{_t(carga["ritmo"])}</b> itens/semana - Cobertura: <b>{_t(carga["cobertura"])}</b> sem. - '
        f'{carga["atrasados"]} atrasados - {carga["sem_prev"]} sem previsão.', est["pequeno"])
    bloco = Table([[[Paragraph(_ic(ctx, "💡") + "O que fazer agora", est["sub"]), t_dec],
                    [Paragraph(_ic(ctx, "📈") + "Carga das próximas semanas", est["sub"]), graf, rodape_carga]]],
                  colWidths=[largura * 0.58, largura * 0.42])
    bloco.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
    out.append(KeepTogether([bloco]))
    out.append(Spacer(1, 2 * mm))
    out.append(Paragraph(
        "Semáforos: vermelho = ação imediata - amarelo = atencao - verde = dentro do esperado. Números calculados com as mesmas "
        "regras do Painel e de Gestão Operação.", est["obs"]))
    return out


def sec_alertas(ctx, d, largura):
    est = _estilos()
    moeda = ctx["moeda"]
    atr = d["atrasados"]
    ra = atr["resumo"]
    n_crit = sum(1 for g in d["gargalos"] if g["cor"] == "vermelho")
    tiles = [
        {"emoji": "🔴", "titulo": "Pedidos atrasados", "valor": str(ra["total"]), "sub": f'{moeda(ra["valor_em_risco"])} em risco', "cor": "vermelho" if ra["total"] else "verde"},
        {"emoji": "🟡", "titulo": "Vencendo em 3 dias", "valor": str(len(d["vencendo"])), "sub": "prazo do cliente", "cor": "amarelo" if d["vencendo"] else "verde"},
        {"emoji": "🏭", "titulo": "Estações criticas", "valor": str(n_crit), "sub": f'de {len(d["gargalos"])} estações', "cor": "vermelho" if n_crit else "verde"},
        {"emoji": "💵", "titulo": "Faturamento atrasado", "valor": moeda(d["fat_pend_total"]), "sub": f'{d["fat_pend_itens"]} itens previstos e não faturados', "cor": "amarelo" if d["fat_pend_total"] else "verde"},
    ]
    out = [Paragraph(_ic(ctx, "🚨", 12) + "Alertas e Gargalos", est["secao"]), _kpi_grid(ctx, tiles, largura), Spacer(1, 3 * mm)]

    out.append(Paragraph(_ic(ctx, "🔴") + f'Pedidos atrasados ({ra["total"]}) - causa, datas e ação recomendada', est["sub"]))
    cab = [Paragraph(c, est["cab"]) for c in ("Pedido", "Cliente", "Estação", "Prazo cliente", "Prev. PCP", "Atraso", "Valor", "Frete/UF", "Causa", "Ação recomendada")]
    linhas = [cab]
    for p in sorted(atr["pedidos"], key=lambda p: (-(p["dias_atraso"] or 0), -(p["valor"] or 0))):
        linhas.append([
            Paragraph(f'<b>{_t(p["pedido_venda"])}</b>', est["pequeno"]), Paragraph(_t((p["cliente"] or "")[:24]), est["pequeno"]),
            Paragraph(_t(p["estacoes"] or "-"), est["pequeno"]), Paragraph(_d(p["prazo_cliente"]), est["pequeno"]),
            Paragraph(_d(p["prevista_pcp"]), est["pequeno"]), _celula_cor(f'{p["dias_atraso"]}d', "vermelho", est["centro"]),
            Paragraph(_t(moeda(p["valor"])), ParagraphStyle("pd", parent=est["pequeno"], alignment=2)),
            Paragraph(_t(f'{p["frete"] or "-"}/{p["uf"] or "-"}'), est["pequeno"]),
            Paragraph(_ic(ctx, p.get("motivo_emoji", ""), 8) + _t(p["motivo_titulo"]), est["pequeno"]),
            Paragraph(_t((p.get("acao") or "")[:140]), est["pequeno"]),
        ])
    if len(linhas) == 1:
        linhas.append([Paragraph("Nenhum pedido atrasado.", est["normal"])] + [""] * 9)
    out.append(_tabela(linhas, [8, 14, 9, 8, 8, 6, 10, 8, 15, 25], largura))

    out.append(CondPageBreak(60 * mm))
    out.append(Paragraph(_ic(ctx, "🟡") + f'Vencendo nos próximos 3 dias ({len(d["vencendo"])})', est["sub"]))
    cab = [Paragraph(c, est["cab"]) for c in ("Pedido", "Cliente", "Prazo do cliente", "Itens abertos")]
    linhas = [cab]
    for p in d["vencendo"]:
        abertos = [i for i in p.itens if i.status_producao != "FINALIZADO"]
        linhas.append([Paragraph(f'<b>{_t(p.pedido_venda or "-")}</b>', est["normal"]), Paragraph(_t(p.cliente), est["normal"]),
                       Paragraph(_d(p.data_cliente), est["normal"]), Paragraph(str(len(abertos)), est["centro"])])
    if len(linhas) == 1:
        linhas.append([Paragraph("Nenhum pedido vencendo nos próximos dias.", est["normal"]), "", "", ""])
    out.append(_tabela(linhas, [14, 40, 16, 12], largura * 0.7))

    out.append(CondPageBreak(70 * mm))
    out.append(Paragraph(_ic(ctx, "🏭") + "Gargalos por estação (fila, atraso, espera, lead time, capacidade)", est["sub"]))
    cab = [Paragraph(c, est["cab"]) for c in ("Estação", "Abertos", "Pend.", "Prod.", "Atras.", "Cliente vencido", "Risco", "Espera (d)", "Lead time (d)", "Valor parado", "Ritmo/sem", "Cobertura")]
    linhas = [cab]
    extra = []
    for i, g in enumerate(d["gargalos"], start=1):
        linhas.append([
            Paragraph(_ic(ctx, g.get("emoji", "")) + _t(g["curto"]), est["normal"]), Paragraph(str(g["abertos"]), est["centro"]),
            Paragraph(str(g["pendente"]), est["centro"]), Paragraph(str(g["producao"]), est["centro"]),
            Paragraph(str(g["atrasados"]), est["centro"]), Paragraph(str(g["cli_vencido"]), est["centro"]),
            _celula_cor(f'{g["risco_pct"]}%', g["cor"], est["centro"]),
            Paragraph(str(g["espera"]) if g["espera"] is not None else "-", est["centro"]),
            Paragraph(str(g["lt"]) if g["lt"] is not None else "-", est["centro"]),
            Paragraph(_t(moeda(g["valor_parado"])), ParagraphStyle("gp", parent=est["normal"], alignment=2)),
            Paragraph(_t(g["ritmo"]) if g["ritmo"] is not None else "-", est["centro"]),
            _celula_cor((f'{g["cobertura"]:.1f} sem'.replace(".", ",")) if g["cobertura"] is not None else "-", g["cap_cor"], est["centro"]),
        ])
        extra.append(("BACKGROUND", (6, i), (6, i), colors.HexColor(FUNDO[g["cor"]])))
    out.append(_tabela(linhas, [18, 7, 6, 6, 6, 9, 7, 8, 9, 13, 8, 9], largura, extra=extra))

    graf_dados = [g for g in d["gargalos"] if g["abertos"] > 0][:10]
    if graf_dados:
        out.append(Spacer(1, 3 * mm))
        dr = Drawing(largura, 8 * mm + 6 * mm * len(graf_dados))
        bc = HorizontalBarChart()
        bc.bars.strokeColor = None
        bc.x, bc.y, bc.width, bc.height = 42 * mm, 3 * mm, largura - 55 * mm, 6 * mm * len(graf_dados)
        bc.data = [[g["abertos"] for g in reversed(graf_dados)], [g["atrasados"] for g in reversed(graf_dados)]]
        bc.bars[0].fillColor = colors.HexColor("#f59f00")
        bc.bars[1].fillColor = colors.HexColor("#dc3545")
        bc.categoryAxis.categoryNames = [g["curto"] for g in reversed(graf_dados)]
        bc.categoryAxis.labels.fontSize = 7
        bc.categoryAxis.labels.fontName = "Helvetica"
        bc.valueAxis.labels.fontSize = 7
        bc.valueAxis.labels.fontName = "Helvetica"
        bc.valueAxis.valueMin = 0
        bc.barLabelFormat = "%d"
        bc.barLabels.fontSize = 7
        bc.barLabels.fontName = "Helvetica"
        dr.add(bc)
        dr.add(String(42 * mm, 6 * mm * len(graf_dados) + 4 * mm, "Itens abertos (laranja) x atrasados (vermelho) por estação", fontName="Helvetica-Bold", fontSize=8, fillColor=colors.HexColor(NAVY)))
        out.append(dr)

    out.append(CondPageBreak(60 * mm))
    out.append(Paragraph(_ic(ctx, "💵") + f'Faturamento previsto e ainda não realizado - {moeda(d["fat_pend_total"])} ({d["fat_pend_itens"]} itens)', est["sub"]))
    cab = [Paragraph(c, est["cab"]) for c in ("Pedido", "Cliente", "Itens", "Previsto desde", "Dias de atraso", "Valor")]
    linhas = [cab]
    for r in d["fat_pend"][:15]:
        linhas.append([Paragraph(f'<b>{_t(r["pedido"])}</b>', est["normal"]), Paragraph(_t((r["cliente"] or "")[:30]), est["normal"]),
                       Paragraph(str(r["itens"]), est["centro"]), Paragraph(_d(r["mais_antigo"]), est["normal"]),
                       _celula_cor(f'{r["dias"]}d' if r["dias"] is not None else "-", "amarelo", est["centro"]),
                       Paragraph(_t(moeda(r["valor"])), est["direita"])])
    if len(linhas) == 1:
        linhas.append([Paragraph("Sem pendências de faturamento previsto.", est["normal"]), "", "", "", "", ""])
    out.append(_tabela(linhas, [14, 34, 7, 14, 13, 18], largura * 0.8))
    return out


def sec_leadtime_gaps(ctx, gaps, lt_estacao, largura, titulo="Lead Time e Gaps"):
    est = _estilos()
    out = [Paragraph(_ic(ctx, "⏱️", 12) + f"{titulo} - últimos {gaps['meses']} meses", est["secao"])]
    if lt_estacao:
        out.append(Paragraph("Lead time por estação (dias)", est["sub"]))
        cab = [Paragraph(c, est["cab"]) for c in ("Estação", "Fila", "Processamento", "Total", "OPs")]
        linhas = [cab]
        for l in lt_estacao:
            neg = str(l["estacao"]).startswith("TOTAL")
            e = est["negrito"] if neg else est["normal"]
            f = lambda v: (f"{v}d" if v is not None else "-")
            linhas.append([Paragraph(_t(l["estacao"]), e), Paragraph(f(l["fila_media"]), e), Paragraph(f(l["processamento_medio"]), e),
                           Paragraph(f(l["lead_time_medio"]), e), Paragraph(str(l["quantidade"]), e)])
        out.append(_tabela(linhas, [34, 14, 18, 14, 10], largura * 0.6))
        out.append(Spacer(1, 3 * mm))

    out.append(Paragraph(_ic(ctx, "🔁") + "Gaps da cadeia (liberação -> NF -> expedição -> entrega) e contra o prazo do cliente", est["sub"]))
    cab = [Paragraph(c, est["cab"]) for c in ("Etapa", "Média", "Médiana", "Máximo", "Pedidos", "No prazo", "Leitura")]
    linhas = [cab]
    extra = []
    for i, e in enumerate(gaps["etapas"], start=1):
        media = e["media"]
        cor = "cinza"
        if media is not None and e["vs_prazo"]:
            cor = "verde" if media <= 0 else ("amarelo" if media <= 3 else "vermelho")
        linhas.append([
            Paragraph(_ic(ctx, e["emoji"]) + _t(e["titulo"]), est["normal"]),
            _celula_cor((f'{media:+.1f}d' if e["vs_prazo"] else f'{media:.1f}d').replace(".", ",") if media is not None else "-", cor, est["centro"]),
            Paragraph(str(e["mediana"]) + "d" if e["mediana"] is not None else "-", est["centro"]),
            Paragraph(str(e["max"]) + "d" if e["max"] is not None else "-", est["centro"]),
            Paragraph(str(e["n"]), est["centro"]),
            Paragraph((f'{e["pct_ok"]}%' if e["vs_prazo"] and e["pct_ok"] is not None else "-"), est["centro"]),
            Paragraph(_t(e["dica"]), est["pequeno"]),
        ])
        if cor != "cinza":
            extra.append(("BACKGROUND", (1, i), (1, i), colors.HexColor(FUNDO[cor])))
    out.append(_tabela(linhas, [30, 8, 8, 8, 8, 8, 30], largura, extra=extra))
    a = gaps["aderencia_pcp"]
    if a["n"]:
        out.append(Spacer(1, 2 * mm))
        out.append(Paragraph(
            f'<b>Aderência ao planejamento do PCP</b> (liberação real x prevista, {a["n"]} itens): media {_t(a["media"])}d '
            f'(positivo = liberou depois do previsto) - {a["pct_ok"]}% liberados no prazo previsto.', est["normal"]))
    return out


def sec_faturamento(ctx, d, largura):
    est = _estilos()
    moeda = ctx["moeda"]
    base = d["base"]
    previsto, realizado = base["previsto_total"], base["realizado_total"]
    pct = round(100 * realizado / previsto) if previsto else None
    tiles = [
        {"emoji": "🎯", "titulo": f'Previsto {d["mes_label"]}', "valor": moeda(previsto), "sub": "planejamento PCP", "cor": "azul"},
        {"emoji": "💵", "titulo": "Realizado no mês", "valor": moeda(realizado), "sub": (f'{pct}% do previsto' if pct is not None else "") + f' - {base["itens_realizados"]} itens', "cor": "verde"},
        {"emoji": "⏳", "titulo": "Previsto e não faturado", "valor": moeda(d["fat_pend_total"]), "sub": f'{d["fat_pend_itens"]} itens com liberação vencida', "cor": "amarelo" if d["fat_pend_total"] else "verde"},
        {"emoji": "📦", "titulo": "Carteira em aberto", "valor": moeda(d["carteira_total"]), "sub": "valor parado nas estações", "cor": "cinza"},
    ]
    out = [Paragraph(_ic(ctx, "💵", 12) + f'Faturamento e Projeção - {d["mes_label"]}', est["secao"]), _kpi_grid(ctx, tiles, largura), Spacer(1, 3 * mm)]

    meses = d["meses_ano"]
    dr = Drawing(largura, 62 * mm)
    bc = VerticalBarChart()
    bc.bars.strokeColor = None
    bc.x, bc.y, bc.width, bc.height = 14 * mm, 10 * mm, largura - 24 * mm, 40 * mm
    bc.data = [[m["projetado"] / 1000 for m in meses], [m["realizado"] / 1000 for m in meses]]
    bc.bars[0].fillColor = colors.HexColor("#6ea8fe")
    bc.bars[1].fillColor = colors.HexColor("#198754")
    bc.categoryAxis.categoryNames = [m["label"] for m in meses]
    bc.categoryAxis.labels.fontSize = 7
    bc.categoryAxis.labels.fontName = "Helvetica"
    bc.valueAxis.labels.fontSize = 7
    bc.valueAxis.labels.fontName = "Helvetica"
    bc.valueAxis.valueMin = 0
    dr.add(bc)
    dr.add(String(14 * mm, 55 * mm, f'Projetado pelo PCP (azul) x faturado (verde) por mês - {d["ano"]} (R$ mil)', fontName="Helvetica-Bold", fontSize=8, fillColor=colors.HexColor(NAVY)))
    out += [dr]

    out.append(Paragraph(_ic(ctx, "📈") + "Projeção dos próximos 6 meses (o que o PCP ainda tem em aberto)", est["sub"]))
    cab = [Paragraph(c, est["cab"]) for c in ("Mês", "Pedidos em aberto", "Valor em aberto", "Pedidos planejados", "Valor planejado")]
    linhas = [cab]
    for f in d["futuros"]:
        linhas.append([Paragraph(f'<b>{_t(f["label"])}</b>', est["normal"]), Paragraph(str(f["pedidos_aberto"]), est["centro"]),
                       Paragraph(_t(moeda(f["valor_aberto"])), est["direita"]), Paragraph(str(f["pedidos_total"]), est["centro"]),
                       Paragraph(_t(moeda(f["valor_total"])), est["direita"])])
    out.append(_tabela(linhas, [14, 16, 22, 16, 22], largura * 0.75))

    out.append(CondPageBreak(70 * mm))
    out.append(Paragraph("Previsto do mês por região e por estação", est["sub"]))

    def _quebra(linhas_q, titulo):
        cab_q = [Paragraph(titulo, est["cab"]), Paragraph("Previsto", est["cab"])]
        rows = [cab_q] + [[Paragraph(_t(l["chave"]), est["normal"]), Paragraph(_t(moeda(l["valor"])), est["direita"])] for l in linhas_q[:8]]
        if len(rows) == 1:
            rows.append([Paragraph("Sem dados", est["normal"]), ""])
        return _tabela(rows, [60, 40], largura * 0.47)

    duas = Table([[_quebra(base["previsto_por_regiao"], "Região"), _quebra(base["previsto_por_estacao"], "Estação")]],
                 colWidths=[largura * 0.5, largura * 0.5])
    duas.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
    out.append(duas)

    par = base["previsto_pareto_clientes"]
    out.append(Spacer(1, 3 * mm))
    out.append(Paragraph(f'Principais clientes - curva 80/20: {par["n_principais"]} de {par["n_clientes"]} clientes respondem por 80% do previsto', est["sub"]))
    cab = [Paragraph(c, est["cab"]) for c in ("Cliente", "Previsto", "% do mês", "% acumulado")]
    linhas = [cab]
    extra = []
    for i, l in enumerate(par["linhas"][:10], start=1):
        linhas.append([Paragraph(_t(l["chave"]), est["normal"]), Paragraph(_t(moeda(l["valor"])), est["direita"]),
                       Paragraph(f'{l["pct"]}%'.replace(".", ","), est["centro"]), Paragraph(f'{l["pct_acumulado"]}%'.replace(".", ","), est["centro"])])
        if l["principal"]:
            extra.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#fff3cd")))
    out.append(_tabela(linhas, [40, 22, 14, 16], largura * 0.7, extra=extra, zebra=False))

    out.append(CondPageBreak(60 * mm))
    out.append(Paragraph(_ic(ctx, "⏳") + f'Faturamento previsto e ainda não realizado - {moeda(d["fat_pend_total"])}', est["sub"]))
    cab = [Paragraph(c, est["cab"]) for c in ("Pedido", "Cliente", "Itens", "Previsto desde", "Dias de atraso", "Valor")]
    linhas = [cab]
    for r in d["fat_pend"]:
        linhas.append([Paragraph(f'<b>{_t(r["pedido"])}</b>', est["normal"]), Paragraph(_t((r["cliente"] or "")[:30]), est["normal"]),
                       Paragraph(str(r["itens"]), est["centro"]), Paragraph(_d(r["mais_antigo"]), est["normal"]),
                       _celula_cor(f'{r["dias"]}d' if r["dias"] is not None else "-", "amarelo", est["centro"]),
                       Paragraph(_t(moeda(r["valor"])), est["direita"])])
    if len(linhas) == 1:
        linhas.append([Paragraph("Sem pendências de faturamento previsto.", est["normal"]), "", "", "", "", ""])
    out.append(_tabela(linhas, [14, 34, 7, 14, 13, 18], largura * 0.8))
    return out


# ---------------------------------------------------------------------------
# Geradores públicos
# ---------------------------------------------------------------------------

def gerar_comando(ctx, d):
    return _documento(ctx, "Centro de Comando - Resumo executivo", "Gestão Produção e Operação - visao de 1 pagina com semaforos, gargalos e decisoes",
                      lambda c, w: sec_comando(c, d, w))


def gerar_alertas(ctx, d):
    return _documento(ctx, "Alertas e Gargalos", "Tudo que exige ação agora: atrasados, vencendo, estações criticas e faturamento atrasado",
                      lambda c, w: sec_alertas(c, d, w))


def gerar_faturamento(ctx, d):
    return _documento(ctx, "Faturamento e Projeção", f'Previsto x realizado, projeção de 6 meses e faturamento atrasado - {d["mes_label"]}',
                      lambda c, w: sec_faturamento(c, d, w))


def gerar_leadtime(ctx, gaps, lt_estacao):
    return _documento(ctx, "Lead Time e Gaps", "Lead time por estação e gaps da cadeia liberação -> NF -> expedição -> entrega",
                      lambda c, w: sec_leadtime_gaps(c, gaps, lt_estacao, w))


def gerar_pacote(ctx, comando, alertas, gaps, lt_estacao, faturamento):
    def construir(c, w):
        out = sec_comando(c, comando, w)
        for sec in (lambda: sec_alertas(c, alertas, w), lambda: sec_leadtime_gaps(c, gaps, lt_estacao, w), lambda: sec_faturamento(c, faturamento, w)):
            out.append(PageBreak())
            out += sec()
        return out
    return _documento(ctx, "Pacote Gerencial", "Resumo executivo - Alertas e Gargalos - Lead Time e Gaps - Faturamento e Projeção", construir)
