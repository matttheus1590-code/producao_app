"""Painel semanal ilustrativo (PDF) da Listagem Geral — pedido do Bruno,
29/10/2026: "relatório SOMENTE UMA FOLHA, na horizontal, ilustração de todos
os pedidos da semana, uma linha por pedido, principalmente com datas e se o
PCP já liberou a liberação efetiva (quando fica verde); sem detalhar itens,
só valor total do pedido, modalidade de frete e região; para imprimir, bater
o olho e ver a semana de forma rápida; diversas semanas = uma semana por
folha".

Módulo independente (não importa o app): recebe as semanas já agregadas por
pedido (ver _dados_painel_semanal_visual no app.py) e só apresenta. Só
glifos Latin-1 (nada de setas/✓, que a fonte padrão desenha como quadrado
preto)."""

import io
import os
from datetime import date

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas as _canvas
from reportlab.platypus import (
    BaseDocTemplate, Frame, HRFlowable, Image, PageBreak, PageTemplate, Paragraph, Spacer, Table, TableStyle,
)

NAVY = "#1b2a4a"
VERDE = "#198754"
VERDE_CLARO = "#e3f5e8"
VERMELHO = "#c0392b"
AMARELO = "#d99a00"
CINZA = "#6c757d"
AZUL = "#0d6efd"
BORDA = colors.HexColor("#d5dbe5")

_DIAS = ["seg", "ter", "qua", "qui", "sex", "sáb", "dom"]
_MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro",
          "outubro", "novembro", "dezembro"]

# situação -> (cor de fundo da célula, cor do texto, rótulo curto, ordem na legenda)
_SIT = {
    "liberado": (VERDE, "#ffffff", "LIBERADO"),
    "parcial": ("#7bc47f", "#10331a", "PARCIAL"),
    "pendente": ("#dbe5f4", NAVY, "A LIBERAR"),
    "atrasado": (VERMELHO, "#ffffff", "ATRASADO"),
    "sem_previsao": ("#9aa1a9", "#ffffff", "SEM PREVISÃO"),
}


def _d(d, ano_ref):
    """dd/mm dia-da-semana (com ano de 2 dígitos só quando difere da semana)."""
    if not d:
        return "—"
    base = d.strftime("%d/%m")
    if ano_ref and d.year != ano_ref:
        base += d.strftime("/%y")
    return f"{base} {_DIAS[d.weekday()]}"


def _frete_curto(frete):
    """CIF/FOB como estão; outros valores (ex.: EXPORTAÇÃO) abreviados pra
    não invadirem a coluna ao lado."""
    f = (frete or "—").strip().upper() or "—"
    return f if len(f) <= 4 else f[:3] + "."


def _cortar(txt, fonte, tam, largura):
    """Corta com reticências para caber em UMA linha (sem quebra)."""
    txt = (txt or "—").strip() or "—"
    if stringWidth(txt, fonte, tam) <= largura:
        return txt
    while len(txt) > 1 and stringWidth(txt + "...", fonte, tam) > largura:
        txt = txt[:-1]
    return txt.rstrip() + "..."


class _CanvasNumerado(_canvas.Canvas):
    """Canvas que sabe o total de folhas ("Folha 2 de 4") e desenha o rodapé."""

    rodape_txt = ""

    def __init__(self, *a, **k):
        super().__init__(*a, **k)
        self._salvas = []

    def showPage(self):
        self._salvas.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total = len(self._salvas)
        for estado in self._salvas:
            self.__dict__.update(estado)
            self._rodape(total)
            super().showPage()
        super().save()

    def _rodape(self, total):
        larg, _ = landscape(A4)
        self.setStrokeColor(BORDA)
        self.setLineWidth(0.5)
        self.line(8 * mm, 9.2 * mm, larg - 8 * mm, 9.2 * mm)
        self.setFont("Helvetica", 6.4)
        self.setFillColor(colors.HexColor("#6b7280"))
        self.drawString(8 * mm, 5.6 * mm, self.rodape_txt)
        self.drawRightString(larg - 8 * mm, 5.6 * mm, f"Folha {self._pageNumber} de {total}")


def gerar(semanas, agora, logo_path, moeda, filtros_txt="", hoje=None):
    """semanas = [{"titulo","periodo_txt","mes_txt","ano","pedidos":[...]}, ...]
    Cada pedido: pedido_venda, cliente, frete, regiao, uf, data_inclusao,
    prazo_cliente, lib_prevista, lib_real, situacao, itens_total,
    itens_liberados, valor, parcial_semana, folga_dias."""
    hoje = hoje or agora.date()
    buf = io.BytesIO()
    pag = landscape(A4)
    largura_pag, altura_pag = pag
    mg_l = mg_r = 8 * mm
    mg_t = 7 * mm
    mg_b = 11.5 * mm
    largura = largura_pag - mg_l - mg_r
    altura = altura_pag - mg_t - mg_b

    class _Canvas(_CanvasNumerado):
        rodape_txt = (f"Painel semanal da Listagem Geral · emitido em {agora.strftime('%d/%m/%Y')} às {agora.strftime('%H:%M')} "
                      f"(horário de Brasília)" + (f" · {filtros_txt}" if filtros_txt else ""))

    doc = BaseDocTemplate(buf, pagesize=pag, leftMargin=mg_l, rightMargin=mg_r, topMargin=mg_t, bottomMargin=mg_b,
                          title="Painel semanal — Listagem Geral", author="Gestão da Produção")
    frame = Frame(mg_l, mg_b, largura, altura, leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0, id="f")
    doc.addPageTemplates([PageTemplate(id="p", frames=[frame])])

    est = getSampleStyleSheet()
    base = ParagraphStyle("b", parent=est["Normal"], fontSize=7, leading=8.4)
    titulo = ParagraphStyle("t", parent=est["Title"], fontSize=19, leading=21, alignment=0, spaceAfter=0,
                            textColor=colors.HexColor(NAVY))
    sub = ParagraphStyle("sub", parent=base, fontSize=9, leading=11.5, textColor=colors.HexColor("#3b4559"))
    nota = ParagraphStyle("nota", parent=base, fontSize=6.5, leading=8, textColor=colors.HexColor("#5b6472"))

    def P(t, st=None):
        return Paragraph(t, st or base)

    story = []
    n_semanas = len(semanas)

    for idx, sem in enumerate(semanas):
        peds = sem["pedidos"]
        ano = sem.get("ano")
        flow_topo = []

        # ---------------- agregados da semana ----------------
        n = len(peds)
        valor = sum(p["valor"] for p in peds)
        # Bruno, 29/10/2026: "não quero que cite atraso, somente liberado e a
        # liberar" — só 2 situações; qualquer pedido sem liberação efetiva
        # completa conta como "a liberar".
        lib = [p for p in peds if p["situacao"] == "liberado"]
        pend = [p for p in peds if p["situacao"] != "liberado"]
        v_lib = sum(p["valor"] for p in lib)
        v_pend = valor - v_lib
        v_cif = sum(p["valor"] for p in peds if (p["frete"] or "").upper() == "CIF")
        v_fob = sum(p["valor"] for p in peds if (p["frete"] or "").upper() == "FOB")
        pct_lib = (v_lib / valor * 100) if valor else 0
        pct_ped = (len(lib) / n * 100) if n else 0

        # ---------------- cabeçalho ----------------
        logo = (Image(logo_path, width=42 * mm, height=42 * mm * (63 / 261)) if logo_path and os.path.exists(logo_path) else P(""))
        bloco = [
            P(f'PAINEL SEMANAL &nbsp;·&nbsp; {sem["titulo"]}', titulo),
            P(f'<b>{sem["periodo_txt"]}</b>' + (f' &nbsp;·&nbsp; {sem["mes_txt"]}' if sem.get("mes_txt") else ""), sub),
        ]
        cab = Table([[logo, bloco]], colWidths=[48 * mm, largura - 48 * mm])
        cab.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                 ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 0),
                                 ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
        flow_topo += [cab, Spacer(1, 1.3 * mm), HRFlowable(width="100%", thickness=1.4, color=colors.HexColor(NAVY)),
                      Spacer(1, 1.8 * mm)]

        # ---------------- KPIs ----------------
        def kpi(valor_txt, rot, cor, sub_txt=""):
            return [
                P(f'<font color="{cor}"><b>{valor_txt}</b></font>', ParagraphStyle("kv", parent=base, fontSize=14, leading=16, alignment=1)),
                P(rot, ParagraphStyle("kr", parent=base, fontSize=6.7, leading=7.8, alignment=1, textColor=colors.HexColor("#4a5365"))),
                P(sub_txt or "&nbsp;", ParagraphStyle("ks", parent=base, fontSize=6.5, leading=7.6, alignment=1, textColor=colors.HexColor("#6b7280"))),
            ]

        cards = [
            kpi(str(n), "PEDIDOS NA SEMANA", NAVY, f"{len(set((p['cliente'] or '').strip().upper() for p in peds))} cliente(s)"),
            kpi(moeda(valor), "VALOR TOTAL", NAVY, f"CIF {moeda(v_cif)} · FOB {moeda(v_fob)}"),
            kpi(f"{len(lib)}", "LIBERADOS (PCP)", VERDE, f"{moeda(v_lib)} · {pct_lib:.0f}% do valor"),
            kpi(f"{len(pend)}", "A LIBERAR", NAVY, f"{moeda(v_pend)} · {100 - pct_lib:.0f}% do valor" if valor else moeda(v_pend)),
        ]
        larg_card = largura / len(cards)
        kp = Table([[c for c in cards]], colWidths=[larg_card] * len(cards))
        ks = [("BOX", (0, 0), (-1, -1), 0.6, BORDA), ("INNERGRID", (0, 0), (-1, -1), 0.6, BORDA),
              ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f6f8fb")), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
              ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]
        ks.append(("BACKGROUND", (2, 0), (2, 0), colors.HexColor(VERDE_CLARO)))
        kp.setStyle(TableStyle(ks))
        flow_topo.append(kp)
        flow_topo.append(Spacer(1, 1.6 * mm))

        # ---------------- barra de progresso (por valor) ----------------
        if valor > 0:
            segs = [(v_lib, VERDE, "Liberado"), (v_pend, "#b9c8e3", "A liberar")]
            segs = [s for s in segs if s[0] > 0]
            larg_barra = largura - 62 * mm
            ws = [max(s[0] / valor * larg_barra, 0.1) for s in segs]
            rot = P(f'<b>Progresso da liberação (valor)</b><br/><font size="6.4" color="#5b6472">{pct_lib:.0f}% liberado pelo PCP · '
                    f'{pct_ped:.0f}% dos pedidos</font>', ParagraphStyle("pr", parent=base, fontSize=7.2, leading=8.6))
            barra = Table([[""] * len(segs)], colWidths=ws, rowHeights=[8])
            barra.setStyle(TableStyle([("BACKGROUND", (i, 0), (i, 0), colors.HexColor(s[1])) for i, s in enumerate(segs)] +
                                      [("BOX", (0, 0), (-1, -1), 0.4, colors.HexColor("#8a93a3")),
                                       ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                                       ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
            linha = Table([[rot, barra]], colWidths=[62 * mm, larg_barra])
            linha.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                       ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 0),
                                       ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
            flow_topo += [linha, Spacer(1, 1.5 * mm)]

        # ---------------- região (texto compacto) ----------------
        reg = {}
        for p in peds:
            r = reg.setdefault(p["regiao"] or "—", [0, 0.0])
            r[0] += 1
            r[1] += p["valor"]
        if reg:
            txt = " &nbsp;|&nbsp; ".join(f'<b>{k}</b> {v[0]} ped. · {moeda(v[1])}'
                                          for k, v in sorted(reg.items(), key=lambda kv: -kv[1][1]))
            flow_topo += [P(f'<b>Por região:</b> {txt}', ParagraphStyle("rg", parent=base, fontSize=6.9, leading=8.4)),
                          Spacer(1, 1.6 * mm)]

        # medir o topo para ajustar a altura das linhas à folha
        h_topo = 0
        for f in flow_topo:
            _, h = f.wrap(largura, 9999)
            h_topo += h + (getattr(f, "getSpaceBefore", lambda: 0)() or 0)
        story += flow_topo

        # ---------------- tabela de pedidos ----------------
        cab_txt = ["PEDIDO", "CLIENTE", "FRETE", "REGIÃO / UF", "INCLUSÃO", "PRAZO CLIENTE", "LIB. PREVISTA PCP",
                   "LIB. EFETIVA", "FOLGA", "SITUAÇÃO", "VALOR TOTAL"]
        pesos = [8.5, 25, 5.2, 12, 9.5, 10, 10.5, 10, 6.5, 13, 12.5]
        cw = [p / sum(pesos) * largura for p in pesos]
        h_total = 9 * mm  # linha de total
        disponivel = altura - h_topo - 13 * mm - h_total
        n_lin = max(len(peds), 1)
        h_cab = 15
        h_lin = (disponivel - h_cab) / n_lin
        h_lin = max(7.2, min(19, h_lin))
        fs = max(5.3, min(9.2, h_lin * 0.58))
        fn, fb = "Helvetica", "Helvetica-Bold"

        dados = [cab_txt]
        estilo = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(NAVY)), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), fb), ("FONTSIZE", (0, 0), (-1, 0), 6.4),
            ("ALIGN", (0, 0), (-1, 0), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("FONTNAME", (0, 1), (-1, -1), fn), ("FONTSIZE", (0, 1), (-1, -1), fs),
            ("ALIGN", (0, 1), (1, -1), "LEFT"), ("ALIGN", (2, 1), (-2, -1), "CENTER"), ("ALIGN", (-1, 1), (-1, -1), "RIGHT"),
            ("LINEBELOW", (0, 1), (-1, -1), 0.3, BORDA),
            ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ("FONTNAME", (0, 1), (0, -1), fb),
            ("LEADING", (0, 0), (-1, 0), 7.6), ("LEADING", (0, 1), (-1, -1), fs * 1.1),
        ]
        tem_parcial_semana = False
        for i, p in enumerate(peds, start=1):
            sit = "liberado" if p["situacao"] == "liberado" else "pendente"
            bg, fg, rotulo = _SIT[sit]
            val_txt = moeda(p["valor"]) + ("*" if p["parcial_semana"] else "")
            tem_parcial_semana = tem_parcial_semana or p["parcial_semana"]
            folga = p["folga_dias"]
            if folga is None:
                folga_txt, folga_cor = "—", CINZA
            elif folga < 0:
                folga_txt, folga_cor = f"{folga}d", VERMELHO
            elif folga <= 3:
                folga_txt, folga_cor = f"+{folga}d", AMARELO
            else:
                folga_txt, folga_cor = f"+{folga}d", VERDE
            dados.append([
                _cortar(p["pedido_venda"], fb, fs, cw[0] - 6),
                _cortar(p["cliente"], fn, fs, cw[1] - 6),
                _frete_curto(p["frete"]),
                _cortar(f'{p["regiao"]} · {p["uf"]}' if p["uf"] and p["uf"] != "—" else p["regiao"], fn, fs, cw[3] - 6),
                _d(p["data_inclusao"], ano), _d(p["prazo_cliente"], ano), _d(p["lib_prevista"], ano),
                _d(p["lib_real"], ano) if (p["lib_real"] and sit == "liberado") else "—",
                folga_txt, rotulo, val_txt,
            ])
            if sit == "liberado":
                estilo.append(("BACKGROUND", (0, i), (8, i), colors.HexColor(VERDE_CLARO)))
                estilo.append(("BACKGROUND", (10, i), (10, i), colors.HexColor(VERDE_CLARO)))
                estilo.append(("FONTNAME", (7, i), (7, i), fb))
                estilo.append(("TEXTCOLOR", (7, i), (7, i), colors.HexColor("#0f5a2d")))
            elif i % 2 == 0:
                estilo.append(("BACKGROUND", (0, i), (8, i), colors.HexColor("#f6f8fb")))
                estilo.append(("BACKGROUND", (10, i), (10, i), colors.HexColor("#f6f8fb")))
            estilo.append(("BACKGROUND", (9, i), (9, i), colors.HexColor(bg)))
            estilo.append(("TEXTCOLOR", (9, i), (9, i), colors.HexColor(fg)))
            estilo.append(("FONTNAME", (9, i), (9, i), fb))
            estilo.append(("FONTSIZE", (9, i), (9, i), max(fs - 0.6, 5.6)))
            estilo.append(("TEXTCOLOR", (8, i), (8, i), colors.HexColor(folga_cor)))
            estilo.append(("FONTNAME", (8, i), (8, i), fb))
            frete = (p["frete"] or "").upper()
            if frete == "CIF":
                estilo.append(("TEXTCOLOR", (2, i), (2, i), colors.HexColor("#0b5ed7")))
                estilo.append(("FONTNAME", (2, i), (2, i), fb))
            elif frete == "FOB":
                estilo.append(("TEXTCOLOR", (2, i), (2, i), colors.HexColor("#b35c00")))
                estilo.append(("FONTNAME", (2, i), (2, i), fb))
        if not peds:
            dados.append(["Nenhum pedido planejado nesta semana"] + [""] * 10)
            estilo.append(("SPAN", (0, 1), (-1, 1)))
            estilo.append(("ALIGN", (0, 1), (-1, 1), "CENTER"))

        # linha de total
        dados.append(["TOTAL DA SEMANA", "", f"{n} pedido(s)", "", "", "", "", f"{len(lib)} liberados / {len(pend)} a liberar", "", "", moeda(valor)])
        it = len(dados) - 1
        estilo += [
            ("BACKGROUND", (0, it), (-1, it), colors.HexColor("#dbe5f4")), ("FONTNAME", (0, it), (-1, it), fb),
            ("TEXTCOLOR", (0, it), (-1, it), colors.HexColor(NAVY)), ("LINEABOVE", (0, it), (-1, it), 1.1, colors.HexColor(NAVY)),
            ("SPAN", (0, it), (1, it)), ("SPAN", (2, it), (6, it)), ("SPAN", (7, it), (9, it)),
            ("ALIGN", (2, it), (6, it), "CENTER"), ("ALIGN", (7, it), (9, it), "CENTER"), ("FONTSIZE", (0, it), (-1, it), fs + 0.6),
        ]
        alturas = [h_cab] + [h_lin] * (len(dados) - 2) + [max(h_lin + 2, 13)]
        tabela = Table(dados, colWidths=cw, rowHeights=alturas, repeatRows=1)
        tabela.setStyle(TableStyle(estilo))
        story.append(tabela)

        # ---------------- legenda ----------------
        story.append(Spacer(1, 1.6 * mm))
        leg = (f'<font color="{VERDE}"><b>LIBERADO</b></font> = PCP já lançou a liberação efetiva (linha verde) · '
               f'<b>A LIBERAR</b> = ainda sem liberação efetiva · <b>FOLGA</b> = prazo do cliente menos a liberação '
               f'(efetiva, ou a prevista enquanto não liberou); vermelho = negativa, amarelo = até 3 dias.'
               + (' · <b>*</b> pedido com itens também em outras semanas: valor = parte planejada nesta semana.' if tem_parcial_semana else ""))
        story.append(P(leg, nota))

        if idx < n_semanas - 1:
            story.append(PageBreak())

    doc.build(story, canvasmaker=_Canvas)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# MÊS INTEIRO EM UMA FOLHA (pedido do Bruno, 18/11/2026): "visualizar todo o mês dentro de uma
# folha, uma linha representando cada pedido, agrupando por cores o que representa cada semana…
# com isso consigo ter uma visão micro e macro… gerando o faturamento e todas as colunas já
# existentes: folga, situação, frete etc." — sempre UMA página A4 horizontal: a altura das linhas
# (e a fonte) se ajusta à quantidade de pedidos.
# ---------------------------------------------------------------------------
# cor forte (faixa/legenda), cor clara (fundo das linhas) por semana; 0 = sem semana definida
CORES_SEMANA = {
    1: ("#2f6fd6", "#e4edfb"),
    2: ("#e08a00", "#fdf0da"),
    3: ("#7a4fc2", "#efe8fa"),
    4: ("#0f9488", "#dcf3f0"),
    5: ("#cf3d78", "#fbe4ed"),
    0: ("#6c757d", "#eceef0"),
}


def _cor_semana(sem):
    k = sem.get("cor")
    if k is None:
        k = sem.get("num") or 0
    return CORES_SEMANA.get(k if k in CORES_SEMANA else ((k - 1) % 5) + 1, CORES_SEMANA[0])


def gerar_mes(semanas, agora, logo_path, moeda, titulo_mes="", filtros_txt="", hoje=None):
    """semanas = lista no mesmo formato de `gerar` (já em ordem); cada semana pode trazer "cor" (1..5) para
    forçar a cor (usado quando semanas de meses diferentes vão juntas). Sai SEMPRE em 1 folha."""
    hoje = hoje or agora.date()
    buf = io.BytesIO()
    pag = landscape(A4)
    largura_pag, altura_pag = pag
    mg_l = mg_r = 8 * mm
    mg_t = 6 * mm
    mg_b = 11.5 * mm
    largura = largura_pag - mg_l - mg_r
    altura = altura_pag - mg_t - mg_b

    class _Canvas(_CanvasNumerado):
        rodape_txt = (f"Painel semanal · mês inteiro em 1 folha · emitido em {agora.strftime('%d/%m/%Y')} às {agora.strftime('%H:%M')} "
                      f"(horário de Brasília)" + (f" · {filtros_txt}" if filtros_txt else ""))

    doc = BaseDocTemplate(buf, pagesize=pag, leftMargin=mg_l, rightMargin=mg_r, topMargin=mg_t, bottomMargin=mg_b,
                          title="Painel semanal — mês inteiro", author="Gestão da Produção")
    frame = Frame(mg_l, mg_b, largura, altura, leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0, id="f")
    doc.addPageTemplates([PageTemplate(id="p", frames=[frame])])

    est = getSampleStyleSheet()
    base = ParagraphStyle("b", parent=est["Normal"], fontSize=7, leading=8.4)
    titulo = ParagraphStyle("t", parent=est["Title"], fontSize=17, leading=19, alignment=0, spaceAfter=0,
                            textColor=colors.HexColor(NAVY))
    sub = ParagraphStyle("sub", parent=base, fontSize=8.6, leading=10.5, textColor=colors.HexColor("#3b4559"))
    nota = ParagraphStyle("nota", parent=base, fontSize=6.2, leading=7.4, textColor=colors.HexColor("#5b6472"))

    def P(t, st=None):
        return Paragraph(t, st or base)

    fn, fb = "Helvetica", "Helvetica-Bold"
    todos = [p for s in semanas for p in s["pedidos"]]
    n = len(todos)
    valor = sum(p["valor"] for p in todos)
    lib = [p for p in todos if p["situacao"] == "liberado"]
    pend = [p for p in todos if p["situacao"] != "liberado"]
    v_lib = sum(p["valor"] for p in lib)
    v_pend = valor - v_lib
    v_cif = sum(p["valor"] for p in todos if (p["frete"] or "").upper() == "CIF")
    v_fob = sum(p["valor"] for p in todos if (p["frete"] or "").upper() == "FOB")
    pct_lib = (v_lib / valor * 100) if valor else 0

    story = []
    flow_topo = []
    # ---------------- cabeçalho ----------------
    logo = (Image(logo_path, width=38 * mm, height=38 * mm * (63 / 261)) if logo_path and os.path.exists(logo_path) else P(""))
    bloco = [P("PAINEL SEMANAL &nbsp;·&nbsp; MÊS INTEIRO", titulo),
             P(f"<b>{titulo_mes}</b> &nbsp;·&nbsp; {len(semanas)} semana(s) do Planejamento PCP &nbsp;·&nbsp; "
               f"{n} pedido(s) &nbsp;·&nbsp; <b>faturamento {moeda(valor)}</b>", sub)]
    cab = Table([[logo, bloco]], colWidths=[44 * mm, largura - 44 * mm])
    cab.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                             ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 0),
                             ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
    flow_topo += [cab, Spacer(1, 1 * mm), HRFlowable(width="100%", thickness=1.3, color=colors.HexColor(NAVY)), Spacer(1, 1.4 * mm)]

    # ---------------- faturamento por semana (macro): 1 coluna por semana + total do mês ----------------
    rot_w = 30 * mm
    cols = len(semanas) + 1
    larg_col = (largura - rot_w) / cols
    sm = ParagraphStyle("sm", parent=base, fontSize=6.6, leading=7.8, alignment=1)
    cab_row = [P("<b>FATURAMENTO</b>", ParagraphStyle("fh", parent=base, fontSize=6.6, leading=7.8))]
    l_ped, l_val, l_lib, l_pen, l_fr = ([P("<b>Pedidos</b>", sm)], [P("<b>Valor total</b>", sm)], [P("<b>Liberado (PCP)</b>", sm)],
                                         [P("<b>A liberar</b>", sm)], [P("<b>CIF / FOB</b>", sm)])
    estilo_f = [("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("BOX", (0, 0), (-1, -1), 0.5, BORDA), ("INNERGRID", (0, 0), (-1, -1), 0.3, BORDA),
                ("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
                ("LEFTPADDING", (0, 0), (-1, -1), 2), ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                ("BACKGROUND", (0, 1), (0, -1), colors.HexColor("#f6f8fb"))]
    for j, sem in enumerate(semanas, start=1):
        forte, claro = _cor_semana(sem)
        r = sem.get("resumo") or {}
        peds = sem["pedidos"]
        v = sum(p["valor"] for p in peds)
        vl = sum(p["valor"] for p in peds if p["situacao"] == "liberado")
        nl = sum(1 for p in peds if p["situacao"] == "liberado")
        rotulo = sem["titulo"].replace("SEMANA ", "SEM ") if sem.get("num") else "SEM SEMANA"
        if sem.get("mes_txt") and len(semanas) > 0 and sem.get("rotulo_mes"):
            rotulo += f" · {sem['rotulo_mes']}"
        cab_row.append(P(f'<font color="#ffffff"><b>{rotulo}</b></font>', sm))
        l_ped.append(P(str(len(peds)), sm))
        l_val.append(P(f"<b>{moeda(v)}</b>", sm))
        l_lib.append(P(f'<font color="{VERDE}">{moeda(vl)}</font> ({(vl / v * 100) if v else 0:.0f}%)', sm))
        l_pen.append(P(moeda(v - vl), sm))
        l_fr.append(P(f'{moeda(sum(p["valor"] for p in peds if (p["frete"] or "").upper() == "CIF"))} / '
                      f'{moeda(sum(p["valor"] for p in peds if (p["frete"] or "").upper() == "FOB"))}',
                      ParagraphStyle("sm2", parent=sm, fontSize=5.9, leading=7)))
        estilo_f.append(("BACKGROUND", (j, 0), (j, 0), colors.HexColor(forte)))
        estilo_f.append(("BACKGROUND", (j, 1), (j, -1), colors.HexColor(claro)))
    cab_row.append(P('<font color="#ffffff"><b>TOTAL DO MÊS</b></font>', sm))
    l_ped.append(P(f"<b>{n}</b>", sm))
    l_val.append(P(f"<b>{moeda(valor)}</b>", sm))
    l_lib.append(P(f'<font color="{VERDE}"><b>{moeda(v_lib)}</b></font> ({pct_lib:.0f}%)', sm))
    l_pen.append(P(f"<b>{moeda(v_pend)}</b>", sm))
    l_fr.append(P(f"{moeda(v_cif)} / {moeda(v_fob)}", ParagraphStyle("sm3", parent=sm, fontSize=5.9, leading=7)))
    estilo_f.append(("BACKGROUND", (cols, 0), (cols, 0), colors.HexColor(NAVY)))
    estilo_f.append(("BACKGROUND", (cols, 1), (cols, -1), colors.HexColor("#dbe5f4")))
    fat = Table([cab_row, l_ped, l_val, l_lib, l_pen, l_fr], colWidths=[rot_w] + [larg_col] * cols)
    fat.setStyle(TableStyle(estilo_f))
    flow_topo += [fat, Spacer(1, 1.6 * mm)]

    h_topo = 0
    for f in flow_topo:
        _, h = f.wrap(largura, 9999)
        h_topo += h + (getattr(f, "getSpaceBefore", lambda: 0)() or 0)
    story += flow_topo

    # ---------------- tabela única: faixa colorida por semana + 1 linha por pedido ----------------
    cab_txt = ["SEM", "PEDIDO", "CLIENTE", "FRETE", "REGIÃO / UF", "INCLUSÃO", "PRAZO CLIENTE", "LIB. PREVISTA PCP",
               "LIB. EFETIVA", "FOLGA", "SITUAÇÃO", "VALOR TOTAL"]
    pesos = [4.2, 8.3, 23, 5.2, 11.5, 9.2, 9.6, 10.2, 9.6, 6.3, 11.2, 12.2]
    cw = [p / sum(pesos) * largura for p in pesos]
    n_faixas = sum(1 for s in semanas if s["pedidos"])
    n_linhas = max(n, 1) + n_faixas + 1  # + linha de total
    h_cab = 12
    h_legenda = 7 * mm
    disponivel = altura - h_topo - h_legenda - 2 * mm
    h_lin = (disponivel - h_cab) / n_linhas
    h_lin = max(1.8, min(15, h_lin))  # sem piso prático: a folha é UMA só, mesmo com muitos pedidos (a letra encolhe)
    fs = max(1.8, min(8.6, h_lin * 0.62))
    fs_cab = max(4.2, min(6.4, h_cab * 0.5))

    dados = [cab_txt]
    estilo = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(NAVY)), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), fb), ("FONTSIZE", (0, 0), (-1, 0), fs_cab), ("LEADING", (0, 0), (-1, 0), fs_cab * 1.15),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTNAME", (0, 1), (-1, -1), fn), ("FONTSIZE", (0, 1), (-1, -1), fs), ("LEADING", (0, 1), (-1, -1), fs * 1.08),
        ("ALIGN", (0, 1), (0, -1), "CENTER"), ("ALIGN", (1, 1), (2, -1), "LEFT"), ("ALIGN", (3, 1), (-2, -1), "CENTER"),
        ("ALIGN", (-1, 1), (-1, -1), "RIGHT"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2.2), ("RIGHTPADDING", (0, 0), (-1, -1), 2.2),
        ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("FONTNAME", (1, 1), (1, -1), fb),
    ]
    alturas = [h_cab]
    tem_parcial = False
    for sem in semanas:
        peds = sem["pedidos"]
        if not peds:
            continue
        forte, claro = _cor_semana(sem)
        rf = len(dados)
        v_s = sum(p["valor"] for p in peds)
        vl_s = sum(p["valor"] for p in peds if p["situacao"] == "liberado")
        nl_s = sum(1 for p in peds if p["situacao"] == "liberado")
        nome = sem["titulo"] + (f' · {sem["mes_txt"]}' if sem.get("mes_txt") and sem.get("rotulo_mes") else "")
        faixa = (f'{nome}  ·  {sem["periodo_txt"]}  ·  {len(peds)} pedido(s)  ·  valor {moeda(v_s)}  ·  '
                 f'liberado {moeda(vl_s)} ({(vl_s / v_s * 100) if v_s else 0:.0f}%)  ·  a liberar {moeda(v_s - vl_s)}')
        dados.append([faixa] + [""] * 11)
        estilo += [("SPAN", (0, rf), (-1, rf)), ("BACKGROUND", (0, rf), (-1, rf), colors.HexColor(forte)),
                   ("TEXTCOLOR", (0, rf), (-1, rf), colors.white), ("FONTNAME", (0, rf), (-1, rf), fb),
                   ("ALIGN", (0, rf), (-1, rf), "LEFT"), ("FONTSIZE", (0, rf), (-1, rf), fs + 0.3)]
        alturas.append(h_lin)
        marca = str(sem["num"]) if sem.get("num") else "—"
        for p in peds:
            i = len(dados)
            sit = "liberado" if p["situacao"] == "liberado" else "pendente"
            bg, fg, rotulo = _SIT[sit]
            val_txt = moeda(p["valor"]) + ("*" if p["parcial_semana"] else "")
            tem_parcial = tem_parcial or p["parcial_semana"]
            folga = p["folga_dias"]
            if folga is None:
                folga_txt, folga_cor = "—", CINZA
            elif folga < 0:
                folga_txt, folga_cor = f"{folga}d", VERMELHO
            elif folga <= 3:
                folga_txt, folga_cor = f"+{folga}d", AMARELO
            else:
                folga_txt, folga_cor = f"+{folga}d", VERDE
            ano = sem.get("ano")
            dados.append([
                marca,
                _cortar(p["pedido_venda"], fb, fs, cw[1] - 4.4),
                _cortar(p["cliente"], fn, fs, cw[2] - 4.4),
                _frete_curto(p["frete"]),
                _cortar(f'{p["regiao"]} · {p["uf"]}' if p["uf"] and p["uf"] != "—" else p["regiao"], fn, fs, cw[4] - 4.4),
                _d(p["data_inclusao"], ano), _d(p["prazo_cliente"], ano), _d(p["lib_prevista"], ano),
                _d(p["lib_real"], ano) if (p["lib_real"] and sit == "liberado") else "—",
                folga_txt, rotulo, val_txt,
            ])
            alturas.append(h_lin)
            estilo.append(("BACKGROUND", (1, i), (-1, i), colors.HexColor(claro)))
            estilo.append(("BACKGROUND", (0, i), (0, i), colors.HexColor(forte)))
            estilo.append(("TEXTCOLOR", (0, i), (0, i), colors.white))
            estilo.append(("FONTNAME", (0, i), (0, i), fb))
            estilo.append(("LINEBELOW", (0, i), (-1, i), 0.25, colors.white))
            if sit == "liberado":
                estilo.append(("FONTNAME", (8, i), (8, i), fb))
                estilo.append(("TEXTCOLOR", (8, i), (8, i), colors.HexColor("#0f5a2d")))
            estilo.append(("BACKGROUND", (10, i), (10, i), colors.HexColor(bg)))
            estilo.append(("TEXTCOLOR", (10, i), (10, i), colors.HexColor(fg)))
            estilo.append(("FONTNAME", (10, i), (10, i), fb))
            estilo.append(("FONTSIZE", (10, i), (10, i), max(fs - 0.4, 1.8)))
            estilo.append(("TEXTCOLOR", (9, i), (9, i), colors.HexColor(folga_cor)))
            estilo.append(("FONTNAME", (9, i), (9, i), fb))
            frete = (p["frete"] or "").upper()
            if frete == "CIF":
                estilo += [("TEXTCOLOR", (3, i), (3, i), colors.HexColor("#0b5ed7")), ("FONTNAME", (3, i), (3, i), fb)]
            elif frete == "FOB":
                estilo += [("TEXTCOLOR", (3, i), (3, i), colors.HexColor("#b35c00")), ("FONTNAME", (3, i), (3, i), fb)]
    if not todos:
        dados.append(["Nenhum pedido planejado neste período"] + [""] * 11)
        estilo += [("SPAN", (0, 1), (-1, 1)), ("ALIGN", (0, 1), (-1, 1), "CENTER")]
        alturas.append(h_lin)
    # total do mês
    dados.append(["TOTAL DO MÊS", "", "", f"{n} pedido(s)", "", "", "", "", f"{len(lib)} liberados / {len(pend)} a liberar", "", "", moeda(valor)])
    it = len(dados) - 1
    estilo += [
        ("BACKGROUND", (0, it), (-1, it), colors.HexColor("#dbe5f4")), ("FONTNAME", (0, it), (-1, it), fb),
        ("TEXTCOLOR", (0, it), (-1, it), colors.HexColor(NAVY)), ("LINEABOVE", (0, it), (-1, it), 1.0, colors.HexColor(NAVY)),
        ("SPAN", (0, it), (2, it)), ("SPAN", (3, it), (7, it)), ("SPAN", (8, it), (10, it)),
        ("ALIGN", (0, it), (2, it), "LEFT"), ("ALIGN", (3, it), (10, it), "CENTER"), ("FONTSIZE", (0, it), (-1, it), fs + 0.6),
    ]
    alturas.append(h_lin + 1)
    tabela = Table(dados, colWidths=cw, rowHeights=alturas)
    tabela.setStyle(TableStyle(estilo))
    story.append(tabela)

    story.append(Spacer(1, 1.2 * mm))
    leg = (f'Cada cor = uma semana do Planejamento PCP (faixa de título + coluna SEM). '
           f'<font color="{VERDE}"><b>LIBERADO</b></font> = PCP já lançou a liberação efetiva · <b>A LIBERAR</b> = ainda sem liberação efetiva · '
           f'<b>FOLGA</b> = prazo do cliente menos a liberação (efetiva, ou a prevista enquanto não liberou); vermelho = negativa, amarelo = até 3 dias.'
           + (' · <b>*</b> pedido com itens também em outras semanas: valor = parte planejada na semana.' if tem_parcial else ""))
    story.append(P(leg, nota))

    doc.build(story, canvasmaker=_Canvas)
    return buf.getvalue()
