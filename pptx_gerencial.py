# -*- coding: utf-8 -*-
"""Gerador de apresentações PowerPoint (.pptx) gerenciais — 4PIPE Solutions.

Pedido do Bruno (27/10/2026): "um gerador de apresentação automático em
PowerPoint, já projetando números e resultados" nas áreas Planejamento
Mensal PCP/Operação, Estações (conclusões e risco), Painel Diário e Detalhe
do Pedido 360, em dois modos escolhidos no botão: EXECUTIVA (enxuta, alta
direção) e COMPLETA (gerencial, com detalhamento).

Este módulo NÃO importa o app (sem import circular): só recebe dados já
calculados (dicts / objetos com os mesmos atributos do app) e devolve os
bytes do .pptx. Os gráficos são NATIVOS do PowerPoint (editáveis), não
imagens. Layout 16:9, fonte Calibri, paleta 4PIPE.
"""
import io
import os
import re as _re
from collections import Counter, OrderedDict
from datetime import date

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt
from lxml import etree

# ---------------------------------------------------------------- paleta
# Tema 4PIPE conforme o modelo "KPIs Operacionais" enviado pelo Bruno
# (27/10/2026): faixa azul-marinho com logo + barra laranja, pílula laranja do
# período, cards brancos com sombra suave, barras marinho/azul-petróleo com
# destaque laranja, rodapé marinho com "CONFIDENCIAL".
NAVY = "1B3A6B"        # marinho principal (faixas, títulos, barras)
NAVY_2 = "16305C"      # marinho profundo (painel da capa)
ACENTO = "C84B1E"      # laranja 4PIPE (pílulas, barra vertical, destaque)
AZUL = "0D6E8A"        # azul-petróleo (série secundária / "em produção")
VERDE = "2EAE4D"
VERMELHO = "C0392B"
LARANJA = "E07B1F"     # risco "Alto" (distinto do laranja da marca)
AMARELO = "EFA72E"
CINZA = "6B84A8"       # texto secundário / rótulos
CINZA_CLARO = "F4F6FA"
BORDA = "E2E8F2"
TEXTO = "1A2B4A"
CLARO = "AABBCC"       # texto claro sobre marinho
FUNDO = "F5F7FA"
BRANCO = "FFFFFF"

FONTE = "Calibri"
W, H = 13.333, 7.5
MARGEM = 0.55

NIVEL_COR = {3: VERMELHO, 2: LARANJA, 1: AMARELO, 0: VERDE}
NIVEL_TXT = {3: "Crítico", 2: "Alto", 1: "Médio", 0: "No prazo"}


def rgb(hex_):
    return RGBColor.from_string(hex_)


def tom_claro(hex_, f=0.88):
    """Mistura a cor com branco (fundo suave de card)."""
    r, g, b = int(hex_[0:2], 16), int(hex_[2:4], 16), int(hex_[4:6], 16)
    r, g, b = [int(c + (255 - c) * f) for c in (r, g, b)]
    return "%02X%02X%02X" % (r, g, b)


# ---------------------------------------------------------------- formatos
def fmt_int(v):
    try:
        return f"{int(round(v)):,}".replace(",", ".")
    except (TypeError, ValueError):
        return "—"


def fmt_brl(v, compacto=False):
    if v is None:
        return "—"
    v = float(v)
    if compacto:
        a = abs(v)
        if a >= 1_000_000:
            return "R$ " + f"{v / 1_000_000:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".") + " mi"
        if a >= 10_000:
            return "R$ " + f"{v / 1_000:,.0f}".replace(",", ".") + " mil"
        if a >= 100:
            return "R$ " + f"{v:,.0f}".replace(",", ".")
    s = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    if s.endswith(",00"):
        s = s[:-3]
    return "R$ " + s


def fmt_pct(v, casas=0):
    if v is None:
        return "—"
    return f"{v:.{casas}f}%".replace(".", ",")


def fmt_data(d, curto=False):
    if not d:
        return "—"
    return d.strftime("%d/%m") if curto else d.strftime("%d/%m/%Y")


def plural(n, s, p=None):
    return f"{fmt_int(n)} {s if n == 1 else (p or s + 's')}"


_MES_ABREV_EXT = {"jan": "Janeiro", "fev": "Fevereiro", "mar": "Março", "abr": "Abril", "mai": "Maio", "jun": "Junho",
                  "jul": "Julho", "ago": "Agosto", "set": "Setembro", "out": "Outubro", "nov": "Novembro", "dez": "Dezembro"}


def _mes_extenso(label):
    """'Out/2026' | 'Outubro/2026' -> 'Outubro 2026' (texto da pílula laranja)."""
    try:
        m, a = str(label).replace(" ", "").split("/")[:2]
        return f"{_MES_ABREV_EXT.get(m[:3].lower(), m)} {a}"
    except Exception:  # noqa: BLE001
        return str(label or "")


def _periodo_curto(titulo_periodo):
    """Resume o título do período das Estações para caber na pílula."""
    t = titulo_periodo or ""
    if " + " in t:
        return "Multi-período"
    m = _re.search(r"(\d{1,2})\s*/\s*([A-Za-zçÇãÃéÉ]+)\s*/\s*(\d{4})", t)
    if m and "mês inteiro" not in t:
        return f"Sem {int(m.group(1)):02d} · {m.group(2)[:3].title()} {m.group(3)}"
    m = _re.search(r"([A-Za-zçÇãÃéÉ]+)/(\d{4})", t)
    if m:
        return f"{m.group(1)} {m.group(2)}"
    return t[:18]


def _corta(txt, n):
    txt = txt or ""
    return txt if len(txt) <= n else txt[: n - 1].rstrip() + "…"


# ---------------------------------------------------------------- Deck
class Deck:
    def __init__(self, titulo_curto, gerado_em, logo_path=None, periodo=None, base=None):
        self.periodo = periodo      # texto da pílula laranja (ex.: "Outubro 2026")
        self.base = base            # trecho "Base: ..." do rodapé
        self.prs = Presentation()
        self.prs.slide_width = Inches(W)
        self.prs.slide_height = Inches(H)
        self.titulo_curto = titulo_curto
        self.gerado_em = gerado_em
        self.logo_path = logo_path if (logo_path and os.path.exists(logo_path)) else None
        self._blank = self.prs.slide_layouts[6]
        self.n = 0

    # ------------------------------------------------ primitivas
    def _sombra(self, sh):
        """Sombra suave dos cards do modelo (outerShdw 6B84A8 @18%)."""
        eff = sh._element.spPr.find(qn("a:effectLst"))
        if eff is None:
            eff = etree.SubElement(sh._element.spPr, qn("a:effectLst"))
        sd = etree.SubElement(eff, qn("a:outerShdw"), blurRad="63500", dist="25400", dir="5400000", algn="bl", rotWithShape="0")
        c = etree.SubElement(sd, qn("a:srgbClr"), val="6B84A8")
        etree.SubElement(c, qn("a:alpha"), val="18000")

    def _rect(self, slide, x, y, w, h, fill=None, line=None, forma=MSO_SHAPE.RECTANGLE, raio=None, line_w=0.75, sombra=False):
        sh = slide.shapes.add_shape(forma, Inches(x), Inches(y), Inches(w), Inches(h))
        if fill:
            sh.fill.solid()
            sh.fill.fore_color.rgb = rgb(fill)
        else:
            sh.fill.background()
        if line:
            sh.line.color.rgb = rgb(line)
            sh.line.width = Pt(line_w)
        else:
            sh.line.fill.background()
        sh.shadow.inherit = False
        if raio is not None and forma == MSO_SHAPE.ROUNDED_RECTANGLE:
            sh.adjustments[0] = raio
        sh.text_frame.text = ""
        if sombra:
            self._sombra(sh)
        return sh

    def _fundo(self, slide, cor):
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = rgb(cor)

    def _texto(self, slide, x, y, w, h, texto, tam=14, cor=TEXTO, negrito=False, alinh=PP_ALIGN.LEFT,
               ancora=MSO_ANCHOR.TOP, italico=False, margem=0.0, linha=None):
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = ancora
        tf.margin_left = tf.margin_right = Inches(margem)
        tf.margin_top = tf.margin_bottom = Inches(margem)
        partes = texto if isinstance(texto, list) else [texto]
        for i, parte in enumerate(partes):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = alinh
            if linha:
                p.line_spacing = linha
            runs = parte if isinstance(parte, list) else [(parte, {})]
            for r_txt, r_fmt in runs:
                r = p.add_run()
                r.text = str(r_txt)
                f = r.font
                f.name = FONTE
                f.size = Pt(r_fmt.get("tam", tam))
                f.bold = r_fmt.get("negrito", negrito)
                f.italic = r_fmt.get("italico", italico)
                f.color.rgb = rgb(r_fmt.get("cor", cor))
        return tb

    def _notas(self, slide, texto):
        if texto:
            slide.notes_slide.notes_text_frame.text = texto

    # ------------------------------------------------ slides base
    def capa(self, titulo, subtitulo, meta=None, rotulo="4PIPE SOLUTIONS", area="GESTÃO\nOPERAÇÕES", pilula=None):
        """Capa no modelo 4PIPE: painel marinho profundo à esquerda com o logo e o
        título espaçado, barra laranja, título grande à direita e pílula laranja."""
        s = self.prs.slides.add_slide(self._blank)
        self.n += 1
        self._fundo(s, NAVY)
        self._rect(s, 0, 0, 4.8, H, fill=NAVY_2)
        self._rect(s, 4.8, 0, 0.08, H, fill=ACENTO)
        c = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(9.6), Inches(-1.7), Inches(5.6), Inches(5.6))
        c.fill.solid()
        c.fill.fore_color.rgb = rgb(ACENTO)
        # transparência 10% (alpha 10000) como no modelo
        etree.SubElement(c.fill._xPr.find(qn("a:solidFill")).find(qn("a:srgbClr")), qn("a:alpha"), val="10000")
        c.line.fill.background()
        c.shadow.inherit = False
        if self.logo_path:
            self._rect(s, 0.47, 0.56, 2.48, 0.67, fill=BRANCO, forma=MSO_SHAPE.ROUNDED_RECTANGLE, raio=0.2)
            s.shapes.add_picture(self.logo_path, Inches(0.55), Inches(0.61), height=Inches(0.56))
        else:
            self._texto(s, 0.47, 0.7, 4, 0.5, rotulo, tam=16, cor=CLARO, negrito=True)
        linhas = ["RELATÓRIO DE"] + [x for x in area.split("\n")]
        y = 3.4
        for ln in linhas:
            self._texto(s, 0.47, y, 4.1, 0.4, " ".join(ln).replace("   ", "   "), tam=14, cor=CLARO, negrito=True)
            y += 0.4 if ln != "RELATÓRIO DE" else 0.4
        partes = titulo.split("\n")
        self._texto(s, 5.3, 1.75, 7.4, 1.0, partes[0], tam=48, cor=BRANCO, negrito=True, ancora=MSO_ANCHOR.BOTTOM)
        if len(partes) > 1:
            self._texto(s, 5.35, 2.8, 7.4, 0.8, " ".join(partes[1:]), tam=30, cor="AABBD0", ancora=MSO_ANCHOR.TOP)
        self._rect(s, 5.37, 3.95, max(3.2, min(7.2, 0.11 * len(subtitulo) + 0.8)), 0.56, fill=ACENTO, forma=MSO_SHAPE.ROUNDED_RECTANGLE, raio=0.5)
        self._texto(s, 5.37, 3.95, max(3.2, min(7.2, 0.11 * len(subtitulo) + 0.8)), 0.56, pilula or subtitulo, tam=16, cor=BRANCO, negrito=True, alinh=PP_ALIGN.CENTER, ancora=MSO_ANCHOR.MIDDLE)
        if meta:
            ms = meta if isinstance(meta, list) else [meta]
            self._texto(s, 5.4, H - 1.45, 7.6, 0.6, [[(m, {})] for m in ms[:2]], tam=13, cor=CLARO, ancora=MSO_ANCHOR.BOTTOM)
        self._texto(s, 5.4, H - 0.75, 7.0, 0.35, "4PIPE Solutions — hidropig.com.br", tam=13, cor=CINZA)
        return s

    def slide(self, titulo, manchete=None, notas=None):
        """Slide de conteúdo no modelo 4PIPE: faixa marinho com logo (caixa branca),
        barra laranja, título em caixa-alta, subtítulo claro e pílula laranja do
        período; rodapé marinho com 'CONFIDENCIAL'."""
        s = self.prs.slides.add_slide(self._blank)
        self.n += 1
        self._fundo(s, FUNDO)
        self._rect(s, 0, 0, W, 1.09, fill=NAVY)
        if self.logo_path:
            self._rect(s, 0.2, 0.24, 2.19, 0.61, fill=BRANCO, forma=MSO_SHAPE.ROUNDED_RECTANGLE, raio=0.2)
            s.shapes.add_picture(self.logo_path, Inches(0.27), Inches(0.28), height=Inches(0.53))
            x_bar = 2.6
        else:
            x_bar = 0.4
        self._rect(s, x_bar, 0.18, 0.07, 0.72, fill=ACENTO)
        larg_txt = W - x_bar - 0.3 - (2.2 if self.periodo else 0.4)
        self._texto(s, x_bar + 0.25, 0.14, larg_txt, 0.5, (titulo or "").upper(), tam=(21 if len(titulo or "") <= 40 else (18 if len(titulo or "") <= 52 else 16)), cor=BRANCO, negrito=True, ancora=MSO_ANCHOR.MIDDLE)
        if manchete:
            self._texto(s, x_bar + 0.27, 0.62, larg_txt, 0.42, manchete, tam=12, cor=CLARO, ancora=MSO_ANCHOR.TOP)
        if self.periodo:
            self._rect(s, W - 2.35, 0.3, 2.05, 0.48, fill=ACENTO, forma=MSO_SHAPE.ROUNDED_RECTANGLE, raio=0.2)
            self._texto(s, W - 2.35, 0.3, 2.05, 0.48, self.periodo, tam=14, cor=BRANCO, negrito=True, alinh=PP_ALIGN.CENTER, ancora=MSO_ANCHOR.MIDDLE)
        # rodapé
        self._rect(s, 0, H - 0.27, W, 0.27, fill=NAVY)
        rod = f"4PIPE Solutions  |  {self.titulo_curto}" + (f"  |  Base: {self.base}" if self.base else "") + f"  |  Gerado em {self.gerado_em}"
        self._texto(s, 0.4, H - 0.27, 9.8, 0.27, rod, tam=10, cor=CLARO, ancora=MSO_ANCHOR.MIDDLE)
        self._texto(s, W - 3.0, H - 0.27, 2.6, 0.27, f"CONFIDENCIAL  ·  {self.n}", tam=10, cor=ACENTO, negrito=True, alinh=PP_ALIGN.RIGHT, ancora=MSO_ANCHOR.MIDDLE)
        self._notas(s, notas)
        return s

    def kpis(self, s, itens, y=1.65, h=1.35, x=MARGEM, w=None, gap=0.22):
        """Cards de indicador no modelo: branco, borda suave, sombra, barra de cor à esquerda.
        itens: [{rotulo, valor, sub, cor}]"""
        w = w or (W - 2 * MARGEM)
        n = len(itens)
        cw = (w - gap * (n - 1)) / n
        for i, it in enumerate(itens):
            cx = x + i * (cw + gap)
            cor = it.get("cor", NAVY)
            self._rect(s, cx, y, cw, h, fill=BRANCO, line=BORDA, forma=MSO_SHAPE.ROUNDED_RECTANGLE, raio=0.06, sombra=True)
            self._rect(s, cx, y + 0.05, 0.055, h - 0.1, fill=cor)
            self._texto(s, cx + 0.22, y + 0.12, cw - 0.35, 0.3, it["rotulo"].upper(), tam=11, cor=CINZA, negrito=True)
            tam_v = it.get("tam", 30 if len(str(it["valor"])) <= 9 else 24)
            self._texto(s, cx + 0.22, y + 0.4, cw - 0.35, max(0.4, h - 0.82), it["valor"], tam=tam_v, cor=cor, negrito=True, ancora=MSO_ANCHOR.MIDDLE)
            if it.get("sub"):
                self._texto(s, cx + 0.22, y + h - 0.4, cw - 0.35, 0.32, it["sub"], tam=11, cor=CINZA, italico=True, ancora=MSO_ANCHOR.MIDDLE)

    def cartao(self, s, x, y, w, h, titulo=None, fill=BRANCO):
        self._rect(s, x, y, w, h, fill=fill, line=BORDA, forma=MSO_SHAPE.ROUNDED_RECTANGLE, raio=0.04, sombra=True)
        if titulo:
            self._texto(s, x + 0.25, y + 0.14, w - 0.5, 0.3, titulo.upper(), tam=12, cor=ACENTO, negrito=True)

    def mensagens(self, s, x, y, w, h, msgs, titulo="Leitura gerencial", tam=14, passo=None):
        """Lista de mensagens com marcador colorido. msgs: [(cor_hex, texto)]"""
        self.cartao(s, x, y, w, h, titulo)
        passo = passo or min(0.78, (h - 0.6) / max(1, len(msgs)))
        yy = y + 0.55
        for cor, txt in msgs:
            dot = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 0.22), Inches(yy + 0.07), Inches(0.17), Inches(0.17))
            dot.fill.solid()
            dot.fill.fore_color.rgb = rgb(cor)
            dot.line.fill.background()
            dot.shadow.inherit = False
            self._texto(s, x + 0.55, yy, w - 0.8, passo, txt, tam=tam, cor=TEXTO)
            yy += passo

    def pilula(self, s, x, y, w, h, texto, cor, tam=11):
        self._rect(s, x, y, w, h, fill=cor, forma=MSO_SHAPE.ROUNDED_RECTANGLE, raio=0.5)
        self._texto(s, x, y, w, h, texto, tam=tam, cor=BRANCO, negrito=True, alinh=PP_ALIGN.CENTER, ancora=MSO_ANCHOR.MIDDLE)

    def barra(self, s, x, y, w, h, partes, fundo="E3E7EE"):
        """Barra horizontal segmentada. partes: [(valor, cor_hex)]"""
        self._rect(s, x, y, w, h, fill=fundo, forma=MSO_SHAPE.ROUNDED_RECTANGLE, raio=0.5)
        total = sum(v for v, _ in partes) or 0
        if not total:
            return
        cx = x
        for v, cor in partes:
            if v <= 0:
                continue
            seg = w * v / total
            self._rect(s, cx, y, seg, h, fill=cor)
            cx += seg

    # ------------------------------------------------ tabela
    def tabela(self, s, cabecalho, linhas, larguras, x=MARGEM, y=1.6, tam=11, alt_linha=0.34, cores=None,
               alinh=None, altura_cab=0.38, negrito_ultima=False):
        """Tabela nativa. `cores`: {(linha, col): hex_texto} ou {(linha, None): hex_fundo}.
        `alinh`: lista de 'l'|'c'|'r' por coluna."""
        n_l, n_c = len(linhas) + 1, len(cabecalho)
        total_w = sum(larguras)
        gf = s.shapes.add_table(n_l, n_c, Inches(x), Inches(y), Inches(total_w), Inches(altura_cab + alt_linha * len(linhas)))
        tbl = gf.table
        # remove o estilo padrão azul do PowerPoint
        tblPr = tbl._tbl.tblPr
        tblPr.set("firstRow", "0")
        tblPr.set("bandRow", "0")
        for i, lw in enumerate(larguras):
            tbl.columns[i].width = Inches(lw)
        tbl.rows[0].height = Inches(altura_cab)
        for r in range(1, n_l):
            tbl.rows[r].height = Inches(alt_linha)
        alinh = alinh or ["l"] * n_c
        mapa = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}

        def pinta(cel, txt, fundo, cor_t, neg, al, tam_):
            cel.fill.solid()
            cel.fill.fore_color.rgb = rgb(fundo)
            cel.margin_left = cel.margin_right = Inches(0.08)
            cel.margin_top = cel.margin_bottom = Inches(0.03)
            cel.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cel.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = mapa[al]
            for r_ in list(p.runs):
                r_._r.getparent().remove(r_._r)
            run = p.add_run()
            run.text = "" if txt is None else str(txt)
            run.font.name = FONTE
            run.font.size = Pt(tam_)
            run.font.bold = neg
            run.font.color.rgb = rgb(cor_t)

        for c, t in enumerate(cabecalho):
            pinta(tbl.cell(0, c), t, NAVY, BRANCO, True, alinh[c], tam)
        cores = cores or {}
        for r, linha in enumerate(linhas, start=1):
            fundo = cores.get((r - 1, None)) or (BRANCO if r % 2 else "F5F7FA")
            ultima = negrito_ultima and r == len(linhas)
            if ultima:
                fundo = "E2E8F2"
            for c, v in enumerate(linha):
                pinta(tbl.cell(r, c), v, fundo, cores.get((r - 1, c), TEXTO), ultima or (r - 1, c) in cores, alinh[c], tam)
        return gf

    # ------------------------------------------------ gráficos
    def _estilo_chart(self, chart, legenda=True, tam=11):
        chart.font.name = FONTE
        chart.font.size = Pt(tam)
        chart.font.color.rgb = rgb(TEXTO)
        chart.has_legend = legenda
        if legenda:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False
            chart.legend.font.size = Pt(tam)

    def grafico(self, s, tipo, categorias, series, x, y, w, h, empilhado=False, rotulos=True, fmt="#,##0",
                legenda=None, titulo=None, tam=11, horizontal=False, cor_pontos=None, gap=60, eixo_valor=False, destaque=None):
        """tipo: 'col' | 'linha' | 'rosca'. series: [(nome, [valores], cor_hex)]"""
        cd = CategoryChartData()
        cd.categories = list(categorias)
        for nome, vals, _cor in series:
            cd.add_series(nome, [0 if v is None else v for v in vals])
        if tipo == "rosca":
            ct = XL_CHART_TYPE.DOUGHNUT
        elif tipo == "linha":
            ct = XL_CHART_TYPE.LINE_MARKERS
        elif horizontal:
            ct = XL_CHART_TYPE.BAR_STACKED if empilhado else XL_CHART_TYPE.BAR_CLUSTERED
        else:
            ct = XL_CHART_TYPE.COLUMN_STACKED if empilhado else XL_CHART_TYPE.COLUMN_CLUSTERED
        gf = s.shapes.add_chart(ct, Inches(x), Inches(y), Inches(w), Inches(h), cd)
        ch = gf.chart
        self._estilo_chart(ch, legenda=(len(series) > 1 or tipo == "rosca") if legenda is None else legenda, tam=tam)
        if titulo:
            ch.has_title = True
            ch.chart_title.text_frame.text = titulo
            r = ch.chart_title.text_frame.paragraphs[0].runs[0]
            r.font.size = Pt(tam + 2)
            r.font.bold = True
            r.font.name = FONTE
            r.font.color.rgb = rgb(NAVY)
        else:
            ch.has_title = False
        plot = ch.plots[0]
        if tipo == "rosca":
            plot.has_data_labels = rotulos
            if rotulos:
                dl = plot.data_labels
                dl.number_format = "0"
                dl.number_format_is_linked = False
                dl.font.size = Pt(tam)
                dl.font.bold = True
                dl.font.color.rgb = rgb(BRANCO)
                dl.show_percentage = False
                dl.show_value = True
            ser = plot.series[0]
            for i, cor in enumerate(cor_pontos or []):
                pt = ser.points[i]
                pt.format.fill.solid()
                pt.format.fill.fore_color.rgb = rgb(cor)
            return ch
        if tipo != "linha":
            plot.gap_width = gap
            if empilhado:
                plot.overlap = 100
        for i, (_n, _v, cor) in enumerate(series):
            ser = plot.series[i]
            if tipo == "linha":
                ser.format.line.color.rgb = rgb(cor)
                ser.format.line.width = Pt(2.5)
                ser.marker.format.fill.solid()
                ser.marker.format.fill.fore_color.rgb = rgb(cor)
                ser.marker.format.line.color.rgb = rgb(cor)
                ser.smooth = False
            else:
                ser.format.fill.solid()
                ser.format.fill.fore_color.rgb = rgb(cor)
        if destaque is not None and tipo == "col" and not empilhado and len(series) == 1:
            vals = [v or 0 for v in series[0][1]]
            idxs = [max(range(len(vals)), key=lambda k: vals[k])] if (destaque == "max" and vals) else (destaque if destaque != "max" else [])
            for k in idxs:
                pt = plot.series[0].points[k]
                pt.format.fill.solid()
                pt.format.fill.fore_color.rgb = rgb(ACENTO)
        if rotulos:
            plot.has_data_labels = True
            dl = plot.data_labels
            dl.number_format = fmt
            dl.number_format_is_linked = False
            dl.font.size = Pt(tam)
            dl.font.bold = True
            dl.font.name = FONTE
            if tipo == "linha":
                dl.position = XL_LABEL_POSITION.ABOVE
            elif empilhado:
                dl.position = XL_LABEL_POSITION.CENTER
                dl.font.color.rgb = rgb(BRANCO)
                dl.font.bold = True
            else:
                dl.position = XL_LABEL_POSITION.OUTSIDE_END
        ca = ch.category_axis
        ca.tick_labels.font.size = Pt(tam)
        ca.format.line.color.rgb = rgb(BORDA)
        ca.has_major_gridlines = False
        va = ch.value_axis
        va.has_major_gridlines = eixo_valor
        va.visible = eixo_valor
        if eixo_valor:
            va.major_gridlines.format.line.color.rgb = rgb(BORDA)
            va.tick_labels.font.size = Pt(tam - 1)
            va.tick_labels.number_format = fmt
            va.tick_labels.number_format_is_linked = False
        if horizontal:
            ca.reverse_order = True
        return ch

    # ------------------------------------------------ fechamento
    def proximos_passos(self, passos, titulo="Pontos de atenção e próximos passos", manchete=None):
        """passos: [(cor_hex, titulo, detalhe)] — até 5."""
        s = self.slide(titulo, manchete)
        n = len(passos)
        y0, alt = 1.4, min(0.98, 5.3 / max(1, n))
        for i, (cor, t, det) in enumerate(passos):
            y = y0 + i * (alt + 0.1)
            self._rect(s, MARGEM, y, W - 2 * MARGEM, alt, fill=BRANCO, line=BORDA, forma=MSO_SHAPE.ROUNDED_RECTANGLE, raio=0.06, sombra=True)
            self._rect(s, MARGEM, y + 0.05, 0.06, alt - 0.1, fill=cor)
            c = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(MARGEM + 0.3), Inches(y + alt / 2 - 0.25), Inches(0.5), Inches(0.5))
            c.fill.solid()
            c.fill.fore_color.rgb = rgb(cor)
            c.line.fill.background()
            c.shadow.inherit = False
            self._texto(s, MARGEM + 0.3, y + alt / 2 - 0.25, 0.5, 0.5, str(i + 1), tam=16, cor=BRANCO, negrito=True, alinh=PP_ALIGN.CENTER, ancora=MSO_ANCHOR.MIDDLE)
            self._texto(s, MARGEM + 1.05, y + 0.08, W - 2 * MARGEM - 1.3, 0.35, t, tam=16, cor=NAVY, negrito=True)
            self._texto(s, MARGEM + 1.05, y + 0.45, W - 2 * MARGEM - 1.3, alt - 0.5, det, tam=13, cor=TEXTO)
        return s

    def bytes(self):
        buf = io.BytesIO()
        self.prs.save(buf)
        return buf.getvalue()


# =====================================================================
#  PAINEL DIÁRIO — OPERAÇÃO
# =====================================================================
_RITMO = {"adiantado": ("Adiantado", VERDE), "no_ritmo": ("No ritmo", AZUL), "atrasado": ("Abaixo do ritmo", VERMELHO)}
_SEMAFORO_EST = {"critico": ("CRÍTICA", VERMELHO, 3), "atencao": ("ATENÇÃO", LARANJA, 2), "ok": ("CONTROLADA", VERDE, 0)}
_CAP_NIVEL = {"ok": ("Folga", VERDE), "apertado": ("Apertado", AMARELO), "sobrecarga": ("Sobrecarga", VERMELHO), "sem_base": ("Sem base", CINZA)}


def _dias_mes(hoje):
    from calendar import monthrange
    return monthrange(hoje.year, hoje.month)[1]


def _proj_linear(realizado, hoje, teto):
    """Extrapola o realizado do mês pelo calendário (dia/dias do mês), limitado ao planejado."""
    frac = hoje.day / _dias_mes(hoje)
    if not frac or not realizado:
        return realizado or 0
    return min(realizado / frac, teto) if teto else realizado / frac


def montar_painel_diario(ctx, modo, gerado_em, logo_path=None, hoje=None):
    """Apresentação do Painel Diário. `ctx` = dict de _contexto_painel."""
    hoje = hoje or date.today()
    completa = modo == "completa"
    d = Deck("Painel Diário · Operação", gerado_em, logo_path,
             periodo=_mes_extenso((ctx.get("resumo_mes_atual_pcp") or {}).get("mes_label") or ctx.get("mes_atual_label") or hoje.strftime("%m/%Y")),
             base="pedidos em tempo real (Painel de Gestão)")
    res = ctx.get("resumo") or {}
    atual = ctx.get("resumo_mes_atual_pcp") or {}
    seg = ctx.get("resumo_mes_seguinte_pcp") or {}
    ant = ctx.get("resumo_mes_anterior_pcp") or {}
    prev = (ctx.get("preview_pcp_semanal") or {})
    prev_res = prev.get("resumo") or {}
    risco = ctx.get("risco_estacoes") or {"estacoes": [], "total": {}}
    rt = risco.get("total") or {}
    cap = ctx.get("capacidade_estacoes") or {"estacoes": [], "total": {}}
    a360 = ctx.get("atrasados_360") or {"resumo": {}, "pedidos": []}
    a360r = a360.get("resumo") or {}
    mes_lbl = atual.get("mes_label") or ctx.get("mes_atual_label") or ""
    seg_lbl = seg.get("mes_label") or ""

    n_atrasados = ctx.get("atrasados") or 0
    n_vencendo = ctx.get("vencendo") or 0

    d.capa("Painel Diário\nOperação", f"{_mes_extenso(mes_lbl)}" + (f" + Projeção {_mes_extenso(seg_lbl).split(' ')[0].upper()[:3]}" if seg_lbl else ""),
           area="GESTÃO\nOPERAÇÕES",
           meta=[f"Base: Painel de Gestão  |  Referência {hoje.strftime('%d/%m/%Y')}", f"Gerado em {gerado_em}  ·  Apresentação {'completa' if completa else 'executiva'}"])

    # ---- 2. Resumo executivo ----
    msgs = []
    ativos = (res.get("total") or 0) - (res.get("finalizado") or 0)
    if n_atrasados:
        msgs.append((VERMELHO, f"{plural(n_atrasados, 'pedido atrasado', 'pedidos atrasados')} de {fmt_int(ativos)} em andamento — {fmt_brl(a360r.get('valor_em_risco'), True)} em risco (atraso médio {fmt_int(a360r.get('atraso_medio') or 0)} dias)."))
    else:
        msgs.append((VERDE, f"Nenhum pedido atrasado entre os {fmt_int(ativos)} em andamento."))
    if atual.get("pedidos_total"):
        pct_r = 100 * (atual.get("valor_finalizado") or 0) / atual["valor_total"] if atual.get("valor_total") else 0
        rt_txt, rt_cor = _RITMO.get(prev_res.get("ritmo"), ("", CINZA))
        msgs.append((rt_cor if prev_res.get("ritmo") else AZUL,
                     f"{mes_lbl}: {fmt_brl(atual.get('valor_finalizado'), True)} realizados de {fmt_brl(atual.get('valor_total'), True)} planejados ({fmt_pct(pct_r)}), com {fmt_pct(prev_res.get('pct_mes_decorrido'))} do mês decorrido"
                     + (f" — {rt_txt.lower()}." if prev_res.get("ritmo") else ".")))
    if seg.get("pedidos_total"):
        msgs.append((AZUL, f"{seg_lbl}: {fmt_brl(seg.get('valor_total'), True)} já planejados em {plural(seg['pedidos_total'], 'pedido')}."))
    elif seg_lbl:
        msgs.append((AMARELO, f"{seg_lbl}: ainda sem pedidos no Planejamento PCP — carteira do próximo mês por montar."))
    if risco.get("n_criticas") or risco.get("n_atencao"):
        msgs.append((LARANJA if not risco.get("n_criticas") else VERMELHO,
                     f"{plural(risco.get('n_criticas') or 0, 'estação crítica', 'estações críticas')} e {plural(risco.get('n_atencao') or 0, 'em atenção', 'em atenção')}; {fmt_brl(rt.get('valor_risco'), True)} em itens com risco de prazo."))
    s = d.slide("Resumo executivo", f"Retrato da operação em {hoje.strftime('%d/%m/%Y')}",
                notas="KPIs vêm do Painel Diário (mesma base da tela). Atrasado = item com liberação prevista vencida e ainda aberto.")
    d.kpis(s, [
        {"rotulo": "Pedidos em andamento", "valor": fmt_int(ativos), "sub": f"{fmt_int(res.get('finalizado') or 0)} já finalizados", "cor": NAVY},
        {"rotulo": "Valor da carteira", "valor": fmt_brl(res.get("valor_total"), True), "sub": "todos os pedidos cadastrados", "cor": AZUL},
        {"rotulo": "Atrasados", "valor": fmt_int(n_atrasados), "sub": f"{fmt_brl(a360r.get('valor_em_risco'), True)} em risco", "cor": VERMELHO if n_atrasados else VERDE},
        {"rotulo": "Vencendo (7 dias)", "valor": fmt_int(n_vencendo), "sub": "prazo do cliente próximo", "cor": LARANJA if n_vencendo else VERDE},
        {"rotulo": "Backlog", "valor": fmt_int(ctx.get("backlog") or 0), "sub": "pedidos aguardando início", "cor": NAVY_2},
    ])
    d.mensagens(s, MARGEM, 3.2, W - 2 * MARGEM, 3.7, msgs[:4], tam=15)

    # ---- 3. Projeção de faturamento ----
    tot = atual.get("valor_total") or 0
    real = atual.get("valor_finalizado") or 0
    proj = _proj_linear(real, hoje, tot)
    s = d.slide(f"Projeção de faturamento — {mes_lbl}",
                f"{fmt_brl(real, True)} realizados; ao ritmo atual o mês fecha em ~{fmt_brl(proj, True)} (teto planejado {fmt_brl(tot, True)})",
                notas=("Projeção linear: realizado ÷ fração do mês decorrida, limitada ao total planejado no PCP. "
                       "Planejado = soma dos pedidos com Planejamento PCP no mês. Não é meta, é leitura de ritmo."))
    d.grafico(s, "col", ["Realizado", "Projeção ao ritmo atual", "Planejado (PCP)"],
              [("Valor", [real, proj, tot], NAVY)], MARGEM, 1.6, 6.3, 4.9, fmt='"R$ "#,##0', tam=12, gap=70, cor_pontos=None)
    ch_s = s.shapes[-1].chart
    for idx, cor in enumerate((VERDE, AZUL, NAVY)):
        pt = ch_s.plots[0].series[0].points[idx]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = rgb(cor)
    pct_real = 100 * real / tot if tot else 0
    d.kpis(s, [
        {"rotulo": f"{mes_lbl} · pedidos", "valor": f"{fmt_int(atual.get('pedidos_finalizados') or 0)} / {fmt_int(atual.get('pedidos_total') or 0)}", "sub": f"{fmt_pct(pct_real)} do valor realizado", "cor": VERDE, "tam": 26},
        {"rotulo": "Em aberto no mês", "valor": fmt_brl(atual.get("valor_em_aberto"), True), "sub": plural(atual.get("pedidos_em_aberto") or 0, "pedido"), "cor": LARANJA, "tam": 26},
    ], y=1.65, h=1.3, x=7.2, w=5.6)
    d.kpis(s, [
        {"rotulo": f"{seg_lbl or 'Mês seguinte'} · planejado", "valor": fmt_brl(seg.get("valor_total"), True), "sub": plural(seg.get("pedidos_total") or 0, "pedido") + " no backlog PCP", "cor": AZUL, "tam": 26},
        {"rotulo": f"{ant.get('mes_label') or 'Mês anterior'} · fechado", "valor": fmt_brl(ant.get("valor_finalizado"), True), "sub": plural(ant.get("pedidos_finalizados") or 0, "pedido") + " finalizados", "cor": NAVY_2, "tam": 26},
    ], y=3.15, h=1.3, x=7.2, w=5.6)
    rt_txt, rt_cor = _RITMO.get(prev_res.get("ritmo"), ("Sem leitura", CINZA))
    d.cartao(s, 7.2, 4.65, 5.6, 1.85, "Ritmo do mês")
    d.pilula(s, 7.4, 5.05, 2.1, 0.38, rt_txt.upper(), rt_cor, tam=12)
    d._texto(s, 9.65, 5.0, 3.0, 0.5, f"{fmt_pct(prev_res.get('pct'))} concluído vs {fmt_pct(prev_res.get('pct_mes_decorrido'))} do mês", tam=13, cor=TEXTO, ancora=MSO_ANCHOR.MIDDLE)
    d._texto(s, 7.4, 5.6, 5.2, 0.8, f"Aderência ao planejado nas semanas já encerradas: {fmt_pct(prev_res.get('aderencia'))} ({fmt_int(prev_res.get('concluido_encerrado') or 0)} de {fmt_int(prev_res.get('planejado_encerrado') or 0)} itens).", tam=12, cor=CINZA)

    # ---- 4. Planejamento semanal PCP ----
    sem = prev.get("semanas") or []
    if sem:
        s = d.slide(f"Planejamento semanal PCP — {prev_res.get('mes_label') or mes_lbl}",
                    f"{fmt_int(prev_res.get('n_concluidos') or 0)} de {fmt_int(prev_res.get('n_pedidos') or 0)} pedidos concluídos ({fmt_pct(prev_res.get('pct'))}); {fmt_int(prev_res.get('n_abertos') or 0)} em aberto",
                    notas="Cada barra é uma semana do Planejamento PCP. Semana encerrada com pedidos abertos = atraso de planejamento.")
        d.grafico(s, "col", [x.get("rotulo_curto") or x.get("rotulo") for x in sem],
                  [("Concluídos", [x.get("n_concluidos") or 0 for x in sem], VERDE), ("Em aberto", [x.get("n_abertos") or 0 for x in sem], LARANJA)],
                  MARGEM, 1.6, 7.6, 5.0, empilhado=True, tam=12, gap=55, fmt="0;-0;;")
        linhas_s = []
        cores_s = {}
        for i, x in enumerate(sem):
            st = "Encerrada" if x.get("encerrada") else ("Em curso" if x.get("atual") else "Futura")
            linhas_s.append([x.get("rotulo_curto") or x.get("rotulo"), st, f"{x.get('pct') or 0}%", fmt_brl(x.get("total"), True)])
            if x.get("encerrada") and (x.get("n_abertos") or 0):
                cores_s[(i, 1)] = VERMELHO
                cores_s[(i, 2)] = VERMELHO
        d.tabela(s, ["Semana", "Situação", "Concl.", "Valor"], linhas_s, [1.4, 1.0, 0.8, 1.2], x=8.4, y=1.7, tam=11, alt_linha=0.42, cores=cores_s, alinh=["l", "l", "c", "r"])

    # ---- 5. Risco por estação ----
    est = sorted(risco.get("estacoes") or [], key=lambda e: (-_SEMAFORO_EST.get(e.get("semaforo"), ("", "", 0))[2], -(e.get("em_risco") or 0), -(e.get("abertos") or 0)))
    if est:
        s = d.slide("Risco de prazos por estação",
                    f"{plural(risco.get('n_criticas') or 0, 'estação crítica', 'estações críticas')} · {plural(risco.get('n_atencao') or 0, 'em atenção', 'em atenção')} · {fmt_brl(rt.get('valor_risco'), True)} em risco",
                    notas="Risco = itens abertos com prazo de cliente vencido/≤7 dias ou liberação PCP vencida. Percentual sobre os itens abertos da estação.")
        linhas_e, cores_e = [], {}
        for i, e in enumerate(est[:9]):
            nome_v, cor_v, _ = _SEMAFORO_EST.get(e.get("semaforo"), ("—", CINZA, 0))
            linhas_e.append([e.get("curto") or e.get("nome"), nome_v, fmt_int(e.get("abertos")), fmt_int(e.get("atrasados")), fmt_int(e.get("cli_vencido")), fmt_int(e.get("cli_7d")), f"{e.get('risco_pct') or 0}%", fmt_brl(e.get("valor_risco"), True)])
            cores_e[(i, 1)] = cor_v
            if e.get("atrasados"):
                cores_e[(i, 3)] = VERMELHO
            if e.get("cli_vencido"):
                cores_e[(i, 4)] = VERMELHO
        linhas_e.append(["TOTAL", "", fmt_int(rt.get("abertos")), fmt_int(rt.get("atrasados")), fmt_int(rt.get("cli_vencido")), fmt_int(rt.get("cli_7d")), f"{rt.get('risco_pct') or 0}%", fmt_brl(rt.get("valor_risco"), True)])
        d.tabela(s, ["Estação", "Situação", "Abertos", "Atrasados", "Cliente vencido", "Cliente ≤7d", "% em risco", "Valor em risco"], linhas_e,
                 [2.2, 1.6, 1.15, 1.3, 1.6, 1.4, 1.4, 1.55], y=1.65, tam=13, alt_linha=0.5, cores=cores_e, alinh=["l", "c", "c", "c", "c", "c", "c", "r"], negrito_ultima=True)

    # ---- 6. Atrasos que exigem decisão ----
    if a360r.get("total"):
        s = d.slide("Atrasos que exigem decisão",
                    f"{plural(a360r.get('total'), 'pedido atrasado', 'pedidos atrasados')} · {fmt_brl(a360r.get('valor_em_risco'), True)} em risco · atraso máximo {fmt_int(a360r.get('atraso_max'))} dias",
                    notas="Motivo = causa raiz identificada pelo sistema (sem planejamento PCP, prazo PCP vencido, qualidade, tratativa...). 'Renegociar' = prazo do cliente inviável mesmo liberando hoje.")
        motivos = a360r.get("por_motivo") or []
        if motivos:
            d.grafico(s, "col", [m["titulo"] for m in motivos], [("Pedidos", [m["n"] for m in motivos], AZUL)], MARGEM, 1.6, 5.2, 3.2,
                      horizontal=True, tam=12, gap=45, titulo="Atrasos por motivo", destaque="max")
        d.kpis(s, [
            {"rotulo": "A renegociar", "valor": fmt_int(a360r.get("a_renegociar")), "sub": "prazo inviável", "cor": VERMELHO, "tam": 26},
            {"rotulo": "Recuperáveis", "valor": fmt_int(a360r.get("recuperaveis")), "sub": "ainda dá tempo", "cor": VERDE, "tam": 26},
        ], y=5.0, h=1.3, x=MARGEM, w=5.2)
        linhas_a = []
        for p in (a360.get("pedidos") or [])[:6]:
            linhas_a.append([p.get("pedido_venda"), (p.get("cliente") or "")[:26], f"{p.get('dias_atraso')}d", fmt_brl(p.get("valor"), True), p.get("motivo_titulo") or ""])
        cores_a = {(i, 2): VERMELHO for i in range(len(linhas_a))}
        d.tabela(s, ["Pedido", "Cliente", "Atraso", "Valor", "Motivo"], linhas_a, [0.95, 2.0, 0.8, 1.05, 2.0], x=6.15, y=1.7, tam=11, alt_linha=0.5, cores=cores_a, alinh=["l", "l", "c", "r", "l"])

    # ---- 7. Desempenho (OTD / lead time) ----
    refs = (ctx.get("desempenho_360") or {}).get("refs") or {}
    otd_m = [x for x in (ctx.get("otd_mensal_ano") or []) if x.get("percentual") is not None]
    if refs.get("otd") or otd_m:
        s = d.slide("Desempenho de entrega", "OTD (entrega no prazo) e lead time — últimos períodos fechados",
                    notas="OTD = % de pedidos entregues no prazo do cliente. Lead time = dias entre inclusão e entrega.")
        cards = []
        if refs.get("otd"):
            cards.append({"rotulo": f"OTD · {refs['otd'].get('periodo')}", "valor": fmt_pct(refs["otd"].get("valor"), 1), "sub": f"{fmt_int(refs['otd'].get('n'))} pedidos", "cor": VERDE if (refs['otd'].get('valor') or 0) >= 90 else (AMARELO if (refs['otd'].get('valor') or 0) >= 70 else VERMELHO)})
        for chave, nome in (("lt_operacao", "Lead time operação"), ("lt_cif", "Lead time CIF"), ("lt_fob", "Lead time FOB")):
            r = refs.get(chave)
            if r:
                cards.append({"rotulo": f"{nome} · {r.get('periodo')}", "valor": f"{fmt_int(r.get('valor'))} dias", "sub": f"{fmt_int(r.get('n'))} pedidos", "cor": NAVY})
        if cards:
            d.kpis(s, cards[:4], y=1.65, h=1.3)
        if otd_m:
            d.grafico(s, "linha", [x["mes"] for x in otd_m], [("OTD %", [x["percentual"] for x in otd_m], AZUL)], MARGEM, 3.2, W - 2 * MARGEM, 3.7,
                      fmt='0"%"', tam=12, titulo="OTD mensal (%)", eixo_valor=True)

    if completa:
        # ---- Pipeline + backlog por estação ----
        etapas = (ctx.get("pipeline_mes_pcp") or {}).get("etapas") or []
        back = ctx.get("backlog_estacao") or []
        if etapas or back:
            s = d.slide("Pipeline e backlog", f"Fluxo dos pedidos planejados em {(ctx.get('pipeline_mes_pcp') or {}).get('mes_label') or mes_lbl} e fila por estação")
            if etapas:
                d.grafico(s, "col", [e["label"] for e in etapas], [("Pedidos", [e["total"] for e in etapas], AZUL)], MARGEM, 1.6, 6.0, 5.0, tam=12, titulo="Pipeline do mês (pedidos)", gap=50)
            if back:
                d.grafico(s, "col", [b["estacao"] for b in back], [("Itens na fila", [b["quantidade"] for b in back], NAVY_2)], 6.9, 1.6, 5.9, 5.0, horizontal=True, tam=12, titulo="Itens em aberto por estação", gap=45, destaque="max")
        # ---- Capacidade ----
        ce = cap.get("estacoes") or []
        if ce:
            s = d.slide("Capacidade das estações", "Ritmo histórico (últimas semanas) x carga projetada pelo PCP",
                        notas="Ritmo = média de itens concluídos por semana; pico = melhor semana. Carga > ritmo = apertado; > pico = sobrecarga. Cobertura = semanas de fila ao ritmo médio.")
            lin, cor = [], {}
            for i, e in enumerate(ce[:10]):
                nv, cv = _CAP_NIVEL.get(e.get("semaforo"), ("—", CINZA))
                lin.append([e.get("curto") or e.get("nome"), nv, fmt_int(e.get("abertos")), fmt_int(e.get("atrasados")), str(e.get("ritmo")).replace(".", ","), fmt_int(e.get("pico")),
                            (str(e.get("cobertura")).replace(".", ",") + " sem.") if e.get("cobertura") else "—", fmt_int(e.get("sem_prev"))])
                cor[(i, 1)] = cv
                if e.get("atrasados"):
                    cor[(i, 3)] = VERMELHO
            d.tabela(s, ["Estação", "Carga", "Abertos", "Atrasados", "Ritmo/sem.", "Pico", "Cobertura", "Sem previsão"], lin, [2.2, 1.5, 1.2, 1.3, 1.5, 1.0, 1.7, 1.8],
                     y=1.65, tam=13, alt_linha=0.5, cores=cor, alinh=["l", "c", "c", "c", "c", "c", "c", "c"])
        # ---- Top pedidos do mês ----
        top = atual.get("top5") or []
        if top:
            s = d.slide(f"Maiores pedidos de {mes_lbl}", "Top 5 por valor no Planejamento PCP do mês")
            lin = [[p.get("pedido_venda"), (p.get("cliente") or "")[:30], p.get("estado") or "", p.get("frete") or "", fmt_data(p.get("data_prevista_pcp")), fmt_data(p.get("data_solicitada_cliente")),
                    "Concluído" if p.get("finalizado") else "Em aberto", fmt_brl(p.get("valor"), True)] for p in top]
            cor = {(i, 6): (VERDE if p.get("finalizado") else LARANJA) for i, p in enumerate(top)}
            d.tabela(s, ["Pedido", "Cliente", "UF", "Frete", "Prev. PCP", "Prazo cliente", "Situação", "Valor"], lin, [1.2, 3.2, 0.7, 0.9, 1.4, 1.6, 1.5, 1.7], y=1.65, tam=13, alt_linha=0.55, cores=cor, alinh=["l", "l", "c", "c", "c", "c", "c", "r"])
        # ---- Faturamento realizado no ano ----
        fat = [x for x in (ctx.get("faturamento_mensal_ano") or [])]
        if any((x.get("realizado") or 0) for x in fat):
            s = d.slide(f"Faturamento realizado em {ctx.get('ano_graficos') or hoje.year}", "Valor faturado mês a mês")
            d.grafico(s, "col", [x["mes"] for x in fat], [("Realizado", [x.get("realizado") or 0 for x in fat], NAVY)], MARGEM, 1.6, W - 2 * MARGEM, 5.0, fmt='#,##0', tam=12, gap=40)

    # ---- Próximos passos ----
    passos = []
    if n_atrasados:
        passos.append((VERMELHO, f"Tratar {plural(n_atrasados, 'pedido atrasado', 'pedidos atrasados')}", f"{fmt_brl(a360r.get('valor_em_risco'), True)} em risco; {fmt_int(a360r.get('a_renegociar') or 0)} com prazo inviável (renegociar com o cliente) e {fmt_int(a360r.get('recuperaveis') or 0)} ainda recuperáveis."))
    crit = [e.get("curto") or e.get("nome") for e in est if e.get("semaforo") == "critico"]
    if crit:
        passos.append((LARANJA, "Reforçar as estações críticas", ", ".join(crit[:5]) + " — concentram itens com prazo vencido ou liberação vencida; priorizar a fila e revisar capacidade."))
    if rt.get("cli_vencido") or rt.get("cli_7d"):
        passos.append((AMARELO, "Acompanhar prazos de cliente", f"{fmt_int(rt.get('cli_vencido') or 0)} itens com prazo do cliente vencido e {fmt_int(rt.get('cli_7d') or 0)} vencendo em até 7 dias."))
    if rt.get("sem_prev") or cap.get("total", {}).get("sem_prev"):
        n_sp = rt.get("sem_prev") or cap["total"].get("sem_prev")
        passos.append((AZUL, "Completar o Planejamento PCP", f"{fmt_int(n_sp)} itens abertos sem liberação prevista definida."))
    if not seg.get("pedidos_total") and seg_lbl:
        passos.append((AZUL, f"Montar o planejamento de {seg_lbl}", "Nenhum pedido planejado no PCP para o mês seguinte."))
    if not passos:
        passos.append((VERDE, "Manter a cadência atual", "Sem desvios relevantes de prazo no momento."))
    d.proximos_passos(passos[:5], manchete="Gerado automaticamente a partir dos indicadores do Painel — priorize de cima para baixo")
    return d.bytes()


# =====================================================================
#  PLANEJAMENTO MENSAL PCP / OPERAÇÃO
# =====================================================================
def _status_linha(l):
    return "concluido" if l.status_producao == "FINALIZADO" else ("fila" if l.status_producao == "PENDENTE" else "producao")


def _resumo_bloco(linhas, hoje):
    """KPIs de um mês (lista de linhas = 1 item de pedido cada)."""
    pedidos = OrderedDict()
    for l in linhas:
        pedidos.setdefault(l.pedido_id, []).append(l)
    valor = sum(l.venda_total or 0 for l in linhas)
    concl = [l for l in linhas if l.status_producao == "FINALIZADO"]
    abertas = [l for l in linhas if l.status_producao != "FINALIZADO"]
    atrasadas = [l for l in abertas if l.liberacao_prevista and l.liberacao_prevista < hoje]
    cli_venc = [l for l in abertas if l.data_cliente and l.data_cliente < hoje]
    pcp_apos = [l for l in abertas if l.data_cliente and l.liberacao_prevista and l.liberacao_prevista > l.data_cliente and l.data_cliente >= hoje]
    sem_plano = [l for l in abertas if not l.planejamento_semanal]
    n_ped_concl = sum(1 for ls in pedidos.values() if all(x.status_producao == "FINALIZADO" for x in ls))
    return {
        "linhas": linhas, "pedidos": pedidos, "n_pedidos": len(pedidos), "n_itens": len(linhas), "valor": valor,
        "valor_concl": sum(l.venda_total or 0 for l in concl), "valor_aberto": sum(l.venda_total or 0 for l in abertas),
        "n_concl": len(concl), "n_abertas": len(abertas), "atrasadas": atrasadas, "cli_venc": cli_venc, "pcp_apos": pcp_apos,
        "sem_plano": sem_plano, "n_ped_concl": n_ped_concl,
        "pct_itens": (100 * len(concl) / len(linhas)) if linhas else 0,
        "pct_valor": (100 * sum(l.venda_total or 0 for l in concl) / valor) if valor else 0,
    }


def _por_semana(linhas, aux):
    grupos = OrderedDict()
    for l in sorted(linhas, key=lambda x: (aux["chave_semana"](x.planejamento_semanal) or (9999, 99, 0))):
        grupos.setdefault(l.planejamento_semanal or "Sem planejamento", []).append(l)
    return grupos


def _semana_curta(rot, aux):
    ch = aux["chave_semana"](rot) if rot and rot != "Sem planejamento" else None
    if not ch:
        return "Sem plano"
    return f"S{ch[2]:02d} · {aux['meses_pt'][ch[1] - 1][:3]}"


def _top(linhas, chave, n=6):
    soma = Counter()
    for l in linhas:
        k = (chave(l) or "—")
        soma[k] += l.venda_total or 0
    return soma.most_common(n)


def montar_planejamento_mensal(blocos, filtros_txt, modo, gerado_em, aux, logo_path=None, hoje=None):
    """blocos: [{mes_ano:(ano,mes), linhas:[_LinhaListagemGeral], rotulo}] (1 = mês; 2 = mês + backlog).
    aux: {chave_semana, meses_pt (nomes por extenso)}."""
    hoje = hoje or date.today()
    completa = modo == "completa"
    nomes = [f"{aux['meses_pt'][b['mes_ano'][1] - 1]}/{b['mes_ano'][0]}" for b in blocos]
    d = Deck("Planejamento Mensal PCP/Operação", gerado_em, logo_path, periodo=_mes_extenso(nomes[0]), base="Listagem gerencial PCP (Planejamento semanal)")
    R = [_resumo_bloco(b["linhas"], hoje) for b in blocos]
    cores_bl = [VERDE, AZUL]
    mes_atual = (hoje.year, hoje.month)

    d.capa("Planejamento Mensal\nPCP / Operação", _mes_extenso(nomes[0]) + (f" + Backlog {_mes_extenso(nomes[1])}" if len(blocos) > 1 else ""),
           area="PLANEJAMENTO\nPCP / OPERAÇÃO",
           meta=[("Filtros: " + _corta(filtros_txt, 90) if filtros_txt else "Base: Listagem gerencial PCP  |  todos os pedidos do período"), f"Gerado em {gerado_em}  ·  Apresentação {'completa' if completa else 'executiva'}"])

    # ---- Visão geral ----
    r0 = R[0]
    manchete = f"{fmt_brl(r0['valor'], True)} planejados em {nomes[0]}: {fmt_pct(r0['pct_valor'])} do valor já concluído, {plural(len(r0['atrasadas']), 'item atrasado', 'itens atrasados')}"
    s = d.slide("Visão geral do planejamento", manchete,
                notas="Valor = soma do valor dos itens com Planejamento PCP no mês. Concluído = itens FINALIZADOS. Atrasado = item aberto com liberação prevista vencida.")
    if len(blocos) > 1:
        r1 = R[1]
        for i, (r, nome) in enumerate(zip(R[:2], nomes[:2])):
            x = MARGEM + i * 6.25
            cor = cores_bl[i]
            d._rect(s, x, 1.6, 5.95, 1.95, fill=cor, forma=MSO_SHAPE.ROUNDED_RECTANGLE, raio=0.05)
            d._texto(s, x + 0.25, 1.72, 5.5, 0.35, ("MÊS DO RELATÓRIO · " if i == 0 else "BACKLOG PCP · ") + nome.upper(), tam=13, cor=BRANCO, negrito=True)
            d._texto(s, x + 0.25, 2.1, 5.5, 0.85, fmt_brl(r["valor"], True), tam=44, cor=BRANCO, negrito=True, ancora=MSO_ANCHOR.MIDDLE)
            d._texto(s, x + 0.25, 3.0, 5.5, 0.4, f"{plural(r['n_pedidos'], 'pedido')} · {plural(r['n_itens'], 'item', 'itens')}", tam=14, cor="E6EEFF")
        d.kpis(s, [
            {"rotulo": f"Concluído (valor) · {nomes[0]}", "valor": fmt_pct(R[0]["pct_valor"]), "sub": f"{fmt_int(R[0]['n_concl'])} de {fmt_int(R[0]['n_itens'])} itens concluídos", "cor": VERDE},
            {"rotulo": "Itens atrasados", "valor": fmt_int(len(R[0]["atrasadas"])), "sub": f"em {nomes[0]}", "cor": VERMELHO if R[0]["atrasadas"] else VERDE},
            {"rotulo": "Sem planejamento semanal", "valor": fmt_int(len(R[0]["sem_plano"]) + len(R[1]["sem_plano"])), "sub": "itens abertos nos 2 meses", "cor": AMARELO if (R[0]["sem_plano"] or R[1]["sem_plano"]) else VERDE},
            {"rotulo": "Backlog / mês", "valor": fmt_pct(100 * R[1]["valor"] / R[0]["valor"]) if R[0]["valor"] else "—", "sub": f"{nomes[1]} vs {nomes[0]}", "cor": NAVY_2},
        ], y=3.85, h=1.35)
        delta = R[1]["valor"] - R[0]["valor"]
        d._texto(s, MARGEM, 5.55, W - 2 * MARGEM, 1.2,
                 [[(f"Leitura: ", {"negrito": True, "cor": NAVY}), (f"o backlog de {nomes[1]} ({fmt_brl(R[1]['valor'], True)}) está {fmt_brl(abs(delta), True)} {'acima' if delta >= 0 else 'abaixo'} do planejado de {nomes[0]} — "
                   + ("a carteira do mês seguinte já sustenta o ritmo atual." if delta >= 0 else "há espaço para completar o planejamento do mês seguinte."), {})]], tam=15, cor=TEXTO)
    else:
        d.kpis(s, [
            {"rotulo": "Pedidos", "valor": fmt_int(r0["n_pedidos"]), "sub": f"{fmt_int(r0['n_ped_concl'])} concluídos", "cor": NAVY},
            {"rotulo": "Itens", "valor": fmt_int(r0["n_itens"]), "sub": f"{fmt_int(r0['n_concl'])} concluídos", "cor": NAVY_2},
            {"rotulo": "Valor planejado", "valor": fmt_brl(r0["valor"], True), "sub": f"{fmt_brl(r0['valor_concl'], True)} concluídos", "cor": AZUL},
            {"rotulo": "Concluído (valor)", "valor": fmt_pct(r0["pct_valor"]), "sub": "do planejado no mês", "cor": VERDE},
            {"rotulo": "Itens atrasados", "valor": fmt_int(len(r0["atrasadas"])), "sub": "liberação prevista vencida", "cor": VERMELHO if r0["atrasadas"] else VERDE},
        ])
        msgs = []
        msgs.append((VERDE if r0["pct_valor"] >= 50 else AMARELO, f"{fmt_pct(r0['pct_valor'])} do valor de {nomes[0]} já está concluído ({fmt_brl(r0['valor_concl'], True)} de {fmt_brl(r0['valor'], True)})."))
        if r0["atrasadas"]:
            msgs.append((VERMELHO, (f"1 item aberto está atrasado frente à liberação prevista pelo PCP." if len(r0["atrasadas"]) == 1 else f"{fmt_int(len(r0['atrasadas']))} itens abertos estão atrasados frente à liberação prevista pelo PCP.")))
        if r0["sem_plano"]:
            msgs.append((AMARELO, f"{plural(len(r0['sem_plano']), 'item aberto', 'itens abertos')} ainda sem semana de planejamento PCP."))
        if r0["cli_venc"]:
            msgs.append((VERMELHO, f"{plural(len(r0['cli_venc']), 'item aberto', 'itens abertos')} com prazo do cliente já vencido."))
        d.mensagens(s, MARGEM, 3.3, W - 2 * MARGEM, 3.5, msgs[:4], tam=15)

    # ---- Por mês: faturamento por semana ----
    for i, (b, r, nome) in enumerate(zip(blocos, R, nomes)):
        grupos = _por_semana(r["linhas"], aux)
        if not grupos:
            continue
        cats = [_semana_curta(k, aux) for k in grupos]
        v_c = [sum((l.venda_total or 0) for l in ls if l.status_producao == "FINALIZADO") for ls in grupos.values()]
        v_a = [sum((l.venda_total or 0) for l in ls if l.status_producao != "FINALIZADO") for ls in grupos.values()]
        pior = max(grupos.items(), key=lambda kv: sum(l.venda_total or 0 for l in kv[1]))
        s = d.slide(f"Faturamento por semana — {nome}" + (" (backlog PCP)" if i == 1 else ""),
                    f"Semana mais carregada: {_semana_curta(pior[0], aux)} com {fmt_brl(sum(l.venda_total or 0 for l in pior[1]), True)}",
                    notas="Cada barra é uma semana do Planejamento PCP. Verde = valor já concluído; laranja = valor ainda em aberto.")
        d.grafico(s, "col", cats, [("Concluído", v_c, VERDE), ("Em aberto", v_a, LARANJA)], MARGEM, 1.6, 7.4, 5.1, empilhado=True, tam=12, gap=50, fmt='#,##0;-#,##0;;')
        lin, cor = [], {}
        for j, (k, ls) in enumerate(grupos.items()):
            ped = len({l.pedido_id for l in ls})
            cc = sum(1 for l in ls if l.status_producao == "FINALIZADO")
            lin.append([_semana_curta(k, aux), fmt_int(ped), fmt_int(len(ls)), f"{round(100 * cc / len(ls))}%", fmt_brl(sum(l.venda_total or 0 for l in ls), True)])
            ch = aux["chave_semana"](k) if k != "Sem planejamento" else None
            fim = aux["fim_semana"](k) if (ch and k != "Sem planejamento") else None
            if fim and fim < hoje and cc < len(ls):
                cor[(j, 3)] = VERMELHO
            if k == "Sem planejamento":
                cor[(j, 0)] = AMARELO
        lin.append(["TOTAL", fmt_int(r["n_pedidos"]), fmt_int(r["n_itens"]), f"{round(r['pct_itens'])}%", fmt_brl(r["valor"], True)])
        d.tabela(s, ["Semana", "Pedidos", "Itens", "Concl.", "Valor"], lin, [1.35, 0.85, 0.7, 0.75, 1.3], x=8.3, y=1.7, tam=11, alt_linha=0.42, cores=cor, alinh=["l", "c", "c", "c", "r"], negrito_ultima=True)
        if not completa and i == 0 and len(blocos) == 1:
            pass

    # ---- Projeção de fechamento do mês ----
    b0, r0 = blocos[0], R[0]
    if b0["mes_ano"] == mes_atual and r0["valor"]:
        frac = hoje.day / _dias_mes(hoje)
        proj = min(r0["valor_concl"] / frac, r0["valor"]) if frac and r0["valor_concl"] else r0["valor_concl"]
        s = d.slide(f"Projeção de fechamento — {nomes[0]}",
                    f"Ao ritmo atual o mês fecha em ~{fmt_brl(proj, True)} ({fmt_pct(100 * proj / r0['valor'])} do planejado); {fmt_brl(r0['valor'] - proj, True)} dependem de aceleração",
                    notas=("Projeção linear: valor concluído ÷ fração do mês decorrida (dia/dias do mês), limitada ao planejado. "
                           "'Dependem de aceleração' = planejado − projeção."))
        ch = d.grafico(s, "col", ["Concluído até hoje", "Projeção ao ritmo atual", "Planejado PCP"], [("Valor", [r0["valor_concl"], proj, r0["valor"]], NAVY)],
                       MARGEM, 1.6, 7.0, 5.1, fmt='"R$ "#,##0', tam=12, gap=70)
        for idx, cor_ in enumerate((VERDE, AZUL, NAVY)):
            pt = ch.plots[0].series[0].points[idx]
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = rgb(cor_)
        d.kpis(s, [
            {"rotulo": "Mês decorrido", "valor": fmt_pct(100 * frac), "sub": f"dia {hoje.day} de {_dias_mes(hoje)}", "cor": NAVY, "tam": 28},
            {"rotulo": "Valor concluído", "valor": fmt_pct(r0["pct_valor"]), "sub": "do planejado", "cor": VERDE if r0["pct_valor"] >= 100 * frac - 15 else VERMELHO, "tam": 28},
        ], y=1.7, h=1.3, x=7.9, w=4.9)
        folga = r0["valor"] - proj
        msgs = [(VERDE if folga <= 0 else (AMARELO if folga < 0.25 * r0["valor"] else VERMELHO),
                 f"Faltam {fmt_brl(r0['valor_aberto'], True)} em {plural(r0['n_abertas'], 'item aberto', 'itens abertos')} para fechar o planejado."),
                (AZUL, f"Para cumprir 100% é preciso concluir ~{fmt_brl(r0['valor_aberto'] / max(1, _dias_mes(hoje) - hoje.day), True)} por dia até o fim do mês.")]
        if r0["atrasadas"]:
            msgs.append((VERMELHO, (f"1 item com liberação vencida pressiona o fechamento." if len(r0["atrasadas"]) == 1 else f"{fmt_int(len(r0['atrasadas']))} itens com liberação vencida pressionam o fechamento.")))
        d.mensagens(s, 7.9, 3.2, 4.9, 3.5, msgs, titulo="Para fechar o mês", tam=13)

    # ---- Concentração (clientes / UF) ----
    top_cli = _top(r0["linhas"], lambda l: l.cliente, 6)
    top_uf = _top(r0["linhas"], lambda l: l.estado, 6)
    if top_cli:
        pc = 100 * top_cli[0][1] / r0["valor"] if r0["valor"] else 0
        s = d.slide(f"Onde está o faturamento — {nomes[0]}", f"{top_cli[0][0]} concentra {fmt_pct(pc)} do valor planejado",
                    notas="Concentração por cliente e por estado (UF) sobre o valor dos itens planejados no mês.")
        d.grafico(s, "col", [c[:26] for c, _ in top_cli], [("Valor", [v for _, v in top_cli], NAVY)], MARGEM, 1.6, 6.1, 5.1, horizontal=True, fmt='#,##0', tam=12, titulo="Top clientes (R$)", gap=45, destaque="max")
        d.grafico(s, "col", [c for c, _ in top_uf], [("Valor", [v for _, v in top_uf], AZUL)], 6.9, 1.6, 5.9, 5.1, horizontal=True, fmt='#,##0', tam=12, titulo="Top estados / UF (R$)", gap=45, destaque="max")

    if completa:
        for i, (r, nome) in enumerate(zip(R, nomes)):
            if not r["linhas"]:
                continue
            # situação + estação
            cont = Counter(_status_linha(l) for l in r["linhas"])
            est = Counter((aux.get("rotulo_estacao", lambda x: x)(l.estacao) if l.estacao else "Sem estação") for l in r["linhas"] if l.status_producao != "FINALIZADO")
            s = d.slide(f"Situação dos itens — {nome}", f"{fmt_int(cont['concluido'])} concluídos · {fmt_int(cont['producao'])} em produção · {fmt_int(cont['fila'])} pendentes")
            d.grafico(s, "rosca", ["Concluído", "Em produção", "Pendente"], [("Itens", [cont["concluido"], cont["producao"], cont["fila"]], VERDE)], MARGEM, 1.6, 5.2, 5.0,
                      cor_pontos=[VERDE, AZUL, AMARELO], tam=13, titulo="Itens por situação")
            if est:
                top_e = est.most_common(8)
                d.grafico(s, "col", [k for k, _ in top_e], [("Itens abertos", [v for _, v in top_e], NAVY_2)], 6.2, 1.6, 6.6, 5.0, horizontal=True, tam=12, titulo="Itens abertos por estação", gap=45, destaque="max")
            # pedidos (top 10 por valor)
            peds = []
            for pid, ls in r["pedidos"].items():
                peds.append((sum(l.venda_total or 0 for l in ls), ls))
            peds.sort(key=lambda t: -t[0])
            s = d.slide(f"Maiores pedidos — {nome}", "Top 10 por valor planejado no mês")
            lin, cor = [], {}
            for j, (v, ls) in enumerate(peds[:10]):
                l0 = ls[0]
                abertos = [x for x in ls if x.status_producao != "FINALIZADO"]
                prev = max((x.liberacao_prevista for x in abertos if x.liberacao_prevista), default=None)
                sit = "Concluído" if not abertos else f"{len(abertos)} de {len(ls)} abertos"
                lin.append([l0.pedido_venda, (l0.cliente or "")[:28], l0.estado or "", fmt_data(l0.data_cliente), fmt_data(prev), sit, fmt_brl(v, True)])
                cor[(j, 5)] = VERDE if not abertos else LARANJA
                if prev and l0.data_cliente and prev > l0.data_cliente:
                    cor[(j, 4)] = VERMELHO
            d.tabela(s, ["Pedido", "Cliente", "UF", "Prazo cliente", "Última prev. PCP", "Situação", "Valor"], lin, [1.1, 3.3, 0.7, 1.5, 1.8, 2.0, 1.7], y=1.65, tam=12, alt_linha=0.46, cores=cor, alinh=["l", "l", "c", "c", "c", "c", "r"])
            # itens atrasados
            if r["atrasadas"]:
                at = sorted(r["atrasadas"], key=lambda l: l.liberacao_prevista)
                s = d.slide(f"Itens atrasados — {nome}", f"{plural(len(at), 'item aberto', 'itens abertos')} com liberação prevista vencida (mais antigos primeiro)")
                lin = [[l.pedido_venda, (l.cliente or "")[:24], (l.descricao_produto or "")[:30], (aux.get("rotulo_estacao", lambda x: x)(l.estacao) if l.estacao else "—"), fmt_data(l.liberacao_prevista), f"{(hoje - l.liberacao_prevista).days}d", fmt_data(l.data_cliente)] for l in at[:12]]
                cor = {(j, 5): VERMELHO for j in range(len(lin))}
                d.tabela(s, ["Pedido", "Cliente", "Produto", "Estação", "Prev. PCP", "Atraso", "Prazo cliente"], lin, [1.0, 2.5, 3.2, 1.7, 1.4, 1.0, 1.5], y=1.65, tam=11, alt_linha=0.4, cores=cor, alinh=["l", "l", "l", "l", "c", "c", "c"])

    # ---- Próximos passos ----
    passos = []
    for i, (r, nome) in enumerate(zip(R, nomes)):
        if r["cli_venc"]:
            passos.append((VERMELHO, f"{nome}: prazo do cliente vencido", f"{plural(len(r['cli_venc']), 'item aberto', 'itens abertos')} já passaram da data solicitada — renegociar ou priorizar a expedição."))
        if r["pcp_apos"]:
            passos.append((LARANJA, f"{nome}: PCP libera depois do prazo do cliente", f"{plural(len(r['pcp_apos']), 'item', 'itens')} com liberação prevista posterior à data do cliente — antecipar a semana de planejamento ou alinhar com o comercial."))
        if r["atrasadas"] and not r["cli_venc"]:
            passos.append((LARANJA, f"{nome}: liberação prevista vencida", f"{plural(len(r['atrasadas']), 'item', 'itens')} acima da data prevista — replanejar a semana PCP."))
        if r["sem_plano"]:
            passos.append((AMARELO, f"{nome}: itens sem semana de planejamento", f"{plural(len(r['sem_plano']), 'item aberto', 'itens abertos')} sem Planejamento Semanal PCP — definir a semana para ganhar previsibilidade."))
    if len(blocos) > 1 and R[1]["valor"] < 0.5 * R[0]["valor"]:
        passos.append((AZUL, f"Completar o backlog de {nomes[1]}", f"Planejado ({fmt_brl(R[1]['valor'], True)}) abaixo de 50% do mês corrente — avançar a carteira com o comercial."))
    if not passos:
        passos.append((VERDE, "Planejamento sob controle", "Sem itens vencidos, sem lacunas de planejamento e sem conflito com prazo de cliente."))
    d.proximos_passos(passos[:5], manchete="Prioridades geradas a partir do próprio planejamento — de cima para baixo")
    return d.bytes()


# =====================================================================
#  ESTAÇÕES — CONCLUSÕES E GESTÃO DE RISCO DE PRAZOS
# =====================================================================
def _risco_item(it, hoje, aux):
    """(nível 0-3, motivo curto, dias de exposição). Mesmos critérios do PDF
    "Conclusões por Estação" (gestão de risco de prazos) — só itens EM ABERTO:
      CRÍTICO — prazo do cliente vencido, liberação prevista vencida, ou PCP
                libera DEPOIS do prazo do cliente.
      ALTO    — prazo do cliente em até 7 dias, semana do PCP encerrada com o
                item aberto, ou liberação em até 3 dias.
      MÉDIO   — pendente com a semana do PCP acabando em até 3 dias, ou prazo
                do cliente em 8–15 dias.
      NO PRAZO— o resto."""
    ped = it.pedido
    d_cli = (ped.data_cliente - hoje).days if ped.data_cliente else None
    d_prev = (it.liberacao_prevista - hoje).days if it.liberacao_prevista else None
    fim_sem = aux["fim_semana"](it.planejamento_semanal) if it.planejamento_semanal else None
    d_sem = (fim_sem - hoje).days if fim_sem else None
    crit = []
    if d_cli is not None and d_cli < 0:
        crit.append((-d_cli, f"Prazo do cliente venceu há {-d_cli}d"))
    if d_prev is not None and d_prev < 0:
        crit.append((-d_prev, f"Liberação prevista venceu há {-d_prev}d"))
    if ped.data_cliente and it.liberacao_prevista and it.liberacao_prevista > ped.data_cliente and (d_cli is None or d_cli >= 0):
        crit.append(((it.liberacao_prevista - ped.data_cliente).days, f"PCP libera {it.liberacao_prevista.strftime('%d/%m')}, após o prazo do cliente ({ped.data_cliente.strftime('%d/%m')})"))
    if crit:
        crit.sort(reverse=True)
        return 3, crit[0][1], crit[0][0]
    alto = []
    if d_cli is not None and 0 <= d_cli <= 7:
        alto.append(f"Prazo do cliente em {d_cli}d" if d_cli else "Prazo do cliente é hoje")
    if d_sem is not None and d_sem < 0:
        alto.append(f"Semana do PCP encerrou em {fim_sem.strftime('%d/%m')} e o item segue aberto")
    if d_prev is not None and 0 <= d_prev <= 3:
        alto.append(f"Liberação prevista em {d_prev}d" if d_prev else "Liberação prevista é hoje")
    if alto:
        return 2, alto[0], 0
    medio = []
    if it.status_producao == "PENDENTE" and d_sem is not None and 0 <= d_sem <= 3:
        medio.append(f"Semana do PCP encerra em {d_sem}d e ainda não iniciou")
    if d_cli is not None and 7 < d_cli <= 15:
        medio.append(f"Prazo do cliente em {d_cli}d")
    if medio:
        return 1, medio[0], 0
    return 0, "Dentro do prazo", 0


def _veredito_estacao(cont, n_abertos):
    if n_abertos and (cont[3] >= 5 or cont[3] / n_abertos >= 0.3):
        return "CRÍTICA", VERMELHO, 3
    if cont[3] or cont[2]:
        return "ATENÇÃO", LARANJA, 2
    if cont[1]:
        return "MONITORAR", AMARELO, 1
    return "CONTROLADA", VERDE, 0


def montar_conclusoes_estacoes(titulo_periodo, itens_por_estacao, modo, gerado_em, aux, logo_path=None, hoje=None):
    """itens_por_estacao: {nome_estacao: [ItemPedido]} (já filtrado por período/status).
    aux: {chave_semana, fim_semana, rotulo_estacao, meses_pt}."""
    hoje = hoje or date.today()
    completa = modo == "completa"
    rot = aux.get("rotulo_estacao", lambda x: x)
    d = Deck("Conclusões por Estação", gerado_em, logo_path, periodo=_periodo_curto(titulo_periodo), base="Planejamento PCP (itens por estação)")
    com_itens = [n for n, its in itens_por_estacao.items() if its]
    todos = [i for n in com_itens for i in itens_por_estacao[n]]

    def st(i):
        return "concluido" if i.status_producao == "FINALIZADO" else ("fila" if i.status_producao == "PENDENTE" else "producao")

    n_tot = len(todos)
    n_c = sum(1 for i in todos if st(i) == "concluido")
    n_p = sum(1 for i in todos if st(i) == "producao")
    n_f = sum(1 for i in todos if st(i) == "fila")
    abertos = [i for i in todos if st(i) != "concluido"]
    atras = [i for i in abertos if i.liberacao_prevista and i.liberacao_prevista < hoje]
    aval = {i.id: _risco_item(i, hoje, aux) for i in abertos}
    em_risco = [i for i in abertos if aval[i.id][0] >= 2]
    no_prazo_ab = len(abertos) - len(em_risco)
    proj_prazo = n_c + no_prazo_ab

    d.capa("Conclusões\npor Estação", _corta(titulo_periodo, 58),
           area="PRODUÇÃO\nESTAÇÕES",
           meta=[f"{plural(len(com_itens), 'estação', 'estações')} com itens previstos  ·  {plural(n_tot, 'item', 'itens')}", f"Gerado em {gerado_em}  ·  Apresentação {'completa' if completa else 'executiva'}"])

    if not n_tot:
        s = d.slide("Sem itens no período", "Nenhum item com Planejamento PCP encontrado para os filtros escolhidos.")
        return d.bytes()

    # ---- Resumo + projeção ----
    pct_c = 100 * n_c / n_tot
    pct_proj = 100 * proj_prazo / n_tot
    s = d.slide("Resumo do período", f"{fmt_pct(pct_c)} concluído hoje; projeção de {fmt_pct(pct_proj)} dentro do prazo ({fmt_int(proj_prazo)} de {fmt_int(n_tot)} itens)",
                notas=("Projeção no prazo = itens já concluídos + itens abertos classificados como 'Médio' ou 'No prazo' pelos critérios de risco. "
                       "Itens 'Crítico' ou 'Alto' ficam fora da projeção. Critérios completos no slide de gestão de risco."))
    d.kpis(s, [
        {"rotulo": "Itens previstos", "valor": fmt_int(n_tot), "sub": titulo_periodo[:34], "cor": NAVY},
        {"rotulo": "Concluídos", "valor": fmt_int(n_c), "sub": f"{fmt_pct(pct_c)} do previsto", "cor": VERDE},
        {"rotulo": "Em produção", "valor": fmt_int(n_p), "sub": "andamento no chão", "cor": AZUL},
        {"rotulo": "Pendentes", "valor": fmt_int(n_f), "sub": "ainda não iniciaram", "cor": AMARELO},
        {"rotulo": "Atrasados", "valor": fmt_int(len(atras)), "sub": "liberação prevista vencida", "cor": VERMELHO if atras else VERDE},
    ])
    d._texto(s, MARGEM, 3.3, 8, 0.35, "PROJEÇÃO DE CONCLUSÃO NO PRAZO", tam=12, cor=CINZA, negrito=True)
    d.barra(s, MARGEM, 3.7, W - 2 * MARGEM, 0.5, [(n_c, VERDE), (no_prazo_ab, AZUL), (len(em_risco), VERMELHO)])
    leg = [(VERDE, f"Concluído · {n_c}"), (AZUL, f"Aberto, dentro do prazo · {no_prazo_ab}"), (VERMELHO, f"Aberto, em risco (crítico/alto) · {len(em_risco)}")]
    for k, (cor, txt) in enumerate(leg):
        dot = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(MARGEM + k * 4.1), Inches(4.38), Inches(0.17), Inches(0.17))
        dot.fill.solid(); dot.fill.fore_color.rgb = rgb(cor); dot.line.fill.background(); dot.shadow.inherit = False
        d._texto(s, MARGEM + k * 4.1 + 0.26, 4.3, 3.8, 0.32, txt, tam=13, cor=TEXTO, ancora=MSO_ANCHOR.MIDDLE)
    msgs = []
    if em_risco:
        msgs.append((VERMELHO, f"{plural(len(em_risco), 'item', 'itens')} em risco de prazo — {fmt_pct(100 * len(em_risco) / max(1, len(abertos)))} dos itens em aberto."))
    else:
        msgs.append((VERDE, "Nenhum item aberto em risco de prazo no período."))
    ordem_est = sorted(com_itens, key=lambda n: -sum(1 for i in itens_por_estacao[n] if i.status_producao != "FINALIZADO" and aval.get(i.id, (0,))[0] >= 2))
    if ordem_est and any(aval[i.id][0] >= 2 for i in abertos):
        n0 = ordem_est[0]
        k0 = sum(1 for i in itens_por_estacao[n0] if i.status_producao != "FINALIZADO" and aval[i.id][0] >= 2)
        msgs.append((LARANJA, f"{rot(n0)} concentra o maior número de itens em risco ({k0})."))
    if atras:
        msgs.append((VERMELHO, (f"1 item já ultrapassou a liberação prevista pelo PCP." if len(atras) == 1 else f"{fmt_int(len(atras))} itens já ultrapassaram a liberação prevista pelo PCP.")))
    d.mensagens(s, MARGEM, 4.85, W - 2 * MARGEM, 2.05, msgs[:3], titulo="Leitura gerencial", tam=14, passo=0.45)

    # ---- Conclusões por estação ----
    lin, cor = [], {}
    est_ord = sorted(com_itens, key=lambda n: -len(itens_por_estacao[n]))
    for j, n in enumerate(est_ord):
        its = itens_por_estacao[n]
        c = sum(1 for i in its if i.status_producao == "FINALIZADO")
        p = sum(1 for i in its if st(i) == "producao")
        f = sum(1 for i in its if st(i) == "fila")
        a = sum(1 for i in its if i.status_producao != "FINALIZADO" and i.liberacao_prevista and i.liberacao_prevista < hoje)
        lin.append([rot(n), fmt_int(len(its)), fmt_int(c), fmt_int(p), fmt_int(f), fmt_int(a), f"{round(100 * c / len(its))}%"])
        if a:
            cor[(j, 5)] = VERMELHO
        pc = 100 * c / len(its)
        cor[(j, 6)] = VERDE if pc >= 70 else (AMARELO if pc >= 30 else VERMELHO)
    s = d.slide("Conclusões por estação", f"{plural(len(com_itens), 'estação', 'estações')} com itens previstos no período",
                notas="Previsto = itens com Planejamento PCP no período. % = concluídos ÷ previstos da estação.")
    d.grafico(s, "col", [rot(n) for n in est_ord[:8]],
              [("Concluído", [sum(1 for i in itens_por_estacao[n] if i.status_producao == "FINALIZADO") for n in est_ord[:8]], VERDE),
               ("Em produção", [sum(1 for i in itens_por_estacao[n] if st(i) == "producao") for n in est_ord[:8]], AZUL),
               ("Pendente", [sum(1 for i in itens_por_estacao[n] if st(i) == "fila") for n in est_ord[:8]], AMARELO)],
              MARGEM, 1.6, 6.6, 5.1, horizontal=True, empilhado=True, tam=12, gap=45, fmt="0;-0;;")
    d.tabela(s, ["Estação", "Prev.", "Concl.", "Prod.", "Pend.", "Atras.", "% concl."], lin[:11], [1.65, 0.6, 0.65, 0.6, 0.6, 0.65, 0.75], x=7.35, y=1.7, tam=11, alt_linha=0.4,
             cores=cor, alinh=["l", "c", "c", "c", "c", "c", "c"])

    # ---- Evolução por semana ----
    semanas = OrderedDict()
    for i in sorted(todos, key=lambda x: (aux["chave_semana"](x.planejamento_semanal) or (9999, 99, 0))):
        semanas.setdefault(i.planejamento_semanal, []).append(i)
    if len(semanas) > 1:
        def sc(rotulo):
            ch = aux["chave_semana"](rotulo)
            return f"S{ch[2]:02d} · {aux['meses_pt'][ch[1] - 1][:3]}" if ch else rotulo
        cats = [sc(k) for k in semanas]
        s = d.slide("Conclusões por semana do PCP", "Distribuição dos itens previstos em cada semana de planejamento",
                    notas="Semana = Planejamento Semanal PCP do item. Semana encerrada com itens abertos = desvio de planejamento.")
        d.grafico(s, "col", cats,
                  [("Concluído", [sum(1 for i in v if i.status_producao == "FINALIZADO") for v in semanas.values()], VERDE),
                   ("Em produção", [sum(1 for i in v if st(i) == "producao") for v in semanas.values()], AZUL),
                   ("Pendente", [sum(1 for i in v if st(i) == "fila") for v in semanas.values()], AMARELO)],
                  MARGEM, 1.6, 7.6, 5.1, empilhado=True, tam=12, gap=50, fmt="0;-0;;")
        lin, cor = [], {}
        for j, (k, v) in enumerate(semanas.items()):
            c = sum(1 for i in v if i.status_producao == "FINALIZADO")
            fim = aux["fim_semana"](k) if k else None
            enc = bool(fim and fim < hoje)
            lin.append([sc(k), fmt_int(len(v)), fmt_int(c), f"{round(100 * c / len(v))}%", "Encerrada" if enc else ("Em curso" if (fim and (fim - hoje).days <= 6) else "Futura")])
            if enc and c < len(v):
                cor[(j, 3)] = VERMELHO
                cor[(j, 4)] = VERMELHO
        d.tabela(s, ["Semana", "Prev.", "Concl.", "%", "Situação"], lin, [1.2, 0.65, 0.7, 0.65, 1.1], x=8.4, y=1.7, tam=11, alt_linha=0.42, cores=cor, alinh=["l", "c", "c", "c", "c"])

    # ---- Gestão de risco: painel ----
    risco_est = {}
    for n in com_itens:
        ab = [i for i in itens_por_estacao[n] if i.status_producao != "FINALIZADO"]
        av = [(i,) + aval[i.id] for i in ab]
        cont = {k: sum(1 for a in av if a[1] == k) for k in (3, 2, 1, 0)}
        maior = max([a[3] for a in av if a[1] == 3] or [0])
        ver = _veredito_estacao(cont, len(ab))
        risco_est[n] = {"av": av, "cont": cont, "abertos": len(ab), "maior": maior, "ver": ver,
                        "pecas": sum((a[0].quantidade or 0) for a in av if a[1] >= 2)}
    ordem_r = sorted(com_itens, key=lambda n: (-risco_est[n]["ver"][2], -risco_est[n]["cont"][3], -risco_est[n]["cont"][2]))
    n_crit = sum(1 for n in com_itens if risco_est[n]["ver"][2] == 3)
    n_aten = sum(1 for n in com_itens if risco_est[n]["ver"][2] == 2)
    s = d.slide("Gestão de risco de prazos", f"{plural(n_crit, 'estação crítica', 'estações críticas')} · {plural(n_aten, 'em atenção', 'em atenção')} · {plural(len(em_risco), 'item', 'itens')} em risco (crítico ou alto)",
                notas=("CRÍTICO: prazo do cliente vencido, liberação prevista vencida ou PCP liberando depois do prazo do cliente. "
                       "ALTO: prazo do cliente em até 7 dias, semana do PCP encerrada com item aberto ou liberação em até 3 dias. "
                       "MÉDIO: pendente com a semana do PCP acabando em até 3 dias ou prazo do cliente em 8–15 dias. "
                       "Veredito da estação: CRÍTICA se ≥5 itens críticos ou ≥30% dos abertos; ATENÇÃO se há crítico/alto; MONITORAR se há médio; senão CONTROLADA."))
    lin, cor = [], {}
    for j, n in enumerate(ordem_r[:9]):
        r = risco_est[n]
        lin.append([rot(n), r["ver"][0] if r["abertos"] else "Sem abertos", fmt_int(r["abertos"]), fmt_int(r["cont"][3]), fmt_int(r["cont"][2]), fmt_int(r["cont"][1]), fmt_int(r["cont"][0]),
                    f"{r['maior']}d" if r["maior"] else "—", fmt_int(r["pecas"])])
        cor[(j, 1)] = r["ver"][1] if r["abertos"] else CINZA
        if r["cont"][3]:
            cor[(j, 3)] = VERMELHO
        if r["cont"][2]:
            cor[(j, 4)] = LARANJA
    d.tabela(s, ["Estação", "Situação", "Abertos", "Crítico", "Alto", "Médio", "No prazo", "Maior atraso", "Peças em risco"], lin, [2.1, 1.5, 1.05, 1.05, 0.95, 1.0, 1.1, 1.5, 1.6],
             y=1.65, tam=12, alt_linha=0.46, cores=cor, alinh=["l", "c", "c", "c", "c", "c", "c", "c", "c"])
    yl = 1.65 + 0.38 + 0.46 * len(lin) + 0.25
    d._texto(s, MARGEM, yl, W - 2 * MARGEM, 0.9,
             [[("Critérios — ", {"negrito": True, "cor": NAVY}), ("Crítico: ", {"negrito": True, "cor": VERMELHO}), ("prazo do cliente vencido, liberação vencida ou PCP libera após o prazo do cliente.  ", {}),
               ("Alto: ", {"negrito": True, "cor": LARANJA}), ("cliente em ≤7d, semana PCP encerrada com item aberto ou liberação em ≤3d.  ", {}),
               ("Médio: ", {"negrito": True, "cor": AMARELO}), ("pendente com semana PCP acabando em ≤3d ou cliente em 8–15d.", {})]], tam=11, cor=CINZA)

    # ---- Itens críticos ----
    todos_av = [(n, a) for n in com_itens for a in risco_est[n]["av"] if a[1] >= 2]
    todos_av.sort(key=lambda t: (-t[1][1], -t[1][3], rot(t[0])))
    if todos_av:
        s = d.slide("Itens que exigem ação imediata", f"{plural(len(todos_av), 'item', 'itens')} em risco; os {min(10, len(todos_av))} mais graves",
                    notas="Ordenado por gravidade (crítico antes de alto) e dias de exposição.")
        lin, cor = [], {}
        for j, (n, a) in enumerate(todos_av[:10]):
            it, nivel, motivo, _dias = a
            lin.append([rot(n), it.pedido.pedido_venda, (it.pedido.cliente or "")[:22], (it.descricao_produto or "")[:26], NIVEL_TXT[nivel], motivo[:58]])
            cor[(j, 4)] = NIVEL_COR[nivel]
        d.tabela(s, ["Estação", "Pedido", "Cliente", "Produto", "Nível", "Motivo"], lin, [1.5, 0.95, 2.0, 2.4, 0.95, 4.4], y=1.65, tam=11, alt_linha=0.46, cores=cor, alinh=["l", "l", "l", "l", "c", "l"])

    if completa:
        # ---- Matriz Estação x Semana ----
        sem_keys = list(semanas.keys())
        if 2 <= len(sem_keys) <= 6 and len(com_itens) >= 2:
            def sc2(rotulo):
                ch = aux["chave_semana"](rotulo)
                return f"S{ch[2]:02d}/{aux['meses_pt'][ch[1] - 1][:3]}" if ch else rotulo
            s = d.slide("Mapa estação × semana", "Itens previstos (concluídos / total) em cada semana do PCP",
                        notas="Cada célula: concluídos / previstos. Vermelho = semana encerrada com itens ainda abertos.")
            lin, cor = [], {}
            for j, n in enumerate(est_ord):
                linha = [rot(n)]
                for k, sk in enumerate(sem_keys):
                    its = [i for i in itens_por_estacao[n] if i.planejamento_semanal == sk]
                    if not its:
                        linha.append("—")
                        continue
                    c = sum(1 for i in its if i.status_producao == "FINALIZADO")
                    linha.append(f"{c} / {len(its)}")
                    fim = aux["fim_semana"](sk)
                    if fim and fim < hoje and c < len(its):
                        cor[(j, k + 1)] = VERMELHO
                    elif c == len(its):
                        cor[(j, k + 1)] = VERDE
                lin.append(linha)
            larg = [2.4] + [(W - 2 * MARGEM - 2.4) / len(sem_keys)] * len(sem_keys)
            d.tabela(s, ["Estação"] + [sc2(k) for k in sem_keys], lin, larg, y=1.65, tam=13, alt_linha=0.5, cores=cor, alinh=["l"] + ["c"] * len(sem_keys))
        # ---- Um slide por estação ----
        for n in ordem_r:
            r = risco_est[n]
            its = itens_por_estacao[n]
            c = sum(1 for i in its if i.status_producao == "FINALIZADO")
            nome_v, cor_v, _ = r["ver"]
            s = d.slide(f"Estação {rot(n)}", f"{plural(len(its), 'item previsto', 'itens previstos')} · {c} concluídos ({round(100 * c / len(its))}%) · {r['abertos']} em aberto",
                        notas="Veredito conforme os critérios de risco do slide 'Gestão de risco de prazos'.")
            d.pilula(s, W - MARGEM - 2.0, 1.22, 2.0, 0.3, (nome_v if r["abertos"] else "SEM ABERTOS"), cor_v if r["abertos"] else CINZA, tam=11)
            d.kpis(s, [
                {"rotulo": "Previstos", "valor": fmt_int(len(its)), "sub": "no período", "cor": NAVY, "tam": 26},
                {"rotulo": "Concluídos", "valor": fmt_int(c), "sub": f"{round(100 * c / len(its))}%", "cor": VERDE, "tam": 26},
                {"rotulo": "Em risco", "valor": fmt_int(r["cont"][3] + r["cont"][2]), "sub": f"{fmt_int(r['pecas'])} peças", "cor": VERMELHO if (r["cont"][3] + r["cont"][2]) else VERDE, "tam": 26},
            ], y=1.6, h=1.4, x=MARGEM, w=5.2)
            d._texto(s, MARGEM, 3.25, 5.2, 0.3, "DISTRIBUIÇÃO DOS ITENS EM ABERTO POR RISCO", tam=11, cor=CINZA, negrito=True)
            d.barra(s, MARGEM, 3.6, 5.2, 0.38, [(r["cont"][k], NIVEL_COR[k]) for k in (3, 2, 1, 0)])
            for k, nv in enumerate((3, 2, 1, 0)):
                xx = MARGEM + (k % 2) * 2.7
                yy = 4.15 + (k // 2) * 0.38
                dot = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(xx), Inches(yy + 0.07), Inches(0.16), Inches(0.16))
                dot.fill.solid(); dot.fill.fore_color.rgb = rgb(NIVEL_COR[nv]); dot.line.fill.background(); dot.shadow.inherit = False
                d._texto(s, xx + 0.24, yy, 2.4, 0.3, f"{NIVEL_TXT[nv]} · {r['cont'][nv]}", tam=12, cor=TEXTO, ancora=MSO_ANCHOR.MIDDLE)
            piores = sorted([a for a in r["av"] if a[1] >= 1], key=lambda a: (-a[1], -a[3]))[:6]
            if piores:
                lin, cor = [], {}
                for j, a in enumerate(piores):
                    it, nivel, motivo, _ = a
                    lin.append([it.pedido.pedido_venda, (it.pedido.cliente or "")[:18], NIVEL_TXT[nivel], motivo[:46]])
                    cor[(j, 2)] = NIVEL_COR[nivel]
                d.tabela(s, ["Pedido", "Cliente", "Nível", "Motivo"], lin, [0.95, 1.7, 0.85, 3.55], x=6.15, y=1.7, tam=11, alt_linha=0.58, cores=cor, alinh=["l", "l", "c", "l"])
            else:
                d._texto(s, 6.15, 1.8, 6.4, 1.0, "Nenhum item aberto com risco de prazo nesta estação.", tam=15, cor=VERDE, negrito=True)

    # ---- Próximos passos ----
    passos = []
    for n in ordem_r:
        r = risco_est[n]
        if r["ver"][2] == 3:
            passos.append((VERMELHO, f"{rot(n)}: situação crítica", f"{r['cont'][3]} itens críticos e {r['cont'][2]} altos (maior atraso {r['maior']}d, {fmt_int(r['pecas'])} peças em risco) — priorizar fila e revisar capacidade."))
    for n in ordem_r:
        r = risco_est[n]
        if r["ver"][2] == 2 and len(passos) < 4:
            passos.append((LARANJA, f"{rot(n)}: atenção", f"{plural(r['cont'][3] + r['cont'][2], 'item', 'itens')} em risco — acompanhar diariamente até regularizar."))
    sem_prev = sum(1 for i in abertos if not i.liberacao_prevista)
    if sem_prev:
        passos.append((AZUL, "Definir liberação prevista", f"{plural(sem_prev, 'item aberto', 'itens abertos')} sem liberação prevista — sem data não há como medir o risco."))
    if atras:
        passos.append((VERMELHO, "Replanejar itens atrasados", f"{plural(len(atras), 'item', 'itens')} acima da liberação prevista — confirmar nova data com o PCP."))
    if not passos:
        passos.append((VERDE, "Período sob controle", "Nenhuma estação crítica ou em atenção e nenhum item atrasado."))
    d.proximos_passos(passos[:5], manchete="Prioridades geradas pelos critérios de risco — de cima para baixo")
    return d.bytes()


# =====================================================================
#  DETALHE DO PEDIDO 360
# =====================================================================
import re as _re

_EMOJI_RE = _re.compile("[\U00010000-\U0010ffff☀-➿⌀-⏿⬀-⯿️‍]+")
_NIVEL_ALERTA = {"bad": VERMELHO, "warn": LARANJA, "info": AZUL, "ok": VERDE}
_ST_ITEM = {"PENDENTE": ("Pendente", AMARELO), "ANDAMENTO": ("Em andamento", AZUL), "EM TRATATIVA": ("Em tratativa", LARANJA), "FINALIZADO": ("Concluído", VERDE)}
_TOM_PRAZO = {"bad": VERMELHO, "warn": LARANJA, "ok": VERDE, "neutro": CINZA}


def _limpa(txt):
    """Remove emojis (nem todo leitor de .pptx renderiza) e espaços sobrando."""
    return _re.sub(r"\s{2,}", " ", _EMOJI_RE.sub("", txt or "")).strip()


def montar_pedido_360(pedido, v, modo, gerado_em, logo_path=None, hoje=None):
    """pedido: Pedido; v: retorno de _visao_pedido_detalhe (já calculado pelo app)."""
    hoje = hoje or date.today()
    completa = modo == "completa"
    d = Deck(f"Pedido {pedido.pedido_venda}", gerado_em, logo_path, periodo=_corta(f"Pedido {pedido.pedido_venda}", 16), base="Detalhe do Pedido 360")
    itens = v["itens"]
    fin, prazo, pcp, lts = v["fin"], v["prazo"], v["pcp"], v["lts"]
    n = v["n_itens"]
    cont = v["cont"]
    abertos = [x for x in itens if x["obj"].status_producao != "FINALIZADO"]
    concluido = n > 0 and not abertos
    if concluido:
        sit_txt, sit_cor = "CONCLUÍDO", VERDE
    elif prazo.get("tom") == "bad":
        sit_txt, sit_cor = "ATRASADO / PRAZO VENCIDO", VERMELHO
    elif prazo.get("tom") == "warn":
        sit_txt, sit_cor = "PRAZO PRÓXIMO", LARANJA
    elif cont.get("PENDENTE", 0) == n:
        sit_txt, sit_cor = "AGUARDANDO INÍCIO", AMARELO
    else:
        sit_txt, sit_cor = "EM ANDAMENTO", AZUL

    d.capa(f"Pedido\n{pedido.pedido_venda}", _corta(pedido.cliente or "", 46), area="PEDIDO 360\nOPERAÇÃO",
           meta=[f"Vendedor: {pedido.vendedor or '—'}  ·  Frete: {pedido.frete or '—'}  ·  {pedido.cidade or ''}{'/' + v['uf'] if v.get('uf') else ''}{('  ·  ' + v['regiao']) if v.get('regiao') else ''}",
                 f"Incluído em {fmt_data(pedido.data_inclusao_pedido)}  ·  Gerado em {gerado_em}  ·  Apresentação {'completa' if completa else 'executiva'}"])

    # ---- Resumo ----
    s = d.slide("Resumo do pedido", f"{pedido.cliente or ''} · {plural(n, 'item', 'itens')} · {fmt_pct(v['pct_final'])} concluído",
                notas="Dados do Detalhe do Pedido 360. Prazo = data solicitada pelo cliente.")
    d.pilula(s, W - MARGEM - 3.2, 1.2, 3.2, 0.32, sit_txt, sit_cor, tam=11)
    dias_txt = (f"{prazo['dias']}d" if prazo.get("dias") is not None else "—")
    d.kpis(s, [
        {"rotulo": "Valor do pedido", "valor": fmt_brl(fin.get("valor_total"), True), "sub": f"ticket médio {fmt_brl(fin.get('ticket'), True)}", "cor": NAVY},
        {"rotulo": "Itens concluídos", "valor": f"{cont.get('FINALIZADO', 0)} / {n}", "sub": f"{fmt_pct(v['pct_final'])} do pedido", "cor": VERDE},
        {"rotulo": "Prazo do cliente", "valor": fmt_data(prazo.get("data"), curto=True) if prazo.get("data") else "—", "sub": prazo.get("texto") or "", "cor": _TOM_PRAZO.get(prazo.get("tom"), CINZA)},
        {"rotulo": "Próxima liberação PCP", "valor": fmt_data(pcp.get("proxima"), curto=True) if pcp.get("proxima") else "—", "sub": f"última: {fmt_data(pcp.get('ultima'), True)}" if pcp.get("ultima") else "sem previsão", "cor": AZUL},
        {"rotulo": "Itens sem plano PCP", "valor": fmt_int(pcp.get("sem_plano") or 0), "sub": "sem semana definida", "cor": AMARELO if pcp.get("sem_plano") else VERDE},
    ], y=1.6)
    msgs = []
    for a in (v.get("alertas") or [])[:4]:
        cor = _NIVEL_ALERTA.get(a.get("nivel"), AZUL)
        msgs.append((cor, _limpa(a.get("titulo")) + (" — " + _limpa(a.get("texto")) if a.get("texto") else "")))
    if not msgs:
        msgs.append((VERDE, "Sem alertas: pedido dentro do esperado."))
    d.mensagens(s, MARGEM, 3.2, W - 2 * MARGEM, 3.7, [(c, t[:190]) for c, t in msgs], titulo="Alertas do pedido", tam=14, passo=0.78)

    # ---- Jornada ----
    jor = v.get("jornada") or []
    if jor:
        s = d.slide("Jornada do pedido", "Do cadastro à entrega: onde o pedido está agora",
                    notas="Verde = etapa concluída; azul = etapa atual; cinza = etapa futura. 'feitos/total' = itens que já passaram pela etapa.")
        k = len(jor)
        cw = (W - 2 * MARGEM) / k
        d._rect(s, MARGEM + cw / 2, 2.82, cw * (k - 1), 0.06, fill="E2E8F2")
        for i, e in enumerate(jor):
            cx = MARGEM + i * cw
            cor = VERDE if e.get("estado") == "feito" else (AZUL if (e.get("estado") == "atual" or e.get("atual")) else "B8C0CE")
            c = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx + cw / 2 - 0.3), Inches(2.55), Inches(0.6), Inches(0.6))
            c.fill.solid(); c.fill.fore_color.rgb = rgb(cor); c.line.color.rgb = rgb(BRANCO); c.line.width = Pt(3); c.shadow.inherit = False
            d._texto(s, cx + cw / 2 - 0.3, 2.55, 0.6, 0.6, str(i + 1), tam=16, cor=BRANCO, negrito=True, alinh=PP_ALIGN.CENTER, ancora=MSO_ANCHOR.MIDDLE)
            d._texto(s, cx + 0.05, 3.3, cw - 0.1, 0.7, e.get("titulo") or "", tam=13, cor=NAVY if cor != "B8C0CE" else CINZA, negrito=True, alinh=PP_ALIGN.CENTER)
            sub = f"{e.get('feitos', 0)}/{e.get('total', 0)} itens"
            d._texto(s, cx + 0.05, 4.0, cw - 0.1, 0.35, sub, tam=12, cor=TEXTO, alinh=PP_ALIGN.CENTER)
            if e.get("data"):
                d._texto(s, cx + 0.05, 4.35, cw - 0.1, 0.35, fmt_data(e["data"], True), tam=12, cor=CINZA, alinh=PP_ALIGN.CENTER)
            elif e.get("sub"):
                d._texto(s, cx + 0.05, 4.35, cw - 0.1, 0.5, _limpa(e["sub"]), tam=11, cor=CINZA, alinh=PP_ALIGN.CENTER)
        atual = next((e for e in jor if e.get("estado") == "atual" or e.get("atual")), None)
        feitas = sum(1 for e in jor if e.get("estado") == "feito")
        d.cartao(s, MARGEM, 5.3, W - 2 * MARGEM, 1.4)
        d._texto(s, MARGEM + 0.3, 5.45, W - 2 * MARGEM - 0.6, 1.1,
                 [[("Etapa atual: ", {"negrito": True, "cor": NAVY}), (_limpa(atual["titulo"]) if atual else ("Pedido concluído" if concluido else "—"), {})],
                  [(f"{feitas} de {k} etapas concluídas ({round(100 * feitas / k)}%).", {"cor": CINZA})]], tam=16, cor=TEXTO)

    # ---- Itens ----
    s = d.slide("Itens do pedido", f"{cont.get('FINALIZADO', 0)} concluídos · {cont.get('ANDAMENTO', 0) + cont.get('EM TRATATIVA', 0)} em andamento · {cont.get('PENDENTE', 0)} pendentes")
    d.grafico(s, "rosca", ["Concluído", "Em andamento", "Pendente"], [("Itens", [cont.get("FINALIZADO", 0), cont.get("ANDAMENTO", 0) + cont.get("EM TRATATIVA", 0), cont.get("PENDENTE", 0)], VERDE)],
              MARGEM, 1.6, 3.8, 5.0, cor_pontos=[VERDE, AZUL, AMARELO], tam=12)
    lin, cor = [], {}
    limite = 9 if not completa else 9
    for j, x in enumerate(itens[:limite]):
        it = x["obj"]
        st_txt, st_cor = _ST_ITEM.get(it.status_producao, (it.status_producao, CINZA))
        lin.append([(it.descricao_produto or "")[:34], fmt_int(it.quantidade), (aux_rot(it.estacao)) if it.estacao else "—", st_txt, fmt_data(it.liberacao_prevista, True), fmt_brl(x.get("valor"), True)])
        cor[(j, 3)] = st_cor
        if it.status_producao != "FINALIZADO" and it.liberacao_prevista and it.liberacao_prevista < hoje:
            cor[(j, 4)] = VERMELHO
    d.tabela(s, ["Produto", "Qtd", "Estação", "Situação", "Prev. PCP", "Valor"], lin, [2.9, 0.65, 1.3, 1.25, 0.95, 1.2], x=4.55, y=1.7, tam=11, alt_linha=0.46, cores=cor, alinh=["l", "c", "l", "c", "c", "r"])
    if len(itens) > limite:
        d._texto(s, 4.55, 1.7 + 0.38 + 0.46 * limite + 0.1, 8.2, 0.3, f"+ {len(itens) - limite} itens não exibidos (ver apresentação completa ou o sistema).", tam=11, cor=CINZA, italico=True)

    # ---- Prazo e projeção de entrega ----
    logi = v.get("logi") or {}
    ultima = pcp.get("ultima")
    transp = logi.get("transporte_param") if (pedido.frete == "CIF") else None
    from datetime import timedelta
    # item aberto não sai antes de hoje: se a liberação prevista já venceu, a projeção parte de hoje
    base = ultima
    venceu = bool(ultima and abertos and ultima < hoje)
    if venceu:
        base = hoje
    entrega_proj = (base + timedelta(days=transp or 0)) if base else None
    folga = (prazo["data"] - entrega_proj).days if (prazo.get("data") and entrega_proj) else None
    s = d.slide("Prazo e projeção de entrega", None,
                notas=("Entrega projetada = última liberação prevista pelo PCP (+ lead time de transporte parametrizado, quando o frete é CIF). "
                       "Folga = prazo do cliente − entrega projetada. Negativa = prazo comprometido."))
    marcos = [("Pedido incluído", pedido.data_inclusao_pedido, NAVY_2)]
    if pcp.get("proxima"):
        marcos.append(("1ª liberação PCP", pcp["proxima"], AZUL))
    if ultima and ultima != pcp.get("proxima"):
        marcos.append(("Última liberação PCP", ultima, AZUL))
    if entrega_proj and (transp is not None or venceu):
        marcos.append((("Entrega projetada" + (f" (+{transp}d frete)" if transp else "")) + (" — parte de hoje" if venceu else ""), entrega_proj, LARANJA))
    if prazo.get("data"):
        marcos.append(("Prazo do cliente", prazo["data"], _TOM_PRAZO.get(prazo.get("tom"), CINZA)))
    marcos = [m for m in marcos if m[1]]
    marcos.sort(key=lambda m: m[1])
    if len(marcos) >= 2:
        d0, d1 = marcos[0][1], marcos[-1][1]
        span = max(1, (d1 - d0).days)
        x0, x1 = MARGEM + 0.9, W - MARGEM - 0.9
        d._rect(s, x0, 2.78, x1 - x0, 0.06, fill="C9D0DC")
        ult_x = -9
        for i, (nome, dt, cor) in enumerate(marcos):
            xm = x0 + (x1 - x0) * ((dt - d0).days / span)
            c = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(xm - 0.15), Inches(2.66), Inches(0.3), Inches(0.3))
            c.fill.solid(); c.fill.fore_color.rgb = rgb(cor); c.line.color.rgb = rgb(BRANCO); c.line.width = Pt(2); c.shadow.inherit = False
            alto = (i % 2 == 0)
            d._texto(s, xm - 1.1, 2.0 if alto else 3.05, 2.2, 0.62, [[(nome, {"negrito": True, "cor": NAVY, "tam": 12})], [(fmt_data(dt), {"cor": cor, "tam": 13, "negrito": True})]], alinh=PP_ALIGN.CENTER,
                     ancora=MSO_ANCHOR.BOTTOM if alto else MSO_ANCHOR.TOP)
    else:
        d._texto(s, MARGEM, 2.3, W - 2 * MARGEM, 0.8, "Ainda não há datas suficientes (liberação prevista do PCP e prazo do cliente) para projetar a entrega.", tam=16, cor=CINZA)
    if folga is not None:
        cor_f = VERDE if folga >= 3 else (LARANJA if folga >= 0 else VERMELHO)
        titulo_f = f"{'Folga de ' + str(folga) + ' dia(s)' if folga >= 0 else 'Faltam ' + str(-folga) + ' dia(s) para atingir o prazo'}"
        d._rect(s, MARGEM, 4.25, W - 2 * MARGEM, 1.35, fill=tom_claro(cor_f, 0.9), forma=MSO_SHAPE.ROUNDED_RECTANGLE, raio=0.06)
        d._texto(s, MARGEM + 0.3, 4.33, 8, 0.55, titulo_f, tam=24, cor=cor_f, negrito=True, ancora=MSO_ANCHOR.MIDDLE)
        d._texto(s, MARGEM + 0.3, 4.9, W - 2 * MARGEM - 0.6, 0.65,
                 f"Entrega projetada em {fmt_data(entrega_proj)} contra prazo do cliente em {fmt_data(prazo['data'])}" + (" — recomenda-se renegociar a data com o cliente." if folga < 0 else ".") + (" A liberação prevista pelo PCP já venceu; a projeção parte de hoje." if venceu else ""), tam=14, cor=TEXTO)
    elif prazo.get("data"):
        d._texto(s, MARGEM, 4.4, W - 2 * MARGEM, 0.8, f"Prazo do cliente: {fmt_data(prazo['data'])} ({prazo.get('texto')}). Defina a liberação prevista no PCP para calcular a folga.", tam=15, cor=CINZA)
    kp = []
    if lts.get("comercial") is not None:
        kp.append({"rotulo": "Prazo comercial", "valor": f"{lts['comercial']} dias", "sub": "inclusão → data do cliente", "cor": NAVY_2, "tam": 24})
    if lts.get("fila") is not None:
        kp.append({"rotulo": "Tempo em fila", "valor": f"{str(lts['fila']).replace('.', ',')} dias", "sub": "média dos itens", "cor": AMARELO, "tam": 24})
    if lts.get("chao") is not None:
        kp.append({"rotulo": "Tempo de chão", "valor": f"{str(lts['chao']).replace('.', ',')} dias", "sub": "média dos itens", "cor": AZUL, "tam": 24})
    if lts.get("total") is not None:
        kp.append({"rotulo": "Lead time produção", "valor": f"{str(lts['total']).replace('.', ',')} dias", "sub": "fila + chão", "cor": NAVY, "tam": 24})
    if kp:
        d.kpis(s, kp[:4], y=5.8, h=1.25, w=min(W - 2 * MARGEM, 3.3 * len(kp[:4]) + 0.22 * (len(kp[:4]) - 1)))

    if completa:
        # ---- Financeiro ----
        cards = [
            ("Valor do pedido", fin.get("valor_total")), ("Valor em Gestão Operação", fin.get("valor_operacao")), ("Faturado (itens)", fin.get("faturado")),
            ("Valor NF emitida", fin.get("valor_nf")), ("Custo de produção real", fin.get("custo_producao")), ("Frete previsto", fin.get("frete_previsto")),
            ("Frete final", fin.get("frete_final")), ("Custo total", fin.get("custo_total")),
        ]
        cards = [(a, b) for a, b in cards if b is not None]
        if len(cards) >= 2:
            s = d.slide("Financeiro do pedido", f"{fmt_brl(fin.get('valor_total'), True)} em {plural(n, 'item', 'itens')} · {plural(fin.get('qtd_total') or 0, 'unidade')}")
            for i, (rot_, val) in enumerate(cards[:8]):
                col, lin_ = i % 4, i // 4
                d.kpis(s, [{"rotulo": rot_, "valor": fmt_brl(val, True), "sub": "", "cor": NAVY if i else AZUL}], y=1.7 + lin_ * 1.75, h=1.35, x=MARGEM + col * 3.1, w=2.9)
            if fin.get("valor_total") and fin.get("custo_total"):
                marg = fin["valor_total"] - fin["custo_total"]
                d.mensagens(s, MARGEM, 5.2, W - 2 * MARGEM, 1.6, [(VERDE if marg >= 0 else VERMELHO, f"Resultado estimado: {fmt_brl(marg, True)} ({fmt_pct(100 * marg / fin['valor_total'], 1)} do valor do pedido) = valor − custo total.")], titulo="Leitura", tam=14, passo=0.6)
        # ---- Logística ----
        go = v.get("go")
        if go:
            s = d.slide("Logística e entrega", f"Frete {pedido.frete or '—'} · {logi.get('transportadora') or 'transportadora não definida'}")
            otd = logi.get("otd") or {}
            otd_txt = {"no_prazo": "No prazo", "atrasado": "Atrasado", "atrasando": "Atrasando", "adiantado": "Adiantado"}.get(otd.get("status"), "Sem dado")
            otd_cor = {"no_prazo": VERDE, "atrasado": VERMELHO, "atrasando": LARANJA, "adiantado": VERDE}.get(otd.get("status"), CINZA)
            d.kpis(s, [
                {"rotulo": "Status logística", "valor": str(getattr(go, "go_status_logistica", None) or "—")[:16], "sub": "Gestão Operação", "cor": NAVY, "tam": 22},
                {"rotulo": "Expedido em", "valor": fmt_data(getattr(go, "go_data_pedido_expedido", None), True), "sub": "", "cor": AZUL},
                {"rotulo": "Entregue ao cliente", "valor": fmt_data(getattr(go, "go_data_entregue_cliente", None), True), "sub": "", "cor": VERDE},
                {"rotulo": "OTD", "valor": otd_txt, "sub": (f"{otd['dias']}d" if otd.get("dias") is not None else ""), "cor": otd_cor, "tam": 24},
            ], y=1.65)
            nf = [("Nota fiscal", getattr(go, "go_numero_nf", None)), ("Emissão NF", fmt_data(getattr(go, "go_data_emissao_nf", None)) if getattr(go, "go_data_emissao_nf", None) else None),
                  ("Data comprometida", fmt_data(logi.get("comprometida")) if logi.get("comprometida") else None), ("Transporte parametrizado", f"{logi['transporte_param']} dias" if logi.get("transporte_param") is not None else None)]
            lin = [[a, str(b)] for a, b in nf if b]
            if lin:
                d.tabela(s, ["Informação", "Valor"], lin, [3.6, 3.6], y=3.3, tam=13, alt_linha=0.5, alinh=["l", "l"])
        # ---- Linha do tempo ----
        tl = [e for e in (v.get("timeline") or []) if not e.get("hoje")]
        if tl:
            s = d.slide("Linha do tempo", "Marcos do pedido (realizados e previstos)")
            lin, cor = [], {}
            for j, e in enumerate(tl[:12]):
                est = "Atrasado" if e.get("atrasado") else ("Previsto" if e.get("futuro") else "Realizado")
                lin.append([fmt_data(e.get("data")), _limpa(e.get("titulo")), _limpa(e.get("detalhe") or "")[:50], est])
                cor[(j, 3)] = VERMELHO if e.get("atrasado") else (AZUL if e.get("futuro") else VERDE)
            d.tabela(s, ["Data", "Marco", "Detalhe", "Situação"], lin, [1.5, 4.2, 4.6, 1.8], y=1.65, tam=12, alt_linha=0.42, cores=cor, alinh=["l", "l", "l", "c"])
        # ---- Itens detalhados ----
        if len(itens) > 9:
            for pg in range(9, len(itens), 12):
                s = d.slide("Itens do pedido (continuação)", f"Itens {pg + 1} a {min(pg + 12, len(itens))} de {len(itens)}")
                lin, cor = [], {}
                for j, x in enumerate(itens[pg:pg + 12]):
                    it = x["obj"]
                    st_txt, st_cor = _ST_ITEM.get(it.status_producao, (it.status_producao, CINZA))
                    lin.append([(it.descricao_produto or "")[:46], fmt_int(it.quantidade), aux_rot(it.estacao) if it.estacao else "—", st_txt, fmt_data(it.liberacao_prevista, True), fmt_brl(x.get("valor"), True)])
                    cor[(j, 3)] = st_cor
                d.tabela(s, ["Produto", "Qtd", "Estação", "Situação", "Prev. PCP", "Valor"], lin, [5.0, 1.0, 2.0, 1.6, 1.4, 1.5], y=1.65, tam=12, alt_linha=0.4, cores=cor, alinh=["l", "c", "l", "c", "c", "r"])
        # ---- Histórico ----
        hist = v.get("historico") or []
        if hist:
            s = d.slide("Últimas alterações", f"{plural(len(hist), 'alteração', 'alterações')} registradas — as 10 mais recentes")
            lin = [[h["quando"].strftime("%d/%m/%Y %H:%M") if h.get("quando") else "", h.get("quem") or "", (h.get("escopo") or "")[:26], h.get("campo") or "",
                    f"{h.get('antes')} → {h.get('depois')}"[:44]] for h in hist[:10]]
            d.tabela(s, ["Quando", "Quem", "Onde", "Campo", "Mudança"], lin, [1.8, 1.7, 2.3, 2.3, 4.0], y=1.65, tam=11, alt_linha=0.42, alinh=["l", "l", "l", "l", "l"])

    # ---- Próximos passos ----
    passos = []
    for a in (v.get("alertas") or []):
        if a.get("acao") and a.get("nivel") in ("bad", "warn"):
            passos.append((_NIVEL_ALERTA[a["nivel"]], _limpa(a.get("titulo"))[:80], _limpa(a["acao"])))
    if not passos:
        passos.append((VERDE, "Pedido sem pendências", "Nenhuma ação corretiva sugerida pelo sistema." if not concluido else "Pedido concluído."))
    d.proximos_passos(passos[:5], manchete="Ações sugeridas pelo sistema para este pedido")
    return d.bytes()


def aux_rot(nome):
    """Rótulo de estação (substituído pelo app via set_rotulo_estacao)."""
    return _ROTULO(nome) if nome else ""


_ROTULO = lambda n: n


def set_rotulo_estacao(fn):
    global _ROTULO
    _ROTULO = fn
