"""Relatório executivo (PDF) da gestão de risco de PARADOS (pedido do Bruno):
pedidos sem movimentação e OP/produto parado em fila e espera, horizontes de
5/10/20 dias — mesmo formato e mesma lógica de data da gestão de risco de
prazos.

Módulo independente (não importa o app): recebe os dados já calculados por
_dados_gestao_parados_completa() e só apresenta. Nada de glyph fora do
Latin-1 (setas, ✓...) — a fonte padrão do reportlab os desenha como
quadrado preto; ícones vêm de PNGs (função `icone`)."""

import io
import os
from xml.sax.saxutils import escape as esc

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    CondPageBreak, HRFlowable, Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)

NAVY = "#1b2a4a"
VERMELHO = "#c0392b"
AMARELO = "#d99a00"
CINZA = "#6c757d"
VERDE = "#198754"
BORDA = colors.HexColor("#dee2e6")
_DIAS_SEMANA = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]
_ORDEM = {"critico": 0, "atencao": 1, "sem_prazo": 2}
_ROTULO = {"critico": "Prazo crítico", "atencao": "Prazo folgado", "sem_prazo": "Sem prazo"}
_COR = {"critico": VERMELHO, "atencao": AMARELO, "sem_prazo": CINZA}
_ICONES_OK = {"warning", "target", "calendar", "check", "red_circle", "yellow_circle", "green_circle",
              "hourglass", "magnifier", "bar_chart", "factory", "funnel", "clipboard"}


def _dm(d):
    return f"{d.strftime('%d/%m/%Y')} {_DIAS_SEMANA[d.weekday()]}" if d else "—"


def _dm_curto(d):
    return f"{d.strftime('%d/%m')} {_DIAS_SEMANA[d.weekday()]}" if d else "—"


def _faixa(dias, horizontes):
    """Maior horizonte já atingido: 'há 20d+', 'há 10d+', 'há 5d+'."""
    atingidos = [h for h in horizontes if dias >= h]
    return max(atingidos) if atingidos else min(horizontes)


def gerar(dados, agora, logo_path, icone_base, moeda):
    def icone(nome, tamanho=9):
        return icone_base(nome, tamanho) if nome in _ICONES_OK else ""

    hoje = dados["hoje"]
    horizontes = tuple(dados["horizontes"])
    pressao = dados["pressao_dias"]
    pedidos, fila, espera = dados["pedidos"], dados["fila"], dados["espera"]

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=landscape(A4), leftMargin=8 * mm, rightMargin=8 * mm, topMargin=8 * mm, bottomMargin=13 * mm,
        title="Gestão de risco de parados — visão executiva", author="4PIPE Solutions",
    )
    largura = landscape(A4)[0] - doc.leftMargin - doc.rightMargin
    est = getSampleStyleSheet()
    base = ParagraphStyle("b", parent=est["Normal"], fontSize=7, leading=8.4)
    base_b = ParagraphStyle("bb", parent=base, fontName="Helvetica-Bold")
    peq = ParagraphStyle("peq", parent=base, fontSize=6.3, leading=7.6)
    nota = ParagraphStyle("nota", parent=base, fontSize=6.8, leading=8.6, textColor=colors.HexColor("#555"))
    centro = ParagraphStyle("c", parent=base, alignment=1)
    direita = ParagraphStyle("d", parent=base, alignment=2)
    cab_st = ParagraphStyle("cab", parent=base_b, fontSize=6.6, leading=7.8, textColor=colors.HexColor(NAVY))
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
        out = [CondPageBreak(45 * mm), t]
        if sub_txt:
            out += [Spacer(1, 0.8 * mm), P(sub_txt, nota)]
        out.append(Spacer(1, 1.2 * mm))
        return out

    def tabela(cab, rows, pesos, alin=None, extra=None):
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
        for i in range(2, len(rows) + 1, 2):
            stl.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#f8f9fa")))
        stl += extra or []
        t.setStyle(TableStyle(stl))
        return t

    def prazo_txt(l):
        if l["prazo_cliente"] is None:
            return P(cor("sem prazo", CINZA, False), centro)
        d = l["dias_ate_prazo"]
        quando = f"venceu há {-d}d" if d < 0 else ("vence hoje" if d == 0 else f"em {d}d")
        c = VERMELHO if d <= pressao else CINZA
        return P(f'{l["prazo_cliente"].strftime("%d/%m/%Y")}<br/>{cor(quando, c, d <= pressao)}', centro)

    def dias_cel(l):
        h = _faixa(l["dias_parado"], horizontes)
        c = VERMELHO if h >= max(horizontes) else (AMARELO if h >= sorted(horizontes)[1] else CINZA)
        return P(cor(f'{l["dias_parado"]}d', c), centro)

    def prio_cel(l):
        return P(cor(_ROTULO[l["classe"]], _COR[l["classe"]]), centro)

    def agrega(lista, h):
        sel = [l for l in lista if l["dias_parado"] >= h]
        return {
            "n": len(sel),
            "critico": sum(1 for l in sel if l["classe"] == "critico"),
            "valor": sum(l["valor"] for l in sel),
        }

    def ordena(lista):
        return sorted(lista, key=lambda l: (_ORDEM[l["classe"]], -l["dias_parado"]))

    el = []

    # ============================ CABEÇALHO ============================
    logo = Image(logo_path, width=50 * mm, height=50 * mm * (63 / 261)) if os.path.exists(logo_path) else P("")
    bloco = [
        P(f'{icone("hourglass", 17)} GESTÃO DE RISCO DE PARADOS — VISÃO EXECUTIVA', titulo),
        Spacer(1, 1 * mm),
        P(f'{icone("clipboard", 10)} Emitido em <b>{agora.strftime("%d/%m/%Y")}</b> às <b>{agora.strftime("%H:%M")}</b> (horário de Brasília) · '
          f'{icone("calendar", 10)} Posição de <b>{_dm(hoje)}</b> · pedidos sem movimentação e OP/produto em fila e espera · horizontes de '
          f'<b>{" / ".join(str(h) for h in horizontes)} dias</b> (parado há X dias ou mais; quem já passou do horizonte também entra)', sub),
    ]
    cab = Table([[logo, bloco]], colWidths=[56 * mm, largura - 56 * mm])
    cab.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    el += [cab, Spacer(1, 1.2 * mm), HRFlowable(width="100%", thickness=1.5, color=colors.HexColor(NAVY)), Spacer(1, 2 * mm)]

    el.append(P(
        f'{icone("red_circle", 7)} <b>Prazo crítico</b>: prazo do cliente vencido ou nos próximos {pressao} dias · '
        f'{icone("yellow_circle", 7)} <b>Prazo folgado</b>: parado, mas com prazo do cliente além de {pressao} dias · '
        f'{icone("hourglass", 7)} <b>Sem prazo</b>: pedido sem prazo do cliente cadastrado. '
        f'<b>Movimentação</b> = mudança de estação/status, início ou término de produção/inspeção, liberação, nota fiscal ou envio '
        f'(o carimbo de edição NÃO conta); sem nenhuma, vale a data de inclusão do pedido. '
        f'<b>Fila</b> = OP/produto pendente, ainda sem iniciar; <b>Espera</b> = OP/produto em tratativa.', nota))
    el.append(Spacer(1, 2 * mm))

    # ============================ 1. PAINEL EXECUTIVO ============================
    menor = min(horizontes)
    a_ped, a_fila, a_esp = agrega(pedidos, menor), agrega(fila, menor), agrega(espera, menor)
    ped_max = agrega(pedidos, max(horizontes))
    kp = [
        (str(a_ped["n"]), f"Pedidos parados (há {menor}d+)", NAVY),
        (str(a_ped["critico"]), "Pedidos parados com prazo crítico", VERMELHO if a_ped["critico"] else VERDE),
        (str(ped_max["n"]), f"Pedidos parados há {max(horizontes)}d+", VERMELHO if ped_max["n"] else VERDE),
        (str(a_fila["n"]), f"OP/produto na fila (há {menor}d+)", NAVY),
        (str(a_esp["n"]), f"OP/produto em espera (há {menor}d+)", AMARELO if a_esp["n"] else VERDE),
        (moeda(a_ped["valor"]), "Valor dos pedidos parados (itens abertos)", NAVY),
    ]
    topo = [P(v, ParagraphStyle("kv", parent=base_b, fontSize=13 if len(v) < 12 else 10.5, leading=15, alignment=1, textColor=colors.HexColor(c))) for v, _, c in kp]
    rot = [P(esc(r), ParagraphStyle("kr", parent=base, fontSize=6.6, leading=7.8, alignment=1)) for _, r, _ in kp]
    k = Table([topo, rot], colWidths=[largura / len(kp)] * len(kp))
    k.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.5, BORDA), ("INNERGRID", (0, 0), (-1, -1), 0.5, BORDA), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                           ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3), ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f4f7fc"))]))
    el += [k, Spacer(1, 2.2 * mm)]

    # --- matriz horizonte x visão ---
    linhas_m, est_m = [], []
    for h in horizontes:
        grupos = (("Pedidos FOB", [l for l in pedidos if l.get("frete") == "FOB"]),
                  ("Pedidos CIF", [l for l in pedidos if l.get("frete") == "CIF"]),
                  ("Pedidos (todos)", pedidos),
                  ("OP/produto na fila", fila),
                  ("OP/produto em espera", espera))
        for idx, (nome, lista) in enumerate(grupos):
            a = agrega(lista, h)
            negrito = nome == "Pedidos (todos)"
            stb = base_b if negrito else base
            linhas_m.append([
                (T(f"Parado há {h} dias ou mais", base_b) if idx == 0 else P("")),
                T(nome, stb), P(cor(a["n"], NAVY) if a["n"] else "0", centro),
                P(cor(a["critico"], VERMELHO) if a["critico"] else "0", centro),
                T(moeda(a["valor"]), ParagraphStyle("v", parent=stb, alignment=2)),
            ])
            if negrito:
                est_m.append(("BACKGROUND", (0, len(linhas_m)), (-1, len(linhas_m)), colors.HexColor("#e8eef8")))
        est_m.append(("LINEBELOW", (0, len(linhas_m)), (-1, len(linhas_m)), 0.8, colors.HexColor("#8fa8cc")))
    el.append(tabela(["Horizonte", "Visão", "Quantidade", "Com prazo crítico", "Valor (itens abertos)"],
                     linhas_m, [20, 22, 11, 14, 17], alin="LLCCR", extra=est_m))
    el.append(Spacer(1, 2.4 * mm))

    # --- decisões recomendadas ---
    dec = []
    crit_ped = ordena([l for l in pedidos if l["classe"] == "critico"])
    if crit_ped:
        top = "; ".join(
            f"PV {esc(str(l['pedido_venda'] or l['id']))} {esc((l['cliente'] or '')[:22])} ({l['dias_parado']}d parado)" for l in crit_ped[:3])
        dec.append(("red_circle", VERMELHO,
                    f"<b>{len(crit_ped)} pedido(s) parado(s) com prazo do cliente crítico</b> — destravar hoje ou pedir revisão de prazo ao comercial. Mais críticos: {top}."))
    antigos = [l for l in pedidos if l["dias_parado"] >= max(horizontes) and l["classe"] != "critico"]
    if antigos:
        dec.append(("yellow_circle", AMARELO,
                    f"<b>{len(antigos)} pedido(s) parado(s) há {max(horizontes)} dias ou mais</b> sem pressão imediata de prazo — confirmar se ainda são pedidos válidos "
                    "(cancelar, reprogramar ou ocultar da gestão de risco)."))
    fila_crit = ordena([l for l in fila if l["classe"] == "critico"])
    if fila_crit:
        dec.append(("hourglass", VERMELHO,
                    f"<b>{len(fila_crit)} OP/produto na fila com prazo crítico</b> — programar a entrada em produção; estações com mais itens: "
                    + ", ".join(f"{esc(n)} ({q})" for n, q in _top_estacoes(fila_crit)) + "."))
    if espera:
        dec.append(("hourglass", AMARELO,
                    f"<b>{len(espera)} OP/produto em espera (tratativa) há {menor} dias ou mais</b> — cobrar a resolução do responsável; "
                    f"a mais antiga está parada há {max(l['dias_parado'] for l in espera)} dias."))
    if not dec:
        dec.append(("check", VERDE, "<b>Nada parado além dos limites.</b> Operação fluindo dentro do esperado."))
    el += banner("DECISÕES RECOMENDADAS", NAVY, ic="target")
    linhas_dec = [[P(icone(ic, 9), centro), P(txt)] for ic, _c, txt in dec]
    td = Table(linhas_dec, colWidths=[8 * mm, largura - 8 * mm])
    td.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.25, BORDA), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                            ("TOPPADDING", (0, 0), (-1, -1), 2.4), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.4)]))
    el += [td, Spacer(1, 2.4 * mm)]

    # ============================ 2. PEDIDOS PARADOS ============================
    el += banner(f"PEDIDOS PARADOS SEM MOVIMENTAÇÃO — há {menor} dias ou mais ({len(pedidos)})", NAVY,
                 "Ordenado do prazo mais crítico e do mais antigo para o mais recente. Coluna Faixa = maior horizonte já atingido.", ic="warning")
    if pedidos:
        rows = []
        for l in ordena(pedidos):
            rows.append([
                prio_cel(l),
                T(f"{l['pedido_venda'] or l['id']}", base_b), T((l["cliente"] or "")[:34]),
                T(l["frete"] or "—", centro), T(l["uf"] or "—", centro), dias_cel(l),
                P(f'há {_faixa(l["dias_parado"], horizontes)}d+', centro),
                T(_dm_curto(l["ultimo_mov"]), centro), T(l["ultimo_mov_desc"], peq), T(l["onde"], peq),
                prazo_txt(l), T(f'{l["itens_abertos"]}/{l["itens_total"]}', centro), T(moeda(l["valor"]), direita),
            ])
        el.append(tabela(["Prioridade", "PV", "Cliente", "Frete", "UF", "Parado", "Faixa", "Última mov.", "O que foi", "Onde está", "Prazo cliente",
                          "Itens abertos", "Valor aberto"],
                         rows, [10, 8, 17, 5, 4, 6, 6, 9, 17, 19, 10, 6, 10], alin="CLLCCCCCLLCCR"))
    else:
        el.append(P("Nenhum pedido parado nos horizontes analisados.", nota))
    el.append(Spacer(1, 2.4 * mm))

    # ============================ 3/4. FILA E ESPERA ============================
    def bloco_itens(txt, cor_bg, lista, ic, vazio):
        out = banner(txt, cor_bg, "Contagem a partir da última movimentação do item (ou da entrada do pedido, se nunca se mexeu).", ic=ic)
        if not lista:
            return out + [P(vazio, nota), Spacer(1, 2.4 * mm)]
        rows = []
        for l in ordena(lista):
            rows.append([
                P(cor(_ROTULO[l["classe"]], _COR[l["classe"]]), centro),
                T(f"{l['pedido_venda'] or l['pedido_id']}", base_b), T((l["cliente"] or "")[:28]),
                T((l["produto"] or "")[:44]), T(f"{l['quantidade']:g}" if l.get("quantidade") is not None else "—", centro),
                T(l["estacao"]), dias_cel(l), P(f'há {_faixa(l["dias_parado"], horizontes)}d+', centro),
                T(_dm_curto(l["ultimo_mov"]), centro), T(l["ultimo_mov_desc"], peq), prazo_txt(l),
                T(_dm_curto(l["liberacao_prevista"]) if l.get("liberacao_prevista") else "sem previsão", centro),
            ])
        out.append(tabela(["Prioridade", "PV", "Cliente", "Produto / OP", "Qtd", "Estação", "Parado", "Faixa", "Última mov.", "O que foi",
                           "Prazo cliente", "Liberação PCP"],
                          rows, [10, 8, 15, 21, 4, 12, 6, 6, 9, 17, 10, 9], alin="CLLLCLCCCLCC"))
        out.append(Spacer(1, 2.4 * mm))
        return out

    el += bloco_itens(f"OP/PRODUTO PARADO NA FILA (pendente, sem iniciar) — há {menor} dias ou mais ({len(fila)})", "#34495e", fila, "hourglass",
                      "Nenhuma OP/produto parado na fila nos horizontes analisados.")
    el += bloco_itens(f"OP/PRODUTO PARADO EM ESPERA (em tratativa) — há {menor} dias ou mais ({len(espera)})", "#a87400", espera, "hourglass",
                      "Nenhuma OP/produto parado em espera nos horizontes analisados.")

    # ============================ rodapé de dados ============================
    obs = []
    ref = dados.get("sem_referencia") or {}
    if ref.get("pedido") or ref.get("item"):
        obs.append(f"{ref.get('pedido', 0)} pedido(s) e {ref.get('item', 0)} OP/produto sem nenhuma data de referência (sem inclusão nem movimentação) ficaram fora da análise.")
    if dados.get("ocultos"):
        obs.append(f"{len(dados['ocultos'])} registro(s) oculto(s) pelo PCP não entram nos totais: "
                   + "; ".join(esc(o["titulo"]) for o in dados["ocultos"][:8]) + (" ..." if len(dados["ocultos"]) > 8 else ""))
    for o in obs:
        el.append(P(o, nota))

    carimbo = f'{agora.strftime("%d/%m/%Y")} às {agora.strftime("%H:%M")}'

    def rodape(c, d):
        c.saveState()
        c.setFont("Helvetica", 7.5)
        c.setFillColor(colors.HexColor("#888888"))
        c.drawString(d.leftMargin, 6 * mm, "4PIPE Solutions - Gestão da Produção e Operação - Gestão de risco de parados")
        c.drawRightString(landscape(A4)[0] - d.rightMargin, 6 * mm, f"Emitido em {carimbo} - Página {d.page}")
        c.restoreState()

    doc.build(el, onFirstPage=rodape, onLaterPages=rodape)
    return buf.getvalue()


def _top_estacoes(lista, n=3):
    cont = {}
    for l in lista:
        cont[l["estacao"]] = cont.get(l["estacao"], 0) + 1
    return sorted(cont.items(), key=lambda kv: -kv[1])[:n]
