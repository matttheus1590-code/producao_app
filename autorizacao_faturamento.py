"""Autorização de faturamento do pedido (Bruno, 26/11/2026).

Documento impresso que leva a conclusão do pedido para a expedição: pedido,
cliente, frete, situação, datas (emissão e liberação efetiva), volumes com
medidas e pesos, e uma caixa de observações. Este módulo concentra as regras
(leitura do formulário, formatação das linhas, totais) e o PDF (reportlab), no
mesmo desenho do modelo que a empresa já usava. A página HTML de impressão
usa os mesmos dados (ver templates/autorizacao_faturamento.html).
"""
import io

MESES_EXTENSO = [
    "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro",
]

# (valor gravado, rótulo com emoji para a tela). No documento sai só o valor.
TIPOS_VOLUME = [
    ("VOL.", "📦 Volume (VOL.)"),
    ("CAIXA", "🧰 Caixa"),
    ("PALLET", "🪵 Pallet"),
    ("ENGRADADO", "🏗️ Engradado"),
    ("TAMBOR", "🛢️ Tambor"),
    ("ROLO", "🧻 Rolo"),
]
_TIPOS_VALIDOS = {t for t, _ in TIPOS_VOLUME}
MAX_LINHAS_VOLUME = 30


def data_extenso(d):
    return f"{d.day} de {MESES_EXTENSO[d.month - 1]} de {d.year}" if d else ""


def num_br(texto):
    """'3', '3,5', '1.250,75' ou '3.5' -> float; vazio/ruim -> None; negativo -> None."""
    if texto is None:
        return None
    t = str(texto).strip().replace(" ", "")
    if not t:
        return None
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    try:
        v = float(t)
    except ValueError:
        return None
    return v if v >= 0 else None


def fmt_num(v):
    if v is None:
        return ""
    if abs(v - round(v)) < 1e-9:
        return str(int(round(v)))
    return f"{v:.2f}".rstrip("0").rstrip(".").replace(".", ",")


def ler_volumes(form):
    """Lê as linhas de volume do formulário (listas vol_qtd[], vol_tipo[],
    vol_c[], vol_l[], vol_a[], vol_pb[], vol_pl[]). Linha totalmente vazia é
    ignorada; quantidade inválida/zero vira 1."""
    qtds = form.getlist("vol_qtd[]")
    n = len(qtds)

    def col(nome):
        v = form.getlist(nome)
        return v + [""] * (n - len(v))

    tipos, cs, ls, as_, pbs, pls = (col(x) for x in ("vol_tipo[]", "vol_c[]", "vol_l[]", "vol_a[]", "vol_pb[]", "vol_pl[]"))
    linhas = []
    for i in range(min(n, MAX_LINHAS_VOLUME)):
        c, l, a = num_br(cs[i]), num_br(ls[i]), num_br(as_[i])
        pb, pl = num_br(pbs[i]), num_br(pls[i])
        if all(x is None for x in (c, l, a, pb, pl)):
            continue
        try:
            q = int(float(str(qtds[i]).replace(",", ".")))
        except ValueError:
            q = 1
        tipo = (tipos[i] or "VOL.").strip().upper()
        linhas.append({
            "quantidade": q if q > 0 else 1,
            "tipo": tipo if tipo in _TIPOS_VALIDOS else "VOL.",
            "comprimento": c, "largura": l, "altura": a, "peso_bruto": pb, "peso_liquido": pl,
        })
    return linhas


def texto_linha(v):
    """'1X VOL. 500 X 360 X 360' (medidas só se informadas)."""
    medidas = [fmt_num(v[k]) for k in ("comprimento", "largura", "altura") if v.get(k) is not None]
    base = f"{v['quantidade']}X {v['tipo']}"
    return base + (" " + " X ".join(medidas) if medidas else "")


def montar_dados(pedido, aut, data_liberacao):
    """Dicionário único usado pela tela de impressão e pelo PDF."""
    vols = []
    for v in (aut.volumes if aut else []):
        d = {
            "quantidade": v.quantidade, "tipo": v.tipo, "comprimento": v.comprimento, "largura": v.largura,
            "altura": v.altura, "peso_bruto": v.peso_bruto, "peso_liquido": v.peso_liquido,
        }
        d["texto"] = texto_linha(d)
        d["pb_txt"] = f"{fmt_num(v.peso_bruto)} KGS" if v.peso_bruto is not None else ""
        d["pl_txt"] = f"{fmt_num(v.peso_liquido)} KGS" if v.peso_liquido is not None else ""
        vols.append(d)
    soma = lambda k: sum((x[k] or 0) for x in vols)
    return {
        "pedido_venda": pedido.pedido_venda or f"#{pedido.id}",
        "cliente": (aut.cliente_doc if aut and aut.cliente_doc else pedido.cliente) or "",
        "frete": pedido.frete or "—",
        "situacao": "FINALIZADO (X)",
        "data_emissao": aut.data_emissao if aut else None,
        "data_liberacao": data_liberacao,
        "volumes": vols,
        "observacoes": (aut.observacoes if aut and aut.observacoes else ""),
        "tot_volumes": sum(x["quantidade"] for x in vols),
        "tot_pb": soma("peso_bruto"),
        "tot_pl": soma("peso_liquido"),
    }


def gerar_pdf(dados, logo_path):
    """PDF A4 no desenho do modelo da empresa: moldura, logo + data no topo,
    título, cliente/frete/situação, cabeçalho de colunas, linhas de volume e
    caixa de observações no rodapé."""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.utils import ImageReader, simpleSplit
    from reportlab.pdfgen import canvas

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    c.setTitle(f"Autorização de faturamento - pedido {dados['pedido_venda']}")
    W, H = A4
    m = 28
    c.setLineWidth(1)
    c.rect(m, m, W - 2 * m, H - 2 * m)

    x0 = 85
    xr = W - 85
    y = H - 62
    try:
        img = ImageReader(logo_path)
        iw, ih = img.getSize()
        lw = 150
        c.drawImage(img, x0, y - lw * ih / iw + 8, width=lw, height=lw * ih / iw, mask="auto")
    except Exception:
        c.setFont("Helvetica-Bold", 16)
        c.drawString(x0, y, "4PIPE solutions")
    c.setFont("Helvetica", 10)
    c.drawRightString(xr, y - 8, data_extenso(dados["data_emissao"]))

    y = H - 150
    c.setFont("Helvetica-Bold", 13)
    c.drawString(x0, y, f"AUTORIZAÇÃO DE FATURAMENTO DO PEDIDO: {dados['pedido_venda']}")
    y -= 30
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x0, y, "CLIENTE:")
    c.setFont("Helvetica-BoldOblique", 17)
    c.drawString(x0 + 62, y - 1, dados["cliente"][:46])
    y -= 24
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x0, y, "FRETE:")
    c.setFont("Helvetica-BoldOblique", 11)
    c.drawString(x0 + 46, y, dados["frete"])
    y -= 24
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x0, y, f"SITUAÇÃO: {dados['situacao']}")
    if dados["data_liberacao"]:
        y -= 24
        c.drawString(x0, y, f"LIBERAÇÃO EFETIVA: {dados['data_liberacao'].strftime('%d/%m/%Y')}")

    # colunas (mesma ordem do modelo): QUANTIDADE / TIPO e medidas / PESO BRUTO / PESO LIQ.
    y -= 46
    c.setFont("Helvetica-Bold", 10)
    c.drawString(x0, y, "QUANTIDADE  /  TIPO  /  MEDIDAS (mm)")
    c.drawRightString(xr - 92, y, "PESO BRUTO")
    c.drawRightString(xr, y, "PESO LIQ.")
    y -= 14
    c.setLineWidth(0.6)
    c.line(x0, y, xr, y)
    y -= 26

    caixa_h = 92
    y_min = m + caixa_h + 60
    n = len(dados["volumes"])
    passo = 25 if n <= 14 else max(15, int((y - y_min) / max(n, 1)))
    fonte = 15 if n <= 14 else 11
    for v in dados["volumes"]:
        c.setFont("Helvetica-Bold", fonte)
        c.drawString(x0 + 20, y, v["texto"])
        c.drawRightString(xr - 92, y, v["pb_txt"] or "—")
        c.drawRightString(xr, y, v["pl_txt"] or "—")
        y -= passo
    if dados["volumes"]:
        c.setLineWidth(0.6)
        c.line(x0, y + passo - 12, xr, y + passo - 12)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(x0 + 20, y + passo - 28, f"TOTAL: {dados['tot_volumes']} VOLUME(S)")
        c.drawRightString(xr - 92, y + passo - 28, f"{fmt_num(dados['tot_pb'])} KGS" if dados["tot_pb"] else "—")
        c.drawRightString(xr, y + passo - 28, f"{fmt_num(dados['tot_pl'])} KGS" if dados["tot_pl"] else "—")

    # caixa de observações no rodapé
    bx, bw = m + 10, W - 2 * m - 20
    by = m + 10
    c.setLineWidth(1.5)
    c.rect(bx, by, bw, caixa_h)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(bx + 10, by + caixa_h - 20, "OBSERVAÇÕES:")
    c.setFont("Helvetica", 10)
    ty = by + caixa_h - 38
    linhas_obs = []
    for par in (dados["observacoes"] or "").splitlines():
        linhas_obs += simpleSplit(par, "Helvetica", 10, bw - 20) or [""]
    for linha in linhas_obs[:4]:
        c.drawString(bx + 10, ty, linha)
        ty -= 13
    c.showPage()
    c.save()
    return buf.getvalue()
