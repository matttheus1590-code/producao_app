"""Autorização de faturamento do pedido (Bruno, 26/11/2026).

Documento impresso que leva a conclusão do pedido para a expedição. Fica
EXPOSTO sobre o pedido na expedição, à vista de toda a logística: por isso o
layout (v2, 26/11/2026) é de "cartaz" — nº do pedido gigante, cliente grande,
frete em destaque e faixa verde de COLETA (imediata, ou a partir de dia/hora).
Este módulo concentra as regras (leitura do formulário, formatação das linhas,
totais) e o PDF (reportlab); a página HTML de impressão usa os mesmos dados
(ver templates/autorizacao_faturamento.html).
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


def ler_hora(texto):
    """'14:30' / '9:05' / '1430' -> 'HH:MM'; inválido/vazio -> None."""
    t = (texto or "").strip().replace("h", ":").replace("H", ":")
    if not t:
        return None
    if ":" not in t and t.isdigit() and len(t) in (3, 4):
        t = t[:-2] + ":" + t[-2:]
    try:
        hh, mm = t.split(":")[:2]
        hh, mm = int(hh), int(mm or 0)
    except ValueError:
        return None
    return f"{hh:02d}:{mm:02d}" if 0 <= hh < 24 and 0 <= mm < 60 else None


def ler_volumes(form):
    """Lê as linhas de volume do formulário (listas vol_qtd[], vol_tipo[],
    vol_c[], vol_l[], vol_a[], vol_pb[]). Linha totalmente vazia é ignorada;
    quantidade inválida/zero vira 1."""
    qtds = form.getlist("vol_qtd[]")
    n = len(qtds)

    def col(nome):
        v = form.getlist(nome)
        return v + [""] * (n - len(v))

    tipos, cs, ls, as_, pbs = (col(x) for x in ("vol_tipo[]", "vol_c[]", "vol_l[]", "vol_a[]", "vol_pb[]"))
    linhas = []
    for i in range(min(n, MAX_LINHAS_VOLUME)):
        c, l, a = num_br(cs[i]), num_br(ls[i]), num_br(as_[i])
        pb = num_br(pbs[i])
        if all(x is None for x in (c, l, a, pb)):
            continue
        try:
            q = int(float(str(qtds[i]).replace(",", ".")))
        except ValueError:
            q = 1
        tipo = (tipos[i] or "VOL.").strip().upper()
        linhas.append({
            "quantidade": q if q > 0 else 1,
            "tipo": tipo if tipo in _TIPOS_VALIDOS else "VOL.",
            "comprimento": c, "largura": l, "altura": a, "peso_bruto": pb,
        })
    return linhas


def texto_linha(v):
    """'1X VOL. 500 X 360 X 360' (medidas só se informadas)."""
    medidas = [fmt_num(v[k]) for k in ("comprimento", "largura", "altura") if v.get(k) is not None]
    base = f"{v['quantidade']}X {v['tipo']}"
    return base + (" " + " X ".join(medidas) if medidas else "")


def montar_dados(pedido, aut, data_liberacao=None):
    """Dicionário único usado pela tela de impressão e pelo PDF."""
    vols = []
    for v in (aut.volumes if aut else []):
        d = {
            "quantidade": v.quantidade, "tipo": v.tipo, "comprimento": v.comprimento, "largura": v.largura,
            "altura": v.altura, "peso_bruto": v.peso_bruto,
        }
        d["texto"] = texto_linha(d)
        d["pb_txt"] = f"{fmt_num(v.peso_bruto)} KGS" if v.peso_bruto is not None else ""
        vols.append(d)
    tot_pb = sum((x["peso_bruto"] or 0) for x in vols)
    tot_vol = sum(x["quantidade"] for x in vols)
    data_emissao = aut.data_emissao if aut else None
    hora = (aut.hora_emissao if aut and aut.hora_emissao else "")
    imediata = True if not aut else bool(aut.coleta_imediata)
    if imediata:
        coleta_l1, coleta_l2 = "COLETA LIBERADA", "IMEDIATA"
    else:
        quando = aut.coleta_data.strftime("%d/%m/%Y") + (f" ÀS {aut.coleta_hora}" if aut.coleta_hora else "")
        coleta_l1, coleta_l2 = "COLETA LIBERADA A PARTIR DE", quando
    return {
        "pedido_venda": pedido.pedido_venda or f"#{pedido.id}",
        "cliente": (aut.cliente_doc if aut and aut.cliente_doc else pedido.cliente) or "",
        "frete": (pedido.frete or "—").upper(),
        "situacao": "FINALIZADO (X)",
        "data_emissao": data_emissao,
        "emissao_txt": (data_extenso(data_emissao) + (f" · {hora}" if hora else "")) if data_emissao else "",
        "coleta_imediata": imediata,
        "coleta_l1": coleta_l1,
        "coleta_l2": coleta_l2,
        "volumes": vols,
        "observacoes": (aut.observacoes if aut and aut.observacoes else ""),
        "tot_volumes": tot_vol,
        "tot_pb": tot_pb,
        "tot_txt": (f"{tot_vol} VOLUME" + ("S" if tot_vol != 1 else "") + (f" · {fmt_num(tot_pb)} KGS" if tot_pb else "")) if vols else "",
    }


VERDE = (0 / 255, 156 / 255, 59 / 255)  # verde bandeira, o mesmo do "LIBERADO" no painel


def gerar_pdf(dados, logo_path):
    """PDF A4 "cartaz" para ficar exposto na expedição (layout v2)."""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.utils import ImageReader, simpleSplit
    from reportlab.pdfbase.pdfmetrics import stringWidth
    from reportlab.pdfgen import canvas

    def ajusta(texto, fonte, tamanho_max, largura, tamanho_min=8):
        t = tamanho_max
        while t > tamanho_min and stringWidth(texto, fonte, t) > largura:
            t -= 1
        return t

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    c.setTitle(f"Autorização de faturamento - pedido {dados['pedido_venda']}")
    W, H = A4
    m = 22
    c.setLineWidth(1)
    c.rect(m, m, W - 2 * m, H - 2 * m)
    x0, x1 = m + 22, W - m - 22
    larg = x1 - x0

    # topo: logo + data/hora de emissão
    y = H - m - 22
    try:
        img = ImageReader(logo_path)
        iw, ih = img.getSize()
        lw = 140
        c.drawImage(img, x0, y - lw * ih / iw, width=lw, height=lw * ih / iw, mask="auto")
    except Exception:
        c.setFont("Helvetica-Bold", 16)
        c.drawString(x0, y - 14, "4PIPE solutions")
    c.setFont("Helvetica-Bold", 11)
    c.drawRightString(x1, y - 12, dados["emissao_txt"])

    # título + número do pedido gigante
    y -= 62
    c.setFont("Helvetica-Bold", 15)
    c.drawString(x0, y, "AUTORIZAÇÃO DE FATURAMENTO")
    y -= 10
    caixa_h = 120
    c.setLineWidth(4)
    c.rect(x0, y - caixa_h, larg, caixa_h)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(x0 + 14, y - 26, "PEDIDO")
    num = str(dados["pedido_venda"])
    t = ajusta(num, "Helvetica-Bold", 110, larg - 40, 30)
    c.setFont("Helvetica-Bold", t)
    c.drawCentredString(x0 + larg / 2, y - caixa_h + 26, num)
    y -= caixa_h + 22

    # cliente
    c.setFont("Helvetica-Bold", 10)
    c.setFillGray(0.35)
    c.drawString(x0, y, "CLIENTE")
    c.setFillGray(0)
    y -= 6
    cli = dados["cliente"].upper()
    tam = 26
    linhas = simpleSplit(cli, "Helvetica-Bold", tam, larg)
    while len(linhas) > 2 and tam > 14:
        tam -= 2
        linhas = simpleSplit(cli, "Helvetica-Bold", tam, larg)
    c.setFont("Helvetica-Bold", tam)
    for ln in linhas[:2]:
        y -= tam + 2
        c.drawString(x0, y, ln)
    y -= 18

    # frete (caixa preta) + situação
    bh = 50
    fr = dados["frete"]
    fw = max(110, stringWidth(fr, "Helvetica-Bold", 34) + 36)
    c.setFillGray(0)
    c.rect(x0, y - bh, fw, bh, fill=1, stroke=0)
    c.setFillGray(1)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(x0 + 10, y - 13, "FRETE")
    c.setFont("Helvetica-Bold", 34)
    c.drawString(x0 + 10, y - bh + 8, fr)
    c.setFillGray(0)
    c.setLineWidth(2)
    sx = x0 + fw + 14
    c.rect(sx, y - bh, x1 - sx, bh)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(sx + 10, y - 13, "SITUAÇÃO")
    c.setFont("Helvetica-Bold", 24)
    c.drawString(sx + 10, y - bh + 9, dados["situacao"])
    y -= bh + 16

    # faixa de coleta (verde)
    ch = 92
    c.setFillColorRGB(*VERDE)
    c.rect(x0, y - ch, larg, ch, fill=1, stroke=0)
    c.setFillGray(1)
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(x0 + larg / 2, y - 26, dados["coleta_l1"])
    t2 = ajusta(dados["coleta_l2"], "Helvetica-Bold", 46, larg - 30, 20)
    c.setFont("Helvetica-Bold", t2)
    c.drawCentredString(x0 + larg / 2, y - ch + 18, dados["coleta_l2"])
    c.setFillGray(0)
    y -= ch + 22

    # volumes
    obs_h = 92
    y_min = m + 14 + obs_h + 20
    c.setFont("Helvetica-Bold", 10)
    c.setFillGray(0.35)
    c.drawString(x0, y, "VOLUMES (mm)")
    c.drawRightString(x1, y, "PESO BRUTO")
    c.setFillGray(0)
    y -= 6
    c.setLineWidth(1)
    c.line(x0, y, x1, y)
    n = len(dados["volumes"])
    reserva_total = 46 if dados["volumes"] else 0
    disp = y - y_min - reserva_total
    passo = min(32, max(15, disp / max(n, 1)))
    fonte = 19 if passo >= 30 else (15 if passo >= 22 else 11)
    y -= passo - 4
    for v in dados["volumes"]:
        c.setFont("Helvetica-Bold", fonte)
        c.drawString(x0 + 6, y, v["texto"])
        c.drawRightString(x1, y, v["pb_txt"] or "—")
        y -= passo
    if dados["volumes"]:
        y += passo - 12
        c.setFillGray(0.92)
        c.rect(x0, y - 30, larg, 30, fill=1, stroke=0)
        c.setFillGray(0)
        c.setFont("Helvetica-Bold", 16)
        c.drawString(x0 + 8, y - 21, "TOTAL: " + dados["tot_txt"])

    # observações
    bx, bw = x0 - 12, larg + 24
    by = m + 14
    c.setLineWidth(1.5)
    c.rect(bx, by, bw, obs_h)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(bx + 10, by + obs_h - 18, "OBSERVAÇÕES:")
    c.setFont("Helvetica", 11)
    ty = by + obs_h - 36
    linhas_obs = []
    for par in (dados["observacoes"] or "").splitlines():
        linhas_obs += simpleSplit(par, "Helvetica", 11, bw - 20) or [""]
    for linha in linhas_obs[:4]:
        c.drawString(bx + 10, ty, linha)
        ty -= 14
    c.showPage()
    c.save()
    return buf.getvalue()
