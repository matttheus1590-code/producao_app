"""Autorização de faturamento do pedido (Bruno, 26/11/2026).

Documento impresso que leva a conclusão do pedido para a expedição (fica
exposto sobre o pedido). Layout CLÁSSICO, no desenho do modelo da empresa —
um "cartaz" colorido foi testado em 26/11/2026 e o Bruno preferiu o clássico —
com: nº do pedido em destaque no título, cliente/frete/situação, linha de
COLETA (imediata, ou a partir de dia/hora), data e hora de emissão no topo,
volumes com medidas e peso bruto, e caixa de observações.
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
        coleta_txt = "COLETA LIBERADA IMEDIATA"
    else:
        quando = aut.coleta_data.strftime("%d/%m/%Y") + (f" ÀS {aut.coleta_hora}" if aut.coleta_hora else "")
        coleta_l1, coleta_l2 = "COLETA LIBERADA A PARTIR DE", quando
        coleta_txt = f"COLETA LIBERADA A PARTIR DE: {quando}"
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
        "coleta_txt": coleta_txt,
        "volumes": vols,
        "observacoes": (aut.observacoes if aut and aut.observacoes else ""),
        "tot_volumes": tot_vol,
        "tot_pb": tot_pb,
        "tot_txt": (f"{tot_vol} VOLUME" + ("S" if tot_vol != 1 else "") + (f" · {fmt_num(tot_pb)} KGS" if tot_pb else "")) if vols else "",
    }


def gerar_pdf(dados, logo_path):
    """PDF A4 no desenho clássico do modelo da empresa."""
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
    c.drawRightString(xr, y - 8, dados["emissao_txt"])

    y = H - 150
    c.setFont("Helvetica-Bold", 17)
    c.drawString(x0, y, f"AUTORIZAÇÃO DE FATURAMENTO DO PEDIDO: {dados['pedido_venda']}")
    y -= 32
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x0, y, "CLIENTE:")
    from reportlab.pdfbase.pdfmetrics import stringWidth
    cli = dados["cliente"]
    tam = 17
    while tam > 10 and stringWidth(cli, "Helvetica-BoldOblique", tam) > (xr - x0 - 62):
        tam -= 1  # nome comprido: reduz a fonte para caber dentro da moldura
    c.setFont("Helvetica-BoldOblique", tam)
    c.drawString(x0 + 62, y - 1, cli)
    y -= 24
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x0, y, "FRETE:")
    c.setFont("Helvetica-BoldOblique", 11)
    c.drawString(x0 + 46, y, dados["frete"])
    y -= 24
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x0, y, f"SITUAÇÃO: {dados['situacao']}")
    y -= 24
    c.drawString(x0, y, dados["coleta_txt"])

    y -= 46
    c.setFont("Helvetica-Bold", 10)
    c.drawString(x0, y, "QUANTIDADE  /  TIPO  /  MEDIDAS (mm)")
    c.drawRightString(xr, y, "PESO BRUTO")
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
        c.drawRightString(xr, y, v["pb_txt"] or "—")
        y -= passo
    if dados["volumes"]:
        c.setLineWidth(0.6)
        c.line(x0, y + passo - 12, xr, y + passo - 12)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(x0 + 20, y + passo - 28, "TOTAL: " + dados["tot_txt"])

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
