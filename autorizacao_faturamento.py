"""Autorização de faturamento do pedido (Bruno, 26/11/2026).

Documento impresso que leva a conclusão do pedido para a expedição (fica
exposto sobre o pedido). Layout sóbrio (v3, 26/11/2026 — um "cartaz" colorido
foi testado e descartado): grade com os campos-chave em destaque (PEDIDO DE
VENDA, FRETE, CLIENTE, COLETA LIBERADA e SITUAÇÃO), tabela de volumes com
quantidade, tipo, medidas, PESO BRUTO UNITÁRIO e PESO BRUTO TOTAL
(quantidade × unitário), data/hora de emissão no topo e caixa de observações.
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
        medidas = [fmt_num(d[k]) for k in ("comprimento", "largura", "altura") if d.get(k) is not None]
        d["texto"] = texto_linha(d)
        d["medidas_txt"] = " X ".join(medidas) if medidas else "—"
        d["peso_total"] = (v.peso_bruto * v.quantidade) if v.peso_bruto is not None else None
        d["pu_txt"] = f"{fmt_num(v.peso_bruto)} KGS" if v.peso_bruto is not None else "—"
        d["pt_txt"] = f"{fmt_num(d['peso_total'])} KGS" if d["peso_total"] is not None else "—"
        vols.append(d)
    tot_pb = sum((x["peso_total"] or 0) for x in vols)
    tot_vol = sum(x["quantidade"] for x in vols)
    data_emissao = aut.data_emissao if aut else None
    hora = (aut.hora_emissao if aut and aut.hora_emissao else "")
    imediata = True if not aut else bool(aut.coleta_imediata)
    if imediata:
        coleta_l1, coleta_l2 = "COLETA LIBERADA", "IMEDIATA"
        coleta_txt = "COLETA LIBERADA IMEDIATA"
        coleta_rotulo, coleta_valor = "COLETA LIBERADA", "IMEDIATA"
    else:
        quando = aut.coleta_data.strftime("%d/%m/%Y") + (f" ÀS {aut.coleta_hora}" if aut.coleta_hora else "")
        coleta_l1, coleta_l2 = "COLETA LIBERADA A PARTIR DE", quando
        coleta_txt = f"COLETA LIBERADA A PARTIR DE: {quando}"
        coleta_rotulo = "COLETA LIBERADA A PARTIR DE"
        coleta_valor = aut.coleta_data.strftime("%d/%m/%Y") + (f" às {aut.coleta_hora}" if aut.coleta_hora else "")
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
        "coleta_rotulo": coleta_rotulo,
        "coleta_valor": coleta_valor,
        "volumes": vols,
        "observacoes": (aut.observacoes if aut and aut.observacoes else ""),
        "tot_volumes": tot_vol,
        "tot_pb": tot_pb,
        "tot_pb_txt": (f"{fmt_num(tot_pb)} KGS" if tot_pb else "—"),
        "tot_txt": (f"{tot_vol} VOLUME" + ("S" if tot_vol != 1 else "") + (f" · {fmt_num(tot_pb)} KGS" if tot_pb else "")) if vols else "",
    }


NAVY = (27 / 255, 42 / 255, 74 / 255)  # azul da identidade do sistema
CINZA = 0.93


def gerar_pdf(dados, logo_path):
    """PDF A4 sóbrio, com a grade de campos-chave e a tabela de volumes."""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.utils import ImageReader, simpleSplit
    from reportlab.pdfbase.pdfmetrics import stringWidth
    from reportlab.pdfgen import canvas

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    c.setTitle(f"Autorização de faturamento - pedido {dados['pedido_venda']}")
    W, H = A4
    m = 26
    c.setLineWidth(1)
    c.rect(m, m, W - 2 * m, H - 2 * m)
    x0, x1 = m + 28, W - m - 28
    larg = x1 - x0

    def ajusta(txt, fonte, tmax, largura, tmin=9):
        t = tmax
        while t > tmin and stringWidth(txt, fonte, t) > largura:
            t -= 1
        return t

    def celula(x, ytopo, w, h, rotulo, valor, tmax, fonte="Helvetica-Bold", fundo=None, centro=False):
        if fundo is not None:
            c.setFillGray(fundo)
            c.rect(x, ytopo - h, w, h, fill=1, stroke=0)
            c.setFillGray(0)
        c.setLineWidth(1)
        c.rect(x, ytopo - h, w, h)
        c.setFillColorRGB(*NAVY)
        c.setFont("Helvetica-Bold", 8.5)
        c.drawString(x + 10, ytopo - 15, rotulo)
        c.setFillGray(0)
        t = ajusta(valor, fonte, tmax, w - 20)
        c.setFont(fonte, t)
        base = ytopo - h + (h - 22 - t) / 2 + 4
        if centro:
            c.drawCentredString(x + w / 2, base, valor)
        else:
            c.drawString(x + 10, base, valor)

    # topo: logo + emissão
    y = H - m - 24
    try:
        img = ImageReader(logo_path)
        iw, ih = img.getSize()
        lw = 140
        c.drawImage(img, x0, y - lw * ih / iw + 6, width=lw, height=lw * ih / iw, mask="auto")
    except Exception:
        c.setFont("Helvetica-Bold", 16)
        c.drawString(x0, y - 10, "4PIPE solutions")
    c.setFillGray(0.4)
    c.setFont("Helvetica", 8.5)
    c.drawRightString(x1, y - 4, "EMITIDO EM")
    c.setFillGray(0)
    c.setFont("Helvetica-Bold", 11)
    c.drawRightString(x1, y - 19, dados["emissao_txt"])
    y -= 62
    c.setLineWidth(1.4)
    c.setStrokeColorRGB(*NAVY)
    c.line(x0, y, x1, y)
    c.setStrokeGray(0)

    # título
    y -= 42
    c.setFillColorRGB(*NAVY)
    c.setFont("Helvetica-Bold", 22)
    c.drawCentredString(W / 2, y, "AUTORIZAÇÃO DE FATURAMENTO")
    c.setFont("Helvetica", 9.5)
    c.setFillGray(0.4)
    c.drawCentredString(W / 2, y - 16, "Liberação do pedido para a expedição")
    c.setFillGray(0)

    # grade de campos-chave
    y -= 40
    w1 = larg * 0.6
    celula(x0, y, w1, 82, "PEDIDO DE VENDA", str(dados["pedido_venda"]), 44, fundo=CINZA)
    celula(x0 + w1, y, larg - w1, 82, "FRETE", dados["frete"], 40, fundo=CINZA, centro=True)
    y -= 82
    celula(x0, y, larg, 74, "CLIENTE", dados["cliente"].upper(), 22)
    y -= 74
    w3 = larg * 0.62
    celula(x0, y, w3, 74, dados["coleta_rotulo"], dados["coleta_valor"], 26, fundo=CINZA)
    celula(x0 + w3, y, larg - w3, 74, "SITUAÇÃO", dados["situacao"], 17, centro=True)
    y -= 74

    # tabela de volumes
    y -= 34
    c.setFillColorRGB(*NAVY)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x0, y, "VOLUMES E PESOS")
    c.setFillGray(0)
    y -= 10
    cols = [x0, x0 + 46, x0 + 130, x0 + 282, x0 + 392, x1]  # qtd | tipo | medidas | peso unit. | peso total
    cab_h = 24
    c.setFillColorRGB(*NAVY)
    c.rect(x0, y - cab_h, larg, cab_h, fill=1, stroke=0)
    c.setFillGray(1)
    c.setFont("Helvetica-Bold", 8.5)
    c.drawString(cols[0] + 8, y - 15, "QTD")
    c.drawString(cols[1] + 8, y - 15, "TIPO")
    c.drawString(cols[2] + 8, y - 15, "MEDIDAS (mm)")
    c.drawRightString(cols[4] - 10, y - 15, "PESO BRUTO UNIT.")
    c.drawRightString(cols[5] - 10, y - 15, "PESO BRUTO TOTAL")
    c.setFillGray(0)
    y -= cab_h

    obs_h = 100
    base_min = m + 14 + obs_h + 28
    n = len(dados["volumes"])
    tot_h = 30 if dados["volumes"] else 0
    disp = y - base_min - tot_h
    lin_h = max(14, min(32, disp / max(n, 1)))
    fonte = 13 if lin_h >= 26 else (11 if lin_h >= 19 else 9)
    c.setLineWidth(0.5)
    for i, v in enumerate(dados["volumes"]):
        if i % 2:
            c.setFillGray(0.97)
            c.rect(x0, y - lin_h, larg, lin_h, fill=1, stroke=0)
            c.setFillGray(0)
        by = y - lin_h / 2 - fonte / 3
        c.setFont("Helvetica-Bold", fonte)
        c.drawString(cols[0] + 8, by, f"{v['quantidade']}X")
        c.setFont("Helvetica", fonte)
        c.drawString(cols[1] + 8, by, v["tipo"])
        c.drawString(cols[2] + 8, by, v["medidas_txt"])
        c.drawRightString(cols[4] - 10, by, v["pu_txt"])
        c.setFont("Helvetica-Bold", fonte)
        c.drawRightString(cols[5] - 10, by, v["pt_txt"])
        y -= lin_h
        c.setStrokeGray(0.75)
        c.line(x0, y, x1, y)
        c.setStrokeGray(0)
    if dados["volumes"]:
        c.setFillGray(CINZA)
        c.rect(x0, y - tot_h, larg, tot_h, fill=1, stroke=0)
        c.setFillGray(0)
        c.setLineWidth(1)
        c.rect(x0, y - tot_h, larg, tot_h)
        c.setFont("Helvetica-Bold", 12)
        c.drawString(cols[0] + 8, y - 20, "TOTAL")
        c.drawString(cols[1] + 8, y - 20, f"{dados['tot_volumes']} VOLUME" + ("S" if dados["tot_volumes"] != 1 else ""))
        c.drawRightString(cols[5] - 10, y - 20, dados["tot_pb_txt"])
        y -= tot_h
    else:
        c.setFont("Helvetica-Oblique", 10)
        c.setFillGray(0.45)
        c.drawString(x0 + 8, y - 20, "Nenhum volume informado.")
        c.setFillGray(0)

    # observações (rodapé)
    bx, bw = x0, larg
    by = m + 22
    c.setLineWidth(1)
    c.rect(bx, by, bw, obs_h)
    c.setFillColorRGB(*NAVY)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(bx + 10, by + obs_h - 15, "OBSERVAÇÕES")
    c.setFillGray(0)
    c.setFont("Helvetica", 10.5)
    ty = by + obs_h - 34
    linhas_obs = []
    for par in (dados["observacoes"] or "").splitlines():
        linhas_obs += simpleSplit(par, "Helvetica", 10.5, bw - 20) or [""]
    for linha in linhas_obs[:5]:
        c.drawString(bx + 10, ty, linha)
        ty -= 14
    c.showPage()
    c.save()
    return buf.getvalue()
