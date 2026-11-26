"""Autorização de faturamento do pedido (Bruno, 26/11/2026).

Documento impresso que leva a conclusão do pedido para a expedição (fica
exposto sobre o pedido). Layout CLÁSSICO (o do modelo da empresa), com os blocos espalhados por toda a
folha A4 e os campos-chave em letra grande (pedido, cliente, frete, coleta
liberada). Volumes com PESO BRUTO UNITÁRIO e PESO BRUTO TOTAL da linha
(quantidade × unitário), data/hora de emissão no topo e caixa de observações.
(Em 26/11/2026 o Bruno recusou o "cartaz" colorido e a grade em caixas; ficou o clássico.)
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


def gerar_pdf(dados, logo_path):
    """PDF A4 no desenho clássico, distribuído por toda a folha."""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.utils import ImageReader, simpleSplit
    from reportlab.pdfbase.pdfmetrics import stringWidth
    from reportlab.pdfgen import canvas

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    c.setTitle(f"Autorização de faturamento - pedido {dados['pedido_venda']}")
    W, H = A4
    m = 28
    c.setLineWidth(1)
    c.rect(m, m, W - 2 * m, H - 2 * m)
    x0, x1 = 66, W - 66
    larg = x1 - x0

    def ajusta(txt, fonte, tmax, largura, tmin=10):
        t = tmax
        while t > tmin and stringWidth(txt, fonte, t) > largura:
            t -= 1
        return t

    # ---- espaçamento: a folga da página é dividida em 3 intervalos iguais (g) ----
    # Contas da folha A4: topo+título ocupam até y=656; a caixa de observações começa em y=158.
    n = len(dados["volumes"])
    passo = 46 if n <= 10 else 32          # distância entre cliente / frete / situação / coleta
    lin_h = 30 if n <= 8 else max(10, min(30, (419 - 3 * passo - 36) / n))
    gap = max(12, min(80, (419 - 3 * passo - n * lin_h) / 3))
    h_obs = 130

    y = H - m - 30
    # topo: logo + emissão
    try:
        img = ImageReader(logo_path)
        iw, ih = img.getSize()
        lw = 165
        c.drawImage(img, x0, y - lw * ih / iw + 10, width=lw, height=lw * ih / iw, mask="auto")
    except Exception:
        c.setFont("Helvetica-Bold", 18)
        c.drawString(x0, y - 10, "4PIPE solutions")
    c.setFont("Helvetica-Bold", 11)
    c.drawRightString(x1, y - 8, dados["emissao_txt"])

    # título + nº do pedido (campo-chave)
    y -= 72
    c.setFont("Helvetica-Bold", 16)
    c.drawString(x0, y, "AUTORIZAÇÃO DE FATURAMENTO DO PEDIDO:")
    y -= 44
    c.setFont("Helvetica-Bold", 40)
    c.drawString(x0, y, str(dados["pedido_venda"]))
    c.setLineWidth(1.2)
    c.setStrokeColorRGB(*NAVY)
    c.line(x0, y - 12, x1, y - 12)
    c.setStrokeGray(0)

    # campos-chave, bem espaçados
    y -= 12 + gap + 14
    c.setFont("Helvetica-Bold", 12)
    c.drawString(x0, y, "CLIENTE:")
    cli = dados["cliente"]
    tam = ajusta(cli, "Helvetica-BoldOblique", 24, larg - 78)
    c.setFont("Helvetica-BoldOblique", tam)
    c.drawString(x0 + 78, y - 1, cli)
    y -= passo
    c.setFont("Helvetica-Bold", 12)
    c.drawString(x0, y, "FRETE:")
    c.setFont("Helvetica-BoldOblique", 24)
    c.drawString(x0 + 78, y - 1, dados["frete"])
    y -= passo
    c.setFont("Helvetica-Bold", 12)
    c.drawString(x0, y, "SITUAÇÃO:")
    c.setFont("Helvetica-Bold", 14)
    c.drawString(x0 + 78, y, dados["situacao"])
    y -= passo
    c.setFont("Helvetica-Bold", 12)
    c.drawString(x0, y, dados["coleta_rotulo"] + ":")
    r = stringWidth(dados["coleta_rotulo"] + ":", "Helvetica-Bold", 12)
    c.setFont("Helvetica-Bold", 20)
    c.drawString(x0 + r + 10, y - 2, dados["coleta_valor"])

    # tabela de volumes
    y -= 30 + gap
    cu, ct = x1 - 130, x1  # bordas direitas: peso unitário | peso total da linha
    c.setFont("Helvetica-Bold", 10)
    c.drawString(x0, y, "QUANTIDADE  /  TIPO  /  MEDIDAS (mm)")
    c.drawRightString(cu, y, "PESO BRUTO UNIT.")
    c.drawRightString(ct, y, "PESO BRUTO TOTAL")
    y -= 10
    c.setLineWidth(0.8)
    c.line(x0, y, x1, y)
    y -= lin_h * 0.72
    fonte = 15 if lin_h >= 26 else (12 if lin_h >= 19 else (10 if lin_h >= 14 else 8))
    for v in dados["volumes"]:
        c.setFont("Helvetica-Bold", fonte)
        c.drawString(x0 + 14, y, v["texto"])
        c.drawRightString(cu, y, v["pu_txt"])
        c.drawRightString(ct, y, v["pt_txt"])
        y -= lin_h
    if n:
        yl = y + lin_h * 0.72 - 6
        c.setLineWidth(0.8)
        c.line(x0, yl, x1, yl)
        c.setFont("Helvetica-Bold", 11)
        c.drawString(x0 + 14, yl - 18, f"TOTAL:  {dados['tot_volumes']} VOLUME" + ("S" if dados["tot_volumes"] != 1 else ""))
        c.setFont("Helvetica-Bold", 13)
        c.drawRightString(ct, yl - 19, dados["tot_pb_txt"])
    else:
        c.setFont("Helvetica-Oblique", 10)
        c.drawString(x0 + 14, y, "Nenhum volume informado.")

    # observações (fixa no pé da folha)
    bx, bw = m + 12, W - 2 * m - 24
    by = m + 12
    c.setLineWidth(1.5)
    c.rect(bx, by, bw, h_obs - 12)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(bx + 12, by + h_obs - 12 - 22, "OBSERVAÇÕES:")
    c.setFont("Helvetica", 10.5)
    ty = by + h_obs - 12 - 42
    linhas_obs = []
    for par in (dados["observacoes"] or "").splitlines():
        linhas_obs += simpleSplit(par, "Helvetica", 10.5, bw - 24) or [""]
    for linha in linhas_obs[:5]:
        c.drawString(bx + 12, ty, linha)
        ty -= 14
    c.showPage()
    c.save()
    return buf.getvalue()
