"""Leitura do PDF "Emissão da Confirmação do Pedido" (relatório SIGA/1076 — Pré Nota do ERP) para
alimentar a tela "Novo pedido" (pedido do Bruno, 19/11/2026: "o PCP realiza a digitação manual do
pedido... queria que o PCP incluísse o PDF no novo pedido e ele alimentasse automaticamente").

Módulo independente (não importa o app): recebe os bytes do PDF e devolve um dict com os campos do
formulário + uma lista de avisos para o PCP conferir. NUNCA grava nada — quem decide salvar é o PCP,
depois de revisar a tela.

Campos lidos (o PDF não traz Prioridade, Tipo de pedido nem Status do pedido — esses continuam com o PCP):
  pedido_venda              <- "PEDIDO N. 001076"                         -> "1076"
  data_inclusao_pedido      <- "EMISSAO: 18/11/2026"
  data_cliente              <- menor "Entrega:" entre os itens (prazo pedido pelo cliente)
  cliente / cnpj / cidade / estado <- caixa do cliente (código/loja, nome, endereço, CEP cidade UF, CNPJ)
  frete                     <- "FRETE...: (FOB)"
  vendedor                  <- "VENDEDOR.: 000017 - Nome"
  go_pedido_compra_cliente  <- "PEDIDO(S) DO CLIENTE: 13029"
  go_valor_pedido_operacao  <- "T O T A I S" (valor total do pedido)
  itens[]                   <- cada bloco "IT: Produto: ... Descricao: ... Quant.: ... Prc Unitario: ... Entrega: ..."
"""

import io
import re
from datetime import datetime

_RE_NUM_BR = re.compile(r"^-?\d{1,3}(?:\.\d{3})*(?:,\d+)?$|^-?\d+(?:,\d+)?$")
_RE_CNPJ = re.compile(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}")
_RE_CPF = re.compile(r"\d{3}\.\d{3}\.\d{3}-\d{2}")


_RX_NUM_PEDIDO = r"PEDIDO\s+N\s*[.º°o:]*\s*(\d+)"
_RE_NUM_PEDIDO = re.compile(r"PEDIDO\s+N\s*[.º°o:]*\s*\d+", re.I)
_RE_TITULO = re.compile(r"CONFIRMA[CÇ][AÃ]O\s+DO\s+PEDIDO", re.I)


class PdfPedidoInvalido(Exception):
    """O arquivo não é (ou não parece) o relatório de Confirmação do Pedido."""


def _num_br(txt):
    """'3.675,32' -> 3675.32 ; '2,00' -> 2.0 ; None se não for número."""
    if txt is None:
        return None
    t = str(txt).strip()
    if not t or not _RE_NUM_BR.match(t):
        return None
    try:
        return float(t.replace(".", "").replace(",", "."))
    except ValueError:
        return None


def _brl(v):
    return "R$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _cidade_fmt(txt):
    """'RIO DE JANEIRO' -> 'Rio de Janeiro' (o ERP manda tudo em maiúsculas e sem acento)."""
    miudas = {"de", "da", "do", "das", "dos", "e"}
    pal = txt.strip().lower().split()
    return " ".join(p if (i and p in miudas) else p.capitalize() for i, p in enumerate(pal))


def _data_iso(txt):
    try:
        return datetime.strptime(txt.strip(), "%d/%m/%Y").date().isoformat()
    except (ValueError, AttributeError):
        return None


def _agrupar_linhas(palavras, tol=3):
    """[(top, [palavras da linha ordenadas por x])]"""
    linhas = []
    for w in sorted(palavras, key=lambda w: (round(w["top"]), w["x0"])):
        if linhas and abs(linhas[-1][0] - w["top"]) <= tol:
            linhas[-1][1].append(w)
        else:
            linhas.append([w["top"], [w]])
    return [(t, sorted(ws, key=lambda w: w["x0"])) for t, ws in linhas]


def _bloco_cliente(pagina):
    """Lê a caixa do cliente pela POSIÇÃO das palavras (o texto corrido mistura as 3 caixas do cabeçalho).
    Âncora = o código/loja do cliente ("000230/01"). Devolve (cliente, cnpj, cidade, uf) — o que não achar vem None."""
    palavras = pagina.extract_words()
    ancora = next((w for w in palavras if re.fullmatch(r"\d{4,7}/\d{2}", w["text"])), None)
    if ancora is None:
        return None, None, None, None
    x_ini = ancora["x0"] - 2
    direita = next((w["x0"] for w in palavras if w["text"].upper().startswith("CONFIRMACAO") and w["x0"] > x_ini), None)
    x_fim = (direita - 4) if direita else x_ini + 380
    caixa = [w for w in palavras if x_ini <= w["x0"] < x_fim and ancora["top"] - 3 <= w["top"] <= ancora["top"] + 55]
    linhas = [" ".join(w["text"] for w in ws) for _, ws in _agrupar_linhas(caixa)]
    if not linhas:
        return None, None, None, None

    cliente = re.sub(r"^\d{4,7}/\d{2}\s+", "", linhas[0]).strip() or None
    cnpj = cidade = uf = None
    for lin in linhas[1:]:
        if cnpj is None:
            m = _RE_CNPJ.search(lin) or _RE_CPF.search(lin)
            if m:
                cnpj = m.group(0)
                continue
        if cidade is None:
            m = re.match(r"^\d{5}-?\d{3}\s+(.+?)\s+([A-Z]{2})$", lin.strip())
            if m:
                cidade, uf = (_cidade_fmt(m.group(1)) if m.group(1).isupper() else m.group(1).strip()), m.group(2)
    return cliente, cnpj, cidade, uf


def _itens(texto):
    """Quebra o texto das páginas do pedido em itens, de "Produto:" a "Produto:" (o último termina em "T O T A I S")."""
    marcas = [m.start() for m in re.finditer(r"Produto:\s*\S+", texto)]
    itens = []
    for i, ini in enumerate(marcas):
        fim = marcas[i + 1] if i + 1 < len(marcas) else len(texto)
        bloco = texto[ini:fim]
        corte = re.search(r"T\s*O\s*T\s*A\s*I\s*S", bloco)
        if corte:
            bloco = bloco[:corte.start()]
        itens.append(bloco)
    return itens


def _campo(rx, texto, flags=0):
    m = re.search(rx, texto, flags)
    return m.group(1).strip() if m else None


def ler_pedidos_pdf(dados):
    """bytes do PDF -> LISTA de dicts, um por pedido (ver docstring do módulo). O relatório do ERP aceita
    "do pedido... até o pedido", então 1 arquivo pode trazer vários pedidos (páginas seguidas com o mesmo
    "PEDIDO N." são continuação do mesmo pedido). Levanta PdfPedidoInvalido se não for o relatório esperado."""
    try:
        import pdfplumber
    except ImportError as e:  # pragma: no cover - só se a dependência faltar no servidor
        raise PdfPedidoInvalido("Leitura de PDF indisponível no servidor (falta a biblioteca pdfplumber).") from e

    try:
        pdf = pdfplumber.open(io.BytesIO(dados))
    except Exception as e:
        raise PdfPedidoInvalido("Não consegui abrir o arquivo como PDF.") from e

    with pdf:
        grupos = []  # [[(pagina, texto), ...], ...] — 1 grupo por pedido
        ultimo_num = None
        textos = []
        for pg in pdf.pages:
            txt = pg.extract_text() or ""
            textos.append(txt)
            if not (_RE_TITULO.search(txt) and _RE_NUM_PEDIDO.search(txt)):
                continue
            num = _campo(_RX_NUM_PEDIDO, txt, re.I)
            if grupos and num == ultimo_num:
                grupos[-1].append((pg, txt))
            else:
                grupos.append([(pg, txt)])
            ultimo_num = num
        if not grupos:
            raise PdfPedidoInvalido(_motivo_nao_reconhecido(textos, (pdf.metadata or {}).get("Producer", "")))
        return [_ler_um_pedido(g) for g in grupos]


def _motivo_nao_reconhecido(textos, produtor=""):
    """Mensagem específica pro motivo de o PDF não ter sido reconhecido (o PCP vê na tela): sem texto
    (imagem/escaneado), só a página de parâmetros do ERP (nenhum pedido no intervalo) ou outro relatório."""
    gere = "Gere o relatório \"Confirmação do Pedido\" direto pelo ERP (SIGA/1076) e anexe de novo."
    total = sum(len(t.strip()) for t in textos)
    if total < 30:
        if "print to pdf" in str(produtor).lower():
            return ("Esse PDF foi criado com a impressora \"Microsoft Print to PDF\", que transforma o texto em desenho "
                    "— por isso o sistema não consegue ler. Gere o PDF pelo próprio ERP (exportar/salvar em PDF, "
                    "como nos pedidos anteriores) ou use \"Salvar como PDF\" em vez dessa impressora.")
        return ("Esse PDF não tem texto que o sistema consiga ler (parece imagem, escaneado ou impresso como figura). "
                + gere)
    todo = "\n".join(textos)
    if re.search(r"Pergunta\s*0?1\s*:\s*Do\s+Pedido", todo, re.I):
        m = re.search(r"Do\s+Pedido\s*\?\s*(\d+)", todo, re.I)
        ate = re.search(r"Ate\s+o\s+Pedido\s*\?\s*(\d+)", todo, re.I)
        faixa = f" ({m.group(1)} até {ate.group(1)})" if m and ate else ""
        return (f"O ERP gerou só a página de parâmetros{faixa}, sem nenhum pedido dentro. "
                "Confira se o número do pedido existe/está liberado e gere o relatório de novo.")
    primeira = next((l.strip() for l in todo.splitlines() if l.strip()), "")
    dica = f" (o arquivo começa com: \"{primeira[:70]}\")" if primeira else ""
    return f"Esse PDF não parece o relatório \"Confirmação do Pedido\" do ERP (SIGA/1076){dica}. " + gere


def ler_pedido_pdf(dados):
    """Compatibilidade: só o 1º pedido do arquivo."""
    return ler_pedidos_pdf(dados)[0]


def _ler_um_pedido(paginas):
    pg0, txt0 = paginas[0]
    texto = "\n".join(t for _, t in paginas)
    cliente, cnpj, cidade, uf = _bloco_cliente(pg0)

    avisos = []
    num_ped = _campo(_RX_NUM_PEDIDO, txt0, re.I)
    pedido_venda = (num_ped.lstrip("0") or "0") if num_ped else None
    emissao = _data_iso(_campo(r"EMISSAO:\s*(\d{2}/\d{2}/\d{4})", txt0, re.I))

    ped_cliente = _campo(r"PEDIDO\(S\)\s*DO\s*CLIENTE:[ \t]*([^\n]*)", txt0, re.I)
    vendedor = _campo(r"VENDEDOR\.*:\s*\d*\s*-?\s*(.+?)\s+COMISSAO", txt0, re.I)
    frete_txt = (_campo(r"FRETE\.*:\s*\(?\s*([A-Za-z ]+?)\s*\)?\s+SEGURO", txt0, re.I) or "").upper()
    frete = frete_txt if frete_txt in ("CIF", "FOB") else None
    if frete_txt and frete is None:
        avisos.append(f"Frete \"{frete_txt}\" do PDF não é CIF nem FOB — escolha o frete na tela.")

    itens = []
    datas_entrega = []
    for bloco in _itens(texto):
        desc = _campo(r"Descricao:\s*(.+?)\s+TES:", bloco, re.S) or _campo(r"Descricao:\s*([^\n]+)", bloco)
        desc = re.sub(r"\s+", " ", desc).strip() if desc else None
        qtd = _num_br(_campo(r"Quant\.:\s*([\d.,]+)", bloco))
        preco = _num_br(_campo(r"Prc\s*Unitario:\s*([\d.,]+)", bloco))
        total = _num_br(_campo(r"Vl\.Tot\.C/IPI:\s*([\d.,]+)", bloco))
        entrega = _data_iso(_campo(r"Entrega:\s*(\d{2}/\d{2}/\d{4})", bloco))
        codigo = _campo(r"Produto:\s*(\S+)", bloco)
        if entrega:
            datas_entrega.append(entrega)
        if not desc:
            avisos.append(f"Item {codigo or len(itens) + 1}: descrição não encontrada.")
            continue
        if qtd is not None and preco is not None and total is not None and abs(qtd * preco - total) > 0.05:
            avisos.append(
                f"Item \"{desc[:40]}\": total do PDF ({_brl(total)}) difere de quantidade × preço "
                f"({_brl(qtd * preco)}) — provavelmente IPI/desconto. Foi usado o preço unitário do PDF."
            )
        itens.append({"descricao": desc, "quantidade": qtd, "custo_unitario": preco, "entrega": entrega, "produto_codigo": codigo})

    tot = re.search(r"T\s*O\s*T\s*A\s*I\s*S\s+([\d.,]+)\s+([\d.,]+)", texto)
    valor_total = _num_br(tot.group(1)) if tot else None
    if valor_total is None:
        somas = [i["quantidade"] * i["custo_unitario"] for i in itens if i["quantidade"] is not None and i["custo_unitario"] is not None]
        valor_total = round(sum(somas), 2) if somas else None

    if not itens:
        avisos.append("Nenhum item foi encontrado no PDF — preencha os produtos à mão.")
    if not cliente:
        avisos.append("Não consegui ler o nome do cliente — preencha à mão.")
    if len(set(datas_entrega)) > 1:
        avisos.append("Os itens têm datas de entrega diferentes; em \"Data do cliente\" foi usada a mais próxima. "
                      "Confira (as datas por item aparecem na lista abaixo).")
    if not vendedor:
        avisos.append("Vendedor não encontrado no PDF.")
    if not emissao:
        avisos.append("Data de emissão não encontrada no PDF.")

    return {
        "pedido_venda": pedido_venda,
        "data_inclusao_pedido": emissao,
        "data_cliente": min(datas_entrega) if datas_entrega else None,
        "cliente": cliente,
        "cnpj": cnpj,
        "cidade": cidade,
        "estado": uf,
        "pais": "Brasil",
        "frete": frete,
        "vendedor": vendedor,
        "go_pedido_compra_cliente": ped_cliente or None,
        "go_valor_pedido_operacao": valor_total,
        "itens": itens,
        "avisos": avisos,
    }
