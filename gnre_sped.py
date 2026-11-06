# -*- coding: utf-8 -*-
"""
Lancamento automatico de GNRE no SPED Fiscal (EFD ICMS/IPI).

Para cada PDF "GNRE <numero CT-e>.pdf" da pasta escolhida, acrescenta no D100
(CT-e de saida) correspondente:
    |D195|GNRE|GUIA DE RECOLHIMENTO|
    |D197|MT10001002|NUMERO DA GUIA <n controle>||||<valor principal>||
(credita so o VALOR PRINCIPAL da guia; juros e multa nao entram)
Tipos aceitos: GNRE, DARE-SP (n = campo 09 Numero do DARE, valor = 09 Valor da Receita),
DAE-MG (n = N Documento, valor = linha "ICMS SERVICO DE TRANSPORTE") e
GR-PR (n = Numero SEFA so com os digitos, valor = Valor da Receita).
DAR-1 de MT e ignorado (so aparece no relatorio).
Garante o cadastro 0460 (GNRE), recalcula os totalizadores (0990, D990, bloco 9)
e salva um NOVO arquivo TXT (o original nao e alterado) + relatorio CSV.

Depois e so importar o TXT novo no PVA (Escrituracao Fiscal > Importar).
"""
import csv
import logging
import os
import re
import sys
from collections import OrderedDict, defaultdict

logging.getLogger("pypdf").setLevel(logging.ERROR)   # sem avisos de PDF truncado na tela

COD_OBS = "GNRE"
TXT_OBS_PADRAO = "GUIA DE RECOLHIMENTO"
COD_AJUSTE = "MT10001002"
COD_AJUSTE_ANTIGOS = ("MT10000007",)   # lancamentos antigos de GNRE que o robo corrige

FILHOS_D100 = {"D101", "D110", "D120", "D130", "D140", "D150", "D160", "D161",
               "D162", "D170", "D180", "D190", "D195", "D197"}
ORDEM_BLOCO0_APOS_0460 = ("0500", "0600", "0990")


# ----------------------------------------------------------------- PDF
def ler_pdf(caminho):
    """Texto do PDF. Se o pypdf nao conseguir (PDF truncado/corrompido), tenta pelo pdfium."""
    try:
        from pypdf import PdfReader
        texto = ""
        for pag in PdfReader(caminho).pages:
            texto += (pag.extract_text() or "") + "\n"
        return texto
    except Exception:
        import pypdfium2 as pdfium
        pdf = pdfium.PdfDocument(caminho)
        try:
            return "\n".join(pdf[i].get_textpage().get_text_range() for i in range(len(pdf)))
        finally:
            pdf.close()


def valor_float(s):
    return float(s.replace(".", "").replace(",", "."))


class EhDAR(Exception):
    """O PDF e um DAR-1 (MT), nao uma GNRE: so vai para o relatorio."""


# ----------------------------------------------------------------- tipo de documento
def classificar(texto):
    """GNRE, DARE-SP, DAE-MG, GR-PR, DAR-1 ou None (nao deu para saber)."""
    t = _sem_acento(texto)
    if "dare-sp" in t or "dare sp" in t or "numero do dare" in t or "dare avulso" in t:
        return "DARE-SP"
    if "guia de recolhimento do estado do parana" in t or "numero sefa" in t:
        return "GR-PR"
    if "guia nacional" in t or "gnre" in t:
        return "GNRE"
    if "minas gerais" in t or re.search(r"arrecadacao estadual\s*-?\s*dae", t) or re.search(r"\bdae\b", t):
        return "DAE-MG"
    if re.search(r"\bdar\s*-?\s*(modelo\s*)?[1il]\b", t) or ("mato grosso" in t and "documento de arrecadacao" in t):
        return "DAR-1"
    return None


def extrair_guia(caminho, forcar_ocr=False):
    """Retorna (numero_guia, valor_total, valor_principal, cte_no_pdf, via_ocr, tipo)."""
    texto = "" if forcar_ocr else ler_pdf(caminho)
    tipo = classificar(texto) if texto.strip() else None
    if tipo == "DAR-1":
        raise EhDAR()
    rico = len(re.sub(r"\s", "", texto)) >= 50

    # 1) tenta pelo texto do PDF (rapido e exato)
    if rico and tipo in ("GNRE", None):
        r = extrair_guia_texto(texto)
        if r[0] and r[2] is not None:
            return r + (False, "GNRE")
    if rico and tipo == "DAE-MG":
        r = _dae_texto(texto)
        if r[0] and r[2] is not None:
            return r + (False, tipo)
    if rico and tipo == "DARE-SP":
        r = _dare_texto(texto)
        if r[0] and r[2] is not None:
            return r + (False, tipo)
    if tipo == "GR-PR":
        r = _grpr_texto(texto)
        return r + (False, tipo)

    # 2) PDF escaneado / impresso como imagem: OCR do Windows
    topo = 0.7 if tipo == "DARE-SP" else 0.5
    linhas = _ocr_linhas(caminho, topo)
    texto_ocr = "\n".join(l["t"] for l in linhas)
    if tipo is None:
        tipo = classificar(texto_ocr) or "GNRE"
        if tipo == "DAR-1":
            raise EhDAR()
        if tipo == "DARE-SP" and topo < 0.7:      # o valor da receita fica mais abaixo
            linhas = _ocr_linhas(caminho, 0.7)
            texto_ocr = "\n".join(l["t"] for l in linhas)
    if tipo == "DARE-SP":
        r = _dare_ocr(linhas, texto_ocr)
    elif tipo == "DAE-MG":
        r = _dae_ocr(linhas, texto_ocr)
    else:
        r = _gnre_ocr(linhas, texto_ocr)
    cte_txt = _cte_do_texto(texto) if texto else None   # o texto do PDF, quando existe, e mais confiavel
    if cte_txt:
        r = r[:3] + (cte_txt,)
    return r + (True, tipo)


# ----------------------------------------------------------------- OCR (PDF escaneado)
MONEY = re.compile(r"(\d{1,3}(?:\.\d{3})*,\d{2})(?!\d)")


def _sem_acento(s):
    import unicodedata
    return "".join(c for c in unicodedata.normalize("NFKD", s) if unicodedata.category(c) != "Mn").lower()


def _digitos_ocr(s):
    """Corrige trocas comuns do OCR em campos numericos."""
    return s.translate(str.maketrans({"O": "0", "o": "0", "D": "0", "Q": "0", "l": "1", "I": "1",
                                      "|": "1", "i": "1", "S": "5", "s": "5", "B": "8", "Z": "2"}))


def _renderizar(caminho, topo):
    """Imagem da parte de cima da 1a pagina (`topo` = fracao da altura)."""
    import pypdfium2 as pdfium
    pdf = pdfium.PdfDocument(caminho)
    pag = pdf[0]
    largura, altura = pag.get_size()
    escala = min(4.0, 2500 / largura, 2600 / (altura * topo))
    img = pag.render(scale=escala, crop=(0, altura * (1 - topo), 0, 0)).to_pil().convert("RGBA")
    pdf.close()
    return img


def _ocr_linhas(caminho, topo=0.5):
    """OCR nativo do Windows. Devolve linhas com texto, caixa e palavras (cada uma com caixa)."""
    img = _renderizar(caminho, topo)
    if os.environ.get("GNRE_OCR_TESSERACT"):        # so para testes fora do Windows
        return _ocr_tesseract(img)
    import asyncio
    from winrt.windows.media.ocr import OcrEngine
    from winrt.windows.storage.streams import DataWriter
    from winrt.windows.graphics.imaging import SoftwareBitmap, BitmapPixelFormat

    motor = OcrEngine.try_create_from_user_profile_languages()
    if motor is None:
        raise RuntimeError("OCR do Windows indisponivel neste computador")
    w = DataWriter()
    w.write_bytes(img.tobytes())
    sb = SoftwareBitmap.create_copy_from_buffer(w.detach_buffer(), BitmapPixelFormat.RGBA8, img.width, img.height)

    async def _rodar():
        return await motor.recognize_async(sb)
    res = asyncio.run(_rodar())

    linhas = []
    for ln in res.lines:
        ws = [{"t": p.text, "x0": p.bounding_rect.x, "y0": p.bounding_rect.y,
               "x1": p.bounding_rect.x + p.bounding_rect.width, "y1": p.bounding_rect.y + p.bounding_rect.height}
              for p in ln.words]
        if ws:
            linhas.append(_linha(ln.text, ws))
    linhas.sort(key=lambda l: (l["y0"], l["x0"]))
    return linhas


def _linha(texto, ws):
    return {"t": texto, "w": ws, "x0": min(p["x0"] for p in ws), "y0": min(p["y0"] for p in ws),
            "x1": max(p["x1"] for p in ws), "y1": max(p["y1"] for p in ws)}


def _ocr_tesseract(img):
    import pytesseract
    d = pytesseract.image_to_data(img.convert("RGB"), output_type=pytesseract.Output.DICT)
    grupos = OrderedDict()
    for i, t in enumerate(d["text"]):
        if not t.strip():
            continue
        k = (d["block_num"][i], d["par_num"][i], d["line_num"][i])
        grupos.setdefault(k, []).append({"t": t, "x0": d["left"][i], "y0": d["top"][i],
                                         "x1": d["left"][i] + d["width"][i], "y1": d["top"][i] + d["height"][i]})
    linhas = [_linha(" ".join(p["t"] for p in ws), ws) for ws in grupos.values()]
    linhas.sort(key=lambda l: (l["y0"], l["x0"]))
    return linhas


def _caixa_rotulo(linhas, rotulo, excluir=()):
    """Caixa (x0, y0, x1, y1) das palavras que formam o rotulo, na 1a ocorrencia (1a via)."""
    achados = []
    for l in linhas:
        tl = _sem_acento(l["t"])
        if rotulo not in tl or any(e in tl for e in excluir):
            continue
        ws, norm = l["w"], [_sem_acento(p["t"]) for p in l["w"]]
        melhor = None
        for i in range(len(ws)):
            acc = ""
            for j in range(i, len(ws)):
                acc = (acc + " " + norm[j]).strip()
                if rotulo in acc or rotulo.replace(" ", "") in acc.replace(" ", ""):
                    if melhor is None or j - i < melhor[1] - melhor[0]:
                        melhor = (i, j)
                    break
        i, j = melhor if melhor else (0, len(ws) - 1)
        sel = ws[i:j + 1]
        achados.append((min(p["x0"] for p in sel), min(p["y0"] for p in sel),
                        max(p["x1"] for p in sel), max(p["y1"] for p in sel)))
    return min(achados, key=lambda b: (b[1], b[0])) if achados else None


def _candidatos(linhas, achar):
    """Cada palavra (ou ate 3 palavras seguidas) que contem um valor reconhecido por `achar`."""
    for l in linhas:
        ws = l["w"]
        for i in range(len(ws)):
            for n in (1, 2, 3):
                if i + n > len(ws):
                    break
                v = achar("".join(p["t"] for p in ws[i:i + n]))
                if v is not None:
                    yield ws[i], v
                    break


def _valor_coluna(linhas, rotulo, achar, excluir=(), max_linhas=10, direita=8):
    """Valor na mesma caixa do rotulo: logo abaixo (ou ao lado), na mesma coluna."""
    caixa = _caixa_rotulo(linhas, rotulo, excluir)
    if not caixa:
        return None
    x0, y0, x1, y1 = caixa
    h = max(y1 - y0, 8)
    cands = []
    for w, v in _candidatos(linhas, achar):
        if w["y0"] < y0 - 0.4 * h or w["y0"] > y1 + max_linhas * h:
            continue
        if w["x1"] < x0 - 2 * h or w["x0"] > x1 + direita * h:
            continue
        dy = max(0, w["y0"] - y0)
        dx = 0 if w["x0"] <= x1 + h else w["x0"] - x1
        cands.append((dy + 0.3 * dx, v))
    return min(cands, key=lambda c: c[0])[1] if cands else None


def _valor_mesma_linha(linhas, rotulo, achar, excluir=()):
    """Valor na mesma altura do rotulo, a direita dele (ex.: tabela de receitas do DAE)."""
    caixa = _caixa_rotulo(linhas, rotulo, excluir)
    if not caixa:
        return None
    x0, y0, x1, y1 = caixa
    h = max(y1 - y0, 8)
    cands = []
    for w, v in _candidatos(linhas, achar):
        cy = (w["y0"] + w["y1"]) / 2
        if y0 - 0.6 * h <= cy <= y1 + 0.6 * h and w["x0"] >= x1 - h:
            cands.append((w["x0"] - x1, v))
    return min(cands, key=lambda c: c[0])[1] if cands else None


def _achar_dinheiro(t):
    m = MONEY.search(_digitos_ocr(t.replace(" ", "")).replace("R$", ""))
    return valor_float(m.group(1)) if m else None


def _achar_controle(t):
    m = re.search(r"(?<!\d)(\d{16})(?!\d)", _digitos_ocr(t.replace(" ", "")))
    return m.group(1) if m else None


def _achar_numero_doc(t):
    """Numero de documento com 10 a 16 digitos (DARE, DAE)."""
    m = re.search(r"(?<!\d)(\d{10,16})(?!\d)", _digitos_ocr(t.replace(" ", "")))
    return m.group(1) if m else None


def _achar_inteiro(t):
    m = re.search(r"(\d{3,})", _digitos_ocr(t.replace(" ", "")))
    return int(m.group(1)) if m else None


def _cte_do_texto(texto):
    m = re.search(r"CT-?E\s*[:N°º.]*\s*([\dOoIl]{2,})", texto, re.I)
    return int(_digitos_ocr(m.group(1))) if m else None


def _valor_abaixo(linhas, rotulo, padrao):
    """Acha o valor (regex `padrao`) na caixa do rotulo: mesma linha ou logo abaixo/a direita (GNRE)."""
    alvo = [l for l in linhas if rotulo in _sem_acento(l["t"])]
    if not alvo:
        return None
    lab = min(alvo, key=lambda l: l["y0"])          # 1a ocorrencia = 1a via
    resto = lab["t"][_sem_acento(lab["t"]).find(rotulo) + len(rotulo):]
    m = padrao(resto)
    if m:
        return m
    h = max(lab["y1"] - lab["y0"], 8)
    cands = []
    for l in linhas:
        if l is lab or l["y0"] < lab["y0"] - h * 0.5 or l["y0"] > lab["y1"] + h * 3:
            continue
        if l["x1"] < lab["x0"]:
            continue
        m = padrao(l["t"])
        if m:
            cands.append((abs(l["y0"] - lab["y0"]) + 0.2 * abs(l["x1"] - lab["x1"]), m))
    return min(cands)[1] if cands else None


def _gnre_ocr(linhas, texto):
    controle = _valor_abaixo(linhas, "controle", _achar_controle)
    if not controle:
        for l in linhas:
            controle = _achar_controle(l["t"])
            if controle:
                break
    principal = _valor_abaixo(linhas, "principal", _achar_dinheiro)
    total = _valor_abaixo(linhas, "recolher", _achar_dinheiro)
    cte = _cte_do_texto(texto) or _valor_abaixo(linhas, "documento de origem", _achar_inteiro)
    return controle, total, principal, cte


def extrair_guia_ocr(caminho):
    """GNRE escaneada (usado tambem pelo diagnostico --testar-ocr)."""
    linhas = _ocr_linhas(caminho, 0.5)
    texto = "\n".join(l["t"] for l in linhas)
    if classificar(texto) == "DAR-1":
        raise EhDAR()
    return _gnre_ocr(linhas, texto)


# --- DARE-SP: numero = campo 09 "Numero do DARE"; valor = campo 09 "Valor da Receita" (sem juros/multa)
def _dare_ocr(linhas, texto):
    numero = _valor_coluna(linhas, "numero do dare", _achar_numero_doc, max_linhas=6)
    # conferencia: campo 18 "N do Documento Detalhe" = numero do DARE + "-0001"
    detalhe = None
    for l in linhas:
        for w in l["w"]:
            m = re.fullmatch(r"(\d{12,16})-\d{4}", _digitos_ocr(w["t"]))
            if m and not detalhe:
                detalhe = m.group(1)
    if not numero:
        numero = detalhe
    elif detalhe and detalhe != numero:
        if len(detalhe) != len(numero):
            numero = detalhe               # campo 09 lido pela metade: usa o do detalhe
        else:
            numero = None                  # leitura duvidosa: deixa para lancar manualmente
    principal = _valor_coluna(linhas, "valor da receita", _achar_dinheiro, max_linhas=12)
    total = _valor_coluna(linhas, "valor total", _achar_dinheiro, max_linhas=6, direita=25)
    return numero, total, principal, _cte_do_texto(texto)


def _dare_texto(texto):
    m1 = re.search(r"N[uú]mero do DARE\D{0,40}(\d{12,16})", texto, re.I | re.S)
    m2 = re.search(r"Valor da Receita\D{0,40}([\d.]+,\d{2})", texto, re.I | re.S)
    m3 = re.search(r"Valor Total\D{0,40}([\d.]+,\d{2})", texto, re.I | re.S)
    return (m1.group(1) if m1 else None, valor_float(m3.group(1)) if m3 else None,
            valor_float(m2.group(1)) if m2 else None, _cte_do_texto(texto))


# --- DAE-MG: numero = "N Documento"; valor = linha "315-2 ICMS SERVICO DE TRANSPORTE" (sem juros/multa)
def _dae_ocr(linhas, texto):
    numero = _valor_coluna(linhas, "documento", _achar_numero_doc, excluir=("origem", "arrecada"), max_linhas=4)
    principal = _valor_mesma_linha(linhas, "servico de transporte", _achar_dinheiro)
    total = _valor_mesma_linha(linhas, "total", _achar_dinheiro, excluir=("autentica",))
    return numero, total, principal, _cte_do_texto(texto)


def _dae_texto(texto):
    m1 = re.search(r"N\W{0,2}\s*Documento\D{0,20}(\d{10,16})", texto, re.I | re.S)
    m2 = re.search(r"SERVI[CÇ]O DE TRANSPORTE\s*([\d.]+,\d{2})", texto, re.I)
    m3 = re.search(r"TOTAL\s*([\d.]+,\d{2})", texto)
    return (m1.group(1) if m1 else None, valor_float(m3.group(1)) if m3 else None,
            valor_float(m2.group(1)) if m2 else None, _cte_do_texto(texto))


# --- GR-PR (Guia de Recolhimento do Estado do Parana): numero = "Numero SEFA" (so os digitos);
#     valor = "Valor da Receita (R$)" (sem multa, juros e acrescimo); CT-e = campo "Documento NF: 000000731"
def _grpr_texto(texto):
    m1 = re.search(r"N[uú]mero SEFA:?\s*(\d[\d.]{14,24}\d)", texto, re.I)
    numero = re.sub(r"\D", "", m1.group(1)) if m1 else None
    m2 = re.search(r"Valor da Receita \(R\$\)\s*([\d.]+,\d{2})", texto, re.I)
    m3 = re.search(r"Valor Total \(R\$\)\s*([\d.]+,\d{2})", texto, re.I)
    m4 = re.search(r"Documento\s*NF?\s*:?\s*0*(\d+)", texto, re.I)
    cte = int(m4.group(1)) if m4 else _cte_do_texto(texto)
    return (numero, valor_float(m3.group(1)) if m3 else None,
            valor_float(m2.group(1)) if m2 else None, cte)


def extrair_guia_texto(texto):
    # o PDF tem 1a e 2a via; usa so a primeira
    marcador = texto.find("Guia Nacional de Recolhimento", 40)
    primeira = texto[:marcador] if marcador > 0 else texto

    controle = None
    m = re.search(r"Controle\D{0,40}(\d{16})", primeira, re.S)
    if m:
        controle = m.group(1)
    else:
        nums = re.findall(r"(?<!\d)(\d{16})(?!\d)", primeira)
        controle = nums[0] if nums else None

    valores = [valor_float(v) for v in re.findall(r"R\$\s*([\d.]+,\d{2})", primeira)]
    total = None
    m = re.search(r"Total a Recolher\D{0,20}([\d.]+,\d{2})", primeira, re.S)
    if m:
        total = valor_float(m.group(1))
    elif valores:
        total = max(valores)
    m = re.search(r"Valor Principal\D{0,20}([\d.]+,\d{2})", primeira, re.S)
    if m:
        principal = valor_float(m.group(1))
    else:
        principal = valores[0] if valores else None

    # CT-e: o "N Documento de Origem" (000000528) vale mais que o texto livre "ICMS CTE 528",
    # que as vezes vem digitado errado na emissao da guia
    cte = None
    m = re.search(r"Documento de Origem\s*(\d{3,})", primeira)
    origens = [m.group(1)] if m else re.findall(r"(?m)^\s*(00\d{7})\s*$", primeira)
    if origens and int(origens[0]) > 0:
        cte = int(origens[0])
    else:
        m = re.search(r"CTE\s*(\d+)", primeira, re.I)
        if m:
            cte = int(m.group(1))
    return controle, total, principal, cte


def fmt_valor(v):
    return f"{v:.2f}".replace(".", ",")


# ----------------------------------------------------------------- SPED
def reg(linha):
    p = linha.split("|")
    return p[1] if len(p) > 2 else ""


def recalcular_totais(linhas):
    """Refaz 0990/C990/D990/... e todo o bloco 9."""
    corpo = [l for l in linhas if not reg(l).startswith("9")]

    # X990 de cada bloco
    por_bloco = defaultdict(int)
    for l in corpo:
        por_bloco[reg(l)[0]] += 1
    corpo = [f"|{reg(l)[0]}990|{por_bloco[reg(l)[0]]}|" if reg(l).endswith("990") else l
             for l in corpo]

    contagem = OrderedDict()
    for l in corpo:
        contagem[reg(l)] = contagem.get(reg(l), 0) + 1
    tipos = list(contagem.items()) + [("9001", 1), ("9900", None), ("9990", 1), ("9999", 1)]
    qtd_9900 = len(tipos)
    bloco9 = ["|9001|0|"]
    for r, n in tipos:
        bloco9.append(f"|9900|{r}|{qtd_9900 if r == '9900' else n}|")
    bloco9.append(f"|9990|{len(bloco9) + 2}|")
    total = len(corpo) + len(bloco9) + 1
    bloco9.append(f"|9999|{total}|")
    return corpo + bloco9


def garantir_0460(linhas):
    """Devolve (linhas, texto_da_obs, criado?)."""
    for l in linhas:
        p = l.split("|")
        if reg(l) == "0460" and p[2].strip().upper() == COD_OBS:
            return linhas, p[3].strip() or TXT_OBS_PADRAO, False
    nova = f"|0460|{COD_OBS}|{TXT_OBS_PADRAO}|"
    pos = None
    for i, l in enumerate(linhas):
        if reg(l) == "0460":
            pos = i + 1
    if pos is None:
        pos = next(i for i, l in enumerate(linhas) if reg(l) in ORDEM_BLOCO0_APOS_0460)
    return linhas[:pos] + [nova] + linhas[pos:], TXT_OBS_PADRAO, True


def processar(caminho_txt, pasta_pdf):
    with open(caminho_txt, "r", encoding="latin-1", newline="") as f:
        bruto = f.read()
    quebra = "\r\n" if "\r\n" in bruto else "\n"
    linhas = [l for l in bruto.split(quebra) if l.strip()]
    i9999 = next(i for i, l in enumerate(linhas) if reg(l) == "9999")
    if i9999 < len(linhas) - 1:
        print("Aviso: assinatura digital no final do arquivo foi removida (o arquivo sera alterado).")
    linhas = linhas[:i9999 + 1]

    # --- GNREs da pasta
    relatorio = []
    guias = defaultdict(list)  # cte -> [(controle, valor, arquivo)]
    for nome in sorted(os.listdir(pasta_pdf)):
        m = re.match(r"GNRE\D*(\d+)", nome, re.I)
        if not (m and nome.lower().endswith(".pdf")):
            continue
        cte = int(m.group(1))
        try:
            controle, total, principal, cte_pdf, via_ocr, tipo = extrair_guia(os.path.join(pasta_pdf, nome))
        except EhDAR:
            relatorio.append([nome, cte, "", "", "DAR-1 (nao e GNRE) - fora do SPED, nada a fazer"])
            continue
        except Exception as e:
            relatorio.append([nome, cte, "", "", "ERRO lendo PDF: " + str(e)])
            continue
        obs = [tipo] if tipo != "GNRE" else []
        if via_ocr:
            obs.append("LIDO POR OCR - conferir n da guia e valor")
        if not controle or principal is None:
            relatorio.append([nome, cte, controle or "", fmt_valor(principal) if principal is not None else "",
                              f"ERRO ({tipo}): nao achei n da guia/valor - lancar manualmente"])
            continue
        if cte_pdf and cte_pdf != cte:
            obs.append(f"ATENCAO: PDF cita CTE {cte_pdf}")
        if total is not None and abs(total - principal) > 0.005:
            obs.append(f"(guia com juros/multa: total {fmt_valor(total)}, creditado so o principal)")
        guias[cte].append((controle, principal, nome, "; ".join(obs)))

    linhas, txt_obs, criou_0460 = garantir_0460(linhas)

    # --- insere D195/D197 em cada D100
    saida, i, usados, corrigidos = [], 0, set(), []
    while i < len(linhas):
        l = linhas[i]
        saida.append(l)
        i += 1
        p = l.split("|")
        if reg(l) != "D100" or p[2] != "1" or p[3] != "0":
            continue
        try:
            num = int(p[9])
        except ValueError:
            continue
        grupo = []
        while i < len(linhas) and reg(linhas[i]) in FILHOS_D100:
            grupo.append(linhas[i])
            i += 1
        # corrige D197 de GNRE ja lancados com o codigo antigo (MT10000007 -> MT10001002)
        obs_atual = None
        for k, g in enumerate(grupo):
            pg = g.split("|")
            if reg(g) == "D195":
                obs_atual = pg[2].strip().upper()
            elif reg(g) == "D197" and obs_atual == COD_OBS and pg[2].strip() in COD_AJUSTE_ANTIGOS:
                pg[2] = COD_AJUSTE
                grupo[k] = "|".join(pg)
                corrigidos.append([f"(SPED) CT-e {num}", num, pg[3].replace("NUMERO DA GUIA", "").strip(),
                                   pg[7] if len(pg) > 7 else "", "CORRIGIDO: codigo de ajuste MT10000007 -> " + COD_AJUSTE])
        saida.extend(grupo)
        if num not in guias:
            continue
        usados.add(num)
        ja_tem = any(reg(g) == "D195" and g.split("|")[2].strip().upper() == COD_OBS for g in grupo)
        if ja_tem:
            for c, v, nome, obs in guias[num]:
                relatorio.append([nome, num, c, fmt_valor(v), "PULADO: CT-e ja tem D195 GNRE"])
            continue
        saida.append(f"|D195|{COD_OBS}|{txt_obs}|")
        for c, v, nome, obs in guias[num]:
            saida.append("|".join(["", "D197", COD_AJUSTE, f"NUMERO DA GUIA {c}", "", "", "", fmt_valor(v), "", ""]))
            relatorio.append([nome, num, c, fmt_valor(v), ("OK " + obs).strip()])

    for num in sorted(set(guias) - usados):
        for c, v, nome, obs in guias[num]:
            relatorio.append([nome, num, c, fmt_valor(v), "NAO LANCADO: CT-e nao encontrado no SPED"])

    relatorio.extend(corrigidos)
    saida = recalcular_totais(saida)

    base, _ = os.path.splitext(caminho_txt)
    novo_txt = base + " - COM GNRE.txt"
    for tentativa in range(1, 20):
        try:
            with open(novo_txt, "w", encoding="latin-1", newline="") as f:
                f.write(quebra.join(saida) + quebra)
            break
        except PermissionError:   # arquivo anterior aberto em outro programa
            novo_txt = base + f" - COM GNRE ({tentativa + 1}).txt"
    linhas_rel = sorted(relatorio, key=lambda r: (r[1] if isinstance(r[1], int) else 0))
    novo_csv = base + " - relatorio GNRE.csv"
    for tentativa in range(1, 20):
        try:
            with open(novo_csv, "w", encoding="utf-8-sig", newline="") as f:
                w = csv.writer(f, delimiter=";")
                w.writerow(["Arquivo", "CT-e", "N controle", "Valor ICMS", "Situacao"])
                # n da guia como texto (="...") para o Excel nao cortar zeros nem mostrar 2,9E+15
                w.writerows([[r[0], r[1], f'="{r[2]}"' if r[2] else "", r[3], r[4]] for r in linhas_rel])
            break
        except PermissionError:   # relatorio anterior aberto no Excel
            novo_csv = base + f" - relatorio GNRE ({tentativa + 1}).csv"

    ok = sum(1 for r in relatorio if r[4].startswith("OK"))
    dar = sum(1 for r in relatorio if r[4].startswith("DAR-1"))
    pend = len(relatorio) - ok - dar - len(corrigidos)
    por_tipo = defaultdict(int)
    for r in relatorio:
        if r[4].startswith("OK"):
            por_tipo[next((t for t in ("DARE-SP", "DAE-MG", "GR-PR") if t in r[4]), "GNRE")] += 1
    detalhe = ", ".join(f"{k}: {v}" for k, v in sorted(por_tipo.items()))
    print(f"\nGuias lancadas: {ok} ({detalhe})   DAR-1 ignorados: {dar}   Pendentes (ver relatorio): {pend}")
    if corrigidos:
        print(f"Codigo de ajuste corrigido para {COD_AJUSTE} em {len(corrigidos)} lancamento(s) que ja existiam no SPED.")
    if criou_0460:
        print("Cadastro 0460 GNRE incluido no arquivo.")
    print(f"Arquivo novo:  {novo_txt}")
    print(f"Relatorio:     {novo_csv}")
    return novo_txt, novo_csv


def escolher_pelo_windows():
    import tkinter as tk
    from tkinter import filedialog, messagebox
    raiz = tk.Tk()
    raiz.withdraw()
    desk = os.path.join(os.path.expanduser("~"), "Desktop")
    txt = filedialog.askopenfilename(title="Escolha o arquivo TXT do SPED", initialdir=desk,
                                     filetypes=[("SPED TXT", "*.txt"), ("Todos", "*.*")])
    if not txt:
        sys.exit(0)
    pasta = filedialog.askdirectory(title="Escolha a pasta com os PDFs das GNRE", initialdir=desk)
    if not pasta:
        sys.exit(0)
    novo, rel = processar(txt, pasta)
    messagebox.showinfo("GNRE no SPED", f"Pronto!\n\nArquivo para importar no PVA:\n{novo}\n\nConfira o relatorio:\n{rel}")


def testar_ocr(pasta_pdf):
    """Diagnostico: compara OCR x texto nos PDFs que tem texto."""
    ok = ruim = 0
    for nome in sorted(os.listdir(pasta_pdf)):
        if not re.match(r"GNRE\D*\d+.*\.pdf$", nome, re.I):
            continue
        cam = os.path.join(pasta_pdf, nome)
        texto = ler_pdf(cam)
        try:
            if len(re.sub(r"\s", "", texto)) < 50:
                print(nome, "ESCANEADO ->", extrair_guia_ocr(cam))
                continue
            a = extrair_guia_texto(texto)
            b = extrair_guia_ocr(cam)
        except Exception as e:
            print(nome, "ERRO", e)
            continue
        if a[0] == b[0] and a[2] == b[2]:
            ok += 1
        else:
            ruim += 1
            print(nome, "DIFERENTE texto:", a, " ocr:", b)
    print(f"OCR igual ao texto: {ok}  diferente: {ruim}")


def diagnostico(arquivos):
    """Grava em diag.txt (ao lado do script) o tipo, o resultado e as linhas do OCR de cada PDF."""
    saida = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "diag.txt"), "w", encoding="utf-8")
    for cam in arquivos:
        print("=" * 30, os.path.basename(cam), file=saida)
        try:
            print("RESULTADO:", extrair_guia(cam), file=saida)
        except EhDAR:
            print("RESULTADO: DAR-1", file=saida)
        except Exception as e:
            print("RESULTADO: ERRO", repr(e), file=saida)
        try:
            for l in _ocr_linhas(cam, 0.7):
                print(f"  {int(l['y0']):5d} {int(l['x0']):5d} {int(l['x1']):5d} | {l['t']}", file=saida)
        except Exception as e:
            print("OCR ERRO", repr(e), file=saida)
    saida.close()


if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == "--diag":
        diagnostico(sys.argv[2:])
    elif len(sys.argv) == 3 and sys.argv[1] == "--testar-ocr":
        testar_ocr(sys.argv[2])
    elif len(sys.argv) == 3:
        processar(sys.argv[1], sys.argv[2])
    else:
        escolher_pelo_windows()
