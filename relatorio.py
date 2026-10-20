"""Relatório mensal em .docx: resumo com IA (OpenAI) + montagem no padrão da empresa.

Fluxo:
  1. gerar_rascunho(): manda as atividades do mês e as referências para a OpenAI e recebe
     o texto em Markdown simples (#, ##, ###, ####, listas com "-", tabelas com "|").
  2. O texto é revisado na tela.
  3. montar_docx(): parte do .docx padrão, troca capa/rodapé e substitui o corpo de exemplo
     pelo texto, usando os estilos do próprio padrão. A IA nunca mexe na formatação.

Só usa a biblioteca padrão do Python.
"""
import base64
import io
import json
import os
import posixpath
import re
import urllib.error
import urllib.request
import zipfile
from xml.sax.saxutils import escape

MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
         "agosto", "setembro", "outubro", "novembro", "dezembro"]

OPENAI_URL = "https://api.openai.com/v1/responses"
MODELO_PADRAO = "gpt-5.5"
MAX_CHARS_REFERENCIA = 40000

# trechos de exemplo da capa/rodapé do padrão que são trocados pelos dados do relatório
CAPA_AUTOR = "Nombre Autor - Cargo"
CAPA_AREA = "Datos complementarios"
CAPA_TITULO = "Título principal"
CAPA_SUBTITULO = "Subtítulo texto complementario"
CAPA_REMOVER = ("ForFuture Sans 19 pt",)
RODAPE_TITULO = "Título de documento"
CABECALHO_REMOVER = ("LOGO CLIENTE",)

W_T = re.compile(r"<w:t(?: [^>]*)?>([^<]*)</w:t>")
P_INTERNO = re.compile(r"<w:p[ >](?:(?!<w:p[ >]).)*?</w:p>", re.S)


def nome_mes(mes):
    """'2026-10' -> 'Outubro de 2026'"""
    ano, m = mes.split("-")
    return f"{MESES[int(m) - 1].capitalize()} de {ano}"


def carregar_env(path):
    """Lê KEY=valor de um arquivo .env para os.environ (sem sobrescrever o que já existe)."""
    if not os.path.exists(path):
        return
    with open(path, encoding="utf-8") as f:
        for linha in f:
            linha = linha.strip()
            if not linha or linha.startswith("#") or "=" not in linha:
                continue
            k, v = linha.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


# ---------------------------------------------------------------- leitura de .docx

def texto_xml(xml):
    """Texto de um trecho de XML do Word, com quebras de linha <w:br/> viradas em \\n."""
    xml = re.sub(r"<w:br(?: [^>]*)?/>", "<w:t>\n</w:t>", xml)
    return "".join(W_T.findall(xml)).replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")


def elementos_topo(body):
    """Divide o conteúdo de <w:body> nos seus elementos de primeiro nível."""
    out, i, n = [], 0, len(body)
    while i < n:
        if body[i] != "<":
            i = body.index("<", i)
            continue
        tag = re.match(r"<([\w:]+)", body[i:i + 40]).group(1)
        fim_abertura = body.index(">", i)
        if body[fim_abertura - 1] == "/":
            j = fim_abertura + 1
        else:
            prof, j = 0, i
            for m in re.finditer(r"<(/?)%s(?=[ >/])[^>]*?(/?)>" % re.escape(tag), body[i:]):
                if m.group(1):
                    prof -= 1
                elif not m.group(2):
                    prof += 1
                if prof == 0:
                    j = i + m.end()
                    break
        out.append(body[i:j])
        i = j
    return out


def ler_docx(path):
    with zipfile.ZipFile(path) as z:
        doc = z.read("word/document.xml").decode("utf-8")
        estilos = z.read("word/styles.xml").decode("utf-8") if "word/styles.xml" in z.namelist() else ""
    return doc, estilos


def nomes_estilos(estilos_xml):
    """styleId -> nome interno (ex.: 'Ttulo1' -> 'heading 1')."""
    return {m.group(1): m.group(2).lower() for m in re.finditer(
        r'<w:style [^>]*w:styleId="([^"]+)"[^>]*>.*?<w:name w:val="([^"]+)"', estilos_xml, re.S)}


def docx_para_markdown(path):
    """Resumo textual de um .docx (títulos, parágrafos, listas, tabelas) para servir de referência."""
    doc, estilos = ler_docx(path)
    nomes = nomes_estilos(estilos)
    body = doc[doc.index("<w:body>") + 8: doc.rindex("</w:body>")]
    body = re.sub(r"<mc:AlternateContent>.*?</mc:AlternateContent>", "", body, flags=re.S)  # capa/caixas de texto
    linhas = []
    for el in elementos_topo(body):
        if el.startswith("<w:tbl"):
            for tr in re.findall(r"<w:tr[ >].*?</w:tr>", el, re.S):
                cels = [texto_xml(tc).strip().replace("\n", " ") for tc in re.findall(r"<w:tc>.*?</w:tc>", tr, re.S)]
                linhas.append("| " + " | ".join(cels) + " |")
            linhas.append("")
            continue
        if not el.startswith("<w:p"):
            continue
        texto = texto_xml(el).strip()
        if not texto:
            continue
        m = re.search(r'<w:pStyle w:val="([^"]+)"', el)
        nome = nomes.get(m.group(1), "") if m else ""
        if nome.startswith("toc") or nome.startswith("table of figures"):
            continue
        h = re.match(r"heading (\d)", nome)
        if h:
            linhas += ["", "#" * int(h.group(1)) + " " + texto]
        elif "list" in nome or "<w:numPr>" in el:
            linhas.append("- " + texto)
        else:
            linhas += ["", texto]
    primeiro_titulo = next((i for i, l in enumerate(linhas) if l.startswith("#")), 0)
    return re.sub(r"\n{3,}", "\n\n", "\n".join(linhas[primeiro_titulo:])).strip()


def capa_do_docx(path):
    """Textos da capa (autor, área, título, subtítulo), na ordem em que aparecem no documento."""
    doc, _ = ler_docx(path)
    inicio = doc[doc.index("<w:body>"): doc.find("TOC \\o") if "TOC \\o" in doc else len(doc)]
    textos = []
    for p in P_INTERNO.findall(inicio):
        t = texto_xml(p).strip()
        if t and t not in textos:
            textos.append(t)
    chaves = ("autor", "area", "titulo", "subtitulo")
    return {k: (textos[i] if i < len(textos) else "") for i, k in enumerate(chaves)}


# ---------------------------------------------------------------- OpenAI

INSTRUCOES = """Você escreve o relatório mensal de atividades que uma consultoria entrega ao cliente.

Regras:
- Escreva em português do Brasil, em tom formal e corporativo, no estilo dos relatórios de referência.
- Siga a estrutura das referências (seções, profundidade dos títulos, nível de detalhe, tipo de fechamento).
  Quando houver mais de uma referência, o que se repete entre elas é a estrutura esperada.
- Use SOMENTE os fatos das atividades do mês. Não invente números, nomes, datas, entregas ou resultados.
  Pode agrupar, reorganizar e contextualizar as atividades em texto corrido.
- Não cite horas trabalhadas, IDs de tickets nem as datas de cada lançamento, a menos que sejam relevantes ao texto.
- Não escreva capa, sumário/índice, nome do autor nem título geral do documento: comece direto pela primeira seção.
- Responda apenas com o conteúdo, neste Markdown simples (nada além disto):
  # Título de seção (nível 1)      ## Subseção      ### Nível 3      #### Nível 4
  parágrafos de texto separados por linha em branco
  - item de lista        (use "  - " para subitem)
  | Coluna | Coluna |   (tabela; a primeira linha é o cabeçalho)
  **negrito** só para destacar termos curtos.
  Não numere os títulos (a numeração é automática)."""


def descrever_atividades(atividades):
    linhas = []
    for a in atividades:
        d = "/".join(reversed(a["data"].split("-")))
        extra = ", ".join(x for x in (a.get("projeto"), a.get("categoria"), a.get("ticket")) if x)
        linhas.append(f"- {d}{' [' + extra + ']' if extra else ''}: {a['descricao'].strip()}")
    return "\n".join(linhas)


def chamar_openai(instrucoes, conteudo):
    chave = os.environ.get("OPENAI_API_KEY", "").strip()
    if not chave:
        raise RuntimeError("Configure OPENAI_API_KEY no arquivo .env.")
    corpo = {
        "model": os.environ.get("OPENAI_MODEL", "").strip() or MODELO_PADRAO,
        "instructions": instrucoes,
        "input": [{"role": "user", "content": conteudo}],
    }
    req = urllib.request.Request(
        OPENAI_URL, data=json.dumps(corpo).encode("utf-8"), method="POST",
        headers={"Authorization": f"Bearer {chave}", "Content-Type": "application/json"},
    )
    try:
        with urllib.request.urlopen(req, timeout=600) as r:
            data = json.load(r)
    except urllib.error.HTTPError as e:
        try:
            msg = json.load(e).get("error", {}).get("message", "")
        except ValueError:
            msg = ""
        raise RuntimeError(f"OpenAI respondeu {e.code}: {msg or e.reason}")
    except urllib.error.URLError as e:
        raise RuntimeError(f"Sem conexão com a OpenAI: {e.reason}")
    partes = [c.get("text", "") for item in data.get("output", []) if item.get("type") == "message"
              for c in item.get("content", []) if c.get("type") == "output_text"]
    texto = "".join(partes).strip()
    if not texto:
        raise RuntimeError("A OpenAI não devolveu texto.")
    # às vezes o modelo embrulha tudo em ```markdown ... ```
    return re.sub(r"^```\w*\n(.*?)\n```$", r"\1", texto, flags=re.S).strip()


def gerar_rascunho(atividades, mes, projeto, referencias, pedido=""):
    """referencias: lista de caminhos (.docx/.dotx/.pdf). Devolve o relatório em Markdown."""
    if not atividades:
        raise RuntimeError("Não há atividades registradas nesse mês.")
    conteudo = []
    for path in referencias:
        nome = os.path.basename(path)
        if path.lower().endswith(".pdf"):
            with open(path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode()
            conteudo.append({"type": "input_text", "text": f"Relatório de referência: {nome} (PDF a seguir)"})
            conteudo.append({"type": "input_file", "filename": nome,
                             "file_data": f"data:application/pdf;base64,{b64}"})
        else:
            texto = docx_para_markdown(path)[:MAX_CHARS_REFERENCIA]
            conteudo.append({"type": "input_text",
                             "text": f"Relatório de referência: {nome}\n<<<\n{texto}\n>>>"})
    alvo = f"Escreva o relatório de {nome_mes(mes)}"
    if projeto:
        alvo += f", somente sobre o projeto {projeto}"
    conteudo.append({"type": "input_text",
                     "text": f"{alvo}. Atividades registradas no mês:\n\n{descrever_atividades(atividades)}"})
    if pedido.strip():
        conteudo.append({"type": "input_text", "text": "Orientações adicionais de quem pediu o relatório "
                         "(siga, desde que não contrariem as regras):\n" + pedido.strip()})
    return chamar_openai(INSTRUCOES, conteudo)


# ---------------------------------------------------------------- Markdown -> blocos

def markdown_para_blocos(md):
    """Converte o Markdown simples em blocos: ('h', nível, texto) | ('p', texto) |
    ('li', nível, texto) | ('tabela', [linhas])."""
    blocos, paragrafo, tabela = [], [], []

    def fecha():
        if paragrafo:
            blocos.append(("p", " ".join(paragrafo)))
            paragrafo.clear()
        if tabela:
            blocos.append(("tabela", [r[:] for r in tabela]))
            tabela.clear()

    for linha in md.replace("\r\n", "\n").split("\n"):
        s = linha.strip()
        if s.startswith("|"):
            if paragrafo:
                blocos.append(("p", " ".join(paragrafo)))
                paragrafo.clear()
            cels = [c.strip() for c in s.strip("|").split("|")]
            if not all(re.fullmatch(r":?-{2,}:?", c) for c in cels if c):  # linha |---|---|
                tabela.append(cels)
            continue
        if tabela:
            fecha()
        m = re.match(r"(#{1,4})\s+(.*)", s)
        li = re.match(r"( *)[-*•]\s+(.*)", linha.rstrip())
        num = re.match(r"( *)\d+[.)]\s+(.*)", linha.rstrip())
        if not s:
            fecha()
        elif m:
            fecha()
            blocos.append(("h", len(m.group(1)), re.sub(r"^[\d.]+\s+", "", m.group(2))))
        elif li or num:
            fecha()
            g = li or num
            blocos.append(("li", 2 if len(g.group(1)) >= 2 else 1, g.group(2)))
        else:
            paragrafo.append(s)
    fecha()
    return blocos


# ---------------------------------------------------------------- montagem do .docx

def runs(texto, rpr=""):
    """Texto com **negrito** e *itálico* -> <w:r>...</w:r>."""
    out = []
    for parte in re.split(r"(\*\*[^*]+\*\*|\*[^*]+\*)", texto):
        if not parte:
            continue
        extra = ""
        if parte.startswith("**") and parte.endswith("**") and len(parte) > 4:
            parte, extra = parte[2:-2], "<w:b/>"
        elif parte.startswith("*") and parte.endswith("*") and len(parte) > 2:
            parte, extra = parte[1:-1], "<w:i/>"
        parte = parte.replace("`", "")
        props = f"<w:rPr>{rpr}{extra}</w:rPr>" if (rpr or extra) else ""
        out.append(f'<w:r>{props}<w:t xml:space="preserve">{escape(parte)}</w:t></w:r>')
    return "".join(out)


def paragrafo(texto, estilo=None, ppr_extra="", rpr=""):
    estilo_xml = f'<w:pStyle w:val="{estilo}"/>' if estilo else ""
    ppr = f"<w:pPr>{estilo_xml}{ppr_extra}</w:pPr>" if (estilo_xml or ppr_extra) else ""
    return f"<w:p>{ppr}{runs(texto, rpr)}</w:p>"


def tabela_xml(linhas, est):
    """Tabela no visual do padrão: bordas finas, cabeçalho em negrito, 1ª coluna destacada."""
    ncol = max(len(r) for r in linhas)
    linhas = [r + [""] * (ncol - len(r)) for r in linhas]
    borda = '<w:{0} w:val="single" w:sz="4" w:space="0" w:color="004254" w:themeColor="text1"/>'
    largura = 5000 // ncol
    trs = []
    for i, r in enumerate(linhas):
        tcs = []
        for j, c in enumerate(r):
            bordas = borda.format("top") + (borda.format("bottom") if i == 0 else "")
            shd = ('<w:shd w:val="clear" w:color="auto" w:fill="E3E2DA" w:themeFill="background2"/>'
                   if i > 0 and j == 0 else "")
            negrito = "<w:b/>" if i == 0 or j == 0 else ""
            tcs.append(
                f'<w:tc><w:tcPr><w:tcW w:w="{largura}" w:type="pct"/><w:tcBorders>{bordas}</w:tcBorders>'
                f'{shd}<w:vAlign w:val="center"/></w:tcPr>'
                + paragrafo(c, est["tabela"], '<w:jc w:val="left"/>', negrito) + "</w:tc>")
        cab = "<w:trPr><w:tblHeader/></w:trPr>" if i == 0 else ""
        trs.append(f"<w:tr>{cab}{''.join(tcs)}</w:tr>")
    grid = "".join(f'<w:gridCol w:w="{9800 // ncol}"/>' for _ in range(ncol))
    return (f'<w:tbl><w:tblPr><w:tblStyle w:val="{est["grade"]}"/><w:tblW w:w="5000" w:type="pct"/>'
            '<w:tblBorders><w:insideV w:val="single" w:sz="4" w:space="0" w:color="E3E2DA" w:themeColor="background2"/></w:tblBorders>'
            '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
            f"</w:tblPr><w:tblGrid>{grid}</w:tblGrid>{''.join(trs)}</w:tbl>"
            "<w:p/>")


def estilos_do_padrao(estilos_xml):
    """Descobre os styleIds do padrão pelos nomes internos (independe do idioma do Word)."""
    por_nome = {v: k for k, v in nomes_estilos(estilos_xml).items()}
    return {
        "h": {n: por_nome.get(f"heading {n}") for n in range(1, 5)},
        "li": {1: por_nome.get("list bullet"), 2: por_nome.get("list bullet 2") or por_nome.get("list bullet")},
        "toc": {n: por_nome.get(f"toc {n}") for n in range(1, 4)},
        "tabela": por_nome.get("texto tablas") or por_nome.get("table text"),
        "grade": por_nome.get("table grid") or "TableGrid",
    }


def corpo_xml(blocos, est):
    xml, toc, cont = [], [], [0, 0, 0, 0]
    for b in blocos:
        if b[0] == "h":
            n = b[1]
            cont[n - 1] += 1
            cont[n:] = [0] * (4 - n)
            if n <= 3:
                toc.append((n, ".".join(str(c) for c in cont[:n]), b[2]))
            xml.append(paragrafo(b[2], est["h"][n]))
        elif b[0] == "p":
            xml.append(paragrafo(b[1]))
        elif b[0] == "li":
            xml.append(paragrafo(b[2], est["li"][b[1]]))
        elif b[0] == "tabela":
            xml.append(tabela_xml(b[1], est))
    return "".join(xml), toc


def indice_xml(toc, est):
    """Campo de sumário do Word já preenchido com os títulos (sem números de página).
    O Word atualiza ao abrir (updateFields); no LibreOffice: Ferramentas > Atualizar."""
    inicio = ('<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>'
              '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>'
              '<w:r><w:fldChar w:fldCharType="separate"/></w:r>')
    fim = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    if not toc:
        toc = [(1, "", "Atualize o índice (botão direito > Atualizar campo)")]
    ps = []
    for i, (n, numero, titulo) in enumerate(toc):
        estilo = est["toc"].get(n) or est["toc"].get(1)
        ppr = f'<w:pPr><w:pStyle w:val="{estilo}"/></w:pPr>' if estilo else ""
        texto = f"{numero}  {titulo}".strip()
        ps.append(f'<w:p>{ppr}{inicio if i == 0 else ""}'
                  f'<w:r><w:t xml:space="preserve">{escape(texto)}</w:t></w:r>'
                  f'{fim if i == len(toc) - 1 else ""}</w:p>')
    return "".join(ps)


def trocar_paragrafo(p, novo):
    """Mantém a formatação do parágrafo e do 1º trecho; troca o texto (\\n vira quebra de linha).
    Num título com duas formatações (1ª linha colorida), as linhas seguintes usam a do último trecho."""
    ppr = re.search(r"<w:pPr>.*?</w:pPr>", p, re.S)
    rprs = []
    for r in re.findall(r"<w:r(?: [^>]*)?>.*?</w:r>", p, re.S):
        if "<w:t" in r:
            m = re.search(r"<w:rPr>.*?</w:rPr>", r, re.S)
            rprs.append(m.group(0) if m else "")
    primeiro = rprs[0] if rprs else ""
    ultimo = rprs[-1] if rprs else ""
    linhas = novo.split("\n")
    rs = []
    for i, linha in enumerate(linhas):
        rpr = primeiro if i == 0 else ultimo
        if i:
            rs.append(f"<w:r>{rpr}<w:br/></w:r>")
        rs.append(f'<w:r>{rpr}<w:t xml:space="preserve">{escape(linha)}</w:t></w:r>')
    abertura = re.match(r"<w:p(?: [^>]*)?>", p).group(0)
    return f"{abertura}{ppr.group(0) if ppr else ''}{''.join(rs)}</w:p>"


def substituir_textos(xml, trocas, remover=()):
    """Em cada parágrafo (inclusive dentro de caixas de texto) cujo texto começa com uma das
    chaves, troca o texto. Parágrafos que começam com algo de `remover` ficam vazios."""
    def f(m):
        p = m.group(0)
        t = texto_xml(p).strip()
        for chave, novo in trocas.items():
            if t.startswith(chave):
                return trocar_paragrafo(p, novo)
        if any(t.startswith(r) for r in remover):
            return trocar_paragrafo(p, "")
        return p
    return P_INTERNO.sub(f, xml)


def ativar_atualizacao_campos(settings):
    if "<w:updateFields" in settings:
        return settings
    tag = '<w:updateFields w:val="true"/>'
    # segue a ordem do schema: updateFields vem antes destes elementos
    for antes in ("<w:hdrShapeDefaults", "<w:footnotePr", "<w:endnotePr", "<w:compat", "<w:docVars", "<w:rsids"):
        if antes in settings:
            i = settings.index(antes)
            return settings[:i] + tag + settings[i:]
    return settings.replace("</w:settings>", tag + "</w:settings>")


def podar_partes(arquivos):
    """Remove imagens/gráficos/planilhas do corpo de exemplo que deixaram de ser usados."""
    doc = arquivos["word/document.xml"].decode("utf-8")
    rels_path = "word/_rels/document.xml.rels"
    rels = arquivos[rels_path].decode("utf-8")
    descartaveis = ("/image", "/chart", "/package", "/oleObject", "/diagramData", "/hyperlink")

    def manter(m):
        rid, tipo = m.group(1), m.group(2)
        if tipo.endswith(descartaveis) and f'"{rid}"' not in doc:
            return ""
        return m.group(0)
    rels = re.sub(r'<Relationship Id="([^"]+)" Type="([^"]+)"[^>]*/>', manter, rels)
    arquivos[rels_path] = rels.encode("utf-8")

    # partes alcançáveis a partir de _rels/.rels
    alcancadas, fila = set(), ["", "word/document.xml"]
    while fila:
        parte = fila.pop()
        pasta, nome = posixpath.split(parte)
        rp = posixpath.join(pasta, "_rels", nome + ".rels") if parte else "_rels/.rels"
        if rp not in arquivos:
            continue
        for alvo, modo in re.findall(r'Target="([^"]+)"(?: TargetMode="([^"]+)")?', arquivos[rp].decode("utf-8")):
            if modo == "External":
                continue
            caminho = alvo.lstrip("/") if alvo.startswith("/") else posixpath.normpath(posixpath.join(pasta, alvo))
            if caminho not in alcancadas:
                alcancadas.add(caminho)
                fila.append(caminho)
    for nome in list(arquivos):
        if nome.startswith(("word/media/", "word/charts/", "word/embeddings/")) and nome not in alcancadas \
                and "/_rels/" not in nome:
            del arquivos[nome]
    for nome in list(arquivos):  # rels de partes removidas
        m = re.match(r"(.*)/_rels/(.+)\.rels$", nome)
        if m and m.group(1).startswith("word/charts") and f"{m.group(1)}/{m.group(2)}" not in arquivos:
            del arquivos[nome]
    tipos = arquivos["[Content_Types].xml"].decode("utf-8")
    tipos = re.sub(r'<Override PartName="/([^"]+)"[^>]*/>',
                   lambda m: m.group(0) if m.group(1) in arquivos else "", tipos)
    arquivos["[Content_Types].xml"] = tipos.encode("utf-8")


def montar_docx(padrao_path, markdown, capa):
    """capa: {titulo, subtitulo, autor, area, versao}. Devolve os bytes do .docx."""
    with zipfile.ZipFile(padrao_path) as z:
        arquivos = {i.filename: z.read(i.filename) for i in z.infolist()}
        ordem = [i.filename for i in z.infolist()]
    doc = arquivos["word/document.xml"].decode("utf-8")
    est = estilos_do_padrao(arquivos.get("word/styles.xml", b"").decode("utf-8"))
    if not est["h"][1]:
        raise RuntimeError("O padrão não tem o estilo 'Título 1' (heading 1).")

    a, b = doc.index("<w:body>") + 8, doc.rindex("</w:body>")
    els = elementos_topo(doc[a:b])
    h1 = f'<w:pStyle w:val="{est["h"][1]}"/>'
    i_toc = next((i for i, e in enumerate(els) if "TOC \\o" in e), None)
    i_corpo = next((i for i, e in enumerate(els) if h1 in e and (i_toc is None or i > i_toc)), None)
    if i_corpo is None:
        raise RuntimeError("Não encontrei no padrão onde começa o corpo do documento (primeiro Título 1).")
    # fim do corpo de exemplo: a quebra de seção que antecede a contracapa (ou o fim do documento)
    i_fim = next((i for i in range(i_corpo, len(els)) if "<w:sectPr" in els[i] and els[i].startswith("<w:p")),
                 len(els) - 1 if els[-1].startswith("<w:sectPr") else len(els))

    blocos = markdown_para_blocos(markdown)
    if not any(b[0] == "h" and b[1] == 1 for b in blocos):
        raise RuntimeError("O texto precisa de pelo menos um título de seção (linha começando com '# ').")
    corpo, toc = corpo_xml(blocos, est)
    inicio = "".join(els[:i_toc if i_toc is not None else i_corpo])
    indice = indice_xml(toc, est) if i_toc is not None else ""
    # mantém o último parágrafo antes da quebra de seção sem conteúdo do exemplo
    resto = "".join(els[i_fim:])

    trocas = {CAPA_AUTOR: capa.get("autor", ""), CAPA_AREA: capa.get("area", ""),
              CAPA_TITULO: capa.get("titulo", ""), CAPA_SUBTITULO: capa.get("subtitulo", "")}
    inicio = substituir_textos(inicio, trocas, CAPA_REMOVER)
    doc = doc[:a] + inicio + indice + corpo + resto + doc[b:]
    arquivos["word/document.xml"] = doc.encode("utf-8")

    rodape = f"{capa.get('subtitulo', '')} – Versão {capa.get('versao') or '1.0'}"
    for nome in list(arquivos):
        if re.fullmatch(r"word/(header|footer)\d*\.xml", nome):
            x = arquivos[nome].decode("utf-8")
            x2 = substituir_textos(x, {RODAPE_TITULO: rodape}, CABECALHO_REMOVER)
            if x2 != x:
                arquivos[nome] = x2.encode("utf-8")
    if "word/settings.xml" in arquivos:
        arquivos["word/settings.xml"] = ativar_atualizacao_campos(
            arquivos["word/settings.xml"].decode("utf-8")).encode("utf-8")

    podar_partes(arquivos)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for nome in ordem:
            if nome in arquivos:
                z.writestr(nome, arquivos[nome])
    return buf.getvalue()
