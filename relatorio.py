"""Relatório mensal em .docx: resumo com IA (OpenAI, Anthropic ou compatível) + montagem no padrão da empresa.

Fluxo:
  1. gerar_rascunho(): manda as atividades do mês e as referências para a IA escolhida e recebe
     o texto em Markdown simples (#, ##, ###, ####, listas com "-", tabelas com "|").
  2. O texto é revisado na tela.
  3. montar_docx(): parte do .docx padrão, troca capa/rodapé e substitui o corpo de exemplo
     pelo texto, usando os estilos do próprio padrão. A IA nunca mexe na formatação.

Só usa a biblioteca padrão do Python.
"""
import base64
from datetime import datetime
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

# provedores de IA aceitos. "compativel" = qualquer API no formato /chat/completions da OpenAI
# (Azure OpenAI, OpenRouter, Groq, Mistral, Gemini, DeepSeek, Ollama, LM Studio...)
PROVEDORES = {
    "openai": {"nome": "OpenAI", "modelo": "gpt-5.5", "base_url": "https://api.openai.com/v1", "pdf": True},
    "anthropic": {"nome": "Anthropic (Claude)", "modelo": "claude-sonnet-5-5", "base_url": "https://api.anthropic.com/v1", "pdf": True},
    "compativel": {"nome": "Compatível com OpenAI", "modelo": "", "base_url": "http://localhost:11434/v1", "pdf": False},
}
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
- Atividades marcadas com (ENTREGA) são entregas concluídas no mês: dê destaque a elas
  (por exemplo, numa seção ou tabela de entregas, se as referências tiverem algo parecido).
- Não cite horas trabalhadas, IDs de tickets nem as datas de cada lançamento, a menos que sejam relevantes ao texto.
- Não escreva capa, sumário/índice, nome do autor nem título geral do documento: comece direto pela primeira seção.
- Responda apenas com o conteúdo, em Markdown (nada além disto):
  # Título de seção (nível 1)      ## Subseção      ### Nível 3      #### Nível 4
  parágrafos separados por linha em branco
  - item de lista  (recuo de 2 espaços para subitem)      1. item numerado
  | Coluna | Coluna |  (tabela; a primeira linha é o cabeçalho)
  **negrito** e *itálico* só para destacar termos curtos; > citação, se fizer sentido.
  Não numere os títulos (a numeração é automática)."""


def descrever_atividades(atividades):
    linhas = []
    for a in atividades:
        d = "/".join(reversed(a["data"].split("-")))
        extra = ", ".join(x for x in (a.get("projeto"), a.get("categoria"), a.get("ticket")) if x)
        entrega = " (ENTREGA)" if a.get("entregavel") else ""
        linhas.append(f"- {d}{' [' + extra + ']' if extra else ''}{entrega}: {a['descricao'].strip()}")
    return "\n".join(linhas)


def _post(url, corpo, headers, timeout=600):
    req = urllib.request.Request(url, data=json.dumps(corpo).encode("utf-8"), method="POST",
                                 headers={"Content-Type": "application/json", **headers})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return json.load(r)
    except urllib.error.HTTPError as e:
        try:
            erro = json.load(e).get("error", {})
            msg = erro.get("message", "") if isinstance(erro, dict) else str(erro)
        except ValueError:
            msg = ""
        raise RuntimeError(f"A IA respondeu {e.code}: {msg or e.reason}")
    except urllib.error.URLError as e:
        raise RuntimeError(f"Sem conexão com a IA ({url}): {e.reason}")
    except TimeoutError:
        raise RuntimeError("A IA demorou demais para responder.")


def chamar_llm(cfg, instrucoes, partes, max_tokens=16000):
    """cfg: {provedor, modelo, base_url, chave}. partes: [{tipo: texto, texto} | {tipo: pdf, nome, b64}].
    Devolve o texto da resposta."""
    provedor = cfg.get("provedor") or "openai"
    info = PROVEDORES.get(provedor)
    if not info:
        raise RuntimeError(f"Provedor de IA desconhecido: {provedor}")
    chave, modelo = (cfg.get("chave") or "").strip(), (cfg.get("modelo") or info["modelo"]).strip()
    base = (cfg.get("base_url") or info["base_url"]).rstrip("/")
    if not modelo:
        raise RuntimeError("Informe o modelo da IA em Configurações > Inteligência artificial.")
    if not chave and provedor != "compativel":
        raise RuntimeError("Configure a chave da IA em Configurações > Inteligência artificial.")

    if provedor == "openai":
        conteudo = [{"type": "input_text", "text": p["texto"]} if p["tipo"] == "texto" else
                    {"type": "input_file", "filename": p["nome"], "file_data": f"data:application/pdf;base64,{p['b64']}"}
                    for p in partes]
        data = _post(f"{base}/responses", {"model": modelo, "instructions": instrucoes,
                                           "input": [{"role": "user", "content": conteudo}]},
                     {"Authorization": f"Bearer {chave}"})
        texto = "".join(c.get("text", "") for item in data.get("output", []) if item.get("type") == "message"
                        for c in item.get("content", []) if c.get("type") == "output_text")
    elif provedor == "anthropic":
        conteudo = [{"type": "text", "text": p["texto"]} if p["tipo"] == "texto" else
                    {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": p["b64"]}}
                    for p in partes]
        data = _post(f"{base}/messages", {"model": modelo, "max_tokens": max_tokens, "system": instrucoes,
                                          "messages": [{"role": "user", "content": conteudo}]},
                     {"x-api-key": chave, "anthropic-version": "2023-06-01"})
        texto = "".join(c.get("text", "") for c in data.get("content", []) if c.get("type") == "text")
    else:  # compatível com /chat/completions: só texto (PDFs viram aviso)
        texto_user = "\n\n".join(p["texto"] if p["tipo"] == "texto" else
                                  f"[Referência {p['nome']} é PDF e não pôde ser lida por este provedor; "
                                  "envie em .docx para usá-la.]" for p in partes)
        headers = {"Authorization": f"Bearer {chave}"} if chave else {}
        data = _post(f"{base}/chat/completions", {"model": modelo, "messages": [
            {"role": "system", "content": instrucoes}, {"role": "user", "content": texto_user}]}, headers)
        escolhas = data.get("choices") or [{}]
        texto = (escolhas[0].get("message") or {}).get("content") or ""
    texto = texto.strip()
    if not texto:
        raise RuntimeError("A IA não devolveu texto.")
    # às vezes o modelo embrulha tudo em ```markdown ... ```
    return re.sub(r"^```\w*\n(.*?)\n```$", r"\1", texto, flags=re.S).strip()


def gerar_rascunho(cfg_ia, atividades, mes, projeto, referencias, pedido=""):
    """referencias: lista de caminhos (.docx/.dotx/.pdf). Devolve o relatório em Markdown."""
    if not atividades:
        raise RuntimeError("Não há atividades registradas nesse mês.")
    partes = []
    for path in referencias:
        nome = os.path.basename(path)
        if path.lower().endswith(".pdf"):
            with open(path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode()
            partes.append({"tipo": "texto", "texto": f"Relatório de referência: {nome} (PDF a seguir)"})
            partes.append({"tipo": "pdf", "nome": nome, "b64": b64})
        else:
            texto = docx_para_markdown(path)[:MAX_CHARS_REFERENCIA]
            partes.append({"tipo": "texto", "texto": f"Relatório de referência: {nome}\n<<<\n{texto}\n>>>"})
    if not referencias:
        partes.append({"tipo": "texto", "texto": "Não há relatórios de referência: use uma estrutura clara de "
                       "relatório mensal (visão geral, atividades por frente de trabalho, entregas, próximos passos)."})
    alvo = f"Escreva o relatório de {nome_mes(mes)}"
    if projeto:
        alvo += f", somente sobre o projeto {projeto}"
    partes.append({"tipo": "texto", "texto": f"{alvo}. Atividades registradas no mês:\n\n{descrever_atividades(atividades)}"})
    if pedido.strip():
        partes.append({"tipo": "texto", "texto": "Orientações adicionais de quem pediu o relatório "
                       "(siga, desde que não contrariem as regras):\n" + pedido.strip()})
    return chamar_llm(cfg_ia, INSTRUCOES, partes)


# ---------------------------------------------------------------- Markdown -> blocos
# Um único parser para o .docx e para a pré-visualização da tela (as duas ficam iguais).
#
# Blocos: ('h', nível 1-4, texto) | ('p', texto) | ('li', nível 1-3, texto, número|None)
#         ('tabela', [linhas]) | ('citacao', texto) | ('codigo', texto) | ('hr',)

RE_HR = re.compile(r"(?:\*\s*){3,}|(?:-\s*){3,}|(?:_\s*){3,}")
RE_TITULO = re.compile(r"(#{1,6})\s+(.*?)(?:\s+#+)?\s*$")
RE_ITEM = re.compile(r"^(\s*)([-*+•]|\d{1,3}[.)])\s+(.*)$")


def markdown_para_blocos(md):
    blocos, par, tab, cod = [], [], [], None

    def fecha():
        if par:
            blocos.append(("p", " ".join(par)))
            par.clear()
        if tab:
            blocos.append(("tabela", [r[:] for r in tab]))
            tab.clear()

    for linha in md.replace("\r\n", "\n").replace("\t", "    ").split("\n"):
        s = linha.strip()
        if cod is not None:  # dentro de bloco de código ```
            if s.startswith("```"):
                blocos.append(("codigo", "\n".join(cod)))
                cod = None
            else:
                cod.append(linha)
            continue
        if s.startswith("```"):
            fecha()
            cod = []
            continue
        if s.startswith("|"):
            if par:
                blocos.append(("p", " ".join(par)))
                par.clear()
            cels = [c.strip() for c in re.split(r"(?<!\\)\|", s.strip().strip("|"))]
            if not all(re.fullmatch(r":?-{2,}:?", c) for c in cels if c):  # linha |---|---|
                tab.append([c.replace("\\|", "|") for c in cels])
            continue
        if tab:
            fecha()
        if not s:
            fecha()
            continue
        if RE_HR.fullmatch(s):
            fecha()
            blocos.append(("hr",))
            continue
        m = RE_TITULO.match(s)
        if m:
            fecha()
            texto = re.sub(r"^\d+(\.\d+)*\.?\s+", "", m.group(2))  # numeração é automática
            blocos.append(("h", min(len(m.group(1)), 4), texto))
            continue
        if s.startswith(">"):
            texto = s.lstrip(">").strip()
            if blocos and blocos[-1][0] == "citacao" and not par:
                blocos[-1] = ("citacao", blocos[-1][1] + " " + texto)
            else:
                fecha()
                blocos.append(("citacao", texto))
            continue
        m = RE_ITEM.match(linha)
        if m:
            fecha()
            nivel = min(1 + len(m.group(1)) // 2, 3)
            num = None
            if m.group(2)[0].isdigit():
                ant = next((b for b in reversed(blocos) if b[0] == "li" and b[1] <= nivel), None)
                continua = blocos and blocos[-1][0] == "li" and ant and ant[1] == nivel and ant[3] is not None
                num = ant[3] + 1 if continua else int(m.group(2)[:-1])
            blocos.append(("li", nivel, m.group(3), num))
            continue
        if blocos and blocos[-1][0] == "li" and not par and linha.startswith(" "):
            b = blocos[-1]  # continuação do item da lista
            blocos[-1] = ("li", b[1], b[2] + " " + s, b[3])
            continue
        par.append(s)
    if cod is not None:
        blocos.append(("codigo", "\n".join(cod)))
    fecha()
    return blocos


RE_INLINE = re.compile(
    r"(?P<esc>\\[\\`*_{}\[\]()#+\-.!|~>])"
    r"|(?P<bi>\*\*\*(?=\S)(?P<bi_t>.+?)(?<=\S)\*\*\*)"
    r"|(?P<b>\*\*(?=\S)(?P<b_t>.+?)(?<=\S)\*\*|__(?=\S)(?P<b2_t>.+?)(?<=\S)__)"
    r"|(?P<i>\*(?=[^\s*])(?P<i_t>.+?)(?<=[^\s*])\*|(?<![\w])_(?=\S)(?P<i2_t>.+?)(?<=\S)_(?![\w]))"
    r"|(?P<s>~~(?=\S)(?P<s_t>.+?)(?<=\S)~~)"
    r"|(?P<c>`(?P<c_t>[^`]+)`)"
    r"|(?P<l>\[(?P<l_t>[^\]]+)\]\((?P<l_u>[^)\s]+)\))"
    r"|(?P<u><(?P<u_u>https?://[^>\s]+)>)"
)


def inline(texto, fmt=frozenset(), link=None):
    """Texto Markdown de uma linha -> [(texto, {b,i,s,c}, link)]."""
    out, pos = [], 0
    for m in RE_INLINE.finditer(texto):
        if m.start() > pos:
            out.append((texto[pos:m.start()], fmt, link))
        g = m.lastgroup
        if m.group("esc"):
            out.append((m.group("esc")[1], fmt, link))
        elif m.group("bi"):
            out += inline(m.group("bi_t"), fmt | {"b", "i"}, link)
        elif m.group("b"):
            out += inline(m.group("b_t") or m.group("b2_t"), fmt | {"b"}, link)
        elif m.group("i"):
            out += inline(m.group("i_t") or m.group("i2_t"), fmt | {"i"}, link)
        elif m.group("s"):
            out += inline(m.group("s_t"), fmt | {"s"}, link)
        elif m.group("c"):
            out.append((m.group("c_t"), fmt | {"c"}, link))
        elif m.group("l"):
            out += inline(m.group("l_t"), fmt, m.group("l_u"))
        elif m.group("u"):
            out.append((m.group("u_u"), fmt, m.group("u_u")))
        pos = m.end()
    if pos < len(texto):
        out.append((texto[pos:], fmt, link))
    return [p for p in out if p[0]]


def texto_puro(md_inline):
    return "".join(t for t, _, _ in inline(md_inline))


# ---------------------------------------------------------------- Markdown -> HTML (pré-visualização)

def _html_inline(texto):
    out = []
    for t, fmt, link in inline(texto):
        h = escape(t)
        if "c" in fmt:
            h = f"<code>{h}</code>"
        if "s" in fmt:
            h = f"<s>{h}</s>"
        if "i" in fmt:
            h = f"<i>{h}</i>"
        if "b" in fmt:
            h = f"<b>{h}</b>"
        if link:
            h = f'<a href="{escape(link, {chr(34): "&quot;"})}" target="_blank" rel="noopener">{h}</a>'
        out.append(h)
    return "".join(out)


def markdown_para_html(md):
    """Mesma interpretação do .docx: títulos numerados 1, 1.1, 1.1.1..."""
    html, cont, lista = [], [0, 0, 0, 0], None
    for b in markdown_para_blocos(md):
        tipo = ("ol" if b[3] is not None else "ul") if b[0] == "li" else None
        if lista and (b[0] != "li"):
            html.append(f"</{lista}>")
            lista = None
        if b[0] == "h":
            n = b[1]
            cont[n - 1] += 1
            cont[n:] = [0] * (4 - n)
            html.append(f"<h{n + 1}>{'.'.join(map(str, cont[:n]))}. {_html_inline(b[2])}</h{n + 1}>")
        elif b[0] == "p":
            html.append(f"<p>{_html_inline(b[1])}</p>")
        elif b[0] == "li":
            if lista != tipo:
                if lista:
                    html.append(f"</{lista}>")
                html.append(f"<{tipo}>")
                lista = tipo
            marcador = f'<span class="num">{b[3]}.</span> ' if b[3] is not None else ""
            html.append(f'<li class="l{b[1]}">{marcador}{_html_inline(b[2])}</li>')
        elif b[0] == "citacao":
            html.append(f"<blockquote>{_html_inline(b[1])}</blockquote>")
        elif b[0] == "codigo":
            html.append(f"<pre>{escape(b[1])}</pre>")
        elif b[0] == "hr":
            html.append("<hr>")
        elif b[0] == "tabela":
            cab, *linhas = b[1]
            html.append("<table><tr>" + "".join(f"<th>{_html_inline(c)}</th>" for c in cab) + "</tr>"
                        + "".join("<tr>" + "".join(f"<td>{_html_inline(c)}</td>" for c in r) + "</tr>" for r in linhas)
                        + "</table>")
    if lista:
        html.append(f"</{lista}>")
    return "".join(html)


# ---------------------------------------------------------------- montagem do .docx

class Links:
    """Junta os hiperlinks do texto para criar as relações no document.xml.rels."""
    def __init__(self):
        self.itens = []

    def rid(self, url):
        self.itens.append(url)
        return f"rIdDiario{len(self.itens)}"


_links = Links()


def runs(texto, rpr=""):
    """Texto com Markdown inline (**negrito**, *itálico*, ~~riscado~~, `código`, [link](url)) -> <w:r>..."""
    out = []
    for t, fmt, link in inline(texto):
        extra = ("<w:b/>" if "b" in fmt else "") + ("<w:i/>" if "i" in fmt else "") \
            + ("<w:strike/>" if "s" in fmt else "")
        if "c" in fmt:
            extra += '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:shd w:val="clear" w:color="auto" w:fill="F0F0F0"/>'
        if link:
            extra += '<w:color w:val="0563C1"/><w:u w:val="single"/>'
        props = f"<w:rPr>{rpr}{extra}</w:rPr>" if (rpr or extra) else ""
        r = f'<w:r>{props}<w:t xml:space="preserve">{escape(t)}</w:t></w:r>'
        if link:
            r = f'<w:hyperlink r:id="{_links.rid(link)}" w:history="1">{r}</w:hyperlink>'
        out.append(r)
    return "".join(out)


def paragrafo(texto, estilo=None, ppr_extra="", rpr=""):
    estilo_xml = f'<w:pStyle w:val="{estilo}"/>' if estilo else ""
    ppr = f"<w:pPr>{estilo_xml}{ppr_extra}</w:pPr>" if (estilo_xml or ppr_extra) else ""
    return f"<w:p>{ppr}{runs(texto, rpr)}</w:p>"


def tabela_xml(linhas, est):
    """Tabela no visual do template: bordas finas na cor dos títulos, cabeçalho em negrito."""
    ncol = max(len(r) for r in linhas)
    linhas = [r + [""] * (ncol - len(r)) for r in linhas]
    borda = '<w:{0} w:val="single" w:sz="4" w:space="0" w:color="%s"/>' % (est.get("cor") or "auto")
    largura = 5000 // ncol
    trs = []
    for i, r in enumerate(linhas):
        tcs = []
        for j, c in enumerate(r):
            bordas = borda.format("top") + (borda.format("bottom") if i == 0 else "")
            shd = ""
            negrito = "<w:b/>" if i == 0 else ""
            tcs.append(
                f'<w:tc><w:tcPr><w:tcW w:w="{largura}" w:type="pct"/><w:tcBorders>{bordas}</w:tcBorders>'
                f'{shd}<w:vAlign w:val="center"/></w:tcPr>'
                + paragrafo(c, est["tabela"], '<w:jc w:val="left"/>', negrito) + "</w:tc>")
        cab = "<w:trPr><w:tblHeader/></w:trPr>" if i == 0 else ""
        trs.append(f"<w:tr>{cab}{''.join(tcs)}</w:tr>")
    grid = "".join(f'<w:gridCol w:w="{9800 // ncol}"/>' for _ in range(ncol))
    return (f'<w:tbl><w:tblPr><w:tblStyle w:val="{est["grade"]}"/><w:tblW w:w="5000" w:type="pct"/>'
            '<w:tblBorders><w:insideV w:val="single" w:sz="4" w:space="0" w:color="D9D9D9"/></w:tblBorders>'
            '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
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
        "citacao": por_nome.get("quote") or por_nome.get("intense quote"),
        "cor": cor_titulo_hex(estilos_xml),
    }


def corpo_xml(blocos, est):
    xml, toc, cont = [], [], [0, 0, 0, 0]
    for b in blocos:
        if b[0] == "h":
            n = b[1]
            cont[n - 1] += 1
            cont[n:] = [0] * (4 - n)
            if n <= 3:
                toc.append((n, ".".join(str(c) for c in cont[:n]), texto_puro(b[2])))
            if est["h"][n]:
                xml.append(paragrafo(b[2], est["h"][n]))
            else:  # template sem estilos de título: formata direto
                tam = {1: 32, 2: 28, 3: 24, 4: 22}[n]
                xml.append(paragrafo(b[2], None, '<w:keepNext/><w:spacing w:before="240" w:after="120"/>'
                                     f'<w:outlineLvl w:val="{n - 1}"/>', f'<w:b/><w:sz w:val="{tam}"/>'))
        elif b[0] == "p":
            xml.append(paragrafo(b[1]))
        elif b[0] == "li":
            nivel, numero = b[1], b[3]
            recuo = 360 * nivel
            if numero is not None:  # lista numerada: número escrito, com recuo pendente
                xml.append(paragrafo(f"{numero}.\u00a0" + b[2], None,
                                     f'<w:ind w:left="{recuo + 360}" w:hanging="360"/>'))
            elif est["li"].get(min(nivel, 2)) and nivel <= 2:
                xml.append(paragrafo(b[2], est["li"][nivel]))
            else:
                xml.append(paragrafo(("• ", "◦ ", "▪ ")[nivel - 1] + b[2], None,
                                     f'<w:ind w:left="{recuo + 240}" w:hanging="240"/>'))
        elif b[0] == "citacao":
            xml.append(paragrafo(b[1], est.get("citacao"), "" if est.get("citacao") else
                                 '<w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="A6A6A6"/></w:pBdr>'
                                 '<w:ind w:left="400"/>', "" if est.get("citacao") else "<w:i/>"))
        elif b[0] == "codigo":
            for i, linha in enumerate(b[1].split("\n") or [""]):
                xml.append(f'<w:p><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>'
                           f'<w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:rPr>'
                           f'<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/><w:sz w:val="18"/></w:rPr>'
                           f'<w:t xml:space="preserve">{escape(linha)}</w:t></w:r></w:p>')
            xml.append("<w:p/>")
        elif b[0] == "hr":
            xml.append('<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="A6A6A6"/>'
                       '</w:pBdr></w:pPr></w:p>')
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


_cache_template = {}
MARCADOR = re.compile(r"\{\{\s*(\w+)\s*\}\}")


def preencher_marcadores(xml, valores):
    """Troca {{titulo}}, {{autor}}, {{mes}}... em cada parágrafo (o Word costuma partir o texto
    em vários trechos; o parágrafo é refeito com a formatação do primeiro trecho)."""
    def f(m):
        p = m.group(0)
        if "{" not in p:
            return p
        t = texto_xml(p)
        if not MARCADOR.search(t):
            return p
        novo = MARCADOR.sub(lambda mm: valores.get(mm.group(1).lower(), mm.group(0)), t)
        return trocar_paragrafo(p, novo) if novo != t else p
    return P_INTERNO.sub(f, xml)


def montar_docx(padrao_path, markdown, capa):
    """capa: {titulo, subtitulo, autor, area, versao, mes, projeto}. Devolve os bytes do .docx.

    Funciona com qualquer template .docx/.dotx:
      - marcadores {{titulo}} {{subtitulo}} {{autor}} {{area}} {{versao}} {{mes}} {{projeto}} {{data}}
        são preenchidos no corpo, cabeçalho e rodapé;
      - o texto do relatório entra no parágrafo {{conteudo}}; sem ele, substitui o conteúdo de exemplo
        a partir do primeiro Título 1 (depois do sumário); sem isso, vai para o fim do documento;
      - usa os estilos Título 1–4 e Lista do template; se não existirem, formata direto."""
    chave = (padrao_path, os.path.getmtime(padrao_path))
    if chave not in _cache_template:  # a pré-visualização monta o documento a cada edição
        with zipfile.ZipFile(padrao_path) as z:
            _cache_template.clear()
            _cache_template[chave] = ({i.filename: z.read(i.filename) for i in z.infolist()},
                                      [i.filename for i in z.infolist()])
    arquivos, ordem = dict(_cache_template[chave][0]), _cache_template[chave][1]
    doc = arquivos["word/document.xml"].decode("utf-8")
    est = estilos_do_padrao(arquivos.get("word/styles.xml", b"").decode("utf-8"))

    blocos = markdown_para_blocos(markdown)
    if not any(b[0] == "h" for b in blocos):
        raise RuntimeError("O texto precisa de pelo menos um título de seção (linha começando com '# ').")
    _links.itens = []
    corpo, toc = corpo_xml(blocos, est)
    if _links.itens:  # hiperlinks do texto viram relações do documento
        rels_path = "word/_rels/document.xml.rels"
        rels = arquivos[rels_path].decode("utf-8")
        novos = "".join(
            f'<Relationship Id="rIdDiario{n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
            f'Target="{escape(url, {chr(34): "&quot;"})}" TargetMode="External"/>' for n, url in enumerate(_links.itens, 1))
        arquivos[rels_path] = rels.replace("</Relationships>", novos + "</Relationships>").encode("utf-8")
        if 'xmlns:r="' not in doc[:doc.index("<w:body>")]:
            doc = doc.replace("<w:document ", '<w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ', 1)

    a, b = doc.index("<w:body>") + 8, doc.rindex("</w:body>")
    els = elementos_topo(doc[a:b])
    sect_final = [els.pop()] if els and els[-1].startswith("<w:sectPr") else []
    i_marcador = next((i for i, e in enumerate(els) if re.search(r"\{\{\s*conteudo\s*\}\}", texto_xml(e), re.I)), None)
    i_toc = next((i for i, e in enumerate(els) if "TOC \\o" in e), None)
    h1 = f'<w:pStyle w:val="{est["h"][1]}"/>' if est["h"][1] else None
    i_corpo = next((i for i, e in enumerate(els) if h1 and h1 in e and (i_toc is None or i > i_toc)), None)

    if i_marcador is not None:                       # 1) parágrafo {{conteudo}}
        novo = els[:i_marcador] + [corpo] + els[i_marcador + 1:]
    elif i_corpo is not None:                        # 2) troca o corpo de exemplo do template
        # fim do exemplo: a quebra de seção que antecede a contracapa (ou o fim do documento)
        i_fim = next((i for i in range(i_corpo, len(els)) if "<w:sectPr" in els[i] and els[i].startswith("<w:p")),
                     len(els))
        inicio = els[:i_toc if i_toc is not None else i_corpo]
        indice = [indice_xml(toc, est)] if i_toc is not None else []
        novo = inicio + indice + [corpo] + els[i_fim:]
    else:                                            # 3) template sem corpo: acrescenta no fim
        novo = els + [corpo]
    doc = doc[:a] + "".join(novo + sect_final) + doc[b:]

    valores = {k: str(capa.get(k) or "") for k in ("titulo", "subtitulo", "autor", "area", "versao", "mes", "projeto")}
    valores["versao"] = valores["versao"] or "1.0"
    valores["data"] = datetime.now().strftime("%d/%m/%Y")
    # textos de exemplo do padrão Indra (templates que não usam {{marcadores}})
    trocas = {CAPA_AUTOR: valores["autor"], CAPA_AREA: valores["area"],
              CAPA_TITULO: valores["titulo"], CAPA_SUBTITULO: valores["subtitulo"]}
    doc = preencher_marcadores(substituir_textos(doc, trocas, CAPA_REMOVER), valores)
    arquivos["word/document.xml"] = doc.encode("utf-8")

    rodape = f"{valores['subtitulo']} – Versão {valores['versao']}"
    for nome in list(arquivos):
        if re.fullmatch(r"word/(header|footer)\d*\.xml", nome):
            x = arquivos[nome].decode("utf-8")
            x2 = preencher_marcadores(substituir_textos(x, {RODAPE_TITULO: rodape}, CABECALHO_REMOVER), valores)
            if x2 != x:
                arquivos[nome] = x2.encode("utf-8")
    if "word/settings.xml" in arquivos:
        arquivos["word/settings.xml"] = ativar_atualizacao_campos(
            arquivos["word/settings.xml"].decode("utf-8")).encode("utf-8")
    # um .dotx vira documento comum (.docx) ao ser preenchido
    tipos = arquivos["[Content_Types].xml"].decode("utf-8")
    arquivos["[Content_Types].xml"] = tipos.replace("wordprocessingml.template.main+xml",
                                                    "wordprocessingml.document.main+xml").encode("utf-8")

    podar_partes(arquivos)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for nome in ordem:
            if nome in arquivos:
                z.writestr(nome, arquivos[nome])
    return buf.getvalue()


# ---------------------------------------------------------------- visual do template (pré-visualização)

_cache_estilo = {}


def _estilos_por_id(estilos_xml):
    return {m.group(1): m.group(0) for m in re.finditer(
        r'<w:style [^>]*w:styleId="([^"]+)"[^>]*>.*?</w:style>', estilos_xml, re.S)}


def _rpr_props(xml):
    """Fonte, tamanho, cor, negrito e itálico de um trecho com <w:rPr>."""
    m = re.search(r"<w:rPr>(.*?)</w:rPr>", xml or "", re.S)
    rpr = m.group(1) if m else ""
    out = {}
    f = re.search(r'<w:rFonts [^>]*w:ascii="([^"]+)"', rpr)
    ft = re.search(r'<w:rFonts [^>]*w:asciiTheme="(major|minor)', rpr)
    if f:
        out["fonte"] = f.group(1)
    elif ft:
        out["fonte_tema"] = ft.group(1)
    s = re.search(r'<w:sz w:val="(\d+)"', rpr)
    if s:
        out["tamanho"] = int(s.group(1)) / 2
    c = re.search(r'<w:color [^>]*w:val="([0-9A-Fa-f]{6})"', rpr)
    if c:
        out["cor"] = "#" + c.group(1)
    for tag, chave in (("b", "negrito"), ("i", "italico")):
        t = re.search(rf'<w:{tag}(?: w:val="(\w+)")?/>', rpr)
        if t:
            out[chave] = t.group(1) not in ("0", "false")
    return out


def visual_do_template(path):
    """CSS da pré-visualização a partir dos estilos Normal e Título 1-4 do template."""
    chave = (path, os.path.getmtime(path))
    if chave in _cache_estilo:
        return _cache_estilo[chave]
    with zipfile.ZipFile(path) as z:
        nomes = z.namelist()
        estilos = z.read("word/styles.xml").decode("utf-8") if "word/styles.xml" in nomes else ""
        tema = z.read("word/theme/theme1.xml").decode("utf-8") if "word/theme/theme1.xml" in nomes else ""
    fontes_tema = {k: (re.search(rf"<a:{k}Font>.*?<a:latin typeface=\"([^\"]*)\"", tema, re.S) or [None, None])[1]
                   for k in ("major", "minor")}
    por_id = _estilos_por_id(estilos)
    por_nome = {v: k for k, v in nomes_estilos(estilos).items()}
    padrao = _rpr_props((re.search(r"<w:rPrDefault>.*?</w:rPrDefault>", estilos, re.S) or [""])[0])

    def resolver(sid, vistos=()):
        xml = por_id.get(sid, "")
        base = re.search(r'<w:basedOn w:val="([^"]+)"', xml)
        props = resolver(base.group(1), vistos + (sid,)) if base and base.group(1) not in vistos else dict(padrao)
        props.update(_rpr_props(xml))
        return props

    def css(props, padrao_tam):
        fonte = props.get("fonte") or fontes_tema.get(props.get("fonte_tema", "minor")) or "Calibri"
        regras = [f"font-family: '{fonte}', Carlito, Calibri, sans-serif",
                  f"font-size: {props.get('tamanho', padrao_tam) * 1.33:.1f}px"]
        if props.get("cor"):
            regras.append(f"color: {props['cor']}")
        regras.append(f"font-weight: {700 if props.get('negrito') else 400}")
        regras.append(f"font-style: {'italic' if props.get('italico') else 'normal'}")
        return "; ".join(regras)

    normal = resolver(por_nome.get("normal", "Normal"))
    regras = [f"#preview {{ {css(normal, 11)} }}"]
    cor_titulo = None
    for n in range(1, 5):
        sid = por_nome.get(f"heading {n}")
        if not sid:
            continue
        props = resolver(sid)
        cor_titulo = cor_titulo or props.get("cor")
        regras.append(f"#preview h{n + 1} {{ {css(props, 11)} }}")
    borda = cor_titulo or "#7f7f7f"
    regras.append(f"#preview th, #preview td {{ border-top: 1px solid {borda}; }} #preview th {{ border-bottom: 1px solid {borda}; }}")
    _cache_estilo.clear()
    _cache_estilo[chave] = "\n".join(regras)
    return _cache_estilo[chave]


def cor_titulo_hex(estilos_xml):
    """Cor do Título 1 do template (para as bordas das tabelas), sem o '#'. None se não tiver."""
    por_id = _estilos_por_id(estilos_xml)
    sid = {v: k for k, v in nomes_estilos(estilos_xml).items()}.get("heading 1")
    while sid:
        cor = _rpr_props(por_id.get(sid, "")).get("cor")
        if cor:
            return cor[1:]
        base = re.search(r'<w:basedOn w:val="([^"]+)"', por_id.get(sid, ""))
        sid = base.group(1) if base else None
    return None
