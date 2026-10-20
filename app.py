#!/usr/bin/env python3
"""Diário de Atividades — app local para registrar o que foi feito no dia
e gerar o relatório mensal para o cliente.

Só usa a biblioteca padrão do Python. Os dados ficam em data/atividades.db.
Acesse em http://127.0.0.1:8765
"""
import csv
import io
import json
import os
import re
import shutil
import sqlite3
import unicodedata
import zipfile
from datetime import date, datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, quote, unquote, urlparse

import relatorio

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
DB_PATH = os.path.join(DATA_DIR, "atividades.db")
MODELOS_DIR = os.path.join(BASE_DIR, "modelos")
PERFIS_DIR = os.path.join(MODELOS_DIR, "perfis")   # um subdiretório por modelo de relatório
# tipo de arquivo -> quantos cada modelo guarda (o padrão é um só; referências podem ser várias)
TIPOS = {"padrao": 1, "referencia": 10}
MODELO_MAX_BYTES = 20 * 1024 * 1024
FORMATOS = {  # extensão -> tipo MIME
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "dotx": "application/vnd.openxmlformats-officedocument.wordprocessingml.template",
    "pdf": "application/pdf",
}
CONFIG_ANTIGA = os.path.join(DATA_DIR, "config.json")      # capa única (antes dos modelos múltiplos)
PREFS_PATH = os.path.join(DATA_DIR, "preferencias.json")   # modelo principal
IA_PATH = os.path.join(DATA_DIR, "ia.json")                # provedor/modelo/chave da IA
RASCUNHOS_DIR = os.path.join(DATA_DIR, "rascunhos")
relatorio.carregar_env(os.path.join(BASE_DIR, ".env"))
HOST = os.environ.get("DIARIO_HOST", "127.0.0.1")
PORT = int(os.environ.get("DIARIO_PORT", "8765"))

VERSAO = "4"  # o index.html confere: se o servidor for de outra versão, pede para reiniciar o app
FIELDS = ("data", "projeto", "categoria", "ticket", "descricao", "horas", "entregavel")


def ler_json(path, padrao):
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return padrao


def gravar_json(path, dados):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path + ".tmp", "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)
    os.replace(path + ".tmp", path)


# ---------------------------------------------------------------- banco

def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    os.makedirs(DATA_DIR, exist_ok=True)
    with db() as conn:
        conn.execute(
            """CREATE TABLE IF NOT EXISTS atividades (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                data       TEXT NOT NULL,
                projeto    TEXT NOT NULL DEFAULT '',
                categoria  TEXT NOT NULL DEFAULT '',
                ticket     TEXT NOT NULL DEFAULT '',
                descricao  TEXT NOT NULL,
                horas      REAL NOT NULL DEFAULT 0,
                criado_em  TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
            )"""
        )
        conn.execute("CREATE INDEX IF NOT EXISTS idx_atividades_data ON atividades(data)")
        colunas = {r[1] for r in conn.execute("PRAGMA table_info(atividades)")}
        if "entregavel" not in colunas:  # migração: marca de entrega para o relatório do mês
            conn.execute("ALTER TABLE atividades ADD COLUMN entregavel INTEGER NOT NULL DEFAULT 0")


def clean(payload):
    """Valida e normaliza os campos recebidos do formulário."""
    item = {}
    item["data"] = str(payload.get("data") or date.today().isoformat())
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", item["data"]):
        raise ValueError("Data inválida (use AAAA-MM-DD).")
    item["descricao"] = str(payload.get("descricao") or "").strip()
    if not item["descricao"]:
        raise ValueError("Descreva a atividade.")
    for f in ("projeto", "categoria", "ticket"):
        item[f] = str(payload.get(f) or "").strip()
    try:
        item["horas"] = round(float(str(payload.get("horas") or 0).replace(",", ".")), 2)
    except ValueError:
        raise ValueError("Horas inválidas.")
    if item["horas"] < 0 or item["horas"] > 24:
        raise ValueError("Horas devem estar entre 0 e 24.")
    item["entregavel"] = 1 if str(payload.get("entregavel") or "").lower() in ("1", "true", "sim", "on") else 0
    return item


def month_range(month):
    if not re.fullmatch(r"\d{4}-\d{2}", month or ""):
        raise ValueError("Mês inválido (use AAAA-MM).")
    return f"{month}-01", f"{month}-31"


def atividades_do_mes(mes, projeto=""):
    ini, fim = month_range(mes)
    sql = "SELECT * FROM atividades WHERE data BETWEEN ? AND ?"
    params = [ini, fim]
    if projeto:
        sql += " AND projeto = ?"
        params.append(projeto)
    with db() as conn:
        return [dict(r) for r in conn.execute(sql + " ORDER BY data, id", params)]


# ---------------------------------------------------------------- modelos de relatório
# Cada modelo fica em modelos/perfis/<id>/ com:
#   perfil.json   nome, dados da capa e projetos que usam este modelo
#   padrao/       o .docx/.dotx da empresa (um só)
#   referencia/   relatórios prontos que a IA usa de exemplo (vários)

CAPA_CAMPOS = ("titulo", "subtitulo", "autor", "area", "versao")
CAPA_PADRAO = {"titulo": "", "subtitulo": "Relatório de Atividades | {mes}", "autor": "", "area": "", "versao": "1.0"}
ID_RE = re.compile(r"[a-z0-9][a-z0-9-]{0,48}")


def slug(nome):
    s = unicodedata.normalize("NFKD", nome).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", "-", s).strip("-")[:40] or "modelo"


def perfil_dir(pid):
    if not ID_RE.fullmatch(pid or ""):
        raise ValueError("Modelo inválido.")
    return os.path.join(PERFIS_DIR, pid)


def perfil_existe(pid):
    try:
        return os.path.isfile(os.path.join(perfil_dir(pid), "perfil.json"))
    except ValueError:
        return False


def ids_perfis():
    if not os.path.isdir(PERFIS_DIR):
        return []
    ids = [d for d in os.listdir(PERFIS_DIR) if ID_RE.fullmatch(d) and perfil_existe(d)]
    return sorted(ids, key=lambda i: ler_perfil(i)["nome"].lower())


def perfil_principal():
    ids = ids_perfis()
    pref = ler_json(PREFS_PATH, {}).get("perfil_principal")
    return pref if pref in ids else (ids[0] if ids else None)


def ler_perfil(pid):
    dados = ler_json(os.path.join(perfil_dir(pid), "perfil.json"), {})
    out = {"id": pid, "nome": str(dados.get("nome") or pid)}
    out.update({k: str(dados.get(k, CAPA_PADRAO[k])) for k in CAPA_CAMPOS})
    out["projetos"] = [str(p) for p in dados.get("projetos", []) if str(p).strip()]
    return out


def gravar_perfil(pid, dados):
    atual = ler_perfil(pid) if perfil_existe(pid) else {"nome": pid, **CAPA_PADRAO, "projetos": []}
    novo = {"nome": str(dados.get("nome", atual["nome"])).strip() or atual["nome"]}
    novo.update({k: str(dados.get(k, atual[k])).strip() for k in CAPA_CAMPOS})
    projetos = dados.get("projetos", atual["projetos"])
    if isinstance(projetos, str):
        projetos = projetos.split(",")
    novo["projetos"] = sorted({str(p).strip() for p in projetos if str(p).strip()})
    gravar_json(os.path.join(perfil_dir(pid), "perfil.json"), novo)
    if dados.get("principal"):
        gravar_json(PREFS_PATH, {**ler_json(PREFS_PATH, {}), "perfil_principal": pid})
    return ler_perfil(pid)


def criar_perfil(nome, copiar_de=None):
    nome = (nome or "").strip()
    if not nome:
        raise ValueError("Dê um nome ao modelo.")
    base, pid, n = slug(nome), slug(nome), 2
    while os.path.exists(perfil_dir(pid)):
        pid, n = f"{base}-{n}", n + 1
    dados = {**CAPA_PADRAO, "projetos": []}
    if copiar_de and perfil_existe(copiar_de):  # aproveita a capa de outro modelo
        dados.update({k: ler_perfil(copiar_de)[k] for k in CAPA_CAMPOS})
    for tipo in TIPOS:
        os.makedirs(os.path.join(perfil_dir(pid), tipo), exist_ok=True)
    return gravar_perfil(pid, {**dados, "nome": nome})


def excluir_perfil(pid):
    if not perfil_existe(pid):
        raise ValueError("Modelo não encontrado.")
    if len(ids_perfis()) == 1:
        raise ValueError("Esse é o único modelo. Crie outro antes de excluir.")
    shutil.rmtree(perfil_dir(pid))


def arquivos(pid, tipo):
    """Arquivos aceitos em modelos/perfis/<id>/<tipo>/, do mais recente para o mais antigo."""
    pasta = os.path.join(perfil_dir(pid), tipo)
    if not os.path.isdir(pasta):
        return []
    nomes = [n for n in os.listdir(pasta) if n.rsplit(".", 1)[-1].lower() in FORMATOS]
    return sorted(nomes, key=lambda n: os.stat(os.path.join(pasta, n)).st_mtime, reverse=True)


def arquivo_info(pid, tipo, nome):
    st = os.stat(os.path.join(perfil_dir(pid), tipo, nome))
    return {
        "nome": nome,
        "formato": nome.rsplit(".", 1)[1].lower(),
        "tamanho": st.st_size,
        "enviado_em": datetime.fromtimestamp(st.st_mtime).strftime("%d/%m/%Y %H:%M"),
    }


def perfil_completo(pid):
    p = ler_perfil(pid)
    p["arquivos"] = {}
    for tipo, limite in TIPOS.items():
        infos = [arquivo_info(pid, tipo, n) for n in arquivos(pid, tipo)]
        p["arquivos"][tipo] = (infos[0] if infos else None) if limite == 1 else infos
    return p


def nome_seguro(nome, formato, padrao):
    """Nome original sem caminho nem caracteres estranhos, com a extensão do formato real."""
    base = os.path.splitext(os.path.basename(nome.replace("\\", "/")))[0]
    base = re.sub(r"[^\w\-. ()]+", "_", base).strip(" ._")[:100]
    return f"{base or padrao}.{formato}"


def detectar_formato(conteudo):
    """Descobre pelo conteúdo se é .pdf, .docx ou .dotx (o nome do arquivo não importa)."""
    if conteudo.startswith(b"%PDF-"):
        return "pdf"
    try:
        with zipfile.ZipFile(io.BytesIO(conteudo)) as z:
            if "word/document.xml" not in z.namelist():
                return None
            tipos = z.read("[Content_Types].xml")
    except (zipfile.BadZipFile, KeyError):
        return None
    return "dotx" if b"wordprocessingml.template.main+xml" in tipos else "docx"


def salvar_arquivo(pid, tipo, nome, conteudo):
    """Valida o arquivo e grava no modelo. No padrão, substitui o anterior."""
    if not perfil_existe(pid):
        raise ValueError("Modelo não encontrado.")
    if not conteudo:
        raise ValueError("Arquivo vazio.")
    if len(conteudo) > MODELO_MAX_BYTES:
        raise ValueError("Arquivo muito grande (máx. 20 MB).")
    formato = detectar_formato(conteudo)
    if not formato:
        raise ValueError("Formato não aceito. Envie .docx, .dotx ou .pdf.")
    if tipo == "padrao" and formato == "pdf":
        raise ValueError("O padrão precisa ser .docx ou .dotx (é sobre ele que o relatório é montado).")
    nome = nome_seguro(nome, formato, tipo)
    existentes = arquivos(pid, tipo)
    limite = TIPOS[tipo]
    if limite > 1 and nome not in existentes and len(existentes) >= limite:
        raise ValueError(f"Limite de {limite} arquivos. Remova algum antes de enviar outro.")
    pasta = os.path.join(perfil_dir(pid), tipo)
    os.makedirs(pasta, exist_ok=True)
    destino = os.path.join(pasta, nome)
    with open(destino + ".tmp", "wb") as f:
        f.write(conteudo)
    os.replace(destino + ".tmp", destino)
    if limite == 1:
        for antigo in existentes:
            if antigo != nome:
                os.remove(os.path.join(pasta, antigo))
    # modelo novo sem capa: sugere a capa a partir do próprio arquivo
    p = ler_perfil(pid)
    if tipo == "referencia" and formato in ("docx", "dotx") and not (p["titulo"] or p["autor"]):
        capa = relatorio.capa_do_docx(destino)
        sugestao = {k: capa[k] for k in ("titulo", "autor", "area") if capa.get(k)}
        if sugestao:
            gravar_perfil(pid, sugestao)
    return arquivo_info(pid, tipo, nome)


def migrar_modelos():
    """Versão antiga: um padrão em modelos/padrao, referências em modelos/referencia e capa em
    data/config.json. Move tudo para o modelo 'principal' na primeira execução."""
    if ids_perfis():
        return
    pid = "principal"
    for tipo in TIPOS:
        antiga, nova = os.path.join(MODELOS_DIR, tipo), os.path.join(perfil_dir(pid), tipo)
        os.makedirs(nova, exist_ok=True)
        if os.path.isdir(antiga):
            for n in os.listdir(antiga):
                if n.rsplit(".", 1)[-1].lower() in FORMATOS:
                    shutil.move(os.path.join(antiga, n), os.path.join(nova, n))
    capa = {k: v for k, v in ler_json(CONFIG_ANTIGA, {}).items() if k in CAPA_CAMPOS}
    if not capa:  # sem capa salva: sugere a de um relatório de referência
        for n in arquivos(pid, "referencia"):
            if n.lower().endswith((".docx", ".dotx")):
                c = relatorio.capa_do_docx(os.path.join(perfil_dir(pid), "referencia", n))
                capa = {k: c[k] for k in ("titulo", "autor", "area") if c.get(k)}
                break
    gravar_perfil(pid, {"nome": "Principal", **CAPA_PADRAO, **capa, "principal": True})


def perfil_para(pid, projeto=""):
    """Modelo pedido; senão o associado ao projeto; senão o principal."""
    if pid and perfil_existe(pid):
        return pid
    if projeto:
        for i in ids_perfis():
            if projeto in ler_perfil(i)["projetos"]:
                return i
    principal = perfil_principal()
    if not principal:
        raise ValueError("Crie um modelo de relatório em Configurações.")
    return principal


# ---------------------------------------------------------------- IA (qualquer provedor)

def ler_ia():
    """Configuração salva na tela; o .env vale como padrão (OPENAI_API_KEY / LLM_*)."""
    salvo = ler_json(IA_PATH, {})
    env = os.environ
    provedor = salvo.get("provedor") or env.get("LLM_PROVEDOR") or "openai"
    if provedor not in relatorio.PROVEDORES:
        provedor = "openai"
    chave_env = env.get("LLM_API_KEY") or (env.get("OPENAI_API_KEY") if provedor == "openai" else
                                            env.get("ANTHROPIC_API_KEY") if provedor == "anthropic" else "")
    modelo_env = env.get("LLM_MODELO") or (env.get("OPENAI_MODEL") if provedor == "openai" else "")
    return {
        "provedor": provedor,
        "modelo": salvo.get("modelo") or modelo_env or relatorio.PROVEDORES[provedor]["modelo"],
        "base_url": salvo.get("base_url") or env.get("LLM_BASE_URL") or relatorio.PROVEDORES[provedor]["base_url"],
        "chave": salvo.get("chave") or chave_env or "",
    }


def ia_publica():
    cfg = ler_ia()
    chave = cfg.pop("chave")
    cfg["chave"] = ("••••" + chave[-4:]) if chave else ""
    cfg["ativa"] = bool(chave) or cfg["provedor"] == "compativel"
    cfg["provedores"] = relatorio.PROVEDORES
    return cfg


def salvar_ia(payload):
    atual = ler_json(IA_PATH, {})
    provedor = payload.get("provedor") or atual.get("provedor") or "openai"
    if provedor not in relatorio.PROVEDORES:
        raise ValueError("Provedor desconhecido.")
    novo = {"provedor": provedor, "modelo": str(payload.get("modelo") or "").strip(),
            "base_url": str(payload.get("base_url") or "").strip().rstrip("/")}
    chave = str(payload.get("chave") or "").strip()
    if payload.get("limpar_chave"):
        chave = ""
    elif not chave or chave.startswith("••••"):
        chave = atual.get("chave", "") if atual.get("provedor") == provedor else ""
    novo["chave"] = chave
    gravar_json(IA_PATH, novo)
    try:
        os.chmod(IA_PATH, 0o600)  # a chave fica só para o seu usuário
    except OSError:
        pass
    return ia_publica()


# ---------------------------------------------------------------- rascunhos

def rascunho_path(mes, projeto):
    month_range(mes)
    sufixo = "-" + re.sub(r"[^\w-]+", "_", projeto).strip("_") if projeto else ""
    return os.path.join(RASCUNHOS_DIR, f"{mes}{sufixo}.md")


# ---------------------------------------------------------------- HTTP

class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):  # silencia o log padrão
        pass

    def send(self, status, body, ctype="application/json; charset=utf-8", extra=None):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, ensure_ascii=False)
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def read_json(self):
        length = int(self.headers.get("Content-Length") or 0)
        return json.loads(self.rfile.read(length) or b"{}")

    def entry_id(self, path):
        m = re.fullmatch(r"/api/atividades/(\d+)", path)
        return int(m.group(1)) if m else None

    def rota_perfil(self, path):
        """/api/perfis/<id>[/<tipo>[/<nome>]] -> (id, tipo|None, nome|None|'' se não existe)."""
        m = re.fullmatch(r"/api/perfis/([a-z0-9-]+)(?:/(padrao|referencia)(?:/([^/]+))?)?", path)
        if not m:
            return None
        pid, tipo, nome = m.group(1), m.group(2), unquote(m.group(3)) if m.group(3) else None
        if nome is not None and nome not in arquivos(pid, tipo):
            nome = ""
        return pid, tipo, nome

    def gerar_rascunho(self):
        try:
            body = self.read_json()
            mes, projeto = body.get("mes", ""), body.get("projeto", "")
            destino = rascunho_path(mes, projeto)
            pid = perfil_para(body.get("perfil"), projeto)
            refs = [os.path.join(perfil_dir(pid), "referencia", n) for n in arquivos(pid, "referencia")]
            md = relatorio.gerar_rascunho(ler_ia(), atividades_do_mes(mes, projeto), mes, projeto, refs,
                                          str(body.get("pedido") or ""))
        except (ValueError, json.JSONDecodeError, RuntimeError) as e:
            return self.send(400, {"erro": str(e)})
        os.makedirs(RASCUNHOS_DIR, exist_ok=True)
        with open(destino, "w", encoding="utf-8") as f:
            f.write(md)
        return self.send(200, {"markdown": md, "perfil": pid})

    def baixar_docx(self):
        try:
            body = self.read_json()
            mes, projeto = body.get("mes", ""), body.get("projeto", "")
            month_range(mes)
            pid = perfil_para(body.get("perfil"), projeto)
            p = ler_perfil(pid)
            padrao = arquivos(pid, "padrao")
            if not padrao or not padrao[0].lower().endswith((".docx", ".dotx")):
                raise ValueError(f"O modelo “{p['nome']}” ainda não tem o padrão .docx (Configurações).")
            capa = {k: p[k] for k in CAPA_CAMPOS}
            capa["mes"] = relatorio.nome_mes(mes)
            capa["projeto"] = projeto
            capa["subtitulo"] = capa["subtitulo"].replace("{mes}", capa["mes"])
            conteudo = relatorio.montar_docx(os.path.join(perfil_dir(pid), "padrao", padrao[0]),
                                             str(body.get("markdown") or ""), capa)
        except (ValueError, json.JSONDecodeError, RuntimeError) as e:
            return self.send(400, {"erro": str(e)})
        nome = f"Relatório de Atividades - {relatorio.nome_mes(mes)}{' - ' + projeto if projeto else ''}.docx"
        return self.send(200, conteudo, FORMATOS["docx"],
                         {"Content-Disposition": f"attachment; filename*=UTF-8''{quote(nome)}"})

    # ---------- GET ----------
    def do_GET(self):
        url = urlparse(self.path)
        qs = {k: v[0] for k, v in parse_qs(url.query).items()}
        try:
            if url.path in ("/", "/index.html"):
                with open(os.path.join(BASE_DIR, "index.html"), "rb") as f:
                    return self.send(200, f.read(), "text/html; charset=utf-8")

            if url.path == "/api/atividades":
                if "mes" in qs:
                    ini, fim = month_range(qs["mes"])
                else:
                    ini = fim = qs.get("data") or date.today().isoformat()
                with db() as conn:
                    rows = conn.execute(
                        "SELECT * FROM atividades WHERE data BETWEEN ? AND ? ORDER BY data, id", (ini, fim),
                    ).fetchall()
                return self.send(200, [dict(r) for r in rows])

            if url.path == "/api/perfis":
                return self.send(200, {"perfis": [perfil_completo(i) for i in ids_perfis()],
                                       "principal": perfil_principal()})

            rota = self.rota_perfil(url.path)
            if rota:
                pid, tipo, nome = rota
                if not perfil_existe(pid):
                    return self.send(404, {"erro": "Modelo não encontrado."})
                if not tipo:
                    return self.send(200, perfil_completo(pid))
                if not nome:
                    return self.send(404, {"erro": "Arquivo não encontrado."})
                with open(os.path.join(perfil_dir(pid), tipo, nome), "rb") as f:
                    return self.send(200, f.read(), FORMATOS[nome.rsplit(".", 1)[1].lower()],
                                     {"Content-Disposition": f"attachment; filename*=UTF-8''{quote(nome)}"})

            m = re.fullmatch(r"/static/([\w.-]+\.js)", url.path)  # bibliotecas da pré-visualização (sem internet)
            if m and os.path.isfile(os.path.join(BASE_DIR, "static", m.group(1))):
                with open(os.path.join(BASE_DIR, "static", m.group(1)), "rb") as f:
                    return self.send(200, f.read(), "text/javascript; charset=utf-8")

            if url.path == "/api/versao":
                return self.send(200, {"versao": VERSAO})

            if url.path in ("/api/ia", "/api/relatorio/config"):
                return self.send(200, ia_publica())

            if url.path == "/api/relatorio/rascunho":
                path = rascunho_path(qs.get("mes", ""), qs.get("projeto", ""))
                md = open(path, encoding="utf-8").read() if os.path.exists(path) else ""
                return self.send(200, {"markdown": md})

            if url.path == "/api/sugestoes":
                with db() as conn:
                    out = {}
                    for f in ("projeto", "categoria"):
                        out[f] = [r[0] for r in conn.execute(
                            f"SELECT {f} FROM atividades WHERE {f} <> '' GROUP BY {f} ORDER BY MAX(id) DESC LIMIT 50")]
                return self.send(200, out)

            if url.path == "/api/exportar.csv":
                mes = qs.get("mes", "")
                ini, fim = month_range(mes)
                sql = ("SELECT data, projeto, categoria, ticket, descricao, horas, entregavel "
                       "FROM atividades WHERE data BETWEEN ? AND ?")
                params = [ini, fim]
                if qs.get("projeto"):
                    sql += " AND projeto = ?"
                    params.append(qs["projeto"])
                with db() as conn:
                    rows = conn.execute(sql + " ORDER BY data, id", params).fetchall()
                buf = io.StringIO()
                w = csv.writer(buf, delimiter=";")
                w.writerow(["Data", "Projeto", "Categoria", "Ticket", "Descrição", "Horas", "Entrega"])
                for r in rows:
                    d = "/".join(reversed(r["data"].split("-")))
                    w.writerow([d, r["projeto"], r["categoria"], r["ticket"], r["descricao"],
                                str(r["horas"]).replace(".", ","), "Sim" if r["entregavel"] else ""])
                # BOM para o Excel abrir com acentos corretos
                return self.send(
                    200, "﻿" + buf.getvalue(), "text/csv; charset=utf-8",
                    {"Content-Disposition": f'attachment; filename="atividades-{mes}.csv"'},
                )

            return self.send(404, {"erro": "não encontrado"})
        except ValueError as e:
            return self.send(400, {"erro": str(e)})

    # ---------- POST ----------
    def do_POST(self):
        path = urlparse(self.path).path
        try:
            rota = self.rota_perfil(path)
            if rota and rota[1]:  # upload de arquivo para um modelo
                length = int(self.headers.get("Content-Length") or 0)
                if length > MODELO_MAX_BYTES:
                    return self.send(400, {"erro": "Arquivo muito grande (máx. 20 MB)."})
                nome = parse_qs(urlparse(self.path).query).get("nome", [""])[0]
                return self.send(200, salvar_arquivo(rota[0], rota[1], nome, self.rfile.read(length)))
            if path == "/api/perfis":
                body = self.read_json()
                return self.send(201, perfil_completo(criar_perfil(body.get("nome"), body.get("copiar_de"))["id"]))
            if path == "/api/markdown":  # pré-visualização: mesmo parser e estilos do .docx
                body = self.read_json()
                css = ""
                pid = body.get("perfil")
                if pid and perfil_existe(pid):
                    padrao = arquivos(pid, "padrao")
                    if padrao:
                        try:
                            css = relatorio.visual_do_template(os.path.join(perfil_dir(pid), "padrao", padrao[0]))
                        except (KeyError, zipfile.BadZipFile, OSError):
                            css = ""
                return self.send(200, {"html": relatorio.markdown_para_html(str(body.get("markdown") or "")), "css": css})
            if path == "/api/ia/testar":
                resposta = relatorio.chamar_llm(ler_ia(), "Responda só com a palavra: ok",
                                                [{"tipo": "texto", "texto": "Teste de conexão."}], max_tokens=20)
                return self.send(200, {"ok": True, "resposta": resposta[:200]})
        except (ValueError, json.JSONDecodeError, RuntimeError) as e:
            return self.send(400, {"erro": str(e)})
        if path == "/api/relatorio/rascunho":
            return self.gerar_rascunho()
        if path == "/api/relatorio/rascunho-beacon":  # salvamento ao fechar a aba (navigator.sendBeacon)
            try:
                body = self.read_json()
                destino = rascunho_path(body.get("mes", ""), body.get("projeto", ""))
                os.makedirs(RASCUNHOS_DIR, exist_ok=True)
                with open(destino, "w", encoding="utf-8") as f:
                    f.write(str(body.get("markdown") or ""))
                return self.send(200, {"ok": True})
            except (ValueError, json.JSONDecodeError) as e:
                return self.send(400, {"erro": str(e)})
        if path == "/api/relatorio/docx":
            return self.baixar_docx()
        if path != "/api/atividades":
            return self.send(404, {"erro": "não encontrado"})
        try:
            payload = self.read_json()
            lote = isinstance(payload, list)  # a skill do Claude pode mandar várias de uma vez
            itens = [clean(p) for p in (payload if lote else [payload])]
            if not itens:
                raise ValueError("Nenhuma atividade enviada.")
        except (ValueError, json.JSONDecodeError, AttributeError) as e:
            return self.send(400, {"erro": str(e)})
        with db() as conn:
            ids = [conn.execute(
                f"INSERT INTO atividades ({', '.join(FIELDS)}) VALUES ({', '.join('?' * len(FIELDS))})",
                [item[f] for f in FIELDS],
            ).lastrowid for item in itens]
            rows = [dict(conn.execute("SELECT * FROM atividades WHERE id = ?", (i,)).fetchone()) for i in ids]
        return self.send(201, rows if lote else rows[0])

    # ---------- PUT ----------
    def do_PUT(self):
        path = urlparse(self.path).path
        try:
            rota = self.rota_perfil(path)
            if rota and not rota[1]:
                if not perfil_existe(rota[0]):
                    return self.send(404, {"erro": "Modelo não encontrado."})
                gravar_perfil(rota[0], self.read_json())
                return self.send(200, perfil_completo(rota[0]))
            if path == "/api/ia":
                return self.send(200, salvar_ia(self.read_json()))
            if path == "/api/relatorio/rascunho":
                body = self.read_json()
                destino = rascunho_path(body.get("mes", ""), body.get("projeto", ""))
                os.makedirs(RASCUNHOS_DIR, exist_ok=True)
                with open(destino, "w", encoding="utf-8") as f:
                    f.write(str(body.get("markdown") or ""))
                return self.send(200, {"ok": True})
        except (ValueError, json.JSONDecodeError) as e:
            return self.send(400, {"erro": str(e)})
        eid = self.entry_id(path)
        if eid is None:
            return self.send(404, {"erro": "não encontrado"})
        try:
            item = clean(self.read_json())
        except (ValueError, json.JSONDecodeError) as e:
            return self.send(400, {"erro": str(e)})
        with db() as conn:
            conn.execute(
                f"UPDATE atividades SET {', '.join(f + ' = ?' for f in FIELDS)} WHERE id = ?",
                [item[f] for f in FIELDS] + [eid],
            )
            row = conn.execute("SELECT * FROM atividades WHERE id = ?", (eid,)).fetchone()
        return self.send(200, dict(row)) if row else self.send(404, {"erro": "não encontrado"})

    # ---------- DELETE ----------
    def do_DELETE(self):
        path = urlparse(self.path).path
        try:
            rota = self.rota_perfil(path)
            if rota:
                pid, tipo, nome = rota
                if not tipo:
                    excluir_perfil(pid)
                    return self.send(200, {"ok": True})
                if not nome:
                    return self.send(404, {"erro": "Arquivo não encontrado."})
                os.remove(os.path.join(perfil_dir(pid), tipo, nome))
                return self.send(200, {"ok": True})
        except ValueError as e:
            return self.send(400, {"erro": str(e)})
        eid = self.entry_id(path)
        if eid is None:
            return self.send(404, {"erro": "não encontrado"})
        with db() as conn:
            conn.execute("DELETE FROM atividades WHERE id = ?", (eid,))
        return self.send(200, {"ok": True})


if __name__ == "__main__":
    init_db()
    migrar_modelos()
    print(f"Diário de Atividades rodando em http://{HOST}:{PORT}")
    ThreadingHTTPServer((HOST, PORT), Handler).serve_forever()
