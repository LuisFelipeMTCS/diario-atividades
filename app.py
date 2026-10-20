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
import sqlite3
import zipfile
from datetime import date, datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, quote, unquote, urlparse

import relatorio

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
DB_PATH = os.path.join(DATA_DIR, "atividades.db")
MODELOS_DIR = os.path.join(BASE_DIR, "modelos")
# pasta -> quantos arquivos guarda (o padrão é um só; referências podem ser várias)
MODELOS = {"padrao": 1, "referencia": 10}
MODELO_MAX_BYTES = 20 * 1024 * 1024
FORMATOS = {  # extensão -> tipo MIME
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "dotx": "application/vnd.openxmlformats-officedocument.wordprocessingml.template",
    "pdf": "application/pdf",
}
CONFIG_PATH = os.path.join(DATA_DIR, "config.json")
RASCUNHOS_DIR = os.path.join(DATA_DIR, "rascunhos")
relatorio.carregar_env(os.path.join(BASE_DIR, ".env"))
HOST = os.environ.get("DIARIO_HOST", "127.0.0.1")
PORT = int(os.environ.get("DIARIO_PORT", "8765"))

FIELDS = ("data", "projeto", "categoria", "ticket", "descricao", "horas")


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
    return item


def month_range(month):
    if not re.fullmatch(r"\d{4}-\d{2}", month or ""):
        raise ValueError("Mês inválido (use AAAA-MM).")
    return f"{month}-01", f"{month}-31"


def modelo_dir(tipo):
    return os.path.join(MODELOS_DIR, tipo)


def modelo_arquivos(tipo):
    """Arquivos aceitos em modelos/<tipo>/, do mais recente para o mais antigo."""
    pasta = modelo_dir(tipo)
    if not os.path.isdir(pasta):
        return []
    nomes = [n for n in os.listdir(pasta) if n.rsplit(".", 1)[-1].lower() in FORMATOS]
    return sorted(nomes, key=lambda n: os.stat(os.path.join(pasta, n)).st_mtime, reverse=True)


def modelo_info(tipo, nome):
    st = os.stat(os.path.join(modelo_dir(tipo), nome))
    return {
        "nome": nome,
        "formato": nome.rsplit(".", 1)[1].lower(),
        "tamanho": st.st_size,
        "enviado_em": datetime.fromtimestamp(st.st_mtime).strftime("%d/%m/%Y %H:%M"),
    }


def listar_modelos():
    out = {}
    for tipo, limite in MODELOS.items():
        infos = [modelo_info(tipo, n) for n in modelo_arquivos(tipo)]
        out[tipo] = (infos[0] if infos else None) if limite == 1 else infos
    return out


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


def salvar_modelo(tipo, nome, conteudo):
    """Valida o arquivo e grava em modelos/<tipo>/<nome>. No padrão, substitui o anterior."""
    if not conteudo:
        raise ValueError("Arquivo vazio.")
    if len(conteudo) > MODELO_MAX_BYTES:
        raise ValueError("Arquivo muito grande (máx. 20 MB).")
    formato = detectar_formato(conteudo)
    if not formato:
        raise ValueError("Formato não aceito. Envie .docx, .dotx ou .pdf.")
    nome = nome_seguro(nome, formato, tipo)
    existentes = modelo_arquivos(tipo)
    limite = MODELOS[tipo]
    if limite > 1 and nome not in existentes and len(existentes) >= limite:
        raise ValueError(f"Limite de {limite} arquivos. Remova algum antes de enviar outro.")
    pasta = modelo_dir(tipo)
    os.makedirs(pasta, exist_ok=True)
    destino = os.path.join(pasta, nome)
    with open(destino + ".tmp", "wb") as f:
        f.write(conteudo)
    os.replace(destino + ".tmp", destino)
    if limite == 1:
        for antigo in existentes:
            if antigo != nome:
                os.remove(os.path.join(pasta, antigo))
    return modelo_info(tipo, nome)


# ---------- relatório .docx ----------
CAPA_CAMPOS = ("titulo", "autor", "area", "subtitulo", "versao")


def ler_config():
    """Dados da capa. Na primeira vez, sugere os da referência mais recente (.docx)."""
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, encoding="utf-8") as f:
            cfg = json.load(f)
    else:
        cfg = {"subtitulo": "Relatório de Atividades | {mes}", "versao": "1.0"}
        for nome in modelo_arquivos("referencia"):
            if nome.lower().endswith((".docx", ".dotx")):
                capa = relatorio.capa_do_docx(os.path.join(modelo_dir("referencia"), nome))
                cfg.update({k: capa[k] for k in ("titulo", "autor", "area") if capa.get(k)})
                break
    return {k: str(cfg.get(k, "")) for k in CAPA_CAMPOS}


def salvar_config(payload):
    cfg = {k: str(payload.get(k) or "").strip() for k in CAPA_CAMPOS}
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    return cfg


def rascunho_path(mes, projeto):
    month_range(mes)
    sufixo = "-" + re.sub(r"[^\w-]+", "_", projeto).strip("_") if projeto else ""
    return os.path.join(RASCUNHOS_DIR, f"{mes}{sufixo}.md")


def atividades_do_mes(mes, projeto=""):
    ini, fim = month_range(mes)
    sql = "SELECT * FROM atividades WHERE data BETWEEN ? AND ?"
    params = [ini, fim]
    if projeto:
        sql += " AND projeto = ?"
        params.append(projeto)
    with db() as conn:
        return [dict(r) for r in conn.execute(sql + " ORDER BY data, id", params)]


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):  # silencia o log padrão
        pass

    # ---------- respostas ----------
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

    def modelo_rota(self, path):
        """/api/modelos/<tipo>[/<nome>] -> (tipo, nome existente ou None); None se não casar."""
        m = re.fullmatch(r"/api/modelos/(\w+)(?:/([^/]+))?", path)
        if not m or m.group(1) not in MODELOS:
            return None
        tipo, nome = m.group(1), unquote(m.group(2)) if m.group(2) else None
        if nome is not None and nome not in modelo_arquivos(tipo):
            nome = ""  # pediu um arquivo que não existe
        return tipo, nome

    def gerar_rascunho(self):
        try:
            body = self.read_json()
            mes, projeto = body.get("mes", ""), body.get("projeto", "")
            destino = rascunho_path(mes, projeto)
            refs = [os.path.join(modelo_dir("referencia"), n) for n in modelo_arquivos("referencia")]
            md = relatorio.gerar_rascunho(atividades_do_mes(mes, projeto), mes, projeto, refs)
        except (ValueError, json.JSONDecodeError, RuntimeError) as e:
            return self.send(400, {"erro": str(e)})
        os.makedirs(RASCUNHOS_DIR, exist_ok=True)
        with open(destino, "w", encoding="utf-8") as f:
            f.write(md)
        return self.send(200, {"markdown": md})

    def baixar_docx(self):
        try:
            body = self.read_json()
            mes, projeto = body.get("mes", ""), body.get("projeto", "")
            month_range(mes)
            padrao = modelo_arquivos("padrao")
            if not padrao or not padrao[0].lower().endswith((".docx", ".dotx")):
                raise ValueError("Envie o padrão da empresa em .docx ou .dotx (aba Modelos).")
            capa = ler_config()
            capa["subtitulo"] = capa["subtitulo"].replace("{mes}", relatorio.nome_mes(mes))
            conteudo = relatorio.montar_docx(os.path.join(modelo_dir("padrao"), padrao[0]),
                                             str(body.get("markdown") or ""), capa)
        except (ValueError, json.JSONDecodeError, RuntimeError) as e:
            return self.send(400, {"erro": str(e)})
        nome = f"Relatório de Atividades - {relatorio.nome_mes(mes)}{' - ' + projeto if projeto else ''}.docx"
        return self.send(200, conteudo, FORMATOS["docx"],
                         {"Content-Disposition": f"attachment; filename*=UTF-8''{quote(nome)}"})

    # ---------- rotas ----------
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
                        "SELECT * FROM atividades WHERE data BETWEEN ? AND ? ORDER BY data, id",
                        (ini, fim),
                    ).fetchall()
                return self.send(200, [dict(r) for r in rows])

            if url.path == "/api/modelos":
                return self.send(200, listar_modelos())

            rota = self.modelo_rota(url.path)
            if rota:
                tipo, nome = rota
                if not nome:
                    return self.send(404, {"erro": "Arquivo não encontrado."})
                with open(os.path.join(modelo_dir(tipo), nome), "rb") as f:
                    return self.send(200, f.read(), FORMATOS[nome.rsplit(".", 1)[1].lower()],
                                     {"Content-Disposition": f"attachment; filename*=UTF-8''{quote(nome)}"})

            if url.path == "/api/relatorio/config":
                cfg = ler_config()
                cfg["ia"] = bool(os.environ.get("OPENAI_API_KEY"))
                cfg["modelo"] = os.environ.get("OPENAI_MODEL") or relatorio.MODELO_PADRAO
                return self.send(200, cfg)

            if url.path == "/api/relatorio/rascunho":
                path = rascunho_path(qs.get("mes", ""), qs.get("projeto", ""))
                md = open(path, encoding="utf-8").read() if os.path.exists(path) else ""
                return self.send(200, {"markdown": md})

            if url.path == "/api/sugestoes":
                with db() as conn:
                    out = {}
                    for f in ("projeto", "categoria"):
                        out[f] = [
                            r[0]
                            for r in conn.execute(
                                f"SELECT {f} FROM atividades WHERE {f} <> '' "
                                f"GROUP BY {f} ORDER BY MAX(id) DESC LIMIT 50"
                            )
                        ]
                return self.send(200, out)

            if url.path == "/api/exportar.csv":
                mes = qs.get("mes", "")
                ini, fim = month_range(mes)
                sql = ("SELECT data, projeto, categoria, ticket, descricao, horas "
                       "FROM atividades WHERE data BETWEEN ? AND ?")
                params = [ini, fim]
                if qs.get("projeto"):
                    sql += " AND projeto = ?"
                    params.append(qs["projeto"])
                with db() as conn:
                    rows = conn.execute(sql + " ORDER BY data, id", params).fetchall()
                buf = io.StringIO()
                w = csv.writer(buf, delimiter=";")
                w.writerow(["Data", "Projeto", "Categoria", "Ticket", "Descrição", "Horas"])
                for r in rows:
                    d = "/".join(reversed(r["data"].split("-")))
                    w.writerow([d, r["projeto"], r["categoria"], r["ticket"], r["descricao"],
                                str(r["horas"]).replace(".", ",")])
                # BOM para o Excel abrir com acentos corretos
                return self.send(
                    200, "﻿" + buf.getvalue(), "text/csv; charset=utf-8",
                    {"Content-Disposition": f'attachment; filename="atividades-{mes}.csv"'},
                )

            return self.send(404, {"erro": "não encontrado"})
        except ValueError as e:
            return self.send(400, {"erro": str(e)})

    def do_POST(self):
        path = urlparse(self.path).path
        rota = self.modelo_rota(path)
        if rota:
            length = int(self.headers.get("Content-Length") or 0)
            if length > MODELO_MAX_BYTES:
                return self.send(400, {"erro": "Arquivo muito grande (máx. 20 MB)."})
            nome = parse_qs(urlparse(self.path).query).get("nome", [""])[0]
            try:
                return self.send(200, salvar_modelo(rota[0], nome, self.rfile.read(length)))
            except ValueError as e:
                return self.send(400, {"erro": str(e)})
        if path == "/api/relatorio/rascunho":
            return self.gerar_rascunho()
        if path == "/api/relatorio/docx":
            return self.baixar_docx()
        if path != "/api/atividades":
            return self.send(404, {"erro": "não encontrado"})
        try:
            item = clean(self.read_json())
        except (ValueError, json.JSONDecodeError) as e:
            return self.send(400, {"erro": str(e)})
        with db() as conn:
            cur = conn.execute(
                f"INSERT INTO atividades ({', '.join(FIELDS)}) VALUES ({', '.join('?' * len(FIELDS))})",
                [item[f] for f in FIELDS],
            )
            row = conn.execute("SELECT * FROM atividades WHERE id = ?", (cur.lastrowid,)).fetchone()
        return self.send(201, dict(row))

    def do_PUT(self):
        path = urlparse(self.path).path
        try:
            if path == "/api/relatorio/config":
                return self.send(200, salvar_config(self.read_json()))
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

    def do_DELETE(self):
        path = urlparse(self.path).path
        rota = self.modelo_rota(path)
        if rota:
            tipo, nome = rota
            if not nome:
                return self.send(404, {"erro": "Arquivo não encontrado."})
            os.remove(os.path.join(modelo_dir(tipo), nome))
            return self.send(200, {"ok": True})
        eid = self.entry_id(path)
        if eid is None:
            return self.send(404, {"erro": "não encontrado"})
        with db() as conn:
            conn.execute("DELETE FROM atividades WHERE id = ?", (eid,))
        return self.send(200, {"ok": True})


if __name__ == "__main__":
    init_db()
    print(f"Diário de Atividades rodando em http://{HOST}:{PORT}")
    ThreadingHTTPServer((HOST, PORT), Handler).serve_forever()
