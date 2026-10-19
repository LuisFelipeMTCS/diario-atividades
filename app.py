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
from datetime import date
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
DB_PATH = os.path.join(DATA_DIR, "atividades.db")
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
        if urlparse(self.path).path != "/api/atividades":
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
        eid = self.entry_id(urlparse(self.path).path)
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
        eid = self.entry_id(urlparse(self.path).path)
        if eid is None:
            return self.send(404, {"erro": "não encontrado"})
        with db() as conn:
            conn.execute("DELETE FROM atividades WHERE id = ?", (eid,))
        return self.send(200, {"ok": True})


if __name__ == "__main__":
    init_db()
    print(f"Diário de Atividades rodando em http://{HOST}:{PORT}")
    ThreadingHTTPServer((HOST, PORT), Handler).serve_forever()
