from __future__ import annotations

import csv
import hashlib
import io
import os
import re
import sqlite3
import time
import traceback
import unicodedata
import zipfile
from collections import defaultdict
from datetime import datetime
from functools import wraps
from pathlib import Path
from typing import Any
from xml.etree.ElementTree import Element, SubElement, tostring

import pandas as pd
from flask import Flask, flash, redirect, render_template, request, send_file, session, url_for
from openpyxl import load_workbook
from werkzeug.utils import secure_filename
from jinja2 import DictLoader

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = Path(os.getenv("DATABASE_PATH", BASE_DIR / "fleetflow.db"))
UPLOAD_DIR = BASE_DIR / "uploads"
REPORT_DIR = BASE_DIR / "reports"
SEED_XLSX = BASE_DIR / "BANCO_DE_DADOS.xlsx"
if not SEED_XLSX.exists():
    SEED_XLSX = BASE_DIR / "data" / "BANCO_DE_DADOS.xlsx"
UPLOAD_DIR.mkdir(exist_ok=True)
REPORT_DIR.mkdir(exist_ok=True)

app = Flask(__name__)
app.jinja_loader = DictLoader({'closing_detail.html': '{% extends \'base.html\' %}{% block title %}{{closing.label}}{% endblock %}{% block subtitle %}{{closing.status}} · Confira antes de enviar o XML ao banco.{% endblock %}{% block content %}<div class="actions"><a class="button" href="{{url_for(\'closing_excel\',closing_id=closing.id)}}">Baixar Excel</a><a class="button" href="{{url_for(\'closing_xml\',closing_id=closing.id)}}">Baixar XML SEPA</a></div><section class="panel table-wrap"><table><thead><tr><th>Motorista</th><th>Origem</th><th>IBAN</th><th>Bruto</th><th>Dinheiro</th><th>Comissão</th><th>Combustível</th><th>Desconto</th><th>Reembolso</th><th>Taxa</th><th>Líquido</th><th>Total IBAN</th></tr></thead><tbody>{% for i in items %}<tr class="{{\'problem\' if not i.iban or i.net_before_group<0}}"><td>{{i.driver_name}}</td><td>{{i.origins}}</td><td>{{i.iban or \'SEM IBAN\'}}</td><td>€ {{\'%.2f\'|format(i.gross)}}</td><td>€ {{\'%.2f\'|format(i.cash)}}</td><td>€ {{\'%.2f\'|format(i.commission)}}</td><td>€ {{\'%.2f\'|format(i.fuel)}}</td><td>€ {{\'%.2f\'|format(i.discount)}}</td><td>€ {{\'%.2f\'|format(i.reimbursement)}}</td><td>€ {{\'%.2f\'|format(i.bank_fee)}}</td><td>€ {{\'%.2f\'|format(i.net_before_group-i.bank_fee)}}</td><td><b>€ {{\'%.2f\'|format(i.group_total)}}</b></td></tr>{% endfor %}</tbody></table></section>{% endblock %}\n', 'base.html': '<!doctype html><html lang="pt"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{{ company_name }} · FleetFlow</title><script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"></script><style>\n:root{--bg:#f4f7fb;--card:#fff;--text:#172033;--muted:#718096;--accent:#315efb;--dark:#101828;--border:#e5e9f2;--danger:#d92d20}*{box-sizing:border-box}body{margin:0;font-family:Inter,ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;background:var(--bg);color:var(--text)}.layout{display:grid;grid-template-columns:235px 1fr;min-height:100vh}aside{background:var(--dark);color:#fff;padding:24px 16px}.brand{display:flex;align-items:center;gap:12px;margin:0 8px 28px}.brand>span,.logo-big{display:grid;place-items:center;background:var(--accent);color:#fff;border-radius:13px;font-weight:800}.brand>span{width:42px;height:42px}.brand small{display:block;color:#98a2b3;margin-top:3px}nav{display:grid;gap:6px}nav a{color:#d0d5dd;text-decoration:none;padding:12px 14px;border-radius:10px}nav a:hover{background:#1d2939;color:#fff}main{padding:30px;min-width:0}header h1{margin:0;font-size:28px}header p{margin:6px 0 24px;color:var(--muted)}.metrics{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px;margin-bottom:18px}.metric,.panel{background:var(--card);border:1px solid var(--border);border-radius:16px}.metric{padding:18px}.metric span{display:block;color:var(--muted);font-size:13px}.metric b{display:block;margin-top:8px;font-size:24px}.metric.alert b{color:var(--danger)}.panel{padding:20px;margin-bottom:18px}.panel h2{margin:0 0 14px;font-size:18px}.grid2{display:grid;grid-template-columns:1fr 1fr;gap:18px}.toolbar,.actions{display:flex;gap:10px;align-items:center;margin-bottom:18px}.toolbar input{flex:1}input,select,textarea{width:100%;border:1px solid #d0d5dd;border-radius:10px;padding:11px 12px;font:inherit;background:#fff}textarea{min-height:100px}button,.button{border:0;border-radius:10px;background:var(--accent);color:#fff;padding:11px 16px;font-weight:700;text-decoration:none;cursor:pointer;display:inline-block}.secondary{color:var(--accent);text-decoration:none;padding:10px}.danger{background:var(--danger)}label{display:grid;gap:7px;font-size:13px;font-weight:600}.form-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px}.full{grid-column:1/-1}.table-wrap{overflow:auto}table{width:100%;border-collapse:collapse;font-size:13px}th,td{text-align:left;padding:11px 10px;border-bottom:1px solid var(--border);white-space:nowrap}th{color:var(--muted);font-size:12px;text-transform:uppercase}.link{color:var(--accent);font-weight:700;text-decoration:none}.badge{display:inline-block;border-radius:999px;padding:4px 8px;font-size:11px;font-weight:800;background:#edf2f7}.badge.azul{background:#dbeafe;color:#1d4ed8}.badge.yellow{background:#fef3c7;color:#92400e}.flash{padding:12px 14px;border-radius:10px;margin-bottom:16px;background:#eaf2ff}.flash.danger{background:#fee4e2;color:#b42318}.flash.warning{background:#fef0c7;color:#93370d}.flash.success{background:#dcfae6;color:#067647}.note{color:var(--muted)}tr.problem{background:#fff4ed}.login-body{min-height:100vh;display:grid;place-items:center;background:#101828}.login-card{width:min(410px,calc(100% - 32px));background:#fff;border-radius:20px;padding:32px;display:grid;gap:16px}.login-card h1,.login-card p{margin:0;text-align:center}.login-card p,.login-card small{color:var(--muted);text-align:center}.logo-big{width:58px;height:58px;margin:auto;font-size:20px}@media(max-width:950px){.layout{grid-template-columns:1fr}aside{padding:14px}.brand{margin-bottom:12px}nav{display:flex;overflow:auto}.metrics{grid-template-columns:repeat(2,1fr)}.grid2,.form-grid{grid-template-columns:1fr}main{padding:18px}}@media(max-width:520px){.metrics{grid-template-columns:1fr}.toolbar,.actions{align-items:stretch;flex-direction:column}.toolbar input,.toolbar button,.actions a{width:100%}}\n\n</style></head><body>\n<div class="layout"><aside><div class="brand"><span>FF</span><div><b>FleetFlow</b><small>{{ company_name }}</small></div></div><nav><a href="{{ url_for(\'dashboard\') }}">Dashboard</a><a href="{{ url_for(\'drivers\') }}">Motoristas</a><a href="{{ url_for(\'database_page\') }}">Banco de dados</a><a href="{{ url_for(\'imports_page\') }}">Importações</a><a href="{{ url_for(\'closings_page\') }}">Fechamentos</a><a href="{{ url_for(\'logout\') }}">Sair</a></nav></aside><main><header><div><h1>{% block title %}{% endblock %}</h1><p>{% block subtitle %}{% endblock %}</p></div></header>{% for cat,msg in get_flashed_messages(with_categories=true) %}<div class="flash {{cat}}">{{msg}}</div>{% endfor %}{% block content %}{% endblock %}</main></div></body></html>\n', 'closings.html': '{% extends \'base.html\' %}{% block title %}Fechamentos{% endblock %}{% block subtitle %}Processe depois de atualizar o banco e importar todos os relatórios da semana.{% endblock %}{% block content %}<section class="panel"><h2>Novo fechamento</h2><form method="post" class="toolbar"><input name="label" placeholder="Ex.: Semana 20/07/2026"><button>Processar semana</button></form><p class="note">A taxa de € 1,25 é aplicada uma única vez por IBAN AZUL, após a consolidação.</p></section><section class="panel"><table><thead><tr><th>ID</th><th>Nome</th><th>Data</th><th>Status</th><th></th></tr></thead><tbody>{% for r in rows %}<tr><td>#{{r.id}}</td><td>{{r.label}}</td><td>{{r.created_at}}</td><td>{{r.status}}</td><td><a class="link" href="{{url_for(\'closing_detail\',closing_id=r.id)}}">Conferir</a></td></tr>{% else %}<tr><td colspan="5">Nenhum fechamento.</td></tr>{% endfor %}</tbody></table></section>{% endblock %}\n', 'login.html': '<!doctype html><html lang="pt"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Entrar · FleetFlow</title><link rel="stylesheet" href="{{ url_for(\'static\', filename=\'style.css\') }}"></head><body class="login-body"><form class="login-card" method="post"><div class="logo-big">FF</div><h1>FleetFlow</h1><p>{{ company_name }}</p>{% for cat,msg in get_flashed_messages(with_categories=true) %}<div class="flash {{cat}}">{{msg}}</div>{% endfor %}<label>Utilizador<input name="username" required autofocus></label><label>Senha<input name="password" type="password" required></label><button>Entrar</button><small>Configure ADMIN_USER e ADMIN_PASSWORD no Render.</small></form></body></html>\n', 'database.html': '{% extends \'base.html\' %}{% block title %}Banco de dados{% endblock %}{% block subtitle %}Base online dos motoristas e mapa Arquivos_Pgto.{% endblock %}{% block content %}<div class="metrics"><div class="metric"><span>Motoristas</span><b>{{n}}</b></div><div class="metric"><span>Mapas de leitura</span><b>{{m}}</b></div></div><section class="panel"><h2>Atualizar banco completo</h2><p>Envie um XLSX com as abas <b>Banco_de_Dados</b> e <b>Arquivos_Pgto</b>. A importação substitui o cadastro atual.</p><form method="post" enctype="multipart/form-data"><input type="file" name="database_file" accept=".xlsx" required><button>Importar banco atualizado</button></form></section><section class="panel"><h2>Alterações rápidas</h2><p>Para alterar somente IBAN, banco AZUL/YELLOW, porcentagem, desconto, parceiro ou cartão, use a página <a class="link" href="{{url_for(\'drivers\')}}">Motoristas</a>.</p></section>{% endblock %}\n', 'drivers.html': '{% extends \'base.html\' %}{% block title %}Motoristas{% endblock %}{% block subtitle %}{{total}} cadastros. A tabela mostra até 500 resultados por pesquisa.{% endblock %}{% block content %}<form class="toolbar"><input name="q" value="{{q}}" placeholder="Nome, ID, IBAN, cidade ou parceiro"><button>Pesquisar</button></form><section class="panel table-wrap"><table><thead><tr><th>Motorista</th><th>ID</th><th>Cidade</th><th>Companhia</th><th>IBAN</th><th>Banco</th><th>%</th><th>Parceiro</th><th></th></tr></thead><tbody>{% for d in rows %}<tr><td><b>{{d.name}}</b></td><td>{{d.external_id}}</td><td>{{d.city}}</td><td>{{d.company}}</td><td>{{d.iban}}</td><td><span class="badge {{d.bank_color|lower}}">{{d.bank_color}}</span></td><td>{{\'%.2f\'|format(d.percentage*100)}}%</td><td>{{d.partner}}</td><td><a class="link" href="{{url_for(\'edit_driver\',driver_id=d.id)}}">Editar</a></td></tr>{% endfor %}</tbody></table></section>{% endblock %}\n', 'imports.html': '{% extends \'base.html\' %}{% block title %}Importações da semana{% endblock %}{% block subtitle %}Envie vários CSVs de uma vez e, separadamente, os ficheiros PRIO em XLSX.{% endblock %}{% block content %}<div class="grid2"><section class="panel"><h2>Relatórios das plataformas</h2><form method="post" enctype="multipart/form-data"><input type="hidden" name="kind" value="reports"><input type="file" name="files" accept=".csv" multiple required><button>Importar relatórios</button></form></section><section class="panel"><h2>Cartões combustível</h2><form method="post" enctype="multipart/form-data"><input type="hidden" name="kind" value="fuel"><input type="file" name="files" accept=".xlsx,.xls" multiple required><button>Importar PRIO</button></form></section></div><form method="post" action="{{url_for(\'clear_imports\')}}" onsubmit="return confirm(\'Limpar todos os ficheiros da semana?\')"><button class="danger">Limpar semana importada</button></form><section class="panel table-wrap"><h2>Histórico</h2><table><thead><tr><th>Ficheiro</th><th>Tipo</th><th>Status</th><th>Linhas</th><th>Mensagem</th><th>Data</th></tr></thead><tbody>{% for r in rows %}<tr><td>{{r.filename}}</td><td>{{r.kind}}</td><td><span class="badge {{\'azul\' if r.status==\'OK\' else \'yellow\'}}">{{r.status}}</span></td><td>{{r.rows_count}}</td><td>{{r.message}}</td><td>{{r.created_at}}</td></tr>{% else %}<tr><td colspan="6">Nenhum ficheiro importado.</td></tr>{% endfor %}</tbody></table></section>{% endblock %}\n', 'error.html': '{% extends \'base.html\' %}{% block title %}Erro no sistema{% endblock %}{% block subtitle %}Não foi possível concluir esta operação.{% endblock %}{% block content %}<section class="panel"><h2>Erro interno</h2><p>{{ message }}</p><p><a class="link" href="{{ url_for(\'login\') }}">Voltar ao login</a></p></section>{% endblock %}', 'dashboard.html': '{% extends \'base.html\' %}{% block title %}Dashboard semanal{% endblock %}{% block subtitle %}Visão geral do último fechamento processado.{% endblock %}{% block content %}\n<div class="metrics"><div class="metric"><span>Motoristas</span><b>{{drivers}}</b></div><div class="metric"><span>Ficheiros importados</span><b>{{imports}}</b></div><div class="metric"><span>Bruto</span><b>€ {{\'%.2f\'|format(stats.gross)}}</b></div><div class="metric"><span>Líquido</span><b>€ {{\'%.2f\'|format(stats.net)}}</b></div><div class="metric"><span>Comissão</span><b>€ {{\'%.2f\'|format(stats.commission)}}</b></div><div class="metric"><span>Combustível</span><b>€ {{\'%.2f\'|format(stats.fuel)}}</b></div><div class="metric"><span>Transferências</span><b>{{stats.payments}}</b></div><div class="metric alert"><span>Sem IBAN</span><b>{{stats.missing_iban}}</b></div></div>\n<div class="grid2"><section class="panel"><h2>Ganhos por origem</h2><canvas id="originChart"></canvas></section><section class="panel"><h2>Distribuição financeira</h2><canvas id="moneyChart"></canvas></section></div>\n<section class="panel"><h2>Fechamentos recentes</h2><table><thead><tr><th>ID</th><th>Semana</th><th>Data</th><th>Status</th><th></th></tr></thead><tbody>{% for r in recent %}<tr><td>#{{r.id}}</td><td>{{r.label}}</td><td>{{r.created_at}}</td><td>{{r.status}}</td><td><a class="link" href="{{url_for(\'closing_detail\',closing_id=r.id)}}">Abrir</a></td></tr>{% else %}<tr><td colspan="5">Nenhum fechamento processado.</td></tr>{% endfor %}</tbody></table></section>\n<script>new Chart(document.getElementById(\'originChart\'),{type:\'bar\',data:{labels:{{ origin_labels|tojson }},datasets:[{label:\'Bruto (€)\',data:{{ origin_values|tojson }}}]},options:{responsive:true}});new Chart(document.getElementById(\'moneyChart\'),{type:\'doughnut\',data:{labels:[\'Comissão\',\'Combustível\',\'Descontos\',\'Reembolsos\'],datasets:[{data:[{{stats.commission}},{{stats.fuel}},{{stats.discount}},{{stats.reimbursement}}]}]},options:{responsive:true}});</script>\n{% endblock %}\n', 'driver_edit.html': '{% extends \'base.html\' %}{% block title %}Editar motorista{% endblock %}{% block subtitle %}As alterações passam a valer no próximo processamento.{% endblock %}{% block content %}<form method="post" class="panel form-grid"><label>Nome<input name="name" value="{{d.name}}" required></label><label>ID / e-mail / telefone<input name="external_id" value="{{d.external_id}}"></label><label>Cidade/campanha<input name="city" value="{{d.city}}"></label><label>Companhia<input name="company" value="{{d.company}}"></label><label>IBAN<input name="iban" value="{{d.iban}}"></label><label>Banco<select name="bank_color"><option {{\'selected\' if d.bank_color==\'YELLOW\'}}>YELLOW</option><option {{\'selected\' if d.bank_color==\'AZUL\'}}>AZUL</option></select></label><label>Porcentagem (decimal)<input name="percentage" type="number" step="0.0001" value="{{d.percentage}}"></label><label>Comissão Muryllo<input name="commission_owner" type="number" step="0.01" value="{{d.commission_owner}}"></label><label>Parceiro<input name="partner" value="{{d.partner}}"></label><label>Comissão parceiro<input name="partner_commission" type="number" step="0.01" value="{{d.partner_commission}}"></label><label>Cartão combustível<input name="fuel_card" value="{{d.fuel_card}}"></label><label>Aluguel/categoria<input name="rental_label" value="{{d.rental_label}}"></label><label>Desconto<input name="discount" type="number" step="0.01" value="{{d.discount}}"></label><label>Reembolso<input name="reimbursement" type="number" step="0.01" value="{{d.reimbursement}}"></label><label>Imediata<input name="immediate" type="number" step="0.01" value="{{d.immediate}}"></label><label class="full">Observação<textarea name="observation">{{d.observation}}</textarea></label><div class="full actions"><button>Guardar alterações</button><a class="secondary" href="{{url_for(\'drivers\')}}">Voltar</a></div></form>{% endblock %}\n'})
app.secret_key = os.getenv("SECRET_KEY", "troque-esta-chave-em-producao")
app.config["MAX_CONTENT_LENGTH"] = 80 * 1024 * 1024
COMPANY_NAME = os.getenv("COMPANY_NAME", "Irmãos Fleet")
ADMIN_USER = os.getenv("ADMIN_USER", "admin")
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "admin123")


def db() -> sqlite3.Connection:
    con = sqlite3.connect(DB_PATH, timeout=30)
    con.execute("PRAGMA journal_mode=WAL")
    con.execute("PRAGMA busy_timeout=30000")
    con.row_factory = sqlite3.Row
    return con


def norm(v: Any) -> str:
    s = "" if v is None else str(v).strip()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", s).upper()


def clean_identifier(v: Any) -> str:
    return re.sub(r"[^A-Z0-9@.+_-]", "", norm(v))


def money(v: Any) -> float:
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("€", "").replace(" ", "")
    if not s or s in {"-", "NAN", "NONE"}:
        return 0.0
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def init_db() -> None:
    # O Render pode iniciar mais de um processo ao mesmo tempo.
    # As tentativas abaixo evitam falha temporária por bloqueio do SQLite.
    last_error = None
    for attempt in range(5):
        try:
            with db() as con:
                con.executescript("""
        CREATE TABLE IF NOT EXISTS drivers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            external_id TEXT,
            name TEXT NOT NULL,
            city TEXT,
            company TEXT,
            iban TEXT,
            bank_color TEXT,
            commission_owner REAL DEFAULT 0,
            partner TEXT,
            partner_commission REAL DEFAULT 0,
            percentage REAL DEFAULT 0,
            fuel_card TEXT,
            rental_label TEXT,
            discount REAL DEFAULT 0,
            reimbursement REAL DEFAULT 0,
            immediate REAL DEFAULT 0,
            observation TEXT,
            updated_at TEXT
        );
        CREATE INDEX IF NOT EXISTS idx_driver_external ON drivers(external_id);
        CREATE INDEX IF NOT EXISTS idx_driver_iban ON drivers(iban);
        CREATE TABLE IF NOT EXISTS mappings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name TEXT UNIQUE,
            origin_ref TEXT,
            identifier_column TEXT,
            first_name_column TEXT,
            last_name_column TEXT,
            value_column TEXT,
            cash_column TEXT,
            reimbursement_1 TEXT,
            reimbursement_2 TEXT,
            reimbursement_3 TEXT,
            reimbursement_4 TEXT
        );
        CREATE TABLE IF NOT EXISTS imports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT,
            sha256 TEXT UNIQUE,
            kind TEXT,
            status TEXT,
            rows_count INTEGER DEFAULT 0,
            message TEXT,
            created_at TEXT
        );
        CREATE TABLE IF NOT EXISTS raw_earnings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            import_id INTEGER,
            filename TEXT,
            origin_ref TEXT,
            identifier TEXT,
            display_name TEXT,
            gross REAL DEFAULT 0,
            cash REAL DEFAULT 0,
            reimbursement REAL DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS fuel (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            import_id INTEGER,
            card_number TEXT,
            amount REAL DEFAULT 0,
            source_file TEXT
        );
        CREATE TABLE IF NOT EXISTS closings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            label TEXT,
            created_at TEXT,
            status TEXT
        );
        CREATE TABLE IF NOT EXISTS closing_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            closing_id INTEGER,
            driver_id INTEGER,
            driver_name TEXT,
            iban TEXT,
            bank_color TEXT,
            origins TEXT,
            gross REAL,
            cash REAL,
            commission REAL,
            fuel REAL,
            discount REAL,
            reimbursement REAL,
            immediate REAL,
            bank_fee REAL,
            net_before_group REAL,
            group_total REAL
        );
        """)
            break
        except sqlite3.OperationalError as exc:
            last_error = exc
            time.sleep(1 + attempt)
    else:
        raise RuntimeError(f"Não foi possível inicializar o banco: {last_error}")

    if SEED_XLSX.exists():
        with db() as con:
            count = con.execute("SELECT COUNT(*) FROM drivers").fetchone()[0]
        if count == 0:
            try:
                import_master_workbook(SEED_XLSX)
            except sqlite3.IntegrityError:
                # Outro processo pode ter terminado a importação primeiro.
                pass


def import_master_workbook(path: Path) -> tuple[int, int]:
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb["Banco_de_Dados"]
    headers = [str(c.value or "").strip() for c in next(ws.iter_rows(min_row=1, max_row=1))]
    idx = {h: i for i, h in enumerate(headers)}
    drivers = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        name = row[idx.get("MOTORISTA_BD", 2)] if len(row) > 2 else None
        if not name:
            continue
        drivers.append((
            str(row[idx.get("ID", 1)] or ""), str(name), str(row[idx.get("Cidade", 4)] or ""),
            str(row[idx.get("Companhia", 5)] or ""), str(row[idx.get("IBAN", 6)] or ""),
            str(row[idx.get("Banco", 8)] or ""), money(row[idx.get("Comissao_Muryllo", 9)]),
            str(row[idx.get("Parceiro_2", 10)] or ""), money(row[idx.get("Comissao_2", 11)]),
            money(row[idx.get("Porcentagem", 12)]), str(row[idx.get("Numero_Cartao_Combustivel", 13)] or ""),
            str(row[idx.get("Aluguel", 15)] or ""), money(row[idx.get("Desconto_(-)", 16)]),
            money(row[idx.get("Reembolso_(+)", 17)]), money(row[idx.get("Imediata_(-)", 18)]),
            str(row[idx.get("Observação", 19)] or ""), datetime.now().isoformat(timespec="seconds")
        ))
    with db() as con:
        con.execute("DELETE FROM drivers")
        con.executemany("""INSERT INTO drivers(external_id,name,city,company,iban,bank_color,commission_owner,partner,partner_commission,percentage,fuel_card,rental_label,discount,reimbursement,immediate,observation,updated_at)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", drivers)

    ws2 = wb["Arquivos_Pgto"]
    mappings = []
    for r in ws2.iter_rows(min_row=2, values_only=True):
        if not r[0]:
            continue
        vals = list(r[:11]) + [None] * (11-len(r[:11]))
        mappings.append(tuple("" if v is None else str(v) for v in vals[:11]))
    with db() as con:
        con.execute("DELETE FROM mappings")
        con.executemany("""INSERT INTO mappings(file_name,origin_ref,identifier_column,first_name_column,last_name_column,value_column,cash_column,reimbursement_1,reimbursement_2,reimbursement_3,reimbursement_4)
        VALUES(?,?,?,?,?,?,?,?,?,?,?)""", mappings)
    return len(drivers), len(mappings)


def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get("user"):
            return redirect(url_for("login"))
        return fn(*args, **kwargs)
    return wrapper


def read_csv_flexible(path: Path) -> pd.DataFrame:
    last_error = None
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin1"):
        for sep in (None, ",", ";", "\t"):
            try:
                kwargs = {"encoding": enc, "dtype": str, "keep_default_na": False}
                if sep is None:
                    kwargs.update({"sep": None, "engine": "python"})
                else:
                    kwargs.update({"sep": sep})
                df = pd.read_csv(path, **kwargs)
                if len(df.columns) >= 2:
                    df.columns = [str(c).strip() for c in df.columns]
                    return df
            except Exception as exc:
                last_error = exc
    raise ValueError(f"Não foi possível ler o CSV: {last_error}")


def find_mapping(filename: str, columns: list[str]):
    stem = norm(Path(filename).stem)
    with db() as con:
        maps = con.execute("SELECT * FROM mappings ORDER BY LENGTH(file_name) DESC").fetchall()
    for m in maps:
        if norm(m["file_name"]) == stem or norm(m["file_name"]) in stem:
            return m
    # fallback only when exact file mapping is absent
    nc = {norm(c): c for c in columns}
    if "EMAIL" in nc and any("ADJUSTED EARNINGS" in k for k in nc):
        return {"origin_ref":"BOLT_FOOD","identifier_column":nc["EMAIL"],"first_name_column":nc.get("FIRST NAME",""),"last_name_column":nc.get("LAST NAME",""),"value_column":next(v for k,v in nc.items() if "ADJUSTED EARNINGS" in k and "VAT" in k),"cash_column":"-","reimbursement_1":"-","reimbursement_2":"-","reimbursement_3":"-","reimbursement_4":"-"}
    if "TELEMOVEL" in nc or "TELEMÓVEL" in columns:
        key = nc.get("TELEMOVEL", "Telemóvel")
        val = next((v for k,v in nc.items() if "PAGAMENTO PREVISTO" in k), "")
        return {"origin_ref":"BOLT_TVDE","identifier_column":key,"first_name_column":nc.get("MOTORISTA",""),"last_name_column":"-","value_column":val,"cash_column":"-","reimbursement_1":"-","reimbursement_2":"-","reimbursement_3":"-","reimbursement_4":"-"}
    if "UUID DO MOTORISTA" in nc:
        return {"origin_ref":"UBER EATS","identifier_column":nc["UUID DO MOTORISTA"],"first_name_column":nc.get("NOME PROPRIO DO MOTORISTA",""),"last_name_column":nc.get("APELIDO DO MOTORISTA",""),"value_column":nc.get("PAGO A SI",""),"cash_column":next((v for k,v in nc.items() if "DINHEIRO RECEBIDO" in k),"-"),"reimbursement_1":"-","reimbursement_2":"-","reimbursement_3":"-","reimbursement_4":"-"}
    return None


def col_value(row: pd.Series, col: str) -> Any:
    if not col or col == "-":
        return 0
    if col in row.index:
        return row[col]
    target = norm(col)
    for c in row.index:
        if norm(c) == target:
            return row[c]
    return 0


def process_report_file(path: Path, import_id: int) -> tuple[int, str]:
    df = read_csv_flexible(path)
    mapping = find_mapping(path.name, list(df.columns))
    if not mapping:
        raise ValueError("O nome do ficheiro não está configurado em Arquivos_Pgto e o cabeçalho não foi reconhecido.")
    origin = str(mapping["origin_ref"])
    count = 0
    rows = []
    for _, r in df.iterrows():
        identifier = str(col_value(r, mapping["identifier_column"])).strip()
        if not identifier:
            continue
        first = str(col_value(r, mapping["first_name_column"])).strip() if mapping["first_name_column"] != "-" else ""
        last = str(col_value(r, mapping["last_name_column"])).strip() if mapping["last_name_column"] != "-" else ""
        gross = money(col_value(r, mapping["value_column"]))
        cash = money(col_value(r, mapping["cash_column"]))
        reimb = sum(money(col_value(r, mapping[k])) for k in ("reimbursement_1","reimbursement_2","reimbursement_3","reimbursement_4"))
        rows.append((import_id, path.name, origin, identifier, (first + " " + last).strip(), gross, cash, reimb))
        count += 1
    if not rows and len(df.index) == 0:
        return 0, "Ficheiro válido, mas sem linhas de pagamento."
    with db() as con:
        con.executemany("INSERT INTO raw_earnings(import_id,filename,origin_ref,identifier,display_name,gross,cash,reimbursement) VALUES(?,?,?,?,?,?,?,?)", rows)
    return count, f"{origin}: {count} linhas"


def process_fuel_file(path: Path, import_id: int) -> tuple[int, str]:
    df = pd.read_excel(path, dtype=str).fillna("")
    cols = list(df.columns)
    card_col = next((c for c in cols if "CART" in norm(c) or "PAN" in norm(c)), cols[0] if cols else None)
    amount_col = next((c for c in cols if any(x in norm(c) for x in ("VALOR", "TOTAL", "MONTANTE", "DEBITO"))), cols[-1] if cols else None)
    if not card_col or not amount_col:
        raise ValueError("Não foi possível identificar cartão e valor no ficheiro PRIO.")
    grouped = defaultdict(float)
    for _, r in df.iterrows():
        card = clean_identifier(r[card_col])
        if card:
            grouped[card] += money(r[amount_col])
    with db() as con:
        con.executemany("INSERT INTO fuel(import_id,card_number,amount,source_file) VALUES(?,?,?,?)", [(import_id,k,v,path.name) for k,v in grouped.items()])
    return len(grouped), f"{len(grouped)} cartões consolidados"


def match_driver(identifier: str, origin: str, display_name: str):
    ident = clean_identifier(identifier)
    with db() as con:
        drivers = con.execute("SELECT * FROM drivers").fetchall()
    if "BOLT_FOOD" in norm(origin):
        # the current source workbook stores Bolt Food identifiers in ID where available
        for d in drivers:
            if clean_identifier(d["external_id"]) == ident:
                return d
    if "BOLT_TVDE" in norm(origin):
        digits = re.sub(r"\D", "", identifier)
        for d in drivers:
            dd = re.sub(r"\D", "", str(d["external_id"] or ""))
            if digits and dd and (digits == dd or digits[-9:] == dd[-9:]):
                return d
    for d in drivers:
        if clean_identifier(d["external_id"]) == ident:
            return d
    n = norm(display_name)
    if n:
        exact = [d for d in drivers if norm(d["name"]) == n]
        if len(exact) == 1:
            return exact[0]
    return None


def build_closing(label: str) -> int:
    """Processa o fechamento sem deixar uma linha inválida derrubar a aplicação."""
    created_at = datetime.now().isoformat(timespec="seconds")
    with db() as con:
        cur = con.execute(
            "INSERT INTO closings(label,created_at,status) VALUES(?,?,?)",
            (label, created_at, "PROCESSANDO"),
        )
        closing_id = int(cur.lastrowid)

    try:
        with db() as con:
            earnings = [dict(r) for r in con.execute("SELECT * FROM raw_earnings").fetchall()]
            fuel_rows = [dict(r) for r in con.execute(
                "SELECT card_number,SUM(amount) amount FROM fuel GROUP BY card_number"
            ).fetchall()]
            drivers = [dict(r) for r in con.execute("SELECT * FROM drivers").fetchall()]

        fuel_map = {
            clean_identifier(r.get("card_number")): money(r.get("amount"))
            for r in fuel_rows
            if clean_identifier(r.get("card_number"))
        }

        by_external = defaultdict(list)
        by_name = defaultdict(list)
        for d in drivers:
            ext = clean_identifier(d.get("external_id"))
            if ext:
                by_external[ext].append(d)
            name_key = norm(d.get("name"))
            if name_key:
                by_name[name_key].append(d)

        def find_driver(identifier: Any, origin: Any, display_name: Any):
            ident = clean_identifier(identifier)
            origin_key = norm(origin)
            if "BOLT_TVDE" in origin_key:
                digits = re.sub(r"\D", "", str(identifier or ""))
                if digits:
                    for d in drivers:
                        dd = re.sub(r"\D", "", str(d.get("external_id") or ""))
                        if dd and (digits == dd or digits[-9:] == dd[-9:]):
                            return d
            if ident and len(by_external.get(ident, [])) == 1:
                return by_external[ident][0]
            name_key = norm(display_name)
            if name_key and len(by_name.get(name_key, [])) == 1:
                return by_name[name_key][0]
            return None

        aggregated: dict[int, dict[str, Any]] = {}
        unmatched = 0
        invalid_rows = 0
        for e in earnings:
            try:
                driver = find_driver(e.get("identifier"), e.get("origin_ref"), e.get("display_name"))
                if not driver:
                    unmatched += 1
                    continue
                driver_id = int(driver["id"])
                a = aggregated.setdefault(driver_id, {
                    "driver": driver,
                    "gross": 0.0,
                    "cash": 0.0,
                    "report_reimb": 0.0,
                    "origins": set(),
                })
                a["gross"] += money(e.get("gross"))
                a["cash"] += money(e.get("cash"))
                a["report_reimb"] += money(e.get("reimbursement"))
                origin = str(e.get("origin_ref") or "SEM ORIGEM")
                a["origins"].add(origin)
            except Exception:
                invalid_rows += 1
                app.logger.exception("Linha ignorada durante o fechamento: %r", e)

        items = []
        iban_groups = defaultdict(float)
        temp = []
        for a in aggregated.values():
            d = a["driver"]
            origins_set = a["origins"]
            origins = ", ".join(sorted(origins_set))
            gross = money(a["gross"])
            cash = money(a["cash"])
            pct = money(d.get("percentage"))
            # Aceita tanto 0,15 quanto 15 no banco.
            if pct > 1:
                pct = pct / 100.0
            pct = max(0.0, min(pct, 1.0))
            is_eats = any("EATS" in norm(x) for x in origins_set)
            is_tvde = any("TVDE" in norm(x) for x in origins_set)
            commission_base = gross + cash if is_eats else gross
            commission = commission_base * pct
            fuel_value = fuel_map.get(clean_identifier(d.get("fuel_card")), 0.0)
            discount = money(d.get("discount"))
            reimbursement = (
                money(d.get("reimbursement")) + money(a["report_reimb"])
            ) if is_tvde else 0.0
            immediate = money(d.get("immediate"))
            net = gross - commission - fuel_value - discount - immediate + reimbursement
            if is_eats:
                net -= cash
            iban = re.sub(r"\s+", "", str(d.get("iban") or "")).upper()
            temp.append({
                "driver": d, "origins": origins, "gross": gross, "cash": cash,
                "commission": commission, "fuel": fuel_value, "discount": discount,
                "reimbursement": reimbursement, "immediate": immediate,
                "net": net, "iban": iban,
            })
            if iban:
                iban_groups[iban] += net

        fee_ibans = {
            t["iban"] for t in temp
            if t["iban"] and norm(t["driver"].get("bank_color")) == "AZUL"
        }
        for iban in fee_ibans:
            iban_groups[iban] -= 1.25

        fee_applied = set()
        for t in temp:
            d = t["driver"]
            iban = t["iban"]
            fee = 0.0
            if iban in fee_ibans and iban not in fee_applied:
                fee = 1.25
                fee_applied.add(iban)
            group_total = iban_groups.get(iban, t["net"] - fee)
            items.append((
                closing_id, int(d["id"]), str(d.get("name") or "SEM NOME"), iban,
                str(d.get("bank_color") or ""), t["origins"], t["gross"], t["cash"],
                t["commission"], t["fuel"], t["discount"], t["reimbursement"],
                t["immediate"], fee, t["net"], group_total,
            ))

        status = f"PROCESSADO ({unmatched} não encontrados; {invalid_rows} linhas ignoradas)"
        with db() as con:
            con.execute("DELETE FROM closing_items WHERE closing_id=?", (closing_id,))
            if items:
                con.executemany(
                    """INSERT INTO closing_items(
                    closing_id,driver_id,driver_name,iban,bank_color,origins,gross,cash,
                    commission,fuel,discount,reimbursement,immediate,bank_fee,
                    net_before_group,group_total)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    items,
                )
            con.execute("UPDATE closings SET status=? WHERE id=?", (status, closing_id))
        return closing_id
    except Exception as exc:
        app.logger.exception("Falha ao gerar fechamento %s", closing_id)
        with db() as con:
            con.execute(
                "UPDATE closings SET status=? WHERE id=?",
                (f"ERRO: {type(exc).__name__}: {str(exc)[:300]}", closing_id),
            )
        raise

def generate_xml(closing_id: int) -> bytes:
    with db() as con:
        rows = con.execute("SELECT iban,MAX(driver_name) driver_name,MAX(bank_fee) bank_fee,MAX(group_total) amount FROM closing_items WHERE closing_id=? AND iban<>'' GROUP BY iban HAVING amount>0 ORDER BY driver_name", (closing_id,)).fetchall()
    ns = "urn:iso:std:iso:20022:tech:xsd:pain.001.001.03"
    doc = Element("Document", xmlns=ns)
    init = SubElement(doc, "CstmrCdtTrfInitn")
    grp = SubElement(init, "GrpHdr")
    msg_id = f"FLEET-{closing_id}-{datetime.now():%Y%m%d%H%M%S}"
    SubElement(grp, "MsgId").text = msg_id
    SubElement(grp, "CreDtTm").text = datetime.now().isoformat(timespec="seconds")
    SubElement(grp, "NbOfTxs").text = str(len(rows))
    total = sum(money(r["amount"]) for r in rows)
    SubElement(grp, "CtrlSum").text = f"{total:.2f}"
    initg = SubElement(grp, "InitgPty")
    SubElement(initg, "Nm").text = COMPANY_NAME
    pmt = SubElement(init, "PmtInf")
    SubElement(pmt, "PmtInfId").text = msg_id
    SubElement(pmt, "PmtMtd").text = "TRF"
    SubElement(pmt, "BtchBookg").text = "true"
    SubElement(pmt, "NbOfTxs").text = str(len(rows))
    SubElement(pmt, "CtrlSum").text = f"{total:.2f}"
    pmt_type = SubElement(pmt, "PmtTpInf")
    svc = SubElement(pmt_type, "SvcLvl")
    SubElement(svc, "Cd").text = "SEPA"
    SubElement(pmt, "ReqdExctnDt").text = datetime.now().date().isoformat()
    dbtr = SubElement(pmt, "Dbtr")
    SubElement(dbtr, "Nm").text = COMPANY_NAME
    dbtr_acct = SubElement(pmt, "DbtrAcct")
    dbtr_id = SubElement(dbtr_acct, "Id")
    SubElement(dbtr_id, "IBAN").text = os.getenv("DEBTOR_IBAN", "PT50000000000000000000000")
    dbtr_agt = SubElement(pmt, "DbtrAgt")
    fin = SubElement(dbtr_agt, "FinInstnId")
    SubElement(fin, "BIC").text = os.getenv("DEBTOR_BIC", "NOTPROVIDED")
    SubElement(pmt, "ChrgBr").text = "SLEV"
    for i, r in enumerate(rows, 1):
        tx = SubElement(pmt, "CdtTrfTxInf")
        pid = SubElement(tx, "PmtId")
        SubElement(pid, "EndToEndId").text = f"{msg_id}-{i}"
        amt = SubElement(tx, "Amt")
        instd = SubElement(amt, "InstdAmt", Ccy="EUR")
        instd.text = f"{money(r['amount']):.2f}"
        cdtr_agt = SubElement(tx, "CdtrAgt")
        cdtr_fin = SubElement(cdtr_agt, "FinInstnId")
        SubElement(cdtr_fin, "BIC").text = "NOTPROVIDED"
        cdtr = SubElement(tx, "Cdtr")
        SubElement(cdtr, "Nm").text = str(r["driver_name"])[:70]
        acct = SubElement(tx, "CdtrAcct")
        aid = SubElement(acct, "Id")
        SubElement(aid, "IBAN").text = r["iban"]
        rmt = SubElement(tx, "RmtInf")
        SubElement(rmt, "Ustrd").text = f"Fechamento {closing_id}"
    xml = tostring(doc, encoding="utf-8", xml_declaration=True)
    return xml


@app.context_processor
def inject_globals():
    return {"company_name": COMPANY_NAME}


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        if request.form.get("username") == ADMIN_USER and request.form.get("password") == ADMIN_PASSWORD:
            session["user"] = ADMIN_USER
            return redirect(url_for("dashboard"))
        flash("Utilizador ou senha incorretos.", "danger")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
@login_required
def dashboard():
    with db() as con:
        drivers = con.execute("SELECT COUNT(*) FROM drivers").fetchone()[0]
        imports = con.execute("SELECT COUNT(*) FROM imports WHERE status='OK'").fetchone()[0]
        last = con.execute("SELECT * FROM closings ORDER BY id DESC LIMIT 1").fetchone()
        stats = {"gross":0,"commission":0,"fuel":0,"discount":0,"reimbursement":0,"net":0,"payments":0,"missing_iban":0}
        origins = []
        if last:
            r = con.execute("SELECT COALESCE(SUM(gross),0) gross,COALESCE(SUM(commission),0) commission,COALESCE(SUM(fuel),0) fuel,COALESCE(SUM(discount),0) discount,COALESCE(SUM(reimbursement),0) reimbursement,COALESCE(SUM(net_before_group-bank_fee),0) net,COUNT(DISTINCT iban) payments,SUM(CASE WHEN iban='' THEN 1 ELSE 0 END) missing_iban FROM closing_items WHERE closing_id=?", (last["id"],)).fetchone()
            stats = dict(r)
            origins = con.execute("SELECT origins,SUM(gross) total FROM closing_items WHERE closing_id=? GROUP BY origins ORDER BY total DESC", (last["id"],)).fetchall()
        recent = con.execute("SELECT * FROM closings ORDER BY id DESC LIMIT 8").fetchall()
    origin_labels = [row["origins"] or "Sem origem" for row in origins]
    origin_values = [float(row["total"] or 0) for row in origins]
    return render_template("dashboard.html", drivers=drivers, imports=imports, last=last, stats=stats, origins=origins, recent=recent, origin_labels=origin_labels, origin_values=origin_values)


@app.route("/drivers")
@login_required
def drivers():
    q = request.args.get("q", "").strip()
    with db() as con:
        if q:
            like = f"%{q}%"
            rows = con.execute("SELECT * FROM drivers WHERE name LIKE ? OR external_id LIKE ? OR iban LIKE ? OR city LIKE ? OR partner LIKE ? ORDER BY name LIMIT 500", (like,like,like,like,like)).fetchall()
        else:
            rows = con.execute("SELECT * FROM drivers ORDER BY name LIMIT 500").fetchall()
        total = con.execute("SELECT COUNT(*) FROM drivers").fetchone()[0]
    return render_template("drivers.html", rows=rows, q=q, total=total)


@app.route("/drivers/<int:driver_id>", methods=["GET", "POST"])
@login_required
def edit_driver(driver_id: int):
    with db() as con:
        d = con.execute("SELECT * FROM drivers WHERE id=?", (driver_id,)).fetchone()
        if not d:
            return "Motorista não encontrado", 404
        if request.method == "POST":
            fields = ["external_id","name","city","company","iban","bank_color","partner","fuel_card","rental_label","observation"]
            vals = [request.form.get(f, "") for f in fields]
            nums = [money(request.form.get(f, 0)) for f in ("commission_owner","partner_commission","percentage","discount","reimbursement","immediate")]
            con.execute("""UPDATE drivers SET external_id=?,name=?,city=?,company=?,iban=?,bank_color=?,partner=?,fuel_card=?,rental_label=?,observation=?,commission_owner=?,partner_commission=?,percentage=?,discount=?,reimbursement=?,immediate=?,updated_at=? WHERE id=?""", (*vals,*nums,datetime.now().isoformat(timespec="seconds"),driver_id))
            flash("Cadastro atualizado.", "success")
            return redirect(url_for("edit_driver", driver_id=driver_id))
    return render_template("driver_edit.html", d=d)


@app.route("/database", methods=["GET", "POST"])
@login_required
def database_page():
    if request.method == "POST":
        f = request.files.get("database_file")
        if not f or not f.filename.lower().endswith(".xlsx"):
            flash("Envie um ficheiro XLSX válido.", "danger")
        else:
            p = UPLOAD_DIR / secure_filename(f.filename)
            f.save(p)
            try:
                n, m = import_master_workbook(p)
                flash(f"Banco atualizado: {n} motoristas e {m} configurações de leitura.", "success")
            except Exception as exc:
                flash(f"Erro ao importar banco: {exc}", "danger")
    with db() as con:
        n = con.execute("SELECT COUNT(*) FROM drivers").fetchone()[0]
        m = con.execute("SELECT COUNT(*) FROM mappings").fetchone()[0]
    return render_template("database.html", n=n, m=m)


@app.route("/imports", methods=["GET", "POST"])
@login_required
def imports_page():
    if request.method == "POST":
        kind = request.form.get("kind", "reports")
        files = request.files.getlist("files")
        success = errors = 0
        for f in files:
            if not f or not f.filename:
                continue
            filename = secure_filename(f.filename)
            data = f.read()
            sha = hashlib.sha256(data).hexdigest()
            with db() as con:
                duplicate = con.execute("SELECT id FROM imports WHERE sha256=?", (sha,)).fetchone()
            if duplicate:
                flash(f"{filename}: já foi importado anteriormente; ignorado.", "warning")
                continue
            p = UPLOAD_DIR / f"{datetime.now():%Y%m%d%H%M%S%f}_{filename}"
            p.write_bytes(data)
            with db() as con:
                cur = con.execute("INSERT INTO imports(filename,sha256,kind,status,created_at) VALUES(?,?,?,?,?)", (filename,sha,kind,"PROCESSANDO",datetime.now().isoformat(timespec="seconds")))
                iid = cur.lastrowid
            try:
                if kind == "fuel" or filename.lower().endswith((".xlsx", ".xls")):
                    count, msg = process_fuel_file(p, iid)
                else:
                    count, msg = process_report_file(p, iid)
                with db() as con:
                    con.execute("UPDATE imports SET status='OK',rows_count=?,message=? WHERE id=?", (count,msg,iid))
                success += 1
            except Exception as exc:
                with db() as con:
                    con.execute("UPDATE imports SET status='ERRO',message=? WHERE id=?", (str(exc),iid))
                errors += 1
        flash(f"Importação concluída: {success} ficheiros processados e {errors} com erro.", "success" if errors == 0 else "warning")
        return redirect(url_for("imports_page"))
    with db() as con:
        rows = con.execute("SELECT * FROM imports ORDER BY id DESC LIMIT 100").fetchall()
    return render_template("imports.html", rows=rows)


@app.route("/imports/clear", methods=["POST"])
@login_required
def clear_imports():
    with db() as con:
        con.execute("DELETE FROM raw_earnings")
        con.execute("DELETE FROM fuel")
        con.execute("DELETE FROM imports")
    flash("Importações da semana foram limpas.", "success")
    return redirect(url_for("imports_page"))


@app.route("/closings", methods=["GET", "POST"])
@login_required
def closings_page():
    if request.method == "POST":
        label = request.form.get("label") or f"Semana {datetime.now():%d/%m/%Y}"
        try:
            cid = build_closing(label)
            flash("Fechamento processado. Confira os valores antes de gerar o XML.", "success")
            return redirect(url_for("closing_detail", closing_id=cid))
        except Exception as exc:
            app.logger.exception("Erro no botão Processar semana")
            flash(f"Não foi possível gerar o fechamento: {type(exc).__name__}: {str(exc)}", "danger")
            return redirect(url_for("closings_page"))
    with db() as con:
        rows = con.execute("SELECT * FROM closings ORDER BY id DESC").fetchall()
    return render_template("closings.html", rows=rows)


@app.route("/closings/<int:closing_id>")
@login_required
def closing_detail(closing_id: int):
    with db() as con:
        closing = con.execute("SELECT * FROM closings WHERE id=?", (closing_id,)).fetchone()
        items = con.execute("SELECT * FROM closing_items WHERE closing_id=? ORDER BY driver_name", (closing_id,)).fetchall()
    return render_template("closing_detail.html", closing=closing, items=items)


@app.route("/closings/<int:closing_id>/excel")
@login_required
def closing_excel(closing_id: int):
    with db() as con:
        df = pd.read_sql_query("SELECT driver_name AS Motorista,iban AS IBAN,bank_color AS Banco,origins AS Origens,gross AS Bruto,cash AS Dinheiro,commission AS Comissao,fuel AS Combustivel,discount AS Desconto,reimbursement AS Reembolso,immediate AS Imediata,bank_fee AS Taxa_Banco,net_before_group AS Liquido_Individual,group_total AS Total_IBAN FROM closing_items WHERE closing_id=? ORDER BY driver_name", con, params=(closing_id,))
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Resultado_Geral")
        same = df[df["IBAN"].duplicated(False) & (df["IBAN"] != "")]
        same.to_excel(writer, index=False, sheet_name="Mesmo_IBAN")
        df[df["IBAN"] == ""].to_excel(writer, index=False, sheet_name="Sem_IBAN")
        df[df["Liquido_Individual"] < 0].to_excel(writer, index=False, sheet_name="Negativos")
        rentals = df[df["Desconto"] > 0][["Motorista","Desconto","Origens"]]
        rentals.to_excel(writer, index=False, sheet_name="Descontos_Alugueis")
    out.seek(0)
    return send_file(out, as_attachment=True, download_name=f"fechamento_{closing_id}.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/closings/<int:closing_id>/xml")
@login_required
def closing_xml(closing_id: int):
    xml = generate_xml(closing_id)
    return send_file(io.BytesIO(xml), as_attachment=True, download_name=f"pagamentos_{closing_id}.xml", mimetype="application/xml")



@app.route("/health")
def health():
    try:
        with db() as con:
            con.execute("SELECT 1").fetchone()
        return {"status": "ok", "database": str(DB_PATH)}, 200
    except Exception as exc:
        return {"status": "error", "message": str(exc)}, 500


@app.errorhandler(500)
def internal_error(exc):
    app.logger.error("Erro interno: %s\n%s", exc, traceback.format_exc())
    try:
        return render_template("error.html", message=f"Ocorreu um erro interno: {type(exc).__name__}. Consulte os logs do Render."), 500
    except Exception:
        return f"Erro interno: {type(exc).__name__}: {exc}", 500

@app.errorhandler(413)
def too_large(_):
    flash("O envio excede o limite de 80 MB.", "danger")
    return redirect(url_for("imports_page"))


init_db()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "5000")), debug=True)
