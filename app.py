
from flask import (
    Flask, request, redirect, url_for, render_template_string,
    send_file, flash, session
)
from functools import wraps
from io import BytesIO, StringIO
from zipfile import ZipFile, ZIP_DEFLATED
from datetime import date, datetime, timedelta
import csv
import hmac
import os
import sqlite3
import unicodedata
import uuid
import xml.etree.ElementTree as ET

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from openpyxl import load_workbook

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "troque-esta-chave-no-render")
DB = os.environ.get("DB_PATH", os.path.join(os.path.dirname(__file__), "tvde_manager.db"))
ADMIN_USER = os.environ.get("ADMIN_USER", "admin")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "admin123")


def db():
    c = sqlite3.connect(DB)
    c.row_factory = sqlite3.Row
    return c


def init_db():
    c = db()
    c.executescript("""
    CREATE TABLE IF NOT EXISTS motoristas(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      nome TEXT NOT NULL UNIQUE,
      email TEXT,
      telefone TEXT,
      iban TEXT,
      percentual REAL DEFAULT 0,
      ativo INTEGER DEFAULT 1
    );

    CREATE TABLE IF NOT EXISTS viaturas(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      matricula TEXT NOT NULL UNIQUE,
      marca TEXT,
      modelo TEXT,
      motorista_id INTEGER,
      seguro_validade TEXT,
      ipo_validade TEXT,
      carta_verde_validade TEXT,
      km INTEGER DEFAULT 0
    );

    CREATE TABLE IF NOT EXISTS relatorios(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      plataforma TEXT NOT NULL,
      motorista_id INTEGER NOT NULL,
      semana TEXT NOT NULL,
      bruto REAL DEFAULT 0,
      dinheiro_maos REAL DEFAULT 0,
      comissao REAL DEFAULT 0,
      portagens REAL DEFAULT 0,
      outros_descontos REAL DEFAULT 0,
      reembolsos REAL DEFAULT 0,
      liquido REAL DEFAULT 0,
      criado_em TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS importacoes(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      plataforma TEXT NOT NULL,
      semana TEXT NOT NULL,
      arquivo TEXT NOT NULL,
      linhas INTEGER DEFAULT 0,
      criado_em TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS combustivel(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      semana TEXT NOT NULL,
      motorista_id INTEGER,
      cartao TEXT NOT NULL,
      descricao_cartao TEXT,
      data TEXT,
      posto TEXT,
      litros REAL DEFAULT 0,
      total REAL DEFAULT 0,
      arquivo TEXT,
      criado_em TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS importacoes_combustivel(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      semana TEXT NOT NULL,
      arquivo TEXT NOT NULL,
      lancamentos INTEGER DEFAULT 0,
      total REAL DEFAULT 0,
      pendentes INTEGER DEFAULT 0,
      criado_em TEXT NOT NULL
    );
    
    CREATE TABLE IF NOT EXISTS descontos_extras(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      motorista_id INTEGER NOT NULL,
      semana TEXT NOT NULL,
      grupo TEXT NOT NULL DEFAULT 'TVDE',
      categoria TEXT NOT NULL,
      descricao TEXT,
      valor REAL NOT NULL DEFAULT 0,
      criado_em TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS xml_historico(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      semana TEXT NOT NULL,
      grupo TEXT NOT NULL,
      arquivo TEXT NOT NULL,
      pagamentos INTEGER NOT NULL,
      total REAL NOT NULL,
      criado_em TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS pagamentos_consolidados(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      semana TEXT NOT NULL,
      grupo TEXT NOT NULL,
      arquivo TEXT NOT NULL,
      parte INTEGER NOT NULL,
      iban TEXT NOT NULL,
      nome_pagamento TEXT NOT NULL,
      taxa REAL NOT NULL DEFAULT 0,
      valor_transferido REAL NOT NULL,
      criado_em TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS pagamentos_consolidados_itens(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      pagamento_id INTEGER NOT NULL,
      motorista_id INTEGER,
      motorista_nome TEXT NOT NULL,
      plataforma TEXT,
      bruto REAL DEFAULT 0,
      dinheiro_maos REAL DEFAULT 0,
      comissao REAL DEFAULT 0,
      combustivel REAL DEFAULT 0,
      descontos_extras REAL DEFAULT 0,
      reembolsos REAL DEFAULT 0,
      valor_liquido REAL DEFAULT 0
    );
    """)
    cols = [r[1] for r in c.execute("PRAGMA table_info(motoristas)").fetchall()]
    if "cartao_prio" not in cols:
        c.execute("ALTER TABLE motoristas ADD COLUMN cartao_prio TEXT")

    report_cols = [r[1] for r in c.execute("PRAGMA table_info(relatorios)").fetchall()]
    if "grupo" not in report_cols:
        c.execute("ALTER TABLE relatorios ADD COLUMN grupo TEXT DEFAULT 'TVDE'")
    c.commit()
    c.close()


def norm(s):
    s = str(s or "").strip().lower()
    s = unicodedata.normalize("NFKD", s)
    return "".join(ch for ch in s if not unicodedata.combining(ch))


def num(v):
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).strip().replace("€", "").replace("\xa0", "")
    if "," in t and "." in t:
        if t.rfind(",") > t.rfind("."):
            t = t.replace(".", "").replace(",", ".")
        else:
            t = t.replace(",", "")
    elif "," in t:
        t = t.replace(".", "").replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return 0.0


def money(v):
    return f"€ {float(v or 0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login", next=request.path))
        return fn(*args, **kwargs)
    return wrapper


BASE = """
<!doctype html>
<html lang="pt">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{{ title }} · Irmãos Fleet</title>
<style>
:root{
 --bg:#f4f7fb;--surface:#ffffff;--sidebar:#0d1b2f;--sidebar2:#12243d;
 --text:#172033;--muted:#667085;--line:#e4e9f0;--blue:#2477e8;
 --green:#13a36d;--purple:#7557d9;--orange:#f2a900;--cyan:#3694b7;
 --red:#e84566;--ok:#067647;--warn:#b54708;--bad:#b42318
}
*{box-sizing:border-box}
body{margin:0;font-family:Inter,Arial,sans-serif;background:var(--bg);color:var(--text)}
.app-shell{display:grid;grid-template-columns:245px minmax(0,1fr);min-height:100vh}
.sidebar{background:linear-gradient(180deg,var(--sidebar),var(--sidebar2));color:white;padding:18px 14px;position:sticky;top:0;height:100vh;overflow-y:auto}
.brand{display:flex;gap:10px;align-items:center;padding:4px 6px 20px}
.brand img{width:42px;height:42px;border-radius:50%;object-fit:cover;border:1px solid #ffffff30}
.brand strong{display:block;font-size:16px;letter-spacing:.3px}
.brand small{color:#b7c2d2;font-size:11px}
.nav-title{font-size:10px;text-transform:uppercase;letter-spacing:1px;color:#8190a6;margin:18px 10px 8px}
.nav{display:grid;gap:5px}
.nav a{display:flex;align-items:center;gap:10px;color:#d6dfeb;text-decoration:none;padding:11px 12px;border-radius:8px;font-weight:650;font-size:13px}
.nav a:hover,.nav a.active{background:linear-gradient(90deg,#2276e8,#2d82ed);color:white}
.nav .badge{margin-left:auto;background:#ff8a00;color:white;border-radius:7px;padding:2px 7px;font-size:11px}
.version{position:sticky;top:calc(100vh - 42px);padding:16px 8px 2px;color:#9cacbf;font-size:11px}
.main{min-width:0}
.topbar{height:64px;background:white;border-bottom:1px solid var(--line);display:flex;align-items:center;justify-content:space-between;padding:0 20px;position:sticky;top:0;z-index:10}
.topbar-left{display:flex;align-items:center;gap:14px}
.menu-toggle{border:0;background:transparent;font-size:22px;color:var(--text);padding:6px}
.topbar-right{display:flex;align-items:center;gap:18px;color:var(--text);font-size:13px}
.week-chip{border:1px solid var(--line);border-radius:8px;padding:9px 12px;background:white}
.mobile-nav{display:none}
.content{max-width:1450px;margin:auto;padding:22px}
.page-head{margin-bottom:18px}
.page-head h1{margin:0;font-size:24px}
.page-head p{margin:5px 0 0;color:var(--muted);font-size:13px}
.kpi-grid{display:grid;grid-template-columns:repeat(6,minmax(150px,1fr));gap:14px;margin-bottom:18px}
.kpi{border-radius:8px;padding:17px;color:white;min-height:132px;position:relative;overflow:hidden}
.kpi h3{margin:0;font-size:13px;font-weight:650}.kpi .value{font-size:22px;font-weight:850;margin:12px 0}
.kpi small{display:block;color:#ffffffdc;line-height:1.7}.kpi .icon{position:absolute;right:16px;top:45px;font-size:30px;opacity:.75}
.kpi.green{background:linear-gradient(135deg,#0d9b64,#17b277)}
.kpi.blue{background:linear-gradient(135deg,#176cda,#2c83ed)}
.kpi.purple{background:linear-gradient(135deg,#674cd3,#7d60e3)}
.kpi.orange{background:linear-gradient(135deg,#eda900,#ffb800)}
.kpi.cyan{background:linear-gradient(135deg,#2b83a8,#3da0c3)}
.kpi.red{background:linear-gradient(135deg,#df3d5f,#f04d6c)}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:14px}
.dashboard-grid{display:grid;grid-template-columns:1.2fr 1fr .8fr;gap:14px;margin-bottom:14px}
.dashboard-grid.bottom{grid-template-columns:1fr .9fr 1fr}
.card{background:var(--surface);border:1px solid var(--line);border-radius:8px;margin-bottom:14px;overflow:hidden}
.card-header{display:flex;align-items:center;justify-content:space-between;padding:14px 16px;border-bottom:1px solid var(--line)}
.card-header h2,.card h2{font-size:14px;margin:0}
.card-body{padding:15px}
.card > h2{padding:15px 16px;border-bottom:1px solid var(--line)}
.metric-label{font-size:12px;color:var(--muted)}.metric-value{font-size:26px;font-weight:850;margin-top:7px}
.muted{color:var(--muted)}
table{width:100%;border-collapse:collapse}th,td{padding:10px 12px;border-bottom:1px solid var(--line);text-align:left;font-size:12px}
th{font-size:10px;text-transform:uppercase;color:var(--muted);background:#fbfcfe}
label{display:block;font-size:12px;font-weight:750;color:#344054}
input,select{width:100%;padding:10px;border:1px solid #cfd6e1;border-radius:7px;background:white;margin-top:6px}
button,.btn{border:0;border-radius:7px;background:#172033;color:white;padding:10px 14px;text-decoration:none;display:inline-block;font-weight:750;cursor:pointer;font-size:12px}
.btn.primary{background:#2477e8}.btn.green{background:#12a36d}.btn.purple{background:#7557d9}.btn.orange{background:#ef7d00;color:white}
.actions{display:flex;gap:8px;flex-wrap:wrap}
.flash{background:#ecfdf3;border:1px solid #abefc6;padding:11px;border-radius:7px;margin-bottom:14px}
.ok{color:var(--ok);font-weight:750}.warn{color:var(--warn);font-weight:750}.bad{color:var(--bad);font-weight:750}
.status{display:inline-block;padding:3px 7px;border-radius:999px;font-size:10px;font-weight:750}
.status.ok{background:#e8f7ef;color:#067647}.status.warn{background:#fff1df;color:#b54708}.status.bad{background:#fff0f1;color:#b42318}
.filebox{border:2px dashed #c6cfdb;border-radius:8px;padding:14px;background:#fafcff}
.progress{height:8px;background:#edf1f6;border-radius:999px;overflow:hidden}.progress span{display:block;height:100%;background:#2477e8}
.quick-actions{background:linear-gradient(90deg,#eef6ff,#f8fbff);border:1px solid #cfe0f4;border-radius:8px;padding:14px;display:flex;gap:10px;flex-wrap:wrap}
@media(max-width:1180px){.kpi-grid{grid-template-columns:repeat(3,1fr)}.dashboard-grid,.dashboard-grid.bottom{grid-template-columns:1fr}}
@media(max-width:820px){
 .app-shell{display:block}.sidebar{display:none}.topbar{padding:0 12px}.week-chip{display:none}
 .mobile-nav{display:flex;gap:7px;overflow-x:auto;padding:9px;background:#0d1b2f;position:sticky;top:64px;z-index:9}
 .mobile-nav a{white-space:nowrap;color:white;text-decoration:none;background:#ffffff13;padding:8px 10px;border-radius:7px;font-size:12px}
 .content{padding:13px}.kpi-grid{grid-template-columns:repeat(2,1fr)}
 table{display:block;overflow-x:auto;white-space:nowrap}
}
@media(max-width:520px){.kpi-grid{grid-template-columns:1fr}.topbar-right{gap:8px}}
</style>
</head>
<body>
<div class="app-shell">
<aside class="sidebar">
 <div class="brand">
  <img src="/static/logo.png" onerror="this.style.display='none'">
  <div><strong>IRMÃOS FLEET</strong><small>Sistema de Gestão TVDE</small></div>
 </div>
 <div class="nav">
  <a href="/">⌂ Dashboard</a>
  <a href="/motoristas">♙ Motoristas</a>
  <a href="/viaturas">▣ Viaturas</a>
 </div>
 <div class="nav-title">Importações</div>
 <div class="nav">
  <a href="/importar">⇧ Uber e Bolt</a>
  <a href="/combustivel">⛽ Combustível PRIO</a>
 </div>
 <div class="nav-title">Financeiro</div>
 <div class="nav">
  <a href="/relatorios">▤ Relatórios</a>
  <a href="/pagamentos">⇄ Pagamentos XML</a>
  <a href="/auditoria-pagamentos">◎ Auditoria por IBAN</a>
  <a href="/recibos">▧ Recibos</a>
 </div>
 <div class="nav-title">Pendências</div>
 <div class="nav">
  <a href="/sem-iban">! Sem IBAN <span class="badge">!</span></a>
 </div>
 <div class="nav-title">Conta</div>
 <div class="nav"><a href="/logout">↗ Sair</a></div>
 <div class="version">Versão 1.3</div>
</aside>
<section class="main">
 <header class="topbar">
  <div class="topbar-left"><button class="menu-toggle">☰</button></div>
  <div class="topbar-right"><div class="week-chip">📅 Semana atual</div><div>Olá, {{ session.get('username','Admin') }}</div></div>
 </header>
 <nav class="mobile-nav">
  <a href="/">Dashboard</a><a href="/motoristas">Motoristas</a><a href="/viaturas">Viaturas</a>
  <a href="/importar">Importar</a><a href="/combustivel">PRIO</a><a href="/pagamentos">XML</a><a href="/auditoria-pagamentos">Auditoria</a><a href="/sem-iban">Sem IBAN</a>
 </nav>
 <main class="content">
  <div class="page-head"><h1>{{ title }}</h1><p>Visão geral da operação</p></div>
  {% with messages = get_flashed_messages() %}
  {% for m in messages %}<div class="flash">{{m}}</div>{% endfor %}
  {% endwith %}
  {{ body|safe }}
 </main>
</section>
</div>
</body>
</html>
"""


def render(title, body, **ctx):
    inner = render_template_string(body, **ctx)
    return render_template_string(BASE, title=title, body=inner)


@app.route("/login", methods=["GET", "POST"])
def login():
    error = ""
    if request.method == "POST":
        user_ok = hmac.compare_digest(request.form.get("username", ""), ADMIN_USER)
        pass_ok = hmac.compare_digest(request.form.get("password", ""), ADMIN_PASSWORD)
        if user_ok and pass_ok:
            session["logged_in"] = True
            session["username"] = ADMIN_USER
            return redirect(request.args.get("next") or url_for("dashboard"))
        error = "Utilizador ou senha incorretos."
    return render_template_string("""
<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Entrar · Irmãos Fleet</title>
<style>
*{box-sizing:border-box}body{margin:0;font-family:Arial;background:#07111d;min-height:100vh;display:grid;place-items:center;padding:20px}
.login{width:min(430px,100%);background:#fff;border-radius:22px;padding:26px}
.logo{width:100%;height:190px;object-fit:cover;border-radius:16px;margin-bottom:18px}
h1{margin:0 0 5px;font-size:24px}p{color:#667085;margin:0 0 20px}
label{font-size:13px;font-weight:bold;display:block;margin-top:12px}input{width:100%;padding:13px;border:1px solid #d0d5dd;border-radius:10px;margin-top:6px}
button{width:100%;padding:13px;border:0;border-radius:10px;background:linear-gradient(90deg,#18c8d8,#84f438);font-weight:900;margin-top:18px}
.error{background:#fef3f2;color:#b42318;padding:10px;border-radius:8px;margin-bottom:12px}
small{display:block;text-align:center;color:#98a2b3;margin-top:15px}
</style></head><body><div class="login">
<img class="logo" src="/static/logo.png"><h1>Bem-vindo</h1><p>Acesse o painel administrativo.</p>
{% if error %}<div class="error">{{error}}</div>{% endif %}
<form method="post"><label>Utilizador<input name="username" required autocomplete="username"></label>
<label>Senha<input type="password" name="password" required autocomplete="current-password"></label>
<button>Entrar</button></form><small>Irmãos Fleet · Gestão TVDE</small></div></body></html>
""", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
@login_required
def dashboard():
    c = db()
    counts = {
        "motoristas": c.execute("SELECT COUNT(*) n FROM motoristas WHERE ativo=1").fetchone()["n"],
        "viaturas": c.execute("SELECT COUNT(*) n FROM viaturas").fetchone()["n"],
    }
    totals = c.execute("""SELECT COALESCE(SUM(bruto),0) bruto,
                         COALESCE(SUM(comissao),0) comissao,
                         COALESCE(SUM(dinheiro_maos),0) dinheiro,
                         COALESCE(SUM(liquido),0) liquido
                         FROM relatorios""").fetchone()
    fuel_total = c.execute("SELECT COALESCE(SUM(total),0) total FROM combustivel").fetchone()["total"]
    fuel_count = c.execute("SELECT COUNT(*) n FROM combustivel").fetchone()["n"]
    fuel_pending = c.execute("SELECT COUNT(*) n FROM combustivel WHERE motorista_id IS NULL").fetchone()["n"]
    no_iban = c.execute("SELECT COUNT(*) n FROM motoristas WHERE TRIM(COALESCE(iban,''))=''").fetchone()["n"]
    recent = c.execute("""SELECT r.*,m.nome motorista FROM relatorios r
                          JOIN motoristas m ON m.id=r.motorista_id
                          ORDER BY r.id DESC LIMIT 6""").fetchall()
    alerts = []
    today = date.today()
    for v in c.execute("SELECT * FROM viaturas ORDER BY matricula").fetchall():
        for field, label in [("seguro_validade","Seguro"),("ipo_validade","IPO"),("carta_verde_validade","Carta Verde")]:
            if v[field]:
                try:
                    days = (datetime.strptime(v[field], "%Y-%m-%d").date() - today).days
                    if days <= 20:
                        alerts.append((v["matricula"], label, days))
                except ValueError:
                    pass
    fuel_top = c.execute("""SELECT m.nome motorista,m.cartao_prio,COUNT(f.id) abastecimentos,
                            COALESCE(SUM(f.litros),0) litros,COALESCE(SUM(f.total),0) total
                            FROM combustivel f JOIN motoristas m ON m.id=f.motorista_id
                            GROUP BY m.id ORDER BY total DESC LIMIT 5""").fetchall()
    xml_recent = c.execute("""SELECT * FROM xml_historico ORDER BY id DESC LIMIT 3""").fetchall()
    c.close()

    body = """
<div class="kpi-grid">
 <div class="kpi green"><h3>Faturamento Bruto</h3><div class="value">{{money(totals.bruto)}}</div><small>Operação acumulada</small><div class="icon">▥</div></div>
 <div class="kpi blue"><h3>Comissões</h3><div class="value">{{money(totals.comissao)}}</div><small>Comissão processada</small><div class="icon">%</div></div>
 <div class="kpi purple"><h3>Combustível (PRIO)</h3><div class="value">{{money(fuel_total)}}</div><small>Abastecimentos: {{fuel_count}}</small><div class="icon">⛽</div></div>
 <div class="kpi orange"><h3>Valor a Pagar</h3><div class="value">{{money(totals.liquido)}}</div><small>Pagamento previsto</small><div class="icon">♟</div></div>
 <div class="kpi cyan"><h3>Sem IBAN</h3><div class="value">{{no_iban}}</div><small>Aguardando cadastro</small><div class="icon">⚠</div></div>
 <div class="kpi red"><h3>Pendências</h3><div class="value">{{alerts|length + fuel_pending}}</div><small>Verifique antes do fechamento</small><div class="icon">!</div></div>
</div>

<div class="dashboard-grid">
 <section class="card">
  <div class="card-header"><h2>Importações da Semana</h2></div>
  <div class="card-body">
   <table><tr><th>Origem</th><th>Status</th><th>Detalhes</th></tr>
    <tr><td>Uber</td><td><span class="status ok">Importado</span></td><td>Relatórios disponíveis</td></tr>
    <tr><td>Bolt</td><td><span class="status ok">Importado</span></td><td>Relatórios disponíveis</td></tr>
    <tr><td>Combustível PRIO</td><td><span class="status {{'warn' if fuel_pending else 'ok'}}">{{'Pendências' if fuel_pending else 'Importado'}}</span></td><td>{{fuel_pending}} cartão(ões) não identificado(s)</td></tr>
   </table>
  </div>
 </section>
 <section class="card">
  <div class="card-header"><h2>Resumo de Pagamentos</h2></div>
  <div class="card-body">
   <div style="text-align:center;padding:20px 10px">
    <div style="font-size:12px;color:var(--muted)">Total previsto</div>
    <div style="font-size:30px;font-weight:850;margin:8px 0">{{money(totals.liquido)}}</div>
    <div class="progress"><span style="width:{{ 100 if not no_iban else 80 }}%"></span></div>
    <div style="margin-top:12px;font-size:12px;color:var(--muted)">Com IBAN: {{counts.motoristas-no_iban}} · Sem IBAN: {{no_iban}}</div>
   </div>
  </div>
 </section>
 <section class="card">
  <div class="card-header"><h2>Estatísticas Rápidas</h2></div>
  <div class="card-body">
   <table>
    <tr><td>Total de Motoristas</td><td><b>{{counts.motoristas}}</b></td></tr>
    <tr><td>Viaturas</td><td><b>{{counts.viaturas}}</b></td></tr>
    <tr><td>Relatórios Recentes</td><td><b>{{recent|length}}</b></td></tr>
    <tr><td>Abastecimentos</td><td><b>{{fuel_count}}</b></td></tr>
    <tr><td>Sem IBAN</td><td><b>{{no_iban}}</b></td></tr>
   </table>
  </div>
 </section>
</div>

<div class="dashboard-grid bottom">
 <section class="card">
  <div class="card-header"><h2>Top 5 — Combustível por Motorista</h2></div>
  <div class="card-body">
   <table><tr><th>#</th><th>Motorista</th><th>Cartão</th><th>Total</th><th>Litros</th></tr>
   {% for r in fuel_top %}<tr><td>{{loop.index}}</td><td>{{r.motorista}}</td><td>{{r.cartao_prio}}</td><td><b>{{money(r.total)}}</b></td><td>{{"%.2f"|format(r.litros or 0)}}</td></tr>{% endfor %}
   </table>
  </div>
 </section>
 <section class="card">
  <div class="card-header"><h2>Pendências</h2></div>
  <div class="card-body">
   <table>
    <tr><td>Sem IBAN</td><td><span class="status warn">{{no_iban}}</span></td></tr>
    <tr><td>Documentos vencendo</td><td><span class="status {{'bad' if alerts else 'ok'}}">{{alerts|length}}</span></td></tr>
    <tr><td>Cartões PRIO não encontrados</td><td><span class="status {{'warn' if fuel_pending else 'ok'}}">{{fuel_pending}}</span></td></tr>
   </table>
   <div style="margin-top:14px"><a class="btn primary" href="/sem-iban">Ver pendências</a></div>
  </div>
 </section>
 <section class="card">
  <div class="card-header"><h2>Últimos XMLs Gerados</h2></div>
  <div class="card-body">
   <table><tr><th>Semana</th><th>Arquivo</th><th>Total</th></tr>
   {% for x in xml_recent %}<tr><td>{{x.semana}}</td><td>{{x.arquivo}}</td><td><b>{{money(x.total)}}</b></td></tr>{% endfor %}
   </table>
  </div>
 </section>
</div>

<div class="quick-actions">
 <a class="btn primary" href="/importar">⇧ Importar Uber/Bolt</a>
 <a class="btn green" href="/combustivel">⛽ Importar PRIO</a>
 <a class="btn purple" href="/pagamentos">⇄ Simular Pagamentos</a>
 <a class="btn" href="/auditoria-pagamentos">◎ Auditoria por IBAN</a>
 <a class="btn orange" href="/relatorios">▤ Ver Relatórios</a>
</div>
"""
    return render("Dashboard", body, counts=counts, totals=totals, fuel_total=fuel_total,
                  fuel_count=fuel_count, fuel_pending=fuel_pending, no_iban=no_iban,
                  recent=recent, alerts=alerts, fuel_top=fuel_top, xml_recent=xml_recent, money=money)


def get_or_create_driver(c, name, email="", phone=""):
    name = " ".join(str(name or "").split()).strip()
    if not name:
        return None
    found = c.execute("SELECT id FROM motoristas WHERE lower(nome)=lower(?)", (name,)).fetchone()
    if found:
        return found["id"]
    cur = c.execute("INSERT INTO motoristas(nome,email,telefone) VALUES(?,?,?)", (name, email, phone))
    return cur.lastrowid


def import_uber(c, rows, week):
    count = 0
    for row in rows:
        first = row.get("Nome próprio do motorista", "")
        last = row.get("Apelido do motorista", "")
        name = " ".join((first + " " + last).split())
        if not name:
            continue
        driver_id = get_or_create_driver(c, name)
        bruto = num(row.get("Pago a si : Os seus rendimentos : Tarifa"))
        dinheiro = abs(num(row.get("Pago a si : Saldo da viagem : Pagamentos : Dinheiro recebido")))
        service = abs(num(row.get("Pago a si:Os seus rendimentos:Taxa de serviço")))
        tolls = abs(num(row.get("Pago a si:Saldo da viagem:Reembolsos:Portagem")))
        paid = num(row.get("Pago a si"))
        earnings = num(row.get("Pago a si : Os seus rendimentos"))
        liquid = paid if paid else earnings - dinheiro
        c.execute("""INSERT INTO relatorios(plataforma,motorista_id,semana,bruto,dinheiro_maos,comissao,
                   portagens,outros_descontos,reembolsos,liquido,criado_em)
                   VALUES('Uber',?,?,?,?,?,?,?,?,?,?)""",
                  (driver_id, week, bruto, dinheiro, service, tolls, 0, tolls, liquid,
                   datetime.now().isoformat(timespec="seconds")))
        c.execute("UPDATE relatorios SET grupo='TVDE' WHERE id=last_insert_rowid()")
        count += 1
    return count


def import_bolt(c, rows, week):
    count = 0
    for row in rows:
        name = row.get("Motorista", "")
        if not name:
            continue
        driver_id = get_or_create_driver(c, name, row.get("Email",""), row.get("Telemóvel",""))
        bruto = num(row.get("Ganhos brutos (total)|€"))
        dinheiro = abs(num(row.get("Dinheiro recebido|€")))
        commission = abs(num(row.get("Comissões|€")))
        tolls = abs(num(row.get("Portagens|€")))
        other = abs(num(row.get("Outras taxas|€"))) + abs(num(row.get("Reembolsos aos passageiros|€")))
        reimburse = abs(num(row.get("Reembolsos de despesas|€")))
        predicted = num(row.get("Pagamento previsto|€"))
        liquid = predicted if predicted else num(row.get("Ganhos líquidos|€")) - dinheiro
        c.execute("""INSERT INTO relatorios(plataforma,motorista_id,semana,bruto,dinheiro_maos,comissao,
                   portagens,outros_descontos,reembolsos,liquido,criado_em)
                   VALUES('Bolt',?,?,?,?,?,?,?,?,?,?)""",
                  (driver_id, week, bruto, dinheiro, commission, tolls, other, reimburse, liquid,
                   datetime.now().isoformat(timespec="seconds")))
        c.execute("UPDATE relatorios SET grupo='TVDE' WHERE id=last_insert_rowid()")
        count += 1
    return count



@app.route("/importar", methods=["GET","POST"])
@login_required
def importar():
    if request.method == "POST":
        f = request.files.get("arquivo")
        week = request.form.get("semana","").strip()
        if not f or not f.filename or not week:
            flash("Selecione o CSV e informe a semana.")
            return redirect(url_for("importar"))
        raw = f.read().decode("utf-8-sig", errors="replace")
        rows = list(csv.DictReader(StringIO(raw)))
        headers = set(rows[0].keys()) if rows else set()
        c = db()
        if "UUID do motorista" in headers:
            platform = "Uber"
            qty = import_uber(c, rows, week)
        elif "Motorista" in headers and "Ganhos brutos (total)|€" in headers:
            platform = "Bolt"
            qty = import_bolt(c, rows, week)
        else:
            c.close()
            flash("Formato não reconhecido. Envie o CSV original da Uber ou Bolt.")
            return redirect(url_for("importar"))
        c.execute("INSERT INTO importacoes(plataforma,semana,arquivo,linhas,criado_em) VALUES(?,?,?,?,?)",
                  (platform, week, f.filename, qty, datetime.now().isoformat(timespec="seconds")))
        c.commit()
        c.close()
        flash(f"{platform}: {qty} motoristas importados com sucesso.")
        return redirect(url_for("relatorios"))
    body = """
<div class="card"><h2>Importar relatório semanal</h2>
<p class="muted">O sistema identifica automaticamente se o ficheiro é da Uber ou da Bolt.</p>
<form method="post" enctype="multipart/form-data" class="grid">
 <label>Semana<input type="week" name="semana" required></label>
 <label class="filebox">Relatório CSV<input type="file" name="arquivo" accept=".csv,text/csv" required></label>
 <div style="align-self:end"><button class="btn accent">Importar e processar</button></div>
</form></div>
<div class="card"><h2>Regras aplicadas</h2>
<p><b>Dinheiro em mãos:</b> identificado no CSV e considerado no valor final da plataforma.</p>
<p><b>Uber:</b> usa o campo “Pago a si” como valor final disponibilizado.</p>
<p><b>Bolt:</b> usa “Pagamento previsto” como valor final disponibilizado.</p></div>
"""
    return render("Importar relatórios", body)


@app.route("/relatorios")
@login_required
def relatorios():
    c = db()
    rows = c.execute("""SELECT r.*,m.nome motorista FROM relatorios r JOIN motoristas m ON m.id=r.motorista_id
                        ORDER BY r.id DESC""").fetchall()
    c.close()
    body = """
<div class="card"><div class="actions" style="justify-content:space-between;align-items:center">
<div><h2>Relatórios processados</h2><span class="muted">Uber e Bolt separados por motorista.</span></div>
<a class="btn accent" href="/importar">Nova importação</a></div>
<table><tr><th>Semana</th><th>Motorista</th><th>Plataforma</th><th>Bruto</th><th>Dinheiro em mãos</th><th>Comissão</th><th>Portagens</th><th>Líquido</th><th></th></tr>
{% for r in rows %}<tr><td>{{r.semana}}</td><td>{{r.motorista}}</td><td>{{r.plataforma}}</td>
<td>{{money(r.bruto)}}</td><td>{{money(r.dinheiro_maos)}}</td><td>{{money(r.comissao)}}</td><td>{{money(r.portagens)}}</td>
<td><b>{{money(r.liquido)}}</b></td><td><a class="btn secondary" href="/recibo/{{r.id}}">PDF</a></td></tr>{% endfor %}
</table></div>
"""
    return render("Relatórios", body, rows=rows, money=money)





@app.route("/motoristas", methods=["GET","POST"])
@login_required
def motoristas():
    c = db()
    if request.method == "POST":
        try:
            c.execute("""INSERT INTO motoristas(nome,email,telefone,iban,percentual,cartao_prio)
                         VALUES(?,?,?,?,?,?)""",
                      (request.form["nome"].strip(),request.form.get("email","").strip(),
                       request.form.get("telefone","").strip(),request.form.get("iban","").strip(),
                       num(request.form.get("percentual")),
                       request.form.get("cartao_prio","").replace(" ","").strip()))
            c.commit(); flash("Motorista cadastrado.")
        except sqlite3.IntegrityError:
            flash("Esse motorista já está cadastrado.")
        c.close()
        return redirect(url_for("motoristas"))
    rows = c.execute("SELECT * FROM motoristas ORDER BY nome").fetchall()
    c.close()
    body = """
<div class="card"><h2>Novo motorista</h2><form method="post" class="grid">
<label>Nome<input name="nome" required></label>
<label>E-mail<input name="email" type="email"></label>
<label>Telefone<input name="telefone"></label>
<label>IBAN<input name="iban"></label>
<label>Percentual (%)<input name="percentual" type="number" step="0.01"></label>
<label>Número do cartão PRIO<input name="cartao_prio" inputmode="numeric"></label>
<div style="align-self:end"><button>Cadastrar</button></div>
</form></div>
<div class="card"><h2>Motoristas</h2>
<table><tr><th>Nome</th><th>E-mail</th><th>Telefone</th><th>IBAN</th><th>Cartão PRIO</th><th>%</th></tr>
{% for r in rows %}
<tr><td><b>{{r.nome}}</b></td><td>{{r.email}}</td><td>{{r.telefone}}</td><td>{{r.iban}}</td><td>{{r.cartao_prio or '-'}}</td><td>{{r.percentual}}</td></tr>
{% endfor %}</table></div>
"""
    return render("Motoristas", body, rows=rows)


@app.route("/viaturas", methods=["GET","POST"])
@login_required
def viaturas():
    c = db()
    if request.method == "POST":
        try:
            c.execute("""INSERT INTO viaturas(matricula,marca,modelo,motorista_id,seguro_validade,ipo_validade,carta_verde_validade,km)
                         VALUES(?,?,?,?,?,?,?,?)""",
                      (request.form["matricula"].strip().upper(),request.form.get("marca",""),request.form.get("modelo",""),
                       request.form.get("motorista_id") or None,request.form.get("seguro_validade") or None,
                       request.form.get("ipo_validade") or None,request.form.get("carta_verde_validade") or None,
                       int(num(request.form.get("km")))))
            c.commit(); flash("Viatura cadastrada.")
        except sqlite3.IntegrityError:
            flash("Essa matrícula já está cadastrada.")
        c.close(); return redirect(url_for("viaturas"))
    rows = c.execute("""SELECT v.*,m.nome motorista FROM viaturas v LEFT JOIN motoristas m ON m.id=v.motorista_id
                        ORDER BY v.matricula""").fetchall()
    drivers = c.execute("SELECT * FROM motoristas WHERE ativo=1 ORDER BY nome").fetchall(); c.close()
    body = """
<div class="card"><h2>Nova viatura</h2><form method="post" class="grid">
<label>Matrícula<input name="matricula" required></label><label>Marca<input name="marca"></label><label>Modelo<input name="modelo"></label>
<label>Motorista<select name="motorista_id"><option value="">Sem motorista</option>{% for d in drivers %}<option value="{{d.id}}">{{d.nome}}</option>{% endfor %}</select></label>
<label>Seguro<input type="date" name="seguro_validade"></label><label>IPO<input type="date" name="ipo_validade"></label>
<label>Carta Verde<input type="date" name="carta_verde_validade"></label><label>Quilometragem<input type="number" name="km"></label>
<div style="align-self:end"><button>Cadastrar</button></div></form></div>
<div class="card"><h2>Viaturas</h2><table><tr><th>Matrícula</th><th>Viatura</th><th>Motorista</th><th>Seguro</th><th>IPO</th><th>Carta Verde</th></tr>
{% for r in rows %}<tr><td><b>{{r.matricula}}</b></td><td>{{r.marca}} {{r.modelo}}</td><td>{{r.motorista or '-'}}</td>
<td>{{r.seguro_validade or '-'}}</td><td>{{r.ipo_validade or '-'}}</td><td>{{r.carta_verde_validade or '-'}}</td></tr>{% endfor %}
</table></div>
"""
    return render("Viaturas", body, rows=rows, drivers=drivers)



@app.route("/combustivel", methods=["GET","POST"])
@login_required
def combustivel():
    c = db()
    if request.method == "POST":
        arquivo = request.files.get("arquivo")
        semana = request.form.get("semana","").strip()
        if not arquivo or not arquivo.filename or not semana:
            c.close()
            flash("Selecione o relatório PRIO e informe a semana.")
            return redirect(url_for("combustivel"))

        try:
            wb = load_workbook(BytesIO(arquivo.read()), data_only=True, read_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))

            header_index = None
            headers = None
            for i, row in enumerate(rows):
                normalized = [str(v or "").strip().upper() for v in row]
                if "CARTÃO" in normalized and "TOTAL" in normalized:
                    header_index = i
                    headers = normalized
                    break

            if header_index is None:
                raise ValueError("Não encontrei as colunas CARTÃO e TOTAL.")

            idx = {name: pos for pos, name in enumerate(headers)}
            inserted = pending = 0
            total_imported = 0.0

            for row in rows[header_index + 1:]:
                if not row or not any(v is not None and str(v).strip() for v in row):
                    continue

                card = str(row[idx["CARTÃO"]] or "").strip()
                if card.endswith(".0"):
                    card = card[:-2]
                card = card.replace(" ", "")
                if not card:
                    continue

                total_value = num(row[idx["TOTAL"]])
                if total_value == 0:
                    continue

                driver = c.execute(
                    "SELECT id FROM motoristas WHERE REPLACE(COALESCE(cartao_prio,''),' ','')=?",
                    (card,)
                ).fetchone()
                driver_id = driver["id"] if driver else None
                if driver_id is None:
                    pending += 1

                desc = str(row[idx["DESC. CARTÃO"]] or "") if "DESC. CARTÃO" in idx else ""
                posto = str(row[idx["POSTO"]] or "") if "POSTO" in idx else ""
                litros = num(row[idx["LITROS"]]) if "LITROS" in idx else 0
                data_value = str(row[idx["DATA"]] or "") if "DATA" in idx else ""

                c.execute("""INSERT INTO combustivel(
                    semana,motorista_id,cartao,descricao_cartao,data,posto,litros,total,arquivo,criado_em
                ) VALUES(?,?,?,?,?,?,?,?,?,?)""",
                (semana,driver_id,card,desc,data_value,posto,litros,total_value,
                 arquivo.filename,datetime.now().isoformat(timespec="seconds")))

                inserted += 1
                total_imported += total_value

            c.execute("""INSERT INTO importacoes_combustivel(
                semana,arquivo,lancamentos,total,pendentes,criado_em
            ) VALUES(?,?,?,?,?,?)""",
            (semana,arquivo.filename,inserted,total_imported,pending,
             datetime.now().isoformat(timespec="seconds")))
            c.commit()
            flash(f"PRIO importado: {inserted} lançamentos, {money(total_imported)} e {pending} pendência(s).")
        except Exception as exc:
            c.rollback()
            flash(f"Não foi possível importar o relatório: {exc}")

        c.close()
        return redirect(url_for("combustivel"))

    summary = c.execute("""SELECT semana,COUNT(*) lancamentos,SUM(total) total,
                           SUM(CASE WHEN motorista_id IS NULL THEN 1 ELSE 0 END) pendentes
                           FROM combustivel GROUP BY semana ORDER BY semana DESC""").fetchall()
    by_driver = c.execute("""SELECT f.semana,m.nome motorista,m.cartao_prio,
                            COUNT(*) abastecimentos,SUM(f.litros) litros,SUM(f.total) total
                            FROM combustivel f
                            JOIN motoristas m ON m.id=f.motorista_id
                            GROUP BY f.semana,m.id
                            ORDER BY f.semana DESC,total DESC""").fetchall()
    pending_rows = c.execute("""SELECT * FROM combustivel WHERE motorista_id IS NULL
                                ORDER BY id DESC""").fetchall()
    c.close()

    body = """
<div class="card"><h2>Importar relatório PRIO</h2>
<p class="muted">O sistema identifica o motorista pelo número do cartão e usa sempre a coluna <b>TOTAL</b>.</p>
<form method="post" enctype="multipart/form-data" class="grid">
<label>Semana<input type="week" name="semana" required></label>
<label class="filebox">Arquivo Excel PRIO<input type="file" name="arquivo" accept=".xlsx" required></label>
<div style="align-self:end"><button class="btn accent">Importar combustível</button></div>
</form></div>

<div class="card"><h2>Resumo por semana</h2>
<table><tr><th>Semana</th><th>Lançamentos</th><th>Total</th><th>Pendências</th></tr>
{% for s in summary %}
<tr><td>{{s.semana}}</td><td>{{s.lancamentos}}</td><td><b>{{money(s.total)}}</b></td><td>{{s.pendentes}}</td></tr>
{% endfor %}</table></div>

<div class="card"><h2>Combustível por motorista</h2>
<table><tr><th>Semana</th><th>Motorista</th><th>Cartão</th><th>Abastecimentos</th><th>Litros</th><th>Total</th></tr>
{% for r in by_driver %}
<tr><td>{{r.semana}}</td><td>{{r.motorista}}</td><td>{{r.cartao_prio}}</td><td>{{r.abastecimentos}}</td><td>{{"%.2f"|format(r.litros or 0)}}</td><td><b>{{money(r.total)}}</b></td></tr>
{% endfor %}</table></div>

<div class="card"><h2>Cartões não identificados</h2>
{% if pending_rows %}
<table><tr><th>Semana</th><th>Cartão</th><th>Descrição</th><th>Data</th><th>Total</th></tr>
{% for r in pending_rows %}
<tr><td>{{r.semana}}</td><td class="warn">{{r.cartao}}</td><td>{{r.descricao_cartao}}</td><td>{{r.data}}</td><td>{{money(r.total)}}</td></tr>
{% endfor %}</table>
{% else %}<p class="muted">Nenhuma pendência.</p>{% endif %}</div>
"""
    return render("Combustível", body, summary=summary, by_driver=by_driver, pending_rows=pending_rows, money=money)




@app.route("/sem-iban", methods=["GET", "POST"])
@login_required
def sem_iban():
    c = db()

    if request.method == "POST":
        driver_id = request.form.get("motorista_id")
        iban = iban_clean(request.form.get("iban", ""))
        email = request.form.get("email", "").strip()
        telefone = request.form.get("telefone", "").strip()

        if not driver_id:
            c.close()
            flash("Motorista não identificado.")
            return redirect(url_for("sem_iban"))

        if not iban:
            c.close()
            flash("Informe o IBAN.")
            return redirect(url_for("sem_iban"))

        if not iban_valid(iban):
            c.close()
            flash("O IBAN informado não é válido.")
            return redirect(url_for("sem_iban"))

        duplicate = c.execute(
            "SELECT id,nome FROM motoristas WHERE iban=? AND id<>?",
            (iban, driver_id)
        ).fetchone()

        if duplicate:
            c.close()
            flash(f"Este IBAN já está cadastrado para {duplicate['nome']}. Confira antes de continuar.")
            return redirect(url_for("sem_iban"))

        c.execute(
            "UPDATE motoristas SET iban=?,email=?,telefone=? WHERE id=?",
            (iban, email, telefone, driver_id)
        )
        c.commit()
        c.close()
        flash("IBAN cadastrado. O motorista já pode entrar na próxima simulação de pagamento.")
        return redirect(url_for("sem_iban"))

    rows = c.execute("""
        SELECT
            m.id,m.nome,m.email,m.telefone,m.iban,
            COALESCE(SUM(r.liquido),0) valor_pendente,
            COUNT(r.id) quantidade_relatorios,
            MAX(r.semana) ultima_semana
        FROM motoristas m
        LEFT JOIN relatorios r ON r.motorista_id=m.id
        WHERE TRIM(COALESCE(m.iban,''))=''
        GROUP BY m.id,m.nome,m.email,m.telefone,m.iban
        ORDER BY valor_pendente DESC,m.nome
    """).fetchall()

    invalid_rows = c.execute("""
        SELECT id,nome,email,telefone,iban
        FROM motoristas
        WHERE TRIM(COALESCE(iban,''))<>''
        ORDER BY nome
    """).fetchall()

    invalid_rows = [r for r in invalid_rows if not iban_valid(r["iban"])]
    c.close()

    body = """
<div class="grid">
  <div class="card">
    <div class="metric-label">Pessoas sem IBAN</div>
    <div class="metric-value">{{rows|length}}</div>
  </div>
  <div class="card">
    <div class="metric-label">IBAN inválido</div>
    <div class="metric-value">{{invalid_rows|length}}</div>
  </div>
</div>

<div class="card">
<h2>Cadastro rápido de IBAN</h2>
<p class="muted">
Todos os motoristas importados da Uber ou Bolt sem IBAN aparecem automaticamente aqui.
Ao salvar um IBAN válido, a pessoa deixa esta lista e passa a entrar na simulação de pagamentos.
</p>

{% if rows %}
{% for r in rows %}
<form method="post" class="grid" style="border-bottom:1px solid var(--line);padding:14px 0">
  <input type="hidden" name="motorista_id" value="{{r.id}}">
  <div>
    <label>Motorista</label>
    <div style="margin-top:8px"><b>{{r.nome}}</b></div>
    <div class="muted">{{r.quantidade_relatorios}} relatório(s) · {{r.ultima_semana or 'sem semana'}}</div>
  </div>
  <div>
    <label>Valor pendente</label>
    <div style="margin-top:8px;font-weight:800">{{money(r.valor_pendente)}}</div>
  </div>
  <label>IBAN
    <input name="iban" placeholder="PT50..." required>
  </label>
  <label>E-mail
    <input name="email" type="email" value="{{r.email or ''}}">
  </label>
  <label>Telefone
    <input name="telefone" value="{{r.telefone or ''}}">
  </label>
  <div style="align-self:end">
    <button class="btn accent">Salvar e liberar pagamento</button>
  </div>
</form>
{% endfor %}
{% else %}
<p class="ok">Não há motoristas sem IBAN.</p>
{% endif %}
</div>

<div class="card">
<h2>IBAN inválido</h2>
{% if invalid_rows %}
<table>
<tr><th>Motorista</th><th>IBAN atual</th><th>Contato</th></tr>
{% for r in invalid_rows %}
<tr>
  <td class="bad">{{r.nome}}</td>
  <td>{{r.iban}}</td>
  <td>{{r.email}}<br>{{r.telefone}}</td>
</tr>
{% endfor %}
</table>
<p class="muted">Corrija estes IBANs no cadastro do motorista ou remova-os para que voltem à lista Sem IBAN.</p>
{% else %}
<p class="ok">Nenhum IBAN inválido cadastrado.</p>
{% endif %}
</div>
"""
    return render("Sem IBAN", body, rows=rows, invalid_rows=invalid_rows, money=money)


def iban_clean(value):
    return "".join(ch for ch in str(value or "").upper() if ch.isalnum())


def iban_valid(value):
    iban = iban_clean(value)
    if len(iban) < 15 or len(iban) > 34:
        return False
    if not iban[:2].isalpha() or not iban[2:4].isdigit():
        return False
    rearranged = iban[4:] + iban[:4]
    converted = ""
    for ch in rearranged:
        converted += str(ord(ch) - 55) if ch.isalpha() else ch
    try:
        return int(converted) % 97 == 1
    except ValueError:
        return False


def week_fuel_by_driver(c, week):
    rows = c.execute("""
        SELECT motorista_id, COALESCE(SUM(total),0) total
        FROM combustivel
        WHERE semana=? AND motorista_id IS NOT NULL
        GROUP BY motorista_id
    """, (week,)).fetchall()
    return {r["motorista_id"]: float(r["total"] or 0) for r in rows}


def week_extra_by_driver(c, week, group):
    rows = c.execute("""
        SELECT motorista_id, COALESCE(SUM(valor),0) total
        FROM descontos_extras
        WHERE semana=? AND grupo=?
        GROUP BY motorista_id
    """, (week, group)).fetchall()
    return {r["motorista_id"]: float(r["total"] or 0) for r in rows}


def payment_preview(c, week, group, bank_fee):
    fuel = week_fuel_by_driver(c, week)
    extras = week_extra_by_driver(c, week, group)

    rows = c.execute("""
        SELECT r.motorista_id,m.nome,m.iban,
               COALESCE(SUM(r.bruto),0) bruto,
               COALESCE(SUM(r.dinheiro_maos),0) dinheiro,
               COALESCE(SUM(r.comissao),0) comissao,
               COALESCE(SUM(r.portagens),0) portagens,
               COALESCE(SUM(r.outros_descontos),0) outros,
               COALESCE(SUM(r.reembolsos),0) reembolsos,
               COALESCE(SUM(r.liquido),0) base_liquida
        FROM relatorios r
        JOIN motoristas m ON m.id=r.motorista_id
        WHERE r.semana=? AND COALESCE(r.grupo,'TVDE')=?
        GROUP BY r.motorista_id,m.nome,m.iban
        ORDER BY m.nome
    """, (week, group)).fetchall()

    people = []
    for r in rows:
        driver_id = r["motorista_id"]
        fuel_value = fuel.get(driver_id, 0.0) if group == "TVDE" else 0.0
        extra_value = extras.get(driver_id, 0.0)
        amount_before_fee = float(r["base_liquida"] or 0) - fuel_value - extra_value
        people.append({
            "motorista_id": driver_id,
            "nome": r["nome"],
            "iban": iban_clean(r["iban"]),
            "bruto": float(r["bruto"] or 0),
            "dinheiro": float(r["dinheiro"] or 0),
            "comissao": float(r["comissao"] or 0),
            "portagens": float(r["portagens"] or 0),
            "outros": float(r["outros"] or 0),
            "reembolsos": float(r["reembolsos"] or 0),
            "combustivel": fuel_value,
            "extras": extra_value,
            "antes_taxa": amount_before_fee,
        })

    grouped = {}
    no_iban = []
    invalid_iban = []
    for p in people:
        if not p["iban"]:
            no_iban.append(p)
            continue
        if not iban_valid(p["iban"]):
            invalid_iban.append(p)
            continue
        key = p["iban"]
        if key not in grouped:
            grouped[key] = {
                "iban": key, "nomes": [], "motoristas": [], "detalhes": [],
                "bruto": 0, "dinheiro": 0, "comissao": 0, "portagens": 0,
                "outros": 0, "reembolsos": 0, "combustivel": 0, "extras": 0,
                "antes_taxa": 0
            }
        g = grouped[key]
        g["nomes"].append(p["nome"])
        g["motoristas"].append(p["motorista_id"])
        g["detalhes"].append(dict(p))
        for field in ["bruto","dinheiro","comissao","portagens","outros","reembolsos","combustivel","extras","antes_taxa"]:
            g[field] += p[field]

    valid = []
    negatives = []
    warnings = []
    for g in grouped.values():
        g["taxa"] = float(bank_fee)
        g["valor_final"] = round(g["antes_taxa"] - g["taxa"], 2)
        g["nome_pagamento"] = g["nomes"][0]
        if len(set(n.lower() for n in g["nomes"])) > 1:
            warnings.append(g)
        if g["valor_final"] <= 0:
            negatives.append(g)
        else:
            valid.append(g)

    valid.sort(key=lambda x: x["valor_final"], reverse=True)
    return valid, no_iban, invalid_iban, negatives, warnings


def split_batches(payments, limit=50000.0, desired_parts=3):
    if not payments:
        return []
    desired_parts = max(1, int(desired_parts or 1))
    bins = [{"items": [], "total": 0.0} for _ in range(desired_parts)]

    for payment in sorted(payments, key=lambda x: x["valor_final"], reverse=True):
        candidates = [b for b in bins if b["total"] + payment["valor_final"] <= limit]
        if not candidates:
            bins.append({"items": [], "total": 0.0})
            candidates = [bins[-1]]
        target = min(candidates, key=lambda b: b["total"])
        target["items"].append(payment)
        target["total"] += payment["valor_final"]

    return [b for b in bins if b["items"]]


def xml_bytes(batch, group, part, execution_date, debtor_name, debtor_iban, debtor_bic):
    ns = "urn:iso:std:iso:20022:tech:xsd:pain.001.001.03"
    ET.register_namespace("", ns)
    q = lambda tag: f"{{{ns}}}{tag}"

    root = ET.Element(q("Document"))
    init = ET.SubElement(root, q("CstmrCdtTrfInitn"))
    grp = ET.SubElement(init, q("GrpHdr"))
    msg_id = f"{group}-{datetime.now().strftime('%Y%m%d%H%M%S')}-{part}"
    ET.SubElement(grp, q("MsgId")).text = msg_id
    ET.SubElement(grp, q("CreDtTm")).text = datetime.now().isoformat(timespec="seconds")
    ET.SubElement(grp, q("NbOfTxs")).text = str(len(batch["items"]))
    ET.SubElement(grp, q("CtrlSum")).text = f"{batch['total']:.2f}"
    init_party = ET.SubElement(grp, q("InitgPty"))
    ET.SubElement(init_party, q("Nm")).text = debtor_name

    pmt = ET.SubElement(init, q("PmtInf"))
    ET.SubElement(pmt, q("PmtInfId")).text = msg_id
    ET.SubElement(pmt, q("PmtMtd")).text = "TRF"
    ET.SubElement(pmt, q("BtchBookg")).text = "true"
    ET.SubElement(pmt, q("NbOfTxs")).text = str(len(batch["items"]))
    ET.SubElement(pmt, q("CtrlSum")).text = f"{batch['total']:.2f}"

    pmt_type = ET.SubElement(pmt, q("PmtTpInf"))
    svc = ET.SubElement(pmt_type, q("SvcLvl"))
    ET.SubElement(svc, q("Cd")).text = "SEPA"
    ET.SubElement(pmt, q("ReqdExctnDt")).text = execution_date

    debtor = ET.SubElement(pmt, q("Dbtr"))
    ET.SubElement(debtor, q("Nm")).text = debtor_name
    debtor_acct = ET.SubElement(pmt, q("DbtrAcct"))
    debtor_id = ET.SubElement(debtor_acct, q("Id"))
    ET.SubElement(debtor_id, q("IBAN")).text = iban_clean(debtor_iban)

    debtor_agent = ET.SubElement(pmt, q("DbtrAgt"))
    fin = ET.SubElement(debtor_agent, q("FinInstnId"))
    ET.SubElement(fin, q("BIC")).text = debtor_bic
    ET.SubElement(pmt, q("ChrgBr")).text = "SLEV"

    for idx, item in enumerate(batch["items"], 1):
        tx = ET.SubElement(pmt, q("CdtTrfTxInf"))
        pmt_id = ET.SubElement(tx, q("PmtId"))
        ET.SubElement(pmt_id, q("EndToEndId")).text = f"{group}-{part}-{idx}"
        amt = ET.SubElement(tx, q("Amt"))
        instd = ET.SubElement(amt, q("InstdAmt"), Ccy="EUR")
        instd.text = f"{item['valor_final']:.2f}"

        cdtr = ET.SubElement(tx, q("Cdtr"))
        ET.SubElement(cdtr, q("Nm")).text = item["nome_pagamento"][:70]
        acct = ET.SubElement(tx, q("CdtrAcct"))
        acct_id = ET.SubElement(acct, q("Id"))
        ET.SubElement(acct_id, q("IBAN")).text = item["iban"]

        purpose = ET.SubElement(tx, q("Purp"))
        ET.SubElement(purpose, q("Cd")).text = "SALA"
        rem = ET.SubElement(tx, q("RmtInf"))
        ET.SubElement(rem, q("Ustrd")).text = f"Pagamento {group}"

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


@app.route("/descontos-extra", methods=["POST"])
@login_required
def descontos_extra():
    c = db()
    driver_id = request.form.get("motorista_id")
    week = request.form.get("semana","").strip()
    group = request.form.get("grupo","TVDE").strip()
    category = request.form.get("categoria","Outro").strip()
    description = request.form.get("descricao","").strip()
    value = num(request.form.get("valor"))
    if driver_id and week and value:
        c.execute("""INSERT INTO descontos_extras(
            motorista_id,semana,grupo,categoria,descricao,valor,criado_em
        ) VALUES(?,?,?,?,?,?,?)""",
        (driver_id,week,group,category,description,value,datetime.now().isoformat(timespec="seconds")))
        c.commit()
        flash("Desconto extra adicionado.")
    else:
        flash("Preencha motorista, semana e valor.")
    c.close()
    return redirect(url_for("pagamentos", semana=week, grupo=group))


@app.route("/pagamentos", methods=["GET","POST"])
@login_required
def pagamentos():
    c = db()
    week = request.values.get("semana","").strip()
    group = request.values.get("grupo","TVDE").strip()
    bank_fee = num(request.values.get("taxa","1.25"))
    parts = int(num(request.values.get("partes","3")) or 3)

    valid = no_iban = invalid_iban = negatives = warnings = []
    batches = []
    if week:
        valid, no_iban, invalid_iban, negatives, warnings = payment_preview(c, week, group, bank_fee)
        batches = split_batches(valid, 50000.0, parts)

    drivers = c.execute("SELECT id,nome FROM motoristas ORDER BY nome").fetchall()
    extras = []
    if week:
        extras = c.execute("""SELECT d.*,m.nome motorista FROM descontos_extras d
                              JOIN motoristas m ON m.id=d.motorista_id
                              WHERE d.semana=? AND d.grupo=? ORDER BY d.id DESC""",
                           (week,group)).fetchall()
    c.close()

    body = """
<div class="card"><h2>Preparar pagamentos XML</h2>
<form method="get" class="grid">
<label>Semana<input type="week" name="semana" value="{{week}}" required></label>
<label>Grupo<select name="grupo"><option value="TVDE" {{'selected' if group=='TVDE' else ''}}>TVDE</option><option value="DELIVERY" {{'selected' if group=='DELIVERY' else ''}}>Uber Eats / Bolt Food</option></select></label>
<label>Taxa por transferência<input type="number" step="0.01" name="taxa" value="{{bank_fee}}"></label>
<label>Dividir inicialmente em<input type="number" min="1" name="partes" value="{{parts}}"></label>
<div style="align-self:end"><button>Simular e preparar XML</button></div>
</form></div>

<div class="card"><h2>Adicionar desconto extra</h2>
<form method="post" action="/descontos-extra" class="grid">
<input type="hidden" name="semana" value="{{week}}"><input type="hidden" name="grupo" value="{{group}}">
<label>Motorista<select name="motorista_id" required><option value="">Selecione</option>{% for d in drivers %}<option value="{{d.id}}">{{d.nome}}</option>{% endfor %}</select></label>
<label>Categoria<select name="categoria"><option>Combustível extra</option><option>Adiantamento</option><option>Multa</option><option>Renda da viatura</option><option>Outro</option></select></label>
<label>Descrição<input name="descricao"></label>
<label>Valor<input type="number" step="0.01" name="valor" required></label>
<div style="align-self:end"><button>Adicionar desconto</button></div>
</form></div>

{% if week %}
<div class="grid">
<div class="card"><div class="metric-label">Transferências válidas</div><div class="metric-value">{{valid|length}}</div></div>
<a href="/sem-iban" style="text-decoration:none;color:inherit"><div class="card"><div class="metric-label">Sem IBAN</div><div class="metric-value">{{no_iban|length}}</div><div class="muted">Clique para cadastrar</div></div></a>
<div class="card"><div class="metric-label">IBAN inválido</div><div class="metric-value">{{invalid_iban|length}}</div></div>
<div class="card"><div class="metric-label">Negativos / não pagos</div><div class="metric-value">{{negatives|length}}</div></div>
<div class="card"><div class="metric-label">XMLs previstos</div><div class="metric-value">{{batches|length}}</div></div>
</div>

<div class="card"><h2>Divisão prevista</h2>
<table><tr><th>Parte</th><th>Transferências</th><th>Total</th></tr>
{% for b in batches %}<tr><td>{{loop.index}}</td><td>{{b.items|length}}</td><td><b>{{money(b.total)}}</b></td></tr>{% endfor %}
</table>
</div>

<div class="card">
<h2>Exportar pagamentos em XML</h2>
{% if valid %}
<p class="muted">
Os pagamentos válidos serão exportados. Pessoas sem IBAN, com IBAN inválido ou com valor negativo
ficarão fora dos XMLs e continuarão listadas como pendência.
</p>
<form method="post" action="/gerar-xml" class="grid" style="margin-top:16px">
<input type="hidden" name="semana" value="{{week}}">
<input type="hidden" name="grupo" value="{{group}}">
<input type="hidden" name="taxa" value="{{bank_fee}}">
<input type="hidden" name="partes" value="{{parts}}">
<label>Data de execução<input type="date" name="data_execucao" required></label>
<label>Nome da empresa pagadora<input name="debtor_name" required></label>
<label>IBAN da empresa<input name="debtor_iban" required></label>
<label>BIC/SWIFT<input name="debtor_bic" required></label>
{% if warnings %}
<label style="grid-column:1/-1">
<input type="checkbox" name="confirmar_iban_nomes" value="1" style="width:auto;margin-right:8px">
Confirmo que revisei os IBANs associados a nomes diferentes.
</label>
{% endif %}
<div style="align-self:end"><button class="btn accent">Gerar e baixar XMLs</button></div>
</form>
{% else %}
<p class="bad">Não existem pagamentos válidos para exportar nesta simulação.</p>
<p class="muted">Cadastre ou corrija os IBANs e confira os valores negativos antes de tentar novamente.</p>
{% endif %}
</div>

<div class="card"><h2>Pagamentos consolidados por IBAN</h2>
<table><tr><th>Beneficiário</th><th>IBAN</th><th>Bruto</th><th>Combustível</th><th>Extras</th><th>Taxa</th><th>Final</th></tr>
{% for r in valid %}<tr><td>{{r.nomes|join(', ')}}</td><td>{{r.iban}}</td><td>{{money(r.bruto)}}</td><td>{{money(r.combustivel)}}</td><td>{{money(r.extras)}}</td><td>{{money(r.taxa)}}</td><td><b>{{money(r.valor_final)}}</b></td></tr>{% endfor %}
</table></div>

<div class="card"><div class="actions" style="justify-content:space-between;align-items:center"><h2>Sem IBAN</h2><a class="btn accent" href="/sem-iban">Cadastrar IBAN</a></div>
<table><tr><th>Motorista</th><th>Valor antes da taxa</th></tr>
{% for r in no_iban %}<tr><td class="warn">{{r.nome}}</td><td>{{money(r.antes_taxa)}}</td></tr>{% endfor %}
</table></div>

<div class="card"><h2>IBAN inválido</h2>
<table><tr><th>Motorista</th><th>IBAN</th><th>Valor</th></tr>
{% for r in invalid_iban %}<tr><td class="bad">{{r.nome}}</td><td>{{r.iban}}</td><td>{{money(r.antes_taxa)}}</td></tr>{% endfor %}
</table></div>

<div class="card"><h2>Negativos / não pagos</h2>
<table><tr><th>Motorista(s)</th><th>IBAN</th><th>Valor final</th></tr>
{% for r in negatives %}<tr><td class="bad">{{r.nomes|join(', ')}}</td><td>{{r.iban}}</td><td>{{money(r.valor_final)}}</td></tr>{% endfor %}
</table></div>

<div class="card"><h2>Mesmo IBAN com nomes diferentes</h2>
<table><tr><th>Nomes</th><th>IBAN</th><th>Total</th></tr>
{% for r in warnings %}<tr><td class="warn">{{r.nomes|join(', ')}}</td><td>{{r.iban}}</td><td>{{money(r.valor_final)}}</td></tr>{% endfor %}
</table></div>

<div class="card"><h2>Descontos extras lançados</h2>
<table><tr><th>Motorista</th><th>Categoria</th><th>Descrição</th><th>Valor</th></tr>
{% for e in extras %}<tr><td>{{e.motorista}}</td><td>{{e.categoria}}</td><td>{{e.descricao}}</td><td>{{money(e.valor)}}</td></tr>{% endfor %}
</table></div>
{% endif %}
"""
    return render("Pagamentos XML", body, week=week, group=group, bank_fee=bank_fee, parts=parts,
                  valid=valid, no_iban=no_iban, invalid_iban=invalid_iban, negatives=negatives,
                  warnings=warnings, batches=batches, drivers=drivers, extras=extras, money=money)


@app.route("/gerar-xml", methods=["POST"])
@login_required
def gerar_xml():
    week = request.form.get("semana","").strip()
    group = request.form.get("grupo","TVDE").strip()
    bank_fee = num(request.form.get("taxa","1.25"))
    parts = int(num(request.form.get("partes","3")) or 3)
    execution_date = request.form.get("data_execucao","").strip()
    debtor_name = request.form.get("debtor_name","").strip()
    debtor_iban = request.form.get("debtor_iban","").strip()
    debtor_bic = request.form.get("debtor_bic","").strip()

    if not all([week,execution_date,debtor_name,debtor_iban,debtor_bic]):
        flash("Preencha todos os dados da empresa pagadora.")
        return redirect(url_for("pagamentos",semana=week,grupo=group))

    c = db()
    valid, no_iban, invalid_iban, negatives, warnings = payment_preview(c, week, group, bank_fee)
    confirmar_nomes = request.form.get("confirmar_iban_nomes") == "1"
    if warnings and not confirmar_nomes:
        c.close()
        flash("Confirme que revisou os IBANs associados a nomes diferentes.")
        return redirect(url_for("pagamentos",semana=week,grupo=group,taxa=bank_fee,partes=parts))

    if not valid:
        c.close()
        flash("Não existem pagamentos válidos para gerar XML.")
        return redirect(url_for("pagamentos",semana=week,grupo=group,taxa=bank_fee,partes=parts))

    batches = split_batches(valid, 50000.0, parts)
    output = BytesIO()
    with ZipFile(output, "w", ZIP_DEFLATED) as z:
        for i,batch in enumerate(batches,1):
            filename = f"{group}_{week}_PARTE_{i:02d}.xml"
            content = xml_bytes(batch, group, i, execution_date, debtor_name, debtor_iban, debtor_bic)
            z.writestr(filename, content)
            created_at = datetime.now().isoformat(timespec="seconds")
            c.execute("""INSERT INTO xml_historico(semana,grupo,arquivo,pagamentos,total,criado_em)
                         VALUES(?,?,?,?,?,?)""",
                      (week,group,filename,len(batch["items"]),batch["total"],created_at))

            for payment in batch["items"]:
                cur = c.execute("""INSERT INTO pagamentos_consolidados(
                    semana,grupo,arquivo,parte,iban,nome_pagamento,taxa,valor_transferido,criado_em
                ) VALUES(?,?,?,?,?,?,?,?,?)""",
                (week,group,filename,i,payment["iban"],payment["nome_pagamento"],
                 payment["taxa"],payment["valor_final"],created_at))
                pagamento_id = cur.lastrowid

                detalhes = payment.get("detalhes", [])
                taxa_rateada = payment["taxa"] / len(detalhes) if detalhes else 0
                for detail in detalhes:
                    valor_individual = detail["antes_taxa"] - taxa_rateada
                    plataformas = c.execute("""SELECT GROUP_CONCAT(DISTINCT plataforma)
                                               FROM relatorios
                                               WHERE motorista_id=? AND semana=?
                                               AND COALESCE(grupo,'TVDE')=?""",
                                            (detail["motorista_id"],week,group)).fetchone()[0] or ""
                    c.execute("""INSERT INTO pagamentos_consolidados_itens(
                        pagamento_id,motorista_id,motorista_nome,plataforma,bruto,dinheiro_maos,
                        comissao,combustivel,descontos_extras,reembolsos,valor_liquido
                    ) VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                    (pagamento_id,detail["motorista_id"],detail["nome"],plataformas,
                     detail["bruto"],detail["dinheiro"],detail["comissao"],
                     detail["combustivel"],detail["extras"],detail["reembolsos"],
                     round(valor_individual,2)))
    c.commit()
    c.close()
    output.seek(0)
    return send_file(output,mimetype="application/zip",as_attachment=True,
                     download_name=f"XML_{group}_{week}.zip")



@app.route("/auditoria-pagamentos")
@login_required
def auditoria_pagamentos():
    c = db()
    busca = request.args.get("busca","").strip()
    semana = request.args.get("semana","").strip()
    grupo = request.args.get("grupo","").strip()

    sql = """SELECT p.*,
             (SELECT COUNT(*) FROM pagamentos_consolidados_itens i WHERE i.pagamento_id=p.id) itens
             FROM pagamentos_consolidados p WHERE 1=1"""
    params = []
    if busca:
        sql += """ AND (p.iban LIKE ? OR p.nome_pagamento LIKE ? OR p.arquivo LIKE ?
                   OR EXISTS(SELECT 1 FROM pagamentos_consolidados_itens i
                             WHERE i.pagamento_id=p.id AND i.motorista_nome LIKE ?))"""
        term = f"%{busca}%"
        params.extend([term,term,term,term])
    if semana:
        sql += " AND p.semana=?"
        params.append(semana)
    if grupo:
        sql += " AND p.grupo=?"
        params.append(grupo)
    sql += " ORDER BY p.id DESC"

    rows = c.execute(sql, params).fetchall()
    total = sum(float(r["valor_transferido"] or 0) for r in rows)
    c.close()

    body = """
<div class="card">
 <div class="card-header"><h2>Pesquisar pagamentos consolidados</h2></div>
 <div class="card-body">
  <form method="get" class="grid">
   <label>Nome, IBAN ou XML<input name="busca" value="{{busca}}" placeholder="Pesquisar..."></label>
   <label>Semana<input type="week" name="semana" value="{{semana}}"></label>
   <label>Grupo<select name="grupo"><option value="">Todos</option><option value="TVDE" {{'selected' if grupo=='TVDE' else ''}}>TVDE</option><option value="DELIVERY" {{'selected' if grupo=='DELIVERY' else ''}}>Delivery</option></select></label>
   <div style="align-self:end"><button class="btn primary">Pesquisar</button></div>
  </form>
 </div>
</div>

<div class="kpi-grid" style="grid-template-columns:repeat(3,minmax(180px,1fr))">
 <div class="kpi blue"><h3>Transferências</h3><div class="value">{{rows|length}}</div><small>Pagamentos encontrados</small></div>
 <div class="kpi green"><h3>Valor transferido</h3><div class="value">{{money(total)}}</div><small>Total da pesquisa</small></div>
 <div class="kpi purple"><h3>Composições</h3><div class="value">{{rows|sum(attribute='itens')}}</div><small>Motoristas incluídos</small></div>
</div>

<div class="card">
 <div class="card-header"><h2>Pagamentos por IBAN</h2></div>
 <div class="card-body">
  <table>
   <tr><th>Data/hora</th><th>Semana</th><th>Grupo</th><th>Beneficiário</th><th>IBAN</th><th>XML</th><th>Pessoas</th><th>Total</th><th></th></tr>
   {% for r in rows %}
   <tr>
    <td>{{r.criado_em}}</td><td>{{r.semana}}</td><td>{{r.grupo}}</td><td>{{r.nome_pagamento}}</td>
    <td>{{r.iban}}</td><td>{{r.arquivo}}</td><td>{{r.itens}}</td><td><b>{{money(r.valor_transferido)}}</b></td>
    <td><a class="btn primary" href="/auditoria-pagamentos/{{r.id}}">Abrir composição</a></td>
   </tr>
   {% endfor %}
  </table>
 </div>
</div>
"""
    return render("Auditoria de Pagamentos", body, rows=rows, total=total, busca=busca,
                  semana=semana, grupo=grupo, money=money)


@app.route("/auditoria-pagamentos/<int:payment_id>")
@login_required
def auditoria_pagamento_detalhe(payment_id):
    c = db()
    payment = c.execute("SELECT * FROM pagamentos_consolidados WHERE id=?", (payment_id,)).fetchone()
    if not payment:
        c.close()
        return "Pagamento não encontrado", 404

    items = c.execute("""SELECT * FROM pagamentos_consolidados_itens
                         WHERE pagamento_id=? ORDER BY motorista_nome""",
                      (payment_id,)).fetchall()
    c.close()

    body = """
<div class="grid">
 <div class="card"><div class="card-body"><div class="metric-label">IBAN</div><div style="font-size:18px;font-weight:800;margin-top:8px">{{payment.iban}}</div></div></div>
 <div class="card"><div class="card-body"><div class="metric-label">Valor transferido</div><div class="metric-value">{{money(payment.valor_transferido)}}</div></div></div>
 <div class="card"><div class="card-body"><div class="metric-label">XML</div><div style="font-weight:800;margin-top:8px">{{payment.arquivo}}</div></div></div>
 <div class="card"><div class="card-body"><div class="metric-label">Data e hora</div><div style="font-weight:800;margin-top:8px">{{payment.criado_em}}</div></div></div>
</div>

<div class="card">
 <div class="card-header"><h2>Composição individual do pagamento</h2></div>
 <div class="card-body">
  <table>
   <tr><th>Motorista</th><th>Plataforma</th><th>Bruto</th><th>Dinheiro</th><th>Comissão</th><th>Combustível</th><th>Extras</th><th>Reembolso</th><th>Líquido individual</th></tr>
   {% for i in items %}
   <tr>
    <td><b>{{i.motorista_nome}}</b></td><td>{{i.plataforma}}</td><td>{{money(i.bruto)}}</td>
    <td>{{money(i.dinheiro_maos)}}</td><td>{{money(i.comissao)}}</td><td>{{money(i.combustivel)}}</td>
    <td>{{money(i.descontos_extras)}}</td><td>{{money(i.reembolsos)}}</td><td><b>{{money(i.valor_liquido)}}</b></td>
   </tr>
   {% endfor %}
  </table>
  <div style="margin-top:16px"><a class="btn" href="/auditoria-pagamentos">Voltar</a></div>
 </div>
</div>
"""
    return render("Composição do Pagamento", body, payment=payment, items=items, money=money)


@app.route("/recibos")
@login_required
def recibos():
    c = db()
    rows = c.execute("""SELECT r.*,m.nome motorista FROM relatorios r JOIN motoristas m ON m.id=r.motorista_id
                        ORDER BY r.id DESC""").fetchall(); c.close()
    body = """
<div class="card"><h2>Recibos disponíveis</h2><table><tr><th>Motorista</th><th>Semana</th><th>Plataforma</th><th>Líquido</th><th></th></tr>
{% for r in rows %}<tr><td>{{r.motorista}}</td><td>{{r.semana}}</td><td>{{r.plataforma}}</td><td><b>{{money(r.liquido)}}</b></td>
<td><a class="btn" href="/recibo/{{r.id}}">Abrir PDF</a></td></tr>{% endfor %}</table></div>
"""
    return render("Recibos", body, rows=rows, money=money)


@app.route("/recibo/<int:rid>")
@login_required
def recibo(rid):
    c = db()
    r = c.execute("""SELECT r.*,m.nome motorista,m.iban FROM relatorios r JOIN motoristas m ON m.id=r.motorista_id
                     WHERE r.id=?""",(rid,)).fetchone(); c.close()
    if not r: return "Recibo não encontrado",404
    buffer = BytesIO(); pdf = canvas.Canvas(buffer,pagesize=A4); y=800
    pdf.setFont("Helvetica-Bold",18); pdf.drawString(50,y,"Irmãos Fleet - Recibo semanal"); y-=36
    pdf.setFont("Helvetica",11)
    data=[("Motorista",r["motorista"]),("IBAN",r["iban"] or "-"),("Plataforma",r["plataforma"]),
          ("Semana",r["semana"]),("Valor bruto",money(r["bruto"])),("Dinheiro em mãos","- "+money(r["dinheiro_maos"])),
          ("Comissão","- "+money(r["comissao"])),("Portagens",money(r["portagens"])),
          ("Outros descontos","- "+money(r["outros_descontos"])),("Reembolsos","+ "+money(r["reembolsos"]))]
    for k,v in data:
        pdf.drawString(50,y,k); pdf.drawRightString(540,y,str(v)); y-=24
    y-=8; pdf.line(50,y,540,y); y-=32; pdf.setFont("Helvetica-Bold",14)
    pdf.drawString(50,y,"Valor líquido"); pdf.drawRightString(540,y,money(r["liquido"]))
    pdf.showPage(); pdf.save(); buffer.seek(0)
    return send_file(buffer,mimetype="application/pdf",as_attachment=False,
                     download_name=f"recibo_{r['motorista']}_{r['semana']}.pdf")


init_db()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT",5000)), debug=False)
