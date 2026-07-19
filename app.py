
from flask import (
    Flask, request, redirect, url_for, render_template_string,
    send_file, flash, session
)
from functools import wraps
from io import BytesIO, StringIO
from datetime import date, datetime, timedelta
import csv
import hmac
import os
import sqlite3
import unicodedata

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
    """)
    cols = [r[1] for r in c.execute("PRAGMA table_info(motoristas)").fetchall()]
    if "cartao_prio" not in cols:
        c.execute("ALTER TABLE motoristas ADD COLUMN cartao_prio TEXT")
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
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{{ title }} · Irmãos Fleet</title>
<style>
:root{
 --bg:#f3f6fb;--panel:#fff;--nav:#0b1220;--nav2:#131e31;
 --text:#182230;--muted:#667085;--line:#e6eaf0;--accent:#18c8d8;
 --accent2:#84f438;--danger:#d92d20;--warn:#dc6803;--ok:#079455
}
*{box-sizing:border-box} body{margin:0;font-family:Inter,Arial,sans-serif;background:var(--bg);color:var(--text)}
.shell{display:grid;grid-template-columns:260px 1fr;min-height:100vh}
.sidebar{background:linear-gradient(180deg,var(--nav),var(--nav2));color:#fff;padding:18px;position:sticky;top:0;height:100vh}
.brand{display:flex;align-items:center;gap:12px;margin-bottom:24px}
.brand img{width:54px;height:54px;border-radius:14px;object-fit:cover;border:1px solid #ffffff28}
.brand b{font-size:18px}.brand small{display:block;color:#aeb8c8;margin-top:3px}
.menu{display:grid;gap:7px}.menu a{color:#c9d2df;text-decoration:none;padding:12px;border-radius:10px;font-weight:700}
.menu a:hover,.menu a.active{background:#ffffff12;color:#fff}.logout{position:absolute;left:18px;right:18px;bottom:18px}
.main{min-width:0}.top{height:72px;background:#fff;border-bottom:1px solid var(--line);display:flex;align-items:center;justify-content:space-between;padding:0 24px;position:sticky;top:0;z-index:4}
.top h1{font-size:20px;margin:0}.top .user{color:var(--muted);font-size:14px}
.content{max-width:1280px;margin:auto;padding:24px}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:14px}
.card{background:var(--panel);border:1px solid var(--line);border-radius:16px;padding:18px;margin-bottom:18px}
.metric-label{color:var(--muted);font-size:13px}.metric-value{font-size:28px;font-weight:800;margin-top:8px}
h2{font-size:18px;margin:0 0 14px}.muted{color:var(--muted)}
table{width:100%;border-collapse:collapse}th,td{padding:11px 9px;border-bottom:1px solid var(--line);text-align:left;font-size:14px}
th{font-size:11px;text-transform:uppercase;color:var(--muted)}
label{display:block;font-size:12px;font-weight:800;color:#344054}
input,select{width:100%;margin-top:6px;padding:11px;border:1px solid #d0d5dd;border-radius:10px;background:#fff}
button,.btn{border:0;border-radius:10px;background:#111827;color:#fff;padding:11px 15px;text-decoration:none;font-weight:800;cursor:pointer;display:inline-block}
.btn.secondary{background:#475467}.btn.accent{background:linear-gradient(90deg,var(--accent),var(--accent2));color:#07111d}
.actions{display:flex;gap:8px;flex-wrap:wrap}.ok{color:var(--ok);font-weight:800}.warn{color:var(--warn);font-weight:800}.bad{color:var(--danger);font-weight:800}
.flash{background:#ecfdf3;border:1px solid #abefc6;padding:12px;border-radius:10px;margin-bottom:14px}
.filebox{border:2px dashed #b9c4d1;border-radius:14px;padding:18px;background:#f8fafc}
.mobilebar{display:none}
@media(max-width:820px){
 .shell{display:block}.sidebar{display:none}.top{padding:0 14px}.content{padding:14px}
 .mobilebar{display:flex;overflow-x:auto;gap:8px;background:#0b1220;padding:10px;position:sticky;top:72px;z-index:3}
 .mobilebar a{white-space:nowrap;color:#dbe4ef;text-decoration:none;background:#ffffff12;padding:9px 11px;border-radius:9px;font-size:13px;font-weight:800}
 table{display:block;overflow-x:auto;white-space:nowrap}
}
</style>
</head>
<body>
<div class="shell">
<aside class="sidebar">
 <div class="brand"><img src="/static/logo.png"><div><b>Irmãos Fleet</b><small>Gestão TVDE</small></div></div>
 <nav class="menu">
  <a href="/">▦ Dashboard</a>
  <a href="/importar">⇧ Importar relatórios</a>
  <a href="/relatorios">€ Relatórios</a>
  <a href="/motoristas">● Motoristas</a>
  <a href="/viaturas">▣ Viaturas</a>
  <a href="/recibos">▤ Recibos</a>
  <a href="/combustivel">⛽ Combustível</a>
 </nav>
 <div class="logout"><a class="btn secondary" style="width:100%;text-align:center" href="/logout">Sair</a></div>
</aside>
<section class="main">
 <header class="top"><h1>{{ title }}</h1><div class="user">Administrador · Irmãos Fleet</div></header>
 <nav class="mobilebar">
  <a href="/">Dashboard</a><a href="/importar">Importar</a><a href="/relatorios">Relatórios</a>
  <a href="/motoristas">Motoristas</a><a href="/viaturas">Viaturas</a><a href="/recibos">Recibos</a><a href="/combustivel">Combustível</a><a href="/logout">Sair</a>
 </nav>
 <main class="content">
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
    totals = c.execute("""SELECT COALESCE(SUM(bruto),0) bruto, COALESCE(SUM(dinheiro_maos),0) dinheiro,
                         COALESCE(SUM(liquido),0) liquido FROM relatorios""").fetchone()
    fuel_total = c.execute("SELECT COALESCE(SUM(total),0) total FROM combustivel").fetchone()["total"]
    fuel_pending = c.execute("SELECT COUNT(*) n FROM combustivel WHERE motorista_id IS NULL").fetchone()["n"]
    recent = c.execute("""SELECT r.*,m.nome motorista FROM relatorios r JOIN motoristas m ON m.id=r.motorista_id
                          ORDER BY r.id DESC LIMIT 8""").fetchall()
    alerts = []
    today = date.today()
    for v in c.execute("SELECT * FROM viaturas ORDER BY matricula").fetchall():
        for field, label in [("seguro_validade","Seguro"),("ipo_validade","IPO"),("carta_verde_validade","Carta Verde")]:
            if v[field]:
                try:
                    days = (datetime.strptime(v[field], "%Y-%m-%d").date() - today).days
                    if days <= 10:
                        alerts.append((v["matricula"], label, days))
                except ValueError:
                    pass
    c.close()
    body = """
<div class="grid">
 <div class="card"><div class="metric-label">Motoristas ativos</div><div class="metric-value">{{counts.motoristas}}</div></div>
 <div class="card"><div class="metric-label">Viaturas</div><div class="metric-value">{{counts.viaturas}}</div></div>
 <div class="card"><div class="metric-label">Total bruto importado</div><div class="metric-value">{{money(totals.bruto)}}</div></div>
 <div class="card"><div class="metric-label">Dinheiro em mãos</div><div class="metric-value">{{money(totals.dinheiro)}}</div></div>
 <div class="card"><div class="metric-label">Total líquido</div><div class="metric-value">{{money(totals.liquido)}}</div></div>
 <div class="card"><div class="metric-label">Combustível</div><div class="metric-value">{{money(fuel_total)}}</div><div class="muted">{{fuel_pending}} pendência(s)</div></div>
</div>
<div class="grid" style="grid-template-columns:1.4fr 1fr">
 <section class="card"><h2>Relatórios recentes</h2>
 <table><tr><th>Motorista</th><th>Plataforma</th><th>Semana</th><th>Líquido</th></tr>
 {% for r in recent %}<tr><td>{{r.motorista}}</td><td>{{r.plataforma}}</td><td>{{r.semana}}</td><td><b>{{money(r.liquido)}}</b></td></tr>{% endfor %}
 </table></section>
 <section class="card"><h2>Alertas</h2>
 {% if alerts %}<table><tr><th>Matrícula</th><th>Documento</th><th>Estado</th></tr>
 {% for a in alerts %}<tr><td>{{a[0]}}</td><td>{{a[1]}}</td><td class="{{'bad' if a[2]<0 else 'warn'}}">{{'Vencido' if a[2]<0 else 'Vence em '~a[2]~' dias'}}</td></tr>{% endfor %}</table>
 {% else %}<p class="muted">Nenhum documento vencendo em 10 dias.</p>{% endif %}</section>
</div>
"""
    return render("Dashboard", body, counts=counts, totals=totals, recent=recent, alerts=alerts, money=money, fuel_total=fuel_total, fuel_pending=fuel_pending)


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
