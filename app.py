
from flask import Flask, request, redirect, url_for, render_template_string, send_file, flash
import sqlite3
import os
from datetime import date, datetime
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

app = Flask(__name__)
app.secret_key = "tvde-manager-local"
DB = os.path.join(os.path.dirname(__file__), "tvde_manager.db")

BASE = """
<!doctype html>
<html lang="pt">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{{ title }} - TVDE Manager</title>
<style>
:root{--bg:#f4f6f8;--card:#fff;--dark:#101828;--muted:#667085;--line:#e4e7ec;--ok:#067647;--warn:#b54708;--bad:#b42318}
*{box-sizing:border-box}
body{margin:0;font-family:Arial,Helvetica,sans-serif;background:var(--bg);color:var(--dark)}
header{background:#111827;color:#fff;padding:16px 22px;position:sticky;top:0;z-index:5}
header h1{margin:0 0 10px;font-size:22px}
nav{display:flex;gap:14px;flex-wrap:wrap}
nav a{color:white;text-decoration:none;font-weight:700;font-size:14px}
main{max-width:1180px;margin:24px auto;padding:0 16px}
.card{background:var(--card);border-radius:14px;padding:18px;margin-bottom:18px;box-shadow:0 3px 14px rgba(16,24,40,.06)}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:14px}
.kpi{font-size:30px;font-weight:800;margin-top:8px}
.muted{color:var(--muted)}
table{width:100%;border-collapse:collapse}
th,td{padding:11px 9px;border-bottom:1px solid var(--line);text-align:left;font-size:14px}
th{font-size:12px;color:var(--muted);text-transform:uppercase}
input,select{width:100%;padding:10px;border:1px solid #d0d5dd;border-radius:8px;background:white}
label{font-size:13px;font-weight:700;display:block}
button,.btn{border:0;border-radius:8px;background:#111827;color:#fff;padding:10px 14px;text-decoration:none;display:inline-block;cursor:pointer;font-weight:700}
.btn.secondary{background:#475467}.btn.danger{background:#b42318}
.actions{display:flex;gap:8px;flex-wrap:wrap}
.ok{color:var(--ok);font-weight:700}.warn{color:var(--warn);font-weight:700}.bad{color:var(--bad);font-weight:700}
.flash{background:#ecfdf3;border:1px solid #abefc6;padding:10px;border-radius:8px;margin-bottom:14px}
@media(max-width:700px){table{display:block;overflow-x:auto;white-space:nowrap}}
</style>
</head>
<body>
<header>
<h1>TVDE Manager</h1>
<nav>
<a href="/">Dashboard</a>
<a href="/motoristas">Motoristas</a>
<a href="/viaturas">Viaturas</a>
<a href="/relatorios">Relatórios</a>
<a href="/recibos">Recibos</a>
</nav>
</header>
<main>
{% with messages = get_flashed_messages() %}
{% for message in messages %}<div class="flash">{{ message }}</div>{% endfor %}
{% endwith %}
{{ body|safe }}
</main>
</body>
</html>
"""

def conn():
    c = sqlite3.connect(DB)
    c.row_factory = sqlite3.Row
    return c

def init_db():
    c = conn()
    c.executescript("""
    CREATE TABLE IF NOT EXISTS motoristas(
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      nome TEXT NOT NULL,
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
      combustivel REAL DEFAULT 0,
      portagens REAL DEFAULT 0,
      outros_descontos REAL DEFAULT 0,
      reembolsos REAL DEFAULT 0,
      liquido REAL DEFAULT 0,
      criado_em TEXT NOT NULL
    );
    """)
    c.commit()
    c.close()

def money(v):
    return f"€ {float(v or 0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def render(title, body, **ctx):
    page = render_template_string(body, **ctx)
    return render_template_string(BASE, title=title, body=page)

@app.route("/")
def dashboard():
    c = conn()
    motoristas = c.execute("SELECT COUNT(*) n FROM motoristas WHERE ativo=1").fetchone()["n"]
    viaturas = c.execute("SELECT COUNT(*) n FROM viaturas").fetchone()["n"]
    totais = c.execute("""
      SELECT COALESCE(SUM(bruto),0) bruto,
             COALESCE(SUM(dinheiro_maos),0) dinheiro,
             COALESCE(SUM(liquido),0) liquido
      FROM relatorios
    """).fetchone()
    alertas = []
    hoje = date.today()
    for v in c.execute("SELECT * FROM viaturas ORDER BY matricula").fetchall():
        for campo, nome in [("seguro_validade","Seguro"),("ipo_validade","IPO"),("carta_verde_validade","Carta Verde")]:
            valor = v[campo]
            if valor:
                try:
                    dias = (datetime.strptime(valor, "%Y-%m-%d").date() - hoje).days
                    if dias <= 10:
                        alertas.append((v["matricula"], nome, dias))
                except ValueError:
                    pass
    c.close()
    body = """
    <div class="grid">
      <div class="card"><div class="muted">Motoristas ativos</div><div class="kpi">{{ motoristas }}</div></div>
      <div class="card"><div class="muted">Viaturas</div><div class="kpi">{{ viaturas }}</div></div>
      <div class="card"><div class="muted">Total bruto</div><div class="kpi">{{ bruto }}</div></div>
      <div class="card"><div class="muted">Dinheiro em mãos</div><div class="kpi">{{ dinheiro }}</div></div>
      <div class="card"><div class="muted">Total líquido</div><div class="kpi">{{ liquido }}</div></div>
    </div>
    <div class="card">
      <h2>Alertas de documentos</h2>
      {% if alertas %}
      <table><tr><th>Matrícula</th><th>Documento</th><th>Situação</th></tr>
      {% for a in alertas %}
      <tr><td>{{ a[0] }}</td><td>{{ a[1] }}</td>
      <td class="{{ 'bad' if a[2] < 0 else 'warn' }}">{{ 'Vencido há ' ~ (-a[2]) ~ ' dias' if a[2] < 0 else 'Vence em ' ~ a[2] ~ ' dias' }}</td></tr>
      {% endfor %}</table>
      {% else %}<p class="muted">Nenhum documento vencendo nos próximos 10 dias.</p>{% endif %}
    </div>
    """
    return render("Dashboard", body, motoristas=motoristas, viaturas=viaturas,
                  bruto=money(totais["bruto"]), dinheiro=money(totais["dinheiro"]),
                  liquido=money(totais["liquido"]), alertas=alertas)

@app.route("/motoristas", methods=["GET","POST"])
def motoristas():
    c = conn()
    if request.method == "POST":
        c.execute("""INSERT INTO motoristas(nome,email,telefone,iban,percentual)
                     VALUES(?,?,?,?,?)""",
                  (request.form["nome"].strip(), request.form.get("email","").strip(),
                   request.form.get("telefone","").strip(), request.form.get("iban","").strip(),
                   float(request.form.get("percentual") or 0)))
        c.commit()
        c.close()
        flash("Motorista cadastrado.")
        return redirect(url_for("motoristas"))
    rows = c.execute("SELECT * FROM motoristas ORDER BY nome").fetchall()
    c.close()
    body = """
    <div class="card">
      <h2>Novo motorista</h2>
      <form method="post" class="grid">
        <label>Nome<input name="nome" required></label>
        <label>E-mail<input name="email" type="email"></label>
        <label>Telefone<input name="telefone"></label>
        <label>IBAN<input name="iban"></label>
        <label>Percentual (%)<input name="percentual" type="number" step="0.01"></label>
        <div style="align-self:end"><button>Cadastrar</button></div>
      </form>
    </div>
    <div class="card">
      <h2>Motoristas</h2>
      <table><tr><th>Nome</th><th>Contato</th><th>IBAN</th><th>%</th></tr>
      {% for r in rows %}<tr><td>{{ r.nome }}</td><td>{{ r.email }}<br>{{ r.telefone }}</td><td>{{ r.iban }}</td><td>{{ r.percentual }}</td></tr>{% endfor %}
      </table>
    </div>
    """
    return render("Motoristas", body, rows=rows)

@app.route("/viaturas", methods=["GET","POST"])
def viaturas():
    c = conn()
    if request.method == "POST":
        try:
            c.execute("""INSERT INTO viaturas(matricula,marca,modelo,motorista_id,seguro_validade,ipo_validade,carta_verde_validade,km)
                         VALUES(?,?,?,?,?,?,?,?)""",
                      (request.form["matricula"].strip().upper(), request.form.get("marca",""),
                       request.form.get("modelo",""), request.form.get("motorista_id") or None,
                       request.form.get("seguro_validade") or None, request.form.get("ipo_validade") or None,
                       request.form.get("carta_verde_validade") or None, int(request.form.get("km") or 0)))
            c.commit()
            flash("Viatura cadastrada.")
        except sqlite3.IntegrityError:
            flash("Essa matrícula já está cadastrada.")
        c.close()
        return redirect(url_for("viaturas"))
    rows = c.execute("""SELECT v.*,m.nome motorista FROM viaturas v
                       LEFT JOIN motoristas m ON m.id=v.motorista_id ORDER BY v.matricula""").fetchall()
    motoristas_rows = c.execute("SELECT * FROM motoristas WHERE ativo=1 ORDER BY nome").fetchall()
    c.close()
    body = """
    <div class="card">
      <h2>Nova viatura</h2>
      <form method="post" class="grid">
        <label>Matrícula<input name="matricula" required></label>
        <label>Marca<input name="marca"></label>
        <label>Modelo<input name="modelo"></label>
        <label>Motorista<select name="motorista_id"><option value="">Sem motorista</option>{% for m in motoristas %}<option value="{{m.id}}">{{m.nome}}</option>{% endfor %}</select></label>
        <label>Seguro<input type="date" name="seguro_validade"></label>
        <label>IPO<input type="date" name="ipo_validade"></label>
        <label>Carta Verde<input type="date" name="carta_verde_validade"></label>
        <label>Quilometragem<input type="number" name="km"></label>
        <div style="align-self:end"><button>Cadastrar</button></div>
      </form>
    </div>
    <div class="card">
      <h2>Viaturas</h2>
      <table><tr><th>Matrícula</th><th>Viatura</th><th>Motorista</th><th>Seguro</th><th>IPO</th><th>Carta Verde</th></tr>
      {% for r in rows %}<tr><td><strong>{{r.matricula}}</strong></td><td>{{r.marca}} {{r.modelo}}</td><td>{{r.motorista or '-'}}</td><td>{{r.seguro_validade or '-'}}</td><td>{{r.ipo_validade or '-'}}</td><td>{{r.carta_verde_validade or '-'}}</td></tr>{% endfor %}
      </table>
    </div>
    """
    return render("Viaturas", body, rows=rows, motoristas=motoristas_rows)

@app.route("/relatorios", methods=["GET","POST"])
def relatorios():
    c = conn()
    if request.method == "POST":
        bruto = float(request.form.get("bruto") or 0)
        dinheiro = float(request.form.get("dinheiro_maos") or 0)
        comissao = float(request.form.get("comissao") or 0)
        combustivel = float(request.form.get("combustivel") or 0)
        portagens = float(request.form.get("portagens") or 0)
        outros = float(request.form.get("outros_descontos") or 0)
        reembolsos = float(request.form.get("reembolsos") or 0)
        liquido = bruto - dinheiro - comissao - combustivel - portagens - outros + reembolsos
        c.execute("""INSERT INTO relatorios(plataforma,motorista_id,semana,bruto,dinheiro_maos,comissao,
                     combustivel,portagens,outros_descontos,reembolsos,liquido,criado_em)
                     VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                  (request.form["plataforma"], request.form["motorista_id"], request.form["semana"],
                   bruto,dinheiro,comissao,combustivel,portagens,outros,reembolsos,liquido,
                   datetime.now().isoformat(timespec="seconds")))
        c.commit()
        c.close()
        flash("Relatório salvo e valor líquido calculado.")
        return redirect(url_for("relatorios"))
    rows = c.execute("""SELECT r.*,m.nome motorista FROM relatorios r
                       JOIN motoristas m ON m.id=r.motorista_id ORDER BY r.id DESC""").fetchall()
    motoristas_rows = c.execute("SELECT * FROM motoristas WHERE ativo=1 ORDER BY nome").fetchall()
    c.close()
    body = """
    <div class="card">
      <h2>Novo relatório semanal</h2>
      <p class="muted">Regra: líquido = bruto − dinheiro em mãos − comissão − combustível − portagens − outros descontos + reembolsos.</p>
      <form method="post" class="grid">
        <label>Plataforma<select name="plataforma"><option>Uber</option><option>Bolt</option></select></label>
        <label>Motorista<select name="motorista_id" required>{% for m in motoristas %}<option value="{{m.id}}">{{m.nome}}</option>{% endfor %}</select></label>
        <label>Semana<input type="week" name="semana" required></label>
        <label>Valor bruto<input type="number" step="0.01" name="bruto" required></label>
        <label>Dinheiro em mãos<input type="number" step="0.01" name="dinheiro_maos" value="0"></label>
        <label>Comissão<input type="number" step="0.01" name="comissao" value="0"></label>
        <label>Combustível<input type="number" step="0.01" name="combustivel" value="0"></label>
        <label>Portagens<input type="number" step="0.01" name="portagens" value="0"></label>
        <label>Outros descontos<input type="number" step="0.01" name="outros_descontos" value="0"></label>
        <label>Reembolsos<input type="number" step="0.01" name="reembolsos" value="0"></label>
        <div style="align-self:end"><button>Calcular e salvar</button></div>
      </form>
    </div>
    <div class="card">
      <h2>Relatórios processados</h2>
      <table><tr><th>Semana</th><th>Plataforma</th><th>Motorista</th><th>Bruto</th><th>Dinheiro em mãos</th><th>Líquido</th><th></th></tr>
      {% for r in rows %}<tr><td>{{r.semana}}</td><td>{{r.plataforma}}</td><td>{{r.motorista}}</td>
      <td>{{ money(r.bruto) }}</td><td>{{ money(r.dinheiro_maos) }}</td><td><strong>{{ money(r.liquido) }}</strong></td>
      <td><a class="btn secondary" href="/recibo/{{r.id}}">PDF</a></td></tr>{% endfor %}
      </table>
    </div>
    """
    return render("Relatórios", body, rows=rows, motoristas=motoristas_rows, money=money)

@app.route("/recibos")
def recibos():
    c = conn()
    rows = c.execute("""SELECT r.*,m.nome motorista FROM relatorios r
                       JOIN motoristas m ON m.id=r.motorista_id ORDER BY r.id DESC""").fetchall()
    c.close()
    body = """
    <div class="card"><h2>Recibos disponíveis</h2>
    <table><tr><th>Motorista</th><th>Semana</th><th>Plataforma</th><th>Valor líquido</th><th></th></tr>
    {% for r in rows %}<tr><td>{{r.motorista}}</td><td>{{r.semana}}</td><td>{{r.plataforma}}</td><td>{{money(r.liquido)}}</td>
    <td><a class="btn" href="/recibo/{{r.id}}">Abrir PDF</a></td></tr>{% endfor %}</table></div>
    """
    return render("Recibos", body, rows=rows, money=money)

@app.route("/recibo/<int:rid>")
def recibo(rid):
    c = conn()
    r = c.execute("""SELECT r.*,m.nome motorista,m.iban FROM relatorios r
                     JOIN motoristas m ON m.id=r.motorista_id WHERE r.id=?""",(rid,)).fetchone()
    c.close()
    if not r:
        return "Recibo não encontrado", 404
    buffer = BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=A4)
    y = 800
    pdf.setFont("Helvetica-Bold", 18); pdf.drawString(50,y,"Recibo semanal TVDE"); y-=35
    pdf.setFont("Helvetica", 11)
    linhas = [
      ("Motorista", r["motorista"]), ("IBAN", r["iban"] or "-"), ("Plataforma", r["plataforma"]),
      ("Semana", r["semana"]), ("Valor bruto", money(r["bruto"])),
      ("Dinheiro em mãos", "- " + money(r["dinheiro_maos"])), ("Comissão", "- " + money(r["comissao"])),
      ("Combustível", "- " + money(r["combustivel"])), ("Portagens", "- " + money(r["portagens"])),
      ("Outros descontos", "- " + money(r["outros_descontos"])), ("Reembolsos", "+ " + money(r["reembolsos"]))
    ]
    for k,v in linhas:
        pdf.drawString(50,y,k); pdf.drawRightString(540,y,str(v)); y-=24
    y-=10; pdf.line(50,y,540,y); y-=32
    pdf.setFont("Helvetica-Bold", 14)
    pdf.drawString(50,y,"Valor líquido a pagar")
    pdf.drawRightString(540,y,money(r["liquido"]))
    pdf.showPage(); pdf.save()
    buffer.seek(0)
    return send_file(buffer, mimetype="application/pdf", as_attachment=False,
                     download_name=f"recibo_{r['motorista']}_{r['semana']}.pdf")

init_db()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False)
