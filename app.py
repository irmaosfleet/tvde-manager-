from __future__ import annotations

import csv
import hashlib
import io
import os
import re
import sqlite3
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

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = Path(os.getenv("DATABASE_PATH", BASE_DIR / "fleetflow.db"))
UPLOAD_DIR = BASE_DIR / "uploads"
REPORT_DIR = BASE_DIR / "reports"
SEED_XLSX = BASE_DIR / "data" / "BANCO_DE_DADOS.xlsx"
UPLOAD_DIR.mkdir(exist_ok=True)
REPORT_DIR.mkdir(exist_ok=True)

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "troque-esta-chave-em-producao")
app.config["MAX_CONTENT_LENGTH"] = 80 * 1024 * 1024
COMPANY_NAME = os.getenv("COMPANY_NAME", "Irmãos Fleet")
ADMIN_USER = os.getenv("ADMIN_USER", "admin")
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "admin123")


def db() -> sqlite3.Connection:
    con = sqlite3.connect(DB_PATH)
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
    if SEED_XLSX.exists():
        with db() as con:
            count = con.execute("SELECT COUNT(*) FROM drivers").fetchone()[0]
        if count == 0:
            import_master_workbook(SEED_XLSX)


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
    with db() as con:
        cur = con.execute("INSERT INTO closings(label,created_at,status) VALUES(?,?,?)", (label, datetime.now().isoformat(timespec="seconds"), "PROCESSADO"))
        closing_id = cur.lastrowid
        earnings = con.execute("SELECT * FROM raw_earnings").fetchall()
        fuel_rows = con.execute("SELECT card_number,SUM(amount) amount FROM fuel GROUP BY card_number").fetchall()
    fuel_map = {clean_identifier(r["card_number"]): money(r["amount"]) for r in fuel_rows}
    aggregated: dict[int, dict[str, Any]] = {}
    unmatched = 0
    for e in earnings:
        driver = match_driver(e["identifier"], e["origin_ref"], e["display_name"])
        if not driver:
            unmatched += 1
            continue
        a = aggregated.setdefault(driver["id"], {"driver": driver, "gross":0.0, "cash":0.0, "report_reimb":0.0, "origins":set()})
        a["gross"] += money(e["gross"])
        a["cash"] += money(e["cash"])
        a["report_reimb"] += money(e["reimbursement"])
        a["origins"].add(e["origin_ref"])
    items = []
    iban_groups = defaultdict(float)
    temp = []
    for a in aggregated.values():
        d = a["driver"]
        origins = ", ".join(sorted(a["origins"]))
        gross, cash = a["gross"], a["cash"]
        pct = money(d["percentage"])
        is_eats = any("EATS" in norm(x) for x in a["origins"])
        is_tvde = any("TVDE" in norm(x) for x in a["origins"])
        commission_base = gross + cash if is_eats else gross
        commission = commission_base * pct
        fuel_value = fuel_map.get(clean_identifier(d["fuel_card"]), 0.0)
        discount = money(d["discount"])
        reimbursement = (money(d["reimbursement"]) + a["report_reimb"]) if is_tvde else 0.0
        immediate = money(d["immediate"])
        net = gross - commission - fuel_value - discount - immediate + reimbursement
        if is_eats:
            net -= cash
        iban = str(d["iban"] or "").replace(" ", "").upper()
        temp.append({"driver":d,"origins":origins,"gross":gross,"cash":cash,"commission":commission,"fuel":fuel_value,"discount":discount,"reimbursement":reimbursement,"immediate":immediate,"net":net,"iban":iban})
        if iban:
            iban_groups[iban] += net
    fee_applied = set()
    for t in temp:
        d, iban = t["driver"], t["iban"]
        fee = 0.0
        if iban and norm(d["bank_color"]) == "AZUL" and iban not in fee_applied:
            fee = 1.25
            fee_applied.add(iban)
            iban_groups[iban] -= fee
        group_total = iban_groups.get(iban, t["net"] - fee)
        items.append((closing_id,d["id"],d["name"],iban,d["bank_color"],t["origins"],t["gross"],t["cash"],t["commission"],t["fuel"],t["discount"],t["reimbursement"],t["immediate"],fee,t["net"],group_total))
    with db() as con:
        con.executemany("""INSERT INTO closing_items(closing_id,driver_id,driver_name,iban,bank_color,origins,gross,cash,commission,fuel,discount,reimbursement,immediate,bank_fee,net_before_group,group_total)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", items)
        con.execute("UPDATE closings SET status=? WHERE id=?", (f"PROCESSADO ({unmatched} não encontrados)", closing_id))
    return closing_id


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
    return render_template("dashboard.html", drivers=drivers, imports=imports, last=last, stats=stats, origins=origins, recent=recent)


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
        cid = build_closing(label)
        flash("Fechamento processado. Confira os valores antes de gerar o XML.", "success")
        return redirect(url_for("closing_detail", closing_id=cid))
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


@app.errorhandler(413)
def too_large(_):
    flash("O envio excede o limite de 80 MB.", "danger")
    return redirect(url_for("imports_page"))


init_db()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "5000")), debug=True)
