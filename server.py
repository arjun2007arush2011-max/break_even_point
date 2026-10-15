#!/usr/bin/env python3
"""
Break-Even Point Predictor - Python Backend Server (http.server + sqlite3)
Zero external dependencies; uses Python standard library only.
Serves static frontend files and provides complete REST API endpoints.
"""

import http.server
import socketserver
import json
import sqlite3
import hashlib
import secrets
import os
import mimetypes
from datetime import datetime
from urllib.parse import urlparse

PORT = int(os.environ.get("PORT", 8000))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "breakeven.db")


def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    c.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            email TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            token TEXT,
            created_at TEXT NOT NULL
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS profiles (
            user_id INTEGER PRIMARY KEY,
            firm_name TEXT,
            owner_name TEXT,
            business_type TEXT,
            phone TEXT,
            email TEXT,
            address TEXT,
            updated_at TEXT,
            FOREIGN KEY(user_id) REFERENCES users(id)
        )
    """)

    c.execute("""
        CREATE TABLE IF NOT EXISTS inputs (
            user_id INTEGER PRIMARY KEY,
            fixed_costs REAL,
            selling_price REAL,
            variable_cost REAL,
            expected_volume REAL,
            target_profit REAL,
            currency TEXT,
            unit_label TEXT,
            updated_at TEXT,
            FOREIGN KEY(user_id) REFERENCES users(id)
        )
    """)

    # Seed demo account
    c.execute("SELECT id FROM users WHERE email = 'demo@business.com'")
    if not c.fetchone():
        demo_hash = hashlib.sha256(b"Demo1234").hexdigest()
        now = datetime.utcnow().isoformat()
        c.execute(
            "INSERT INTO users (name, email, password_hash, created_at) VALUES (?, ?, ?, ?)",
            ("Elena Rostova", "demo@business.com", demo_hash, now)
        )
        user_id = c.lastrowid
        c.execute("""
            INSERT INTO profiles (user_id, firm_name, owner_name, business_type, phone, email, address, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (user_id, "Velvet Bean Roastery", "Elena Rostova", "Restaurant / Food", "2065550184", "elena@velvetbean.com", "Pike Place Market, Seattle, WA", now))
        c.execute("""
            INSERT INTO inputs (user_id, fixed_costs, selling_price, variable_cost, expected_volume, target_profit, currency, unit_label, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (user_id, 12500.0, 6.50, 2.10, 3800.0, 5000.0, "$", "cups", now))

    conn.commit()
    conn.close()


def calculate_model(f, p, v, q, t):
    f = max(0.0, float(f or 0))
    p = max(0.0, float(p or 0))
    v = max(0.0, float(v or 0))
    q = max(0.0, float(q or 0))
    t = max(0.0, float(t or 0))

    cm = p - v
    cm_ratio = (cm / p) if p > 0 else 0.0
    is_valid = cm > 0

    be_units = (f / cm) if is_valid else None
    be_revenue = (be_units * p) if is_valid else None

    exp_rev = p * q
    exp_var = v * q
    exp_cost = f + exp_var
    exp_profit = exp_rev - exp_cost

    mos_units = (q - be_units) if is_valid else None
    mos_ratio = ((q - be_units) / q) if (is_valid and q > 0) else 0.0
    mos_revenue = (mos_units * p) if is_valid else None

    target_units = ((f + t) / cm) if is_valid else None
    target_rev = (target_units * p) if is_valid else None

    total_cm = q * cm
    dol = (total_cm / exp_profit) if exp_profit > 0 else None

    return {
        "fixed_costs": f,
        "selling_price": p,
        "variable_cost": v,
        "expected_volume": q,
        "target_profit": t,
        "contribution_margin": round(cm, 2),
        "contribution_margin_ratio": round(cm_ratio * 100, 2),
        "is_valid": is_valid,
        "break_even_units": round(be_units, 2) if be_units is not None else None,
        "break_even_revenue": round(be_revenue, 2) if be_revenue is not None else None,
        "expected_revenue": round(exp_rev, 2),
        "expected_total_cost": round(exp_cost, 2),
        "expected_profit": round(exp_profit, 2),
        "margin_of_safety_units": round(mos_units, 2) if mos_units is not None else None,
        "margin_of_safety_ratio": round(mos_ratio * 100, 2),
        "target_units": round(target_units, 2) if target_units is not None else None,
        "target_revenue": round(target_rev, 2) if target_rev is not None else None,
        "degree_of_operating_leverage": round(dol, 2) if dol is not None else None
    }


class RequestHandler(http.server.BaseHTTPRequestHandler):

    def get_db(self):
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        return conn

    def set_headers(self, status=200, content_type="application/json"):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization, X-Requested-With")
        self.end_headers()

    def do_OPTIONS(self):
        self.set_headers(200, "text/plain")

    def parse_body(self):
        length = int(self.headers.get("Content-Length", 0))
        if length > 0:
            raw = self.rfile.read(length).decode("utf-8")
            try:
                return json.loads(raw)
            except Exception:
                return {}
        return {}

    def get_user_from_token(self, conn):
        auth = self.headers.get("Authorization", "")
        if auth.startswith("Bearer "):
            token = auth[7:].strip()
            c = conn.cursor()
            c.execute("SELECT * FROM users WHERE token = ?", (token,))
            return c.fetchone()
        return None

    def send_json(self, status, data):
        self.set_headers(status, "application/json")
        self.wfile.write(json.dumps(data).encode("utf-8"))

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path

        if path == "/api/health":
            self.send_json(200, {
                "status": "ok",
                "backend": "Python http.server + sqlite3",
                "timestamp": datetime.utcnow().isoformat(),
                "database": "breakeven.db"
            })
            return

        conn = self.get_db()

        if path == "/api/auth/me":
            user = self.get_user_from_token(conn)
            if not user:
                conn.close()
                self.send_json(401, {"error": "Unauthorized", "authenticated": False})
                return

            c = conn.cursor()
            c.execute("SELECT * FROM profiles WHERE user_id = ?", (user["id"],))
            profile = c.fetchone()
            c.execute("SELECT * FROM inputs WHERE user_id = ?", (user["id"],))
            inputs = c.fetchone()
            conn.close()

            self.send_json(200, {
                "authenticated": True,
                "user": {"id": user["id"], "name": user["name"], "email": user["email"]},
                "profile": dict(profile) if profile else {},
                "inputs": dict(inputs) if inputs else {}
            })
            return

        if path == "/api/profile":
            user = self.get_user_from_token(conn)
            if not user:
                conn.close()
                self.send_json(401, {"error": "Unauthorized"})
                return

            c = conn.cursor()
            c.execute("SELECT * FROM profiles WHERE user_id = ?", (user["id"],))
            prof = c.fetchone()
            conn.close()
            self.send_json(200, {"success": True, "profile": dict(prof) if prof else {}})
            return

        if path == "/api/inputs":
            user = self.get_user_from_token(conn)
            if not user:
                conn.close()
                self.send_json(401, {"error": "Unauthorized"})
                return

            c = conn.cursor()
            c.execute("SELECT * FROM inputs WHERE user_id = ?", (user["id"],))
            inp = c.fetchone()
            conn.close()
            self.send_json(200, {"success": True, "inputs": dict(inp) if inp else {}})
            return

        conn.close()

        # Static File Serving
        rel_path = path.lstrip("/")
        if not rel_path or rel_path == "":
            rel_path = "index.html"
        file_path = os.path.join(BASE_DIR, rel_path)

        if os.path.isfile(file_path):
            mime_type, _ = mimetypes.guess_type(file_path)
            if not mime_type:
                mime_type = "application/octet-stream"
            with open(file_path, "rb") as f:
                content = f.read()
            self.set_headers(200, mime_type)
            self.wfile.write(content)
        else:
            self.set_headers(404, "text/plain")
            self.wfile.write(b"404 Not Found")

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path
        body = self.parse_body()
        conn = self.get_db()

        if path == "/api/auth/signup":
            name = (body.get("name") or "").strip()
            email = (body.get("email") or "").strip().lower()
            password = body.get("password") or ""
            pass_hash = body.get("password_hash") or (hashlib.sha256(password.encode()).hexdigest() if password else None)

            if not name or not email or not pass_hash:
                conn.close()
                self.send_json(400, {"error": "Name, email, and password required."})
                return

            c = conn.cursor()
            c.execute("SELECT id FROM users WHERE email = ?", (email,))
            if c.fetchone():
                conn.close()
                self.send_json(409, {"error": "An account with this email already exists."})
                return

            token = secrets.token_hex(24)
            now = datetime.utcnow().isoformat()
            c.execute("INSERT INTO users (name, email, password_hash, token, created_at) VALUES (?, ?, ?, ?, ?)",
                      (name, email, pass_hash, token, now))
            user_id = c.lastrowid

            prof = body.get("profile") or {}
            firm = prof.get("firm_name") or f"{name}'s Business"
            b_type = prof.get("business_type") or "Retail"
            phone = prof.get("phone") or "5551234567"
            p_email = prof.get("email") or email
            addr = prof.get("address") or "Headquarters"

            c.execute("""
                INSERT INTO profiles (user_id, firm_name, owner_name, business_type, phone, email, address, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (user_id, firm, name, b_type, phone, p_email, addr, now))

            inp = body.get("inputs") or {}
            c.execute("""
                INSERT INTO inputs (user_id, fixed_costs, selling_price, variable_cost, expected_volume, target_profit, currency, unit_label, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (user_id, float(inp.get("fixed_costs", 12000)), float(inp.get("selling_price", 40)),
                  float(inp.get("variable_cost", 15)), float(inp.get("expected_volume", 600)),
                  float(inp.get("target_profit", 3000)), inp.get("currency", "$"), inp.get("unit_label", "units"), now))

            conn.commit()
            conn.close()

            self.send_json(201, {
                "success": True,
                "token": token,
                "user": {"id": user_id, "name": name, "email": email},
                "profile": {"firm_name": firm, "owner_name": name, "business_type": b_type, "phone": phone, "email": p_email, "address": addr},
                "inputs": {"fixed_costs": 12000, "selling_price": 40, "variable_cost": 15, "expected_volume": 600, "target_profit": 3000, "currency": "$", "unit_label": "units"}
            })
            return

        if path == "/api/auth/login":
            email = (body.get("email") or "").strip().lower()
            password = body.get("password") or ""
            pass_hash = body.get("password_hash") or (hashlib.sha256(password.encode()).hexdigest() if password else None)

            c = conn.cursor()
            c.execute("SELECT * FROM users WHERE email = ?", (email,))
            user = c.fetchone()

            if not user or user["password_hash"] != pass_hash:
                conn.close()
                self.send_json(401, {"error": "Invalid email or password."})
                return

            token = secrets.token_hex(24)
            c.execute("UPDATE users SET token = ? WHERE id = ?", (token, user["id"]))

            c.execute("SELECT * FROM profiles WHERE user_id = ?", (user["id"],))
            profile = c.fetchone()
            c.execute("SELECT * FROM inputs WHERE user_id = ?", (user["id"],))
            inputs = c.fetchone()

            conn.commit()
            conn.close()

            self.send_json(200, {
                "success": True,
                "token": token,
                "user": {"id": user["id"], "name": user["name"], "email": user["email"]},
                "profile": dict(profile) if profile else {},
                "inputs": dict(inputs) if inputs else {}
            })
            return

        if path == "/api/auth/logout":
            user = self.get_user_from_token(conn)
            if user:
                c = conn.cursor()
                c.execute("UPDATE users SET token = NULL WHERE id = ?", (user["id"],))
                conn.commit()
            conn.close()
            self.send_json(200, {"success": True})
            return

        if path in ("/api/profile", "/api/profile/"):
            user = self.get_user_from_token(conn)
            if not user:
                conn.close()
                self.send_json(401, {"error": "Unauthorized"})
                return

            now = datetime.utcnow().isoformat()
            c = conn.cursor()
            c.execute("""
                UPDATE profiles
                SET firm_name = ?, owner_name = ?, business_type = ?, phone = ?, email = ?, address = ?, updated_at = ?
                WHERE user_id = ?
            """, (body.get("firm_name"), body.get("owner_name"), body.get("business_type"),
                  body.get("phone"), body.get("email"), body.get("address"), now, user["id"]))
            conn.commit()

            c.execute("SELECT * FROM profiles WHERE user_id = ?", (user["id"],))
            updated = c.fetchone()
            conn.close()
            self.send_json(200, {"success": True, "profile": dict(updated) if updated else {}})
            return

        if path in ("/api/inputs", "/api/inputs/"):
            user = self.get_user_from_token(conn)
            if not user:
                conn.close()
                self.send_json(401, {"error": "Unauthorized"})
                return

            now = datetime.utcnow().isoformat()
            c = conn.cursor()
            c.execute("""
                UPDATE inputs
                SET fixed_costs = ?, selling_price = ?, variable_cost = ?, expected_volume = ?, target_profit = ?, currency = ?, unit_label = ?, updated_at = ?
                WHERE user_id = ?
            """, (float(body.get("fixed_costs", 0)), float(body.get("selling_price", 0)),
                  float(body.get("variable_cost", 0)), float(body.get("expected_volume", 0)),
                  float(body.get("target_profit", 0)), body.get("currency", "$"), body.get("unit_label", "units"),
                  now, user["id"]))
            conn.commit()

            c.execute("SELECT * FROM inputs WHERE user_id = ?", (user["id"],))
            updated = c.fetchone()
            conn.close()
            self.send_json(200, {"success": True, "inputs": dict(updated) if updated else {}})
            return

        if path == "/api/calculate":
            conn.close()
            calc = calculate_model(
                body.get("fixed_costs", 0),
                body.get("selling_price", 0),
                body.get("variable_cost", 0),
                body.get("expected_volume", 0),
                body.get("target_profit", 0)
            )
            self.send_json(200, {"success": True, "calculation": calc})
            return

        conn.close()
        self.send_json(404, {"error": "Endpoint not found"})


def run():
    init_db()
    with socketserver.TCPServer(("", PORT), RequestHandler) as httpd:
        print("\n=======================================================")
        print(f" Break-Even Point Predictor - Python Server Active")
        print(f" Local URL: http://localhost:{PORT}/")
        print(f" SQLite DB: {DB_PATH}")
        print(" Press Ctrl+C to stop the server")
        print("=======================================================\n")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            httpd.server_close()


if __name__ == "__main__":
    run()
