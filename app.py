import base64
import io
import json
import os
import sys
import queue
import threading
from datetime import datetime
from flask import Flask, Response, jsonify, render_template, request, send_file, stream_with_context

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

app = Flask(__name__)
BASE_DIR      = os.path.dirname(os.path.abspath(__file__))
SETTINGS_FILE = os.path.join(BASE_DIR, "config", "settings.json")

_last_xlsx: dict[str, bytes] = {}
_cached_companies: list[dict] = []   # cached after first login fetch


def load_settings() -> dict:
    if os.path.exists(SETTINGS_FILE):
        with open(SETTINGS_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_settings(data: dict):
    os.makedirs(os.path.dirname(SETTINGS_FILE), exist_ok=True)
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def get_output_path() -> str:
    s = load_settings()
    path = s.get("output_path", os.path.join(os.path.expanduser("~"), "Desktop", "PM_OUTPUT"))
    os.makedirs(path, exist_ok=True)
    return path


@app.route("/health")
def health():
    return "ok"


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/settings", methods=["GET"])
def get_settings():
    return jsonify(load_settings())


@app.route("/api/settings", methods=["POST"])
def post_settings():
    save_settings(request.json)
    return jsonify({"ok": True})


# ---------------------------------------------------------------------------
# Company list — fetched from portal, cached in memory
# ---------------------------------------------------------------------------

@app.route("/api/companies", methods=["GET"])
def get_companies():
    """Return cached company list. If empty, returns [] — client can trigger /api/companies/fetch."""
    return jsonify(_cached_companies)


@app.route("/api/companies/fetch", methods=["POST"])
def fetch_companies():
    """Login to portal and fetch fresh company list. Stores in memory cache."""
    global _cached_companies
    s = load_settings()
    user     = s.get("user", "")
    password = s.get("password", "")
    if not user or not password:
        return jsonify({"ok": False, "error": "Salve as credenciais primeiro"}), 400

    try:
        from iss_worker import login, get_company_list
        p, browser, context, page = login(user, password)
        companies = get_company_list(page)
        browser.close()
        p.stop()
        _cached_companies = companies
        return jsonify({"ok": True, "count": len(companies), "companies": companies})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


# ---------------------------------------------------------------------------
# ISS Substituto — streaming SSE run
# ---------------------------------------------------------------------------

@app.route("/api/iss/run", methods=["POST"])
def iss_run():
    data          = request.json
    user          = data.get("user", "")
    password      = data.get("password", "")
    target_period = data.get("target_period", "").strip()
    limit         = data.get("limit")
    resume        = data.get("resume", False)
    workers       = data.get("workers", 4)

    if not user or not password:
        return jsonify({"ok": False, "error": "Usuário e senha são obrigatórios"}), 400
    if target_period and (len(target_period) != 6 or not target_period.isdigit()):
        return jsonify({"ok": False, "error": "Período inválido — use MMAAAA (ex: 062026)"}), 400

    from iss_worker import run_iss_report

    q = queue.Queue()

    def _worker():
        try:
            for line in run_iss_report(
                user=user, password=password,
                target_period=target_period or None,
                limit=int(limit) if limit else None,
                resume=resume,
                workers=int(workers),
            ):
                q.put(line)
        except Exception as e:
            q.put(f"ERRO inesperado: {e}")
            q.put("__DONE__:1")

    threading.Thread(target=_worker, daemon=True).start()

    def generate():
        yield ": ok\n\n"
        while True:
            try:
                msg = q.get(timeout=3600)
                if msg.startswith("__DONE__:"):
                    yield f"event: done\ndata: {msg.split(':')[1]}\n\n"
                    break
                elif msg.startswith("__XLSX__:"):
                    _last_xlsx["iss"] = base64.b64decode(msg[len("__XLSX__:"):])
                    yield "event: xlsx_ready\ndata: 1\n\n"
                else:
                    yield f"data: {msg.replace(chr(10), ' ')}\n\n"
            except queue.Empty:
                yield "data: [AVISO] Timeout\n\n"
                yield "event: done\ndata: 1\n\n"
                break

    return Response(stream_with_context(generate()), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.route("/api/iss/download")
def iss_download():
    xlsx = _last_xlsx.get("iss")
    if not xlsx:
        return jsonify({"error": "Nenhum relatório disponível"}), 404
    today = datetime.now().strftime("%Y%m%d_%H%M")
    return send_file(io.BytesIO(xlsx), as_attachment=True,
                     download_name=f"iss_substituto_{today}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---------------------------------------------------------------------------
# NFS-e PM — streaming SSE run
# ---------------------------------------------------------------------------

@app.route("/api/nfse/run", methods=["POST"])
def nfse_run():
    data          = request.json
    user          = data.get("user", "")
    password      = data.get("password", "")
    target_period = data.get("target_period", "").strip()
    selected      = data.get("selected", [])  # list of CMC ids; empty = all

    if not user or not password:
        return jsonify({"ok": False, "error": "Usuário e senha são obrigatórios"}), 400
    if not target_period or len(target_period) != 6 or not target_period.isdigit():
        return jsonify({"ok": False, "error": "Período obrigatório — MMAAAA (ex: 062026)"}), 400

    from nfse_worker import run_nfse_pm

    output_base = get_output_path()
    q = queue.Queue()

    def _worker():
        try:
            for line in run_nfse_pm(
                user=user, password=password,
                target_period=target_period,
                selected_cnpjs=selected if selected else None,
                output_base=output_base,
            ):
                q.put(line)
        except Exception as e:
            q.put(f"ERRO inesperado: {e}")
            q.put("__DONE__:1")

    threading.Thread(target=_worker, daemon=True).start()

    def generate():
        yield ": ok\n\n"
        while True:
            try:
                msg = q.get(timeout=3600)
                if msg.startswith("__DONE__:"):
                    yield f"event: done\ndata: {msg.split(':')[1]}\n\n"
                    break
                elif msg.startswith("__SUMMARY__:"):
                    payload = msg[len("__SUMMARY__:"):]
                    yield f"event: summary\ndata: {payload}\n\n"
                else:
                    yield f"data: {msg.replace(chr(10), ' ')}\n\n"
            except queue.Empty:
                yield "data: [AVISO] Timeout\n\n"
                yield "event: done\ndata: 1\n\n"
                break

    return Response(stream_with_context(generate()), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


# ---------------------------------------------------------------------------
# Guias / Boletos
# ---------------------------------------------------------------------------

@app.route("/api/guias/run", methods=["POST"])
def guias_run():
    data     = request.json
    user     = data.get("user", "")
    password = data.get("password", "")

    if not user or not password:
        return jsonify({"ok": False, "error": "Usuário e senha são obrigatórios"}), 400

    from guias_worker import run_guias

    output_base = get_output_path()
    q = queue.Queue()

    def _worker():
        try:
            for line in run_guias(user=user, password=password, output_base=output_base):
                q.put(line)
        except Exception as e:
            q.put(f"ERRO inesperado: {e}")
            q.put("__DONE__:1")

    threading.Thread(target=_worker, daemon=True).start()

    def generate():
        yield ": ok\n\n"
        while True:
            try:
                msg = q.get(timeout=3600)
                if msg.startswith("__DONE__:"):
                    yield f"event: done\ndata: {msg.split(':')[1]}\n\n"
                    break
                elif msg.startswith("__SUMMARY__:"):
                    payload = msg[len("__SUMMARY__:"):]
                    yield f"event: summary\ndata: {payload}\n\n"
                else:
                    yield f"data: {msg.replace(chr(10), ' ')}\n\n"
            except queue.Empty:
                yield "data: [AVISO] Timeout\n\n"
                yield "event: done\ndata: 1\n\n"
                break

    return Response(stream_with_context(generate()), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


if __name__ == "__main__":
    app.run(debug=False, port=5001, threaded=True, use_reloader=False)
