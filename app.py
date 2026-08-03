from flask import Flask, render_template, request, jsonify, send_file, Response, stream_with_context
import json
import subprocess
import os
import sys
import zipfile
import tempfile
import importlib
import threading
import queue
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
from datetime import datetime

app = Flask(__name__)

BASE_DIR       = os.path.dirname(os.path.abspath(__file__))
COMPANIES_FILE = os.path.join(BASE_DIR, 'config', 'companies.json')
GROUPS_FILE    = os.path.join(BASE_DIR, 'config', 'groups.json')
SETTINGS_FILE  = os.path.join(BASE_DIR, 'config', 'settings.json')

SHEET_CSV_URL = 'https://docs.google.com/spreadsheets/d/1MI4xI6rSWfYVYTtPfXOzNPon-AGq0KXh/export?format=csv'

def find_desktop():
    home = os.path.expanduser('~')
    onedrive_desktop = os.path.join(home, 'OneDrive', 'Desktop')
    if os.path.exists(onedrive_desktop):
        return onedrive_desktop
    return os.path.join(home, 'Desktop')

def auto_patch_settings():
    try:
        with open(SETTINGS_FILE, encoding='utf-8') as f:
            settings = json.load(f)

        changed = False

        correct_ext     = os.path.join(BASE_DIR, 'extension', '2.0.5_0')
        correct_profile = os.path.join(BASE_DIR, 'chrome-profile')

        if settings.get('extension_path') != correct_ext:
            settings['extension_path'] = correct_ext
            changed = True

        if settings.get('profile_path') != correct_profile:
            settings['profile_path'] = correct_profile
            changed = True

        current_dl = settings.get('downloads_path', '')
        if not os.path.exists(current_dl):
            desktop = find_desktop()
            new_dl  = os.path.join(desktop, 'NFESAUTOMATION')
            os.makedirs(new_dl, exist_ok=True)
            settings['downloads_path'] = new_dl
            changed = True

        if changed:
            with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=2, ensure_ascii=False)
            print('[SETTINGS] Paths atualizados automaticamente', flush=True)

    except Exception as e:
        print(f'[AVISO] Falha ao atualizar settings.json: {e}', flush=True)

auto_patch_settings()

def load_companies():
    try:
        import urllib.request
        import csv
        import io
        with urllib.request.urlopen(SHEET_CSV_URL) as response:
            content = response.read().decode('utf-8')
        reader = csv.DictReader(io.StringIO(content))
        companies = []
        for row in reader:
            if row.get('cnpj') and row.get('password'):
                companies.append({
                    'name':       row.get('name', '').strip(),
                    'cnpj':       row.get('cnpj', '').strip(),
                    'password':   row.get('password', '').strip(),
                    'accountant': row.get('accountant', 'Empresas').strip(),
                    'email':      row.get('email', '').strip(),
                })
        return companies
    except Exception as e:
        print(f'[AVISO] Falha ao carregar Google Sheets: {e}, usando companies.json')
        with open(COMPANIES_FILE, encoding='utf-8') as f:
            return json.load(f)

def save_companies(companies):
    with open(COMPANIES_FILE, 'w', encoding='utf-8') as f:
        json.dump(companies, f, indent=2, ensure_ascii=False)

def load_groups():
    with open(GROUPS_FILE, encoding='utf-8') as f:
        return json.load(f)

def save_groups(groups):
    with open(GROUPS_FILE, 'w', encoding='utf-8') as f:
        json.dump(groups, f, indent=2, ensure_ascii=False)

def get_downloads_path():
    with open(SETTINGS_FILE, encoding='utf-8') as f:
        p = json.load(f)['downloads_path']
    if not os.path.exists(p):
        desktop = find_desktop()
        p = os.path.join(desktop, 'NFESAUTOMATION')
        os.makedirs(p, exist_ok=True)
        with open(SETTINGS_FILE, encoding='utf-8') as f:
            settings = json.load(f)
        settings['downloads_path'] = p
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(settings, f, indent=2, ensure_ascii=False)
        print(f'[SETTINGS] downloads_path corrigido para: {p}', flush=True)
    return p

@app.route('/health')
def health():
    return 'ok'

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/companies', methods=['GET'])
def get_companies():
    return jsonify({'companies': load_companies(), 'groups': load_groups()})

@app.route('/api/companies', methods=['POST'])
def add_company():
    companies = load_companies()
    companies.append(request.json)
    save_companies(companies)
    return jsonify({'ok': True})

@app.route('/api/companies/<int:idx>', methods=['DELETE'])
def delete_company(idx):
    companies = load_companies()
    companies.pop(idx)
    save_companies(companies)
    return jsonify({'ok': True})

@app.route('/api/companies/<int:idx>', methods=['PUT'])
def update_company(idx):
    companies = load_companies()
    companies[idx] = request.json
    save_companies(companies)
    return jsonify({'ok': True})

@app.route('/api/groups', methods=['GET'])
def get_groups():
    return jsonify(load_groups())

@app.route('/api/groups', methods=['POST'])
def add_group():
    groups = load_groups()
    group  = request.json
    group['id'] = str(datetime.now().timestamp())
    if 'companies' in group and 'cnpjs' not in group:
        group['cnpjs'] = group.pop('companies')
    groups.append(group)
    save_groups(groups)
    return jsonify({'ok': True, 'id': group['id']})

@app.route('/api/groups/<group_id>', methods=['PUT'])
def update_group(group_id):
    groups = load_groups()
    data   = request.json
    if 'companies' in data and 'cnpjs' not in data:
        data['cnpjs'] = data.pop('companies')
    for i, g in enumerate(groups):
        if g['id'] == group_id:
            groups[i]       = data
            groups[i]['id'] = group_id
            break
    save_groups(groups)
    return jsonify({'ok': True})

@app.route('/api/groups/<group_id>', methods=['DELETE'])
def delete_group(group_id):
    groups = [g for g in load_groups() if g['id'] != group_id]
    save_groups(groups)
    return jsonify({'ok': True})

def stream_subprocess(cmd, q):
    try:
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding='utf-8',
            errors='replace',
            bufsize=1
        )
        for line in proc.stdout:
            line = line.rstrip()
            if line:
                clean_line = line.encode('utf-8', errors='replace').decode('utf-8')
                q.put(('log', clean_line))
        proc.wait()
        q.put(('done', proc.returncode))
    except Exception as e:
        q.put(('log', f'[ERRO] {e}'))
        q.put(('done', 1))

@app.route('/api/run/stream', methods=['POST'])
def run_stream():
    data     = request.json
    start    = data.get('start')
    end      = data.get('end')
    selected = data.get('companies', [])
    mode     = data.get('mode', 'reinf')

    try:
        d1 = datetime.strptime(start, '%d/%m/%Y')
        d2 = datetime.strptime(end,   '%d/%m/%Y')
        if (d2 - d1).days > 31:
            return jsonify({'ok': False, 'error': 'Periodo nao pode exceder 31 dias'}), 400
        if d2 < d1:
            return jsonify({'ok': False, 'error': 'Data final deve ser maior que a inicial'}), 400
    except:
        return jsonify({'ok': False, 'error': 'Datas invalidas'}), 400

    all_companies      = load_companies()
    selected_companies = [c for c in all_companies if c['cnpj'] in selected]

    if not selected_companies:
        return jsonify({'ok': False, 'error': 'Nenhuma empresa selecionada'}), 400

    temp_key  = 'temp_run_all.json' if mode == 'all' else 'temp_run.json'
    temp_file = os.path.join(BASE_DIR, 'config', temp_key)
    with open(temp_file, 'w', encoding='utf-8') as f:
        json.dump({'companies': selected_companies, 'start': start, 'end': end, 'mode': mode}, f)

    q   = queue.Queue()
    cmd = [sys.executable, '-u', os.path.join(BASE_DIR, 'main.py'), '--config', temp_file]
    t   = threading.Thread(target=stream_subprocess, args=(cmd, q), daemon=True)
    t.start()

    def generate():
        yield ": ok\n\n"
        while True:
            try:
                msg_type, payload = q.get(timeout=3600)
                if msg_type == 'log':
                    line = payload.replace('\n', ' ')
                    yield f"data: {line}\n\n"
                elif msg_type == 'done':
                    yield f"event: done\ndata: {payload}\n\n"
                    break
            except queue.Empty:
                yield f"data: [AVISO] Timeout aguardando resposta\n\n"
                yield f"event: done\ndata: 1\n\n"
                break

    return Response(
        stream_with_context(generate()),
        mimetype='text/event-stream',
        headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'}
    )

@app.route('/api/run/zip', methods=['POST'])
def run_zip():
    data     = request.json
    start    = data.get('start')
    end      = data.get('end')
    selected = data.get('companies', [])
    mode     = data.get('mode', 'reinf')

    try:
        d1                 = datetime.strptime(start, '%d/%m/%Y')
        month              = d1.strftime('%m-%Y')
        all_companies      = load_companies()
        selected_companies = [c for c in all_companies if c['cnpj'] in selected]
        selected_names     = [c['name'] for c in selected_companies]
        downloads_path     = get_downloads_path()

        sys.path.insert(0, BASE_DIR)

        if mode == 'reinf':
            try:
                import generate_summary
                importlib.reload(generate_summary)
                generate_summary.generate_summary(filter_names=selected_names)
            except Exception as e:
                import traceback
                with open(os.path.join(BASE_DIR, 'resumo_error.log'), 'w', encoding='utf-8') as f:
                    f.write(traceback.format_exc())

            try:
                import generate_fiscal
                importlib.reload(generate_fiscal)
                generate_fiscal.generate_fiscal_all(filter_names=selected_names)
            except Exception as e:
                import traceback
                with open(os.path.join(BASE_DIR, 'fiscal_error.log'), 'w', encoding='utf-8') as f:
                    f.write(traceback.format_exc())

            resumo_path = os.path.join(downloads_path, 'Empresas', 'resumo_nfse.xlsx')
            zip_name    = f'nfse_{month}.zip'
            zip_path    = os.path.join(tempfile.gettempdir(), f'nfse_{month}_{datetime.now().strftime("%H%M%S")}.zip')

            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for company in selected_companies:
                    safe_name   = company['name'].replace('/', '_').replace('\\', '_').replace(':', '_')
                    company_dir = os.path.join(downloads_path, company['accountant'], safe_name, month)
                    print(f'[ZIP] {company["name"]} -> {company_dir} | existe: {os.path.exists(company_dir)}', flush=True)
                    if os.path.exists(company_dir):
                        for root, dirs, files in os.walk(company_dir):
                            rel_root = os.path.relpath(root, company_dir)
                            if rel_root.split(os.sep)[0] in ['pdfs', 'xmls']:
                                continue
                            for file in files:
                                file_path = os.path.join(root, file)
                                arcname   = os.path.relpath(file_path, downloads_path)
                                zf.write(file_path, arcname)
                if os.path.exists(resumo_path):
                    zf.write(resumo_path, os.path.join('Empresas', 'resumo_nfse.xlsx'))

        else:
            zip_name = f'nfse_geral_{month}.zip'
            zip_path = os.path.join(tempfile.gettempdir(), f'nfse_geral_{month}_{datetime.now().strftime("%H%M%S")}.zip')

            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for company in selected_companies:
                    safe_name   = company['name'].replace('/', '_').replace('\\', '_').replace(':', '_')
                    company_dir = os.path.join(downloads_path, company['accountant'], safe_name, month)
                    notas_dir   = os.path.join(company_dir, 'notas')
                    if os.path.exists(notas_dir):
                        for root, dirs, files in os.walk(notas_dir):
                            for file in files:
                                file_path = os.path.join(root, file)
                                arcname   = os.path.relpath(file_path, downloads_path)
                                zf.write(file_path, arcname)

        return send_file(zip_path, as_attachment=True, download_name=zip_name, mimetype='application/zip')

    except Exception as e:
        import traceback
        print(f'[ERRO ZIP] {traceback.format_exc()}', flush=True)
        return jsonify({'ok': False, 'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=False, port=5000, threaded=True, use_reloader=False)
