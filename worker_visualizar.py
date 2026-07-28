import sys
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.stderr.reconfigure(encoding='utf-8', errors='replace')

import os
import re
import json
import argparse
import shutil
import tempfile
import time
from datetime import date, datetime
from dateutil.relativedelta import relativedelta
from playwright.sync_api import sync_playwright

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.auth import create_browser_for_company, login
from src.navigation import navigate_to_recebidas, navigate_to_emitidas, apply_filter
from src.downloader import wait_for_page_ready
from src.scraper_visualizar import scrape_visualizar, SessionExpiredError
from src.generate_visualizar_excel import generate_visualizar_excel

BASE_URL = "https://www.nfse.gov.br/EmissorNacional"

def get_download_dir(base, accountant, name, month):
    safe_name = name.replace("/", "_").replace("\\", "_").replace(":", "_")
    path = os.path.join(base, accountant, safe_name, month)
    os.makedirs(path, exist_ok=True)
    return path

def log(name, msg):
    timestamp = datetime.now().strftime("%H:%M:%S")
    print(f"[{timestamp}] {name} | {msg}", flush=True)

def scrape_all_pages(page, company_name):
    """Collect permanent 50-digit chaves from Visualizar hrefs on all pages."""
    all_chaves = []
    base_url = page.url
    pg = 1

    while True:
        rows = page.query_selector_all('tr[data-chave]')
        page_chaves = []
        for row in rows:
            if 'nfse-cancelada' in (row.get_attribute('class') or ''):
                continue
            viz_link = row.query_selector("a[href*='/Visualizar/Index/']")
            if viz_link:
                href = viz_link.get_attribute('href') or ''
                m = re.search(r'/Visualizar/Index/(\d{50})', href)
                if m:
                    page_chaves.append(m.group(1))

        log(company_name, f"Pagina {pg}: {len(page_chaves)} notas")
        all_chaves.extend(page_chaves)

        proxima = page.query_selector("a[data-original-title='Próxima']")
        ultima  = page.query_selector("a[data-original-title='Última']")
        if not proxima and not ultima:
            break
        if pg >= 50:
            break

        pg += 1
        next_url = re.sub(r'pg=\d+', f'pg={pg}', base_url) if 'pg=' in base_url else base_url + ('&' if '?' in base_url else '?') + f'pg={pg}'
        page.goto(next_url)
        wait_for_page_ready(page)

    return all_chaves


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", required=True)
    args = parser.parse_args()

    with open(args.config, encoding="utf-8-sig") as f:
        run_config = json.load(f)

    company      = run_config["companies"][0]
    custom_start = run_config.get("start")
    custom_end   = run_config.get("end")
    mode         = run_config.get("mode", "reinf")

    with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "config", "settings.json"), encoding="utf-8") as f:
        settings = json.load(f)

    base_dir = settings["downloads_path"]

    if custom_start:
        d     = datetime.strptime(custom_start, "%d/%m/%Y")
        month = d.strftime("%m-%Y")
    else:
        today     = date.today()
        first_day = (today - relativedelta(months=1)).replace(day=1)
        month     = first_day.strftime("%m-%Y")

    cnpj       = company["cnpj"]
    name       = company["name"]
    password   = company["password"]
    accountant = company["accountant"]

    log(name, f"Iniciando modo visualizar â€” {month}")

    download_dir = get_download_dir(base_dir, accountant, name, month)
    temp_dir     = tempfile.mkdtemp(prefix="nfse_viz_")

    try:
        with sync_playwright() as p:
            log(name, "Abrindo navegador...")
            context = create_browser_for_company(p, temp_dir)
            try:
                log(name, "Fazendo login...")
                page = login(context, cnpj, password, name)
                if not page:
                    log(name, "ERRO â€” Login falhou")
                    sys.exit(1)

                log(name, "Login efetuado")

                if mode == 'emitidas':
                    log(name, "Navegando para Emitidas...")
                    navigate_to_emitidas(page)
                else:
                    log(name, "Navegando para Recebidas...")
                    navigate_to_recebidas(page)

                log(name, "Aplicando filtro...")
                apply_filter(page, custom_start, custom_end)

                log(name, "Coletando chaves permanentes...")
                chaves = scrape_all_pages(page, name)

                if not chaves:
                    log(name, "Nenhuma nota encontrada")
                    page.close()
                    context.close()
                    sys.exit(0)

                log(name, f"{len(chaves)} nota(s) encontradas â€” raspando dados...")

                notas  = []
                failed = 0
                for i, chave in enumerate(chaves, 1):
                    log(name, f"[{i}/{len(chaves)}] Visualizar {chave}")
                    for login_attempt in range(3):
                        try:
                            data = scrape_visualizar(page, chave)
                            if data:
                                notas.append(data)
                            else:
                                failed += 1
                                log(name, f"[AVISO] Falha {chave}")
                            break
                        except SessionExpiredError as e:
                            log(name, f"[AVISO] {e} â€” fazendo re-login ({login_attempt+1}/3)")
                            try:
                                page.close()
                            except:
                                pass
                            page = login(context, cnpj, password, name)
                            if not page:
                                log(name, "ERRO â€” Re-login falhou")
                                break
                            if mode == 'emitidas':
                                navigate_to_emitidas(page)
                            else:
                                navigate_to_recebidas(page)
                            apply_filter(page, custom_start, custom_end)
                    time.sleep(0.5)

                    if i % 20 == 0:
                        log(name, "Keepalive...")

                page.close()
                context.close()

                log(name, f"{len(notas)} notas raspadas | {failed} falhas")

                if notas:
                    federal   = [n for n in notas if n['is_federal'] and not n.get('is_cancelada')]
                    municipal = [n for n in notas if n['is_municipal'] and not n.get('is_cancelada')]
                    log(name, f"Federal: {len(federal)} | Municipal: {len(municipal)}")

                    excel_path = generate_visualizar_excel(name, month, notas, download_dir, company_cnpj=cnpj)
                    log(name, f"Excel: {excel_path}")

                    from src.reconstruct_xml import save_reconstructed_xmls
                    from generate_fiscal import generate_fiscal, generate_fiscal_txt
                    log(name, "Reconstruindo XMLs...")
                    save_reconstructed_xmls(notas, download_dir, federal_only=False)
                    if federal:
                        log(name, "Gerando fiscal XLSX e TXT...")
                        generate_fiscal(name, download_dir, month)
                        generate_fiscal_txt(name, download_dir, month)

                log(name, "Concluido com sucesso")
                sys.exit(0)

            except Exception as e:
                log(name, f"ERRO â€” {e}")
                try:
                    context.close()
                except:
                    pass
                sys.exit(1)

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

if __name__ == "__main__":
    main()
