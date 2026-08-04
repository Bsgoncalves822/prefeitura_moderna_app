"""
Guias / Boletos worker for PM Flask app.

Flow (confirmed from HAR + HTML):
  1. Select company → Acessar Movimento Econômico  (same as ISS/NFS-e)
  2. mostrarConteudo('iss-levantamento_debitos.php')
  3. Parse page HTML for forms with action="emissao_boleto.php?st_cartao=1"
     — these are the downloadable boletos (orange button)
     — greyed-out buttons have no form, just a modal trigger → skip
  4. For each downloadable form, POST emissao_boleto.php?st_cartao=1
     with the hidden field values → returns PDF bytes
  5. Save PDF to output/Guias/{run_date}/{company_name}/boleto_{id}.pdf

dt_corrige (monetary correction date) = today's date at runtime.
"""
import re
import time
import random
from datetime import datetime
from pathlib import Path

from bs4 import BeautifulSoup

BASE = "https://tubarao-sc.prefeituramoderna.com.br/meuiss_new"


def safe_name(s: str) -> str:
    for ch in r'\/:*?"<>|':
        s = s.replace(ch, "_")
    return s.strip().rstrip(".")[:60].strip()


def login(user, password, visible=False):
    from iss_worker import login as _login
    return _login(user, password, visible)


def get_company_list(page):
    from iss_worker import get_company_list as _get
    return _get(page)


def go_to_empresas(page):
    from iss_worker import go_to_empresas as _go
    _go(page)


def get_debitos_page(page, company_id: str, dt_correcao: str) -> str:
    """
    Navigate to Consulta de Débitos for a company and return page HTML.
    dt_correcao: DD/MM/YYYY — used for monetary correction (boleto value).
    """
    go_to_empresas(page)
    page.wait_for_selector("select#clientes", timeout=10000)
    page.select_option("select#clientes", company_id)
    page.wait_for_timeout(800)

    page.click("input[value='Acessar Movimento Econômico']")
    page.wait_for_load_state("networkidle", timeout=30000)
    page.wait_for_timeout(1000)

    page.evaluate("() => { mostrarConteudo('iss-levantamento_debitos.php'); }")
    page.wait_for_timeout(2000)

    # Set correction date to today and status parcela = Aberta (value=1)
    try:
        page.wait_for_selector("#dt_correcao", timeout=8000)
        page.evaluate(
            f"() => {{ document.querySelector('#dt_correcao').value = '{dt_correcao}'; }}"
        )
        # Set parcela status to Aberta
        page.select_option("select#select_parcela", "1")
        page.wait_for_timeout(300)

        # Click Atualizar
        page.click("button:has-text('Atualizar'), input[value='Atualizar']")
        page.wait_for_load_state("networkidle", timeout=15000)
        page.wait_for_timeout(1500)
    except Exception:
        pass  # If filter fails, still try to parse what's there

    return page.content()


def parse_boleto_forms(html: str, dt_correcao: str) -> list[dict]:
    """
    Parse all downloadable boleto forms from the page HTML.
    Returns list of field dicts ready to POST to emissao_boleto.php.
    """
    soup = BeautifulSoup(html, "html.parser")
    forms = soup.find_all("form", attrs={"action": re.compile(r"emissao_boleto\.php")})

    boletos = []
    for form in forms:
        fields = {}
        for inp in form.find_all("input", {"type": "hidden"}):
            name = inp.get("name", "")
            val  = inp.get("value", "")
            if name:
                fields[name] = val

        if not fields.get("id_dividasparcelas"):
            continue

        # Override dt_corrige with today's date
        fields["dt_corrige"] = dt_correcao

        # Build a human-readable label for logging/filename
        ds    = fields.get("ds_divida", "DEBITO")[:30]
        id_dp = fields.get("id_dividasparcelas", "0")
        fields["_label"] = f"{ds}_{id_dp}"

        boletos.append(fields)

    return boletos


def download_boleto_pdf(page, fields: dict) -> bytes | None:
    """POST emissao_boleto.php and return PDF bytes."""
    # Remove internal label before posting
    post_fields = {k: v for k, v in fields.items() if not k.startswith("_")}

    try:
        response = page.request.post(
            f"{BASE}/emissao_boleto.php?st_cartao=1",
            form=post_fields,
        )
        if response.status != 200:
            return None
        body = response.body()
        # PDF starts with %PDF
        if not body or not body[:4] == b"%PDF":
            return None
        return body
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Main generator
# ---------------------------------------------------------------------------

def run_guias(user: str, password: str, output_base: str,
              visible: bool = False):
    """
    Generator — yields log strings.
    Ends with __DONE__:code and __SUMMARY__:json.
    Downloads all available boletos for all companies, saving PDFs to output_base.
    """
    import json

    def log(msg):
        return f"[{datetime.now().strftime('%H:%M:%S')}] {msg}"

    today_str   = datetime.now().strftime("%d/%m/%Y")
    folder_date = datetime.now().strftime("%d-%m-%Y")

    yield log("Iniciando browser...")
    try:
        p, browser, context, page = login(user, password, visible)
    except Exception as e:
        yield log(f"ERRO — Login: {e}")
        yield "__DONE__:1"
        return

    yield log("Login efetuado. Carregando empresas...")
    try:
        companies = get_company_list(page)
    except Exception as e:
        yield log(f"ERRO — Lista de empresas: {e}")
        browser.close(); p.stop()
        yield "__DONE__:1"
        return

    yield log(f"{len(companies)} empresas | Data de correção: {today_str}")

    out_base = Path(output_base) / "Guias" / folder_date
    out_base.mkdir(parents=True, exist_ok=True)

    total   = len(companies)
    stats   = {"with_boletos": 0, "no_boletos": 0, "error": 0, "pdfs": 0}
    summary = []

    try:
        for i, comp in enumerate(companies, 1):
            cid  = comp["id"]
            name = comp["name"]
            yield log(f"[{i}/{total}] {name}")

            try:
                html    = get_debitos_page(page, cid, today_str)
                boletos = parse_boleto_forms(html, today_str)

                if not boletos:
                    yield log(f"  -> Sem boletos disponíveis")
                    stats["no_boletos"] += 1
                    continue

                yield log(f"  -> {len(boletos)} boleto(s) disponível(is)")

                company_dir = out_base / safe_name(name)
                company_dir.mkdir(parents=True, exist_ok=True)

                downloaded = 0
                for b in boletos:
                    label = b.get("_label", b.get("id_dividasparcelas", "boleto"))
                    pdf   = download_boleto_pdf(page, b)
                    if pdf:
                        fname = f"boleto_{safe_name(label)}.pdf"
                        (company_dir / fname).write_bytes(pdf)
                        downloaded += 1
                        yield log(f"    ✓ {label}")
                    else:
                        yield log(f"    ✗ {label} — falha no download")
                    time.sleep(0.5)

                stats["with_boletos"] += 1
                stats["pdfs"]         += downloaded
                summary.append({"name": name, "boletos": downloaded})

            except Exception as e:
                yield log(f"  -> ERRO: {e}")
                stats["error"] += 1

            if i < total:
                time.sleep(random.uniform(0.8, 1.8))

    except GeneratorExit:
        pass
    finally:
        browser.close()
        p.stop()

    yield log(
        f"Concluído — {stats['with_boletos']} empresas com boletos | "
        f"{stats['pdfs']} PDFs baixados | {stats['no_boletos']} sem boletos | "
        f"{stats['error']} erros"
    )
    yield "__DONE__:0"
    yield f"__SUMMARY__:{json.dumps(summary, ensure_ascii=False)}"
