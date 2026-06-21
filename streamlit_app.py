from datetime import datetime
import time
from pathlib import Path
from typing import Optional, Tuple
from html import escape
import base64
import json

import gspread
import pandas as pd
import requests
import streamlit as st
from google.oauth2.service_account import Credentials

# =============================
# KONFIGURACE
# =============================
SERVICE_ACCOUNT_FILE = "google-service-account.json"
SPREADSHEET_ID = "1TxtTH5mkiv5s6Z6NskIKzEg407Hw6e2BqnKWCLBX0c4"

ORDERS_SHEET_NAME = "Objednavky"
STOCK_SHEET_NAME = "Sklad"

WORKER_STATUS_ENDPOINT = "https://crimson-wind-bf3b.lukas-minx.workers.dev/send-status-email"
WORKER_PACKETA_ENDPOINT = "https://crimson-wind-bf3b.lukas-minx.workers.dev/create-packeta"
WORKER_PACKETA_CANCEL_ENDPOINT = "https://crimson-wind-bf3b.lukas-minx.workers.dev/cancel-packeta"

# Bezpečnostní token pro interní Worker endpointy.
# Token už NEPATŘÍ přímo do kódu.
# Lokálně ho ulož do souboru .streamlit/secrets.toml:
# WORKER_ADMIN_TOKEN = "tvuj_token"
# Online ho nastav v secrets prostředí, kde dashboard poběží.
WORKER_ADMIN_TOKEN = ""

LOGO_URL = "https://pc-mobil-servis.s21.cdn-upgates.com/g/g69c0252c113b3-x69b5b2e466335-logoo.png"

# =============================
# FAKTURACE – DOPLŇTE PODLE REALITY
# =============================
SUPPLIER_NAME = "Mobile Protection s.r.o."
SUPPLIER_ADDRESS = "U Dlouhé stěny 4, 586 01 Jihlava"
SUPPLIER_ICO = "17865808"
SUPPLIER_DIC = "CZ17865808"
SUPPLIER_EMAIL = "servis@pcmobilshop.cz"
SUPPLIER_BANK_ACCOUNT = "5343505003/5500"
SUPPLIER_IS_VAT_PAYER = True  # True pokud jste plátci DPH

INVOICE_PREFIX = "FA"
INVOICE_PAYMENT_METHOD = "Dobírka"
INVOICE_VAT_RATE = 0.21

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

ORDER_REQUIRED_COLUMNS = [
    "ID",
    "Datum",
    "Cas",
    "Jmeno",
    "Telefon",
    "Email",
    "ServiceID",
    "Produkt",
    "Oprava",
    "Cena",
    "Poznamka",
    "PacketaPointId",
    "PacketaPointName",
    "PacketaPointValue",
    "Stav",
    "SkladPriObjednani",
    "PosledniEmailTyp",
    "EmailZakaznikOdeslan",
    "EmailAdminOdeslan",
    "InterniPoznamka",
    "DatumPosledniZmenyStavu",
    "TrackingCislo",
    "PacketaPacketId",
    "PacketaLabelUrl",
    "PacketaCreated",
    "PacketaInsuranceValue",
    "PacketaStatus",
    "Firma",
    "ICO",
    "DIC",
]

STOCK_REQUIRED_COLUMNS = [
    "ServiceID",
    "Produkt",
    "Oprava",
    "Dil",
    "PocetSkladem",
    "DostupnostWeb",
    "StatusText",
    "KontaktKdyzNeni",
    "Aktivni",
]

ORDER_STATUSES = [
    "Objednávka přijata",
    "Čekáme na doručení telefonu",
    "Telefon přijat",
    "Čekáme na díl",
    "V opravě",
    "Hotovo",
    "Odesláno zpět",
]

STATUS_TO_WORKER_STATUS = {
    "Telefon přijat": "prijato",
    "V opravě": "v_oprave",
    "Odesláno zpět": "odeslano",
}


# =============================
# GOOGLE SHEETS
# =============================
@st.cache_resource
def get_gspread_client() -> gspread.Client:
    """Vrátí Google Sheets klienta.

    Online nasazení: bere service account ze Streamlit Secrets ze sekce
    [gcp_service_account].

    Lokální záloha: pokud secrets nejsou vyplněné, použije soubor
    google-service-account.json vedle aplikace.
    """
    try:
        service_account_info = dict(st.secrets.get("gcp_service_account", {}))
    except Exception:
        service_account_info = {}

    if service_account_info:
        credentials = Credentials.from_service_account_info(
            service_account_info,
            scopes=SCOPES,
        )
        return gspread.authorize(credentials)

    service_file = Path(SERVICE_ACCOUNT_FILE)
    if not service_file.exists():
        raise FileNotFoundError(
            "Chybí Google service account. Online ho vlož do Streamlit Secrets "
            "jako sekci [gcp_service_account], lokálně jako google-service-account.json."
        )

    credentials = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=SCOPES,
    )
    return gspread.authorize(credentials)


@st.cache_resource
def open_spreadsheet() -> gspread.Spreadsheet:
    if not SPREADSHEET_ID:
        raise ValueError("Chybí SPREADSHEET_ID.")
    client = get_gspread_client()
    return client.open_by_key(SPREADSHEET_ID)


def get_worksheet(sheet_name: str) -> gspread.Worksheet:
    spreadsheet = open_spreadsheet()
    return spreadsheet.worksheet(sheet_name)


def ensure_headers(worksheet: gspread.Worksheet, required_columns: list[str]) -> None:
    existing_headers = worksheet.row_values(1)

    if not existing_headers:
        worksheet.update("A1", [required_columns])
        return

    if existing_headers == required_columns:
        return

    missing = [col for col in required_columns if col not in existing_headers]
    if not missing:
        return

    new_headers = existing_headers + missing
    worksheet.update("A1", [new_headers])


@st.cache_data(ttl=20)
def load_dataframe(sheet_name: str, required_columns: list[str]) -> pd.DataFrame:
    ws = get_worksheet(sheet_name)
    ensure_headers(ws, required_columns)

    records = ws.get_all_records(expected_headers=required_columns)
    df = pd.DataFrame(records)

    if df.empty:
        df = pd.DataFrame(columns=required_columns)

    for col in required_columns:
        if col not in df.columns:
            df[col] = ""

    return df[required_columns].copy()


def clear_data_cache() -> None:
    load_dataframe.clear()


def save_cell_by_header(
    worksheet: gspread.Worksheet,
    row_number: int,
    header_name: str,
    value: str,
) -> None:
    headers = worksheet.row_values(1)
    if header_name not in headers:
        raise ValueError(f"Sloupec '{header_name}' nebyl nalezen.")

    col_index = headers.index(header_name) + 1
    worksheet.update_cell(row_number, col_index, value)


def append_stock_row(
    worksheet: gspread.Worksheet,
    service_id: str,
    produkt: str,
    oprava: str,
    dil: str,
    pocet_skladem: int,
    kontakt_kdyz_neni: str,
    aktivni: str,
) -> None:
    dostupnost_web, status_text = normalize_stock_row(str(pocet_skladem))
    worksheet.append_row([
        service_id,
        produkt,
        oprava,
        dil,
        str(pocet_skladem),
        dostupnost_web,
        status_text,
        kontakt_kdyz_neni,
        aktivni,
    ])


def service_id_exists_in_stock(df: pd.DataFrame, service_id: str) -> bool:
    if df.empty:
        return False
    matches = df["ServiceID"].astype(str).str.strip() == str(service_id).strip()
    return bool(matches.any())


def find_order_row_number_from_df(df: pd.DataFrame, order_id: str) -> Optional[int]:
    if df.empty:
        return None

    matches = df.index[df["ID"].astype(str).str.strip() == str(order_id).strip()]
    if len(matches) == 0:
        return None

    return int(matches[0]) + 2


def find_stock_row_number_from_df(df: pd.DataFrame, service_id: str) -> Optional[int]:
    if df.empty:
        return None

    matches = df.index[df["ServiceID"].astype(str).str.strip() == str(service_id).strip()]
    if len(matches) == 0:
        return None

    return int(matches[0]) + 2

def delete_order_by_id(df: pd.DataFrame, order_id: str) -> tuple[bool, str]:
    """Smaže objednávku z Google Sheets podle ID."""
    ws = get_worksheet(ORDERS_SHEET_NAME)
    row_number = find_order_row_number_from_df(df, order_id)

    if not row_number:
        return False, "Objednávku se nepodařilo najít."

    ws.delete_rows(row_number)
    clear_data_cache()
    return True, f"Objednávka {order_id} byla smazána."


def now_cz_string() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def normalize_stock_row(stock_value: str) -> Tuple[str, str]:
    try:
        count = int(str(stock_value).strip())
    except Exception:
        count = 0

    if count > 0:
        return "ANO", "Díl skladem"
    return "NE", "Momentálně čekáme na naskladnění dílu"


def get_worker_auth_headers() -> dict[str, str]:
    """Vrátí Authorization hlavičku pro interní Worker endpointy.

    Priorita:
    1) st.secrets["WORKER_ADMIN_TOKEN"] pro online nasazení
    2) konstanta WORKER_ADMIN_TOKEN pro lokální spuštění
    """
    token = ""

    try:
        token = str(st.secrets.get("WORKER_ADMIN_TOKEN", "") or "").strip()
    except Exception:
        token = ""

    if not token:
        token = str(WORKER_ADMIN_TOKEN or "").strip()

    if not token or token == "SEM_VLOZ_TEN_TOKEN_Z_CLOUDFLARE":
        raise ValueError(
            "Chybí WORKER_ADMIN_TOKEN. Vlož stejný token jako v Cloudflare Worker secrets."
        )

    return {"Authorization": f"Bearer {token}"}


def send_status_email_via_worker(
    order_row: pd.Series,
    status: str,
    tracking_number: str = "",
) -> Tuple[bool, str]:
    if not WORKER_STATUS_ENDPOINT.strip():
        return False, "WORKER_STATUS_ENDPOINT není vyplněný v kódu aplikace."

    worker_status = STATUS_TO_WORKER_STATUS.get(status)
    if not worker_status:
        return False, f"Pro stav '{status}' se email neposílá."

    customer_name = str(order_row.get("Jmeno", "") or "").strip()
    product = str(order_row.get("Produkt", "") or "").strip()
    repair = str(order_row.get("Oprava", "") or "").strip()
    order_id = str(order_row.get("ID", "") or "").strip()
    customer_email = str(order_row.get("Email", "") or "").strip()

    if not customer_email:
        return False, "Objednávka nemá vyplněný e-mail zákazníka."

    price = str(order_row.get("Cena", "") or "").strip()
    pickup_point = str(
        order_row.get("PacketaPointValue", "") or order_row.get("PacketaPointName", "") or ""
    ).strip()

    payload = {
        "to": customer_email,
        "status": worker_status,
        "customerName": customer_name,
        "deviceModel": product,
        "repairType": repair,
        "orderNumber": order_id,
        "trackingNumber": tracking_number.strip() if tracking_number else "",
        "price": price,
        "pickupPoint": pickup_point,
    }

    try:
        response = requests.post(
            WORKER_STATUS_ENDPOINT,
            json=payload,
            headers=get_worker_auth_headers(),
            timeout=20,
        )
        data = response.json()
    except Exception as exc:
        return False, f"Nepodařilo se zavolat endpoint pro e-mail: {exc}"

    if response.ok and data.get("success"):
        return True, "E-mail zákazníkovi byl odeslán."

    return False, data.get("message") or data.get("error") or "Odeslání e-mailu selhalo."


# =============================
# POMOCNÉ FUNKCE UI
# =============================
def to_int_safe(value) -> int:
    try:
        return int(float(str(value).replace(",", ".").strip()))
    except Exception:
        return 0


def safe_text(value: object) -> str:
    txt = str(value or "").strip()
    return txt if txt else "-"


def build_packeta_tracking_url(tracking_number: object) -> str:
    tracking = str(tracking_number or "").strip()
    if not tracking:
        return ""
    return f"https://tracking.packeta.com/cs/tracking/search?id={tracking}"


def download_packeta_label_pdf(label_url: str) -> tuple[bool, bytes | None, str]:
    """Stáhne PDF štítek přes Worker s Authorization headerem.

    Přímé otevření label_url v prohlížeči po zabezpečení Workeru nefunguje,
    protože prohlížeč neposílá WORKER_ADMIN_TOKEN.
    """
    label_url = str(label_url or "").strip()
    if not label_url:
        return False, None, "Chybí URL štítku."

    try:
        response = requests.get(
            label_url,
            headers=get_worker_auth_headers(),
            timeout=30,
        )
    except Exception as exc:
        return False, None, f"Nepodařilo se stáhnout štítek: {exc}"

    content_type = response.headers.get("Content-Type", "")
    if response.ok and "application/pdf" in content_type.lower():
        return True, response.content, "Štítek byl načten."

    try:
        data = response.json()
        msg = data.get("message") or data.get("error") or str(data)
    except Exception:
        msg = response.text[:500]

    return False, None, f"Worker nevrátil PDF štítek: {msg}"


def clear_packeta_data_for_order(order_id: str) -> tuple[bool, str]:
    """Smaže logistické údaje ze Sheets.

    Pozor: toto zatím nemaže zásilku přímo v systému Packeta.
    Pro skutečné zrušení zásilky bude potřeba doplnit Worker endpoint /cancel-packeta.
    """
    ws = get_worksheet(ORDERS_SHEET_NAME)
    row_number = find_order_row_number_from_df(orders_df, order_id)
    if not row_number:
        return False, "Objednávku se nepodařilo najít v Google Sheets."

    for col in [
        "TrackingCislo",
        "PacketaPacketId",
        "PacketaLabelUrl",
        "PacketaCreated",
        "PacketaStatus",
    ]:
        save_cell_by_header(ws, row_number, col, "")

    clear_data_cache()
    return True, "Údaje o zásilce byly smazány z dashboardu / Google Sheets."


def parse_price(val: object) -> float:
    try:
        cleaned = (
            str(val or "")
            .replace("Kč", "")
            .replace("CZK", "")
            .replace("\xa0", "")
            .replace(" ", "")
            .replace(",", ".")
            .strip()
        )
        return float(cleaned) if cleaned else 0.0
    except Exception:
        return 0.0


def yes_no_badge_html(value: object) -> str:
    txt = str(value or "").strip().upper()
    if txt == "ANO":
        color = "#16a34a"
        label = "ANO"
    elif txt == "NE":
        color = "#dc2626"
        label = "NE"
    else:
        color = "#475569"
        label = safe_text(value)

    return f"""
    <span style="
        display:inline-block;
        padding:5px 10px;
        border-radius:999px;
        background:{color};
        color:white;
        font-size:12px;
        font-weight:700;
        line-height:1;
    ">{label}</span>
    """


def status_color(status: str) -> str:
    mapping = {
        "Objednávka přijata": "#16a34a",
        "Nová objednávka": "#16a34a",  # zpětná kompatibilita se starými řádky v Google Sheets
        "Čekáme na doručení telefonu": "#7c3aed",
        "Telefon přijat": "#0ea5e9",
        "Čekáme na díl": "#dc2626",
        "V opravě": "#9333ea",
        "Hotovo": "#059669",
        "Odesláno zpět": "#2563eb",
        "Dokončeno": "#2563eb",  # starý stav bereme vizuálně jako odesláno zpět
    }
    return mapping.get(status, "#475569")


def status_emoji(status: str) -> str:
    mapping = {
        "Objednávka přijata": "✅",
        "Nová objednávka": "✅",
        "Čekáme na doručení telefonu": "🟣",
        "Telefon přijat": "📱",
        "Čekáme na díl": "🔴",
        "V opravě": "🟣",
        "Hotovo": "✅",
        "Odesláno zpět": "📦",
        "Dokončeno": "📦",
    }
    return mapping.get(status, "•")


def styled_status_text(status: str) -> str:
    label = normalize_status_label(status) if "normalize_status_label" in globals() else str(status or "")
    return f"{status_emoji(label)} {label}"


def badge_html(text: str, color: str) -> str:
    return f"""
    <span style="
        display:inline-block;
        padding:6px 12px;
        border-radius:999px;
        background:{color};
        color:white;
        font-size:12px;
        font-weight:700;
        line-height:1;
    ">{text}</span>
    """


def normalize_status_label(status: object) -> str:
    status_text = str(status or "").strip()
    if status_text == "Nová objednávka":
        return "Objednávka přijata"
    if status_text == "Dokončeno":
        return "Odesláno zpět"
    return status_text


def status_badge_html(status: str) -> str:
    label = normalize_status_label(status)
    return f"""
    <span style="
        display:inline-block;
        padding:6px 12px;
        border-radius:999px;
        background:{status_color(label)};
        color:white;
        font-size:12px;
        font-weight:700;
        line-height:1;
    ">{safe_text(label)}</span>
    """



# =============================
# FAKTURY
# =============================
def format_czk(value: object) -> str:
    amount = parse_price(value)
    if amount <= 0:
        return safe_text(value)
    return f"{int(round(amount)):,} Kč".replace(",", " ")


def build_invoice_number(order_id: object) -> str:
    year = datetime.now().strftime("%Y")
    clean_order = str(order_id or "").strip().replace("SRV-", "")
    if not clean_order:
        clean_order = datetime.now().strftime("%m%d%H%M")
    return f"{INVOICE_PREFIX}-{year}-{clean_order}"


def build_invoice_html(order_row: pd.Series) -> str:
    order_id = safe_text(order_row.get("ID", ""))
    invoice_number = build_invoice_number(order_id)
    issue_date = datetime.now().strftime("%d.%m.%Y")
    due_date = (datetime.now() + pd.Timedelta(days=14)).strftime("%d.%m.%Y")
    payment_method = INVOICE_PAYMENT_METHOD

    customer_name = safe_text(order_row.get("Jmeno", ""))
    customer_email = safe_text(order_row.get("Email", ""))
    customer_phone = safe_text(order_row.get("Telefon", ""))
    company = str(order_row.get("Firma", "") or "").strip()
    ico = str(order_row.get("ICO", "") or "").strip()
    dic = str(order_row.get("DIC", "") or "").strip()

    buyer_title = company if company else customer_name
    buyer_lines = []
    if company:
        buyer_lines.append(company)
        if customer_name != "-":
            buyer_lines.append(customer_name)
    else:
        buyer_lines.append(customer_name)

    if ico:
        buyer_lines.append(f"IČO: {ico}")
    if dic:
        buyer_lines.append(f"DIČ: {dic}")
    if customer_email != "-":
        buyer_lines.append(f"E-mail: {customer_email}")
    if customer_phone != "-":
        buyer_lines.append(f"Telefon: {customer_phone}")

    product = safe_text(order_row.get("Produkt", ""))
    repair = safe_text(order_row.get("Oprava", ""))
    price_raw = order_row.get("Cena", "")
    price_num = parse_price(price_raw)

    if SUPPLIER_IS_VAT_PAYER:
        total_with_vat = price_num
        base_without_vat = total_with_vat / (1 + INVOICE_VAT_RATE) if total_with_vat else 0
        vat_amount = total_with_vat - base_without_vat
        vat_percent = int(round(INVOICE_VAT_RATE * 100))
        unit_price = format_czk(base_without_vat)
        vat_price = format_czk(vat_amount)
        price = format_czk(total_with_vat)
        vat_note = f"Cena je uvedena včetně DPH {vat_percent} %."
    else:
        total_with_vat = price_num
        base_without_vat = price_num
        vat_amount = 0
        vat_percent = 0
        unit_price = format_czk(price_raw)
        vat_price = "0 Kč"
        price = format_czk(price_raw)
        vat_note = "Nejsme plátci DPH."

    item_name = f"Servisní oprava – {product}, {repair}"

    buyer_html = "<br>".join(escape(x) for x in buyer_lines if x)
    supplier_dic_line = f"<br>DIČ: {escape(SUPPLIER_DIC)}" if SUPPLIER_DIC else ""

    html = f"""
<!doctype html>
<html lang="cs">
<head>
  <meta charset="utf-8">
  <title>Faktura {escape(invoice_number)}</title>
  <style>
    @page {{
      size: A4;
      margin: 14mm;
    }}
    * {{
      box-sizing: border-box;
    }}
    body {{
      margin: 0;
      font-family: Arial, Helvetica, sans-serif;
      color: #111827;
      background: #ffffff;
      font-size: 13px;
      line-height: 1.45;
    }}
    .invoice {{
      width: 100%;
      max-width: 800px;
      margin: 0 auto;
      padding: 0;
    }}
    .top {{
      display: flex;
      justify-content: space-between;
      align-items: flex-start;
      border-bottom: 2px solid #17305f;
      padding-bottom: 18px;
      margin-bottom: 22px;
    }}
    .logo {{
      max-height: 46px;
      margin-bottom: 8px;
    }}
    .title {{
      text-align: right;
    }}
    .title h1 {{
      margin: 0;
      font-size: 30px;
      color: #17305f;
      letter-spacing: -0.5px;
    }}
    .title .num {{
      margin-top: 6px;
      font-size: 16px;
      font-weight: 800;
    }}
    .grid {{
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 18px;
      margin-bottom: 22px;
    }}
    .box {{
      border: 1px solid #dbe3ef;
      border-radius: 12px;
      padding: 14px 16px;
      background: #f8fbff;
      min-height: 150px;
    }}
    .box h2 {{
      margin: 0 0 10px 0;
      font-size: 15px;
      color: #17305f;
    }}
    .meta {{
      display: grid;
      grid-template-columns: 1fr 1fr 1fr 1fr;
      gap: 12px;
      margin-bottom: 22px;
    }}
    .meta .item {{
      border: 1px solid #dbe3ef;
      border-radius: 10px;
      padding: 12px;
      background: #ffffff;
    }}
    .label {{
      color: #64748b;
      font-size: 12px;
      margin-bottom: 4px;
    }}
    .value {{
      font-weight: 800;
      color: #111827;
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      margin-top: 12px;
      margin-bottom: 18px;
    }}
    th {{
      background: #17305f;
      color: white;
      text-align: left;
      padding: 10px;
      font-size: 12px;
    }}
    td {{
      border-bottom: 1px solid #e5e7eb;
      padding: 12px 10px;
      vertical-align: top;
    }}
    .right {{
      text-align: right;
    }}
    .total {{
      display: flex;
      justify-content: flex-end;
      margin-top: 10px;
    }}
    .total-box {{
      width: 320px;
      border: 2px solid #17305f;
      border-radius: 12px;
      padding: 14px 16px;
      background: #f8fbff;
    }}
    .total-row {{
      display: flex;
      justify-content: space-between;
      gap: 12px;
      font-size: 18px;
      font-weight: 900;
      color: #17305f;
    }}
    .note {{
      margin-top: 22px;
      color: #475569;
      font-size: 12px;
    }}
    .footer {{
      margin-top: 42px;
      padding-top: 14px;
      border-top: 1px solid #e5e7eb;
      color: #64748b;
      font-size: 12px;
    }}
    .print-controls {{
      margin: 0 auto 18px auto;
      max-width: 800px;
      display: flex;
      gap: 10px;
    }}
    .print-button {{
      display: inline-block;
      padding: 10px 16px;
      border-radius: 10px;
      border: none;
      background: #2563eb;
      color: white;
      font-weight: 800;
      cursor: pointer;
    }}
    @media print {{
      .print-controls {{
        display: none;
      }}
      body {{
        background: white;
      }}
    }}
  </style>
</head>
<body>
  <div class="print-controls">
    <button class="print-button" onclick="window.print()">Vytisknout / uložit jako PDF</button>
  </div>

  <div class="invoice">
    <div class="top">
      <div>
        <img src="{escape(LOGO_URL)}" class="logo" alt="PCmobilSHOP">
        <div style="font-weight:800;color:#17305f;">{escape(SUPPLIER_NAME)}</div>
        <div>{escape(SUPPLIER_ADDRESS)}</div>
        <div>IČO: {escape(SUPPLIER_ICO)}{supplier_dic_line}</div>
        <div>E-mail: {escape(SUPPLIER_EMAIL)}</div>
      </div>
      <div class="title">
        <h1>FAKTURA</h1>
        <div class="num">{escape(invoice_number)}</div>
      </div>
    </div>

    <div class="grid">
      <div class="box">
        <h2>Dodavatel</h2>
        <strong>{escape(SUPPLIER_NAME)}</strong><br>
        {escape(SUPPLIER_ADDRESS)}<br>
        IČO: {escape(SUPPLIER_ICO)}{supplier_dic_line}<br>
        E-mail: {escape(SUPPLIER_EMAIL)}
      </div>

      <div class="box">
        <h2>Odběratel</h2>
        <strong>{escape(buyer_title)}</strong><br>
        {buyer_html}
      </div>
    </div>

    <div class="meta">
      <div class="item">
        <div class="label">Datum vystavení</div>
        <div class="value">{issue_date}</div>
      </div>
      <div class="item">
        <div class="label">Datum splatnosti</div>
        <div class="value">{due_date}</div>
      </div>
      <div class="item">
        <div class="label">Forma úhrady</div>
        <div class="value">{escape(payment_method)}</div>
      </div>
      <div class="item">
        <div class="label">Číslo zakázky</div>
        <div class="value">{escape(order_id)}</div>
      </div>
    </div>

    <table>
      <thead>
        <tr>
          <th>Popis položky</th>
          <th class="right">Množství</th>
          <th class="right">Cena bez DPH</th>
          <th class="right">DPH</th>
          <th class="right">Celkem s DPH</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td>{escape(item_name)}</td>
          <td class="right">1 ks</td>
          <td class="right">{escape(unit_price)}</td>
          <td class="right">{vat_percent} %<br>{escape(vat_price)}</td>
          <td class="right"><strong>{escape(price)}</strong></td>
        </tr>
      </tbody>
    </table>

    <div class="total">
      <div class="total-box">
        <div style="display:flex;justify-content:space-between;margin-bottom:6px;color:#475569;font-size:13px;">
          <span>Základ DPH</span>
          <span>{escape(unit_price)}</span>
        </div>
        <div style="display:flex;justify-content:space-between;margin-bottom:10px;color:#475569;font-size:13px;">
          <span>DPH {vat_percent} %</span>
          <span>{escape(vat_price)}</span>
        </div>
        <div class="total-row">
          <span>Celkem k úhradě</span>
          <span>{escape(price)}</span>
        </div>
        <div style="margin-top:8px;color:#64748b;font-size:12px;">
          Forma úhrady: <strong>{escape(payment_method)}</strong>
        </div>
      </div>
    </div>

    <div class="note">
      <strong>Poznámka:</strong> {escape(vat_note)} Faktura byla vystavena k servisní zakázce {escape(order_id)}. Úhrada probíhá formou dobírky při převzetí zásilky.
    </div>

    <div class="footer">
      Děkujeme za využití servisu PCmobilSHOP.
    </div>
  </div>
</body>
</html>
"""
    return html


def build_invoice_filename(order_row: pd.Series) -> str:
    invoice_number = build_invoice_number(order_row.get("ID", ""))
    return f"{invoice_number}.html".replace("/", "-")

def build_invoice_data_url(order_row: pd.Series) -> str:
    html = build_invoice_html(order_row)
    encoded = base64.b64encode(html.encode("utf-8")).decode("ascii")
    return f"data:text/html;base64,{encoded}"

def invoice_open_button_html(order_row: pd.Series) -> str:
    url = build_invoice_data_url(order_row)
    return f"""
    <a href="{url}" target="_blank" style="
        display:block;
        text-align:center;
        padding:11px 14px;
        border:1px solid #dbe3ef;
        border-radius:12px;
        background:#ffffff;
        color:#0f172a !important;
        font-weight:800;
        text-decoration:none;
        min-height:42px;
        line-height:20px;
    ">🧾 Faktura</a>
    """


# =============================
# SESSION STATE
# =============================
if "selected_order_id" not in st.session_state:
    st.session_state.selected_order_id = None

if "order_visible_columns" not in st.session_state:
    st.session_state.order_visible_columns = [
        "ID",
        "Stav",
        "Jmeno",
        "Firma",
        "Telefon",
        "Produkt",
        "Oprava",
        "Cena",
    ]

if "delete_confirm_order_id" not in st.session_state:
    st.session_state.delete_confirm_order_id = None


# =============================
# PAGE STYLE
# =============================
st.set_page_config(page_title="Servisní dashboard", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
.block-container {
    padding-top: 1.2rem;
    padding-bottom: 2rem;
}
div[data-testid="stMetric"] {
    background: linear-gradient(135deg, #0f172a 0%, #16253f 100%);
    border: 1px solid #24354f;
    padding: 16px 18px;
    border-radius: 16px;
    box-shadow: 0 4px 18px rgba(0, 0, 0, 0.22);
}
div[data-testid="stMetricLabel"] {
    font-weight: 700 !important;
    color: #cbd5e1 !important;
}
div[data-testid="stMetricValue"] {
    color: #ffffff !important;
    font-weight: 800 !important;
}
div[data-testid="stDataFrame"] {
    border-radius: 14px;
    overflow: hidden;
}
</style>
""", unsafe_allow_html=True)


st.markdown("""
<style>
html, body, .stApp, [data-testid="stAppViewContainer"] {
    background: #f6f8fc !important;
    color: #0f172a !important;
}
</style>
""", unsafe_allow_html=True)


try:
    orders_df = load_dataframe(ORDERS_SHEET_NAME, ORDER_REQUIRED_COLUMNS)
    stock_df = load_dataframe(STOCK_SHEET_NAME, STOCK_REQUIRED_COLUMNS)
except Exception as exc:
    st.error(f"Nepodařilo se načíst Google Sheets: {repr(exc)}")
    st.exception(exc)
    st.stop()



# =============================
# MODERNÍ UI
# =============================
st.markdown("""
<style>
/* =============================
   FORCE LIGHT THEME
   ============================= */
html,
body,
.stApp,
[data-testid="stAppViewContainer"],
[data-testid="stHeader"],
[data-testid="stToolbar"],
[data-testid="stDecoration"],
.main,
.main .block-container {
    background: #f6f8fc !important;
    color: #0f172a !important;
}

[data-testid="stHeader"] {
    background: rgba(246, 248, 252, 0.92) !important;
}

.block-container {
    padding-top: 1.1rem !important;
    padding-bottom: 2.4rem !important;
    max-width: 1480px !important;
}

/* =============================
   SIDEBAR
   ============================= */
[data-testid="stSidebar"] {
    background: #ffffff !important;
    border-right: 1px solid #e5eaf3 !important;
}

[data-testid="stSidebar"] > div:first-child {
    padding-top: 1.3rem !important;
}

[data-testid="stSidebar"] * {
    color: #0f172a !important;
}

[data-testid="stSidebar"] img {
    margin-bottom: 14px;
}

.sidebar-user {
    border: 1px solid #e5eaf3;
    border-radius: 16px;
    padding: 14px;
    background: #f8fbff;
    margin-top: 20px;
}

.sidebar-avatar {
    width: 38px;
    height: 38px;
    background: #2563eb;
    color: #ffffff !important;
    border-radius: 999px;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    font-weight: 900;
    margin-right: 10px;
}

/* =============================
   GLOBAL TEXT
   ============================= */
h1, h2, h3, h4, h5, h6,
p, span, div, label,
.stMarkdown, .stMarkdown *,
[data-testid="stMarkdownContainer"],
[data-testid="stMarkdownContainer"] *,
[data-testid="stCaptionContainer"],
[data-testid="stCaptionContainer"] * {
    color: #0f172a !important;
}

hr {
    margin: 1.2rem 0;
    border-color: #e5eaf3 !important;
}

/* =============================
   CARDS
   ============================= */
.mod-card {
    background: #ffffff !important;
    border: 1px solid #e5eaf3 !important;
    border-radius: 18px !important;
    padding: 20px 22px !important;
    box-shadow: 0 8px 24px rgba(15, 23, 42, 0.055) !important;
    margin-bottom: 16px !important;
    color: #0f172a !important;
}

.mod-card,
.mod-card * {
    color: #0f172a !important;
}

.mod-card-title {
    display: flex;
    align-items: center;
    gap: 9px;
    font-size: 19px;
    font-weight: 900;
    color: #0f172a !important;
    margin-bottom: 18px;
}

.mod-icon {
    color: #2563eb !important;
    font-size: 20px;
}

.mod-kv {
    display: grid;
    grid-template-columns: 145px 1fr;
    gap: 11px 16px;
    font-size: 15px;
}

.mod-k {
    color: #64748b !important;
    font-weight: 800;
}

.mod-v {
    color: #0f172a !important;
    font-weight: 800;
    word-break: break-word;
}

.mod-note {
    background: linear-gradient(135deg, #eff6ff 0%, #f8fbff 100%) !important;
    border: 1px solid #dbeafe !important;
    border-radius: 14px !important;
    padding: 16px 18px !important;
    color: #1e40af !important;
    font-weight: 800 !important;
}

.mod-note * {
    color: #1e40af !important;
}

/* =============================
   HERO / PROGRESS
   ============================= */
.hero {
    background: #ffffff !important;
    border: 1px solid #e5eaf3 !important;
    border-radius: 22px !important;
    box-shadow: 0 10px 30px rgba(15, 23, 42, .06) !important;
    padding: 26px 30px !important;
    margin: 10px 0 20px 0 !important;
    color: #0f172a !important;
}

.hero,
.hero * {
    color: #0f172a !important;
}

.hero-grid {
    display: grid;
    grid-template-columns: 1fr;
    gap: 18px;
    align-items: center;
}

.hero-order {
    font-size: 34px;
    line-height: 1.05;
    font-weight: 950;
    color: #0f172a !important;
    margin-bottom: 13px;
}

.pill {
    display: inline-flex;
    align-items: center;
    padding: 7px 12px;
    border-radius: 999px;
    color: #ffffff !important;
    font-weight: 900;
    font-size: 13px;
}

.hero-meta {
    margin-top: 18px;
    color: #475569 !important;
    font-size: 14px;
    line-height: 1.9;
}

.hero-meta strong {
    color: #0f172a !important;
}

.progress {
    display: grid;
    grid-template-columns: repeat(5, 1fr);
    gap: 0;
    align-items: start;
}

.step {
    text-align: center;
    position: relative;
}

.step:not(:last-child)::after {
    content: "";
    position: absolute;
    top: 22px;
    left: 58%;
    right: -42%;
    height: 3px;
    background: #dbeafe;
    z-index: 0;
}

.dot {
    width: 46px;
    height: 46px;
    border-radius: 999px;
    display: flex;
    align-items: center;
    justify-content: center;
    margin: 0 auto 10px auto;
    font-weight: 900;
    position: relative;
    z-index: 1;
}

.step-title {
    font-size: 14px;
    font-weight: 900;
    color: #334155 !important;
}

.step-date {
    font-size: 12px;
    color: #94a3b8 !important;
    margin-top: 4px;
}

/* =============================
   BUTTONS / INPUTS
   ============================= */
.stButton > button,
.stDownloadButton > button,
[data-testid="stLinkButton"] > a {
    border-radius: 12px !important;
    font-weight: 800 !important;
    border: 1px solid #dbe3ef !important;
    min-height: 42px !important;
    background: #ffffff !important;
    color: #0f172a !important;
}

.stButton > button[kind="primary"],
button[kind="primary"] {
    background: #2563eb !important;
    color: #ffffff !important;
    border-color: #2563eb !important;
}

.stTextInput input,
.stTextArea textarea,
.stNumberInput input,
.stSelectbox div[data-baseweb="select"],
.stMultiSelect div[data-baseweb="select"] {
    background: #ffffff !important;
    color: #0f172a !important;
    border-color: #dbe3ef !important;
}

.stTextInput input::placeholder,
.stTextArea textarea::placeholder {
    color: #94a3b8 !important;
}

.stCheckbox label,
.stCheckbox label * {
    color: #0f172a !important;
}

/* =============================
   METRICS
   ============================= */
div[data-testid="stMetric"] {
    background: linear-gradient(135deg, #0f2a55 0%, #173b7a 100%) !important;
    border: 1px solid #214a92 !important;
    padding: 16px 18px !important;
    border-radius: 16px !important;
    box-shadow: 0 8px 24px rgba(37, 99, 235, 0.12) !important;
}

div[data-testid="stMetric"] * {
    color: #ffffff !important;
}

div[data-testid="stMetricLabel"] {
    font-weight: 800 !important;
    color: #dbeafe !important;
}

div[data-testid="stMetricValue"] {
    color: #ffffff !important;
    font-weight: 900 !important;
}

/* =============================
   DATAFRAME
   ============================= */
div[data-testid="stDataFrame"] {
    border-radius: 16px !important;
    overflow: hidden !important;
    border: 1px solid #e5eaf3 !important;
    box-shadow: 0 8px 24px rgba(15, 23, 42, 0.045) !important;
    background: #ffffff !important;
}

/* =============================
   EXPANDERS / ALERTS
   ============================= */
.streamlit-expanderHeader,
.streamlit-expanderHeader * {
    color: #0f172a !important;
    background: #ffffff !important;
}

[data-testid="stExpander"] {
    background: #ffffff !important;
    border: 1px solid #e5eaf3 !important;
    border-radius: 14px !important;
}

[data-testid="stAlert"],
[data-testid="stAlert"] * {
    color: #0f172a !important;
}

/* =============================
   RESPONSIVE
   ============================= */
@media (max-width: 1000px) {
    .hero-grid {
        grid-template-columns: 1fr;
    }
    .progress {
        grid-template-columns: 1fr;
        gap: 12px;
    }
    .step:not(:last-child)::after {
        display: none;
    }
    .mod-kv {
        grid-template-columns: 1fr;
        gap: 5px;
    }
}

/* Detail cards equal height */
.mod-card {
    min-height: 100px;
}
iframe {
    background: #ffffff !important;
    border-radius: 14px !important;
}

.compact-notes {
    margin-top: 12px;
}
.status-action-wrap {
    background: #ffffff !important;
    border: 1px solid #e5eaf3 !important;
    border-radius: 18px !important;
    padding: 18px 20px !important;
    box-shadow: 0 8px 24px rgba(15, 23, 42, 0.055) !important;
    margin-bottom: 18px !important;
}
.status-action-title {
    font-size: 17px;
    font-weight: 900;
    color: #0f172a !important;
    margin-bottom: 12px;
}
.note-compact-title {
    font-size: 14px;
    font-weight: 900;
    color: #0f172a !important;
    margin-bottom: 8px;
}

/* v6 compact detail adjustments */
.hero-grid {
    grid-template-columns: 300px 1fr !important;
}
.status-inline-wrap {
    padding-left: 22px;
    border-left: 1px solid #e5eaf3;
}
.status-inline-title {
    font-size: 14px;
    font-weight: 900;
    color: #64748b !important;
    margin-bottom: 10px;
}
.compact-note-box {
    background:#ffffff !important;
    border:1px solid #e5eaf3 !important;
    border-radius:14px !important;
    padding:12px 14px !important;
    box-shadow:0 5px 18px rgba(15,23,42,.04) !important;
}
.compact-note-title {
    font-size:13px;
    font-weight:900;
    color:#0f172a !important;
    margin-bottom:6px;
}
.compact-note-content {
    background:#eff6ff !important;
    border:1px solid #dbeafe !important;
    border-radius:10px !important;
    padding:10px 12px !important;
    color:#1e40af !important;
    font-weight:800 !important;
    max-height:54px;
    overflow:auto;
}
.compact-note-content * {
    color:#1e40af !important;
}
.compact-internal textarea {
    min-height:54px !important;
    height:54px !important;
}

.packeta-panel {
    background:#ffffff !important;
    border:1px solid #e5eaf3 !important;
    border-radius:18px !important;
    padding:18px 20px !important;
    box-shadow:0 8px 24px rgba(15,23,42,.055) !important;
    margin:16px 0 !important;
}

/* v9 tracking placement */
div[data-testid="stTextInput"] label {
    font-weight: 800 !important;
}

/* Moderní vlastní seznam objednávek */
.order-list-header {
    background:#ffffff !important;
    border:1px solid #e5eaf3 !important;
    border-radius:14px 14px 0 0 !important;
    padding:11px 12px !important;
    font-weight:900 !important;
    color:#64748b !important;
    font-size:12px !important;
}
.order-row-wrap {
    background:#ffffff !important;
    border-left:1px solid #e5eaf3 !important;
    border-right:1px solid #e5eaf3 !important;
    border-bottom:1px solid #eef2f7 !important;
    padding:6px 12px !important;
}
.order-row-wrap:hover {
    background:#f8fbff !important;
}
.order-cell-muted {
    color:#64748b !important;
    font-size:13px !important;
    font-weight:700 !important;
}
.order-cell-strong {
    color:#0f172a !important;
    font-size:13px !important;
    font-weight:900 !important;
}

/* v14 lepší seznam objednávek + sidebar */
[data-testid="stSidebar"] [role="radiogroup"] label {
    border-radius: 12px !important;
    padding: 8px 10px !important;
    margin-bottom: 4px !important;
    transition: background .15s ease;
}
[data-testid="stSidebar"] [role="radiogroup"] label:hover {
    background: #f1f5f9 !important;
}
[data-testid="stSidebar"] [role="radiogroup"] label[data-baseweb="radio"] > div:first-child {
    margin-right: 8px !important;
}

.order-table-shell {
    background: #ffffff !important;
    border: 1px solid #e5eaf3 !important;
    border-radius: 20px !important;
    padding: 16px 18px !important;
    box-shadow: 0 10px 30px rgba(15,23,42,.06) !important;
    margin-top: 12px !important;
}
.order-table-title {
    font-size: 17px;
    font-weight: 950;
    color: #0f172a !important;
    margin-bottom: 12px;
}
.order-list-header {
    background:#f8fbff !important;
    border:1px solid #e5eaf3 !important;
    border-radius:12px !important;
    padding:10px 12px !important;
    font-weight:950 !important;
    color:#475569 !important;
    font-size:12px !important;
    text-transform: uppercase;
    letter-spacing: .02em;
}
.order-row-card {
    background:#ffffff !important;
    border:1px solid #eef2f7 !important;
    border-radius:14px !important;
    padding:8px 10px !important;
    margin:7px 0 !important;
    box-shadow: 0 3px 12px rgba(15,23,42,.025) !important;
}
.order-row-card:hover {
    border-color:#bfdbfe !important;
    box-shadow: 0 8px 22px rgba(37,99,235,.08) !important;
    background:#fbfdff !important;
}
.order-cell-muted {
    color:#64748b !important;
    font-size:13px !important;
    font-weight:750 !important;
}
.order-cell-strong {
    color:#0f172a !important;
    font-size:13px !important;
    font-weight:900 !important;
}

/* v16 seznam objednávek – kompaktnější a čitelnější */
.order-table-shell {
    background: #ffffff !important;
    border: 1px solid #e5eaf3 !important;
    border-radius: 18px !important;
    padding: 14px 16px !important;
    box-shadow: 0 8px 24px rgba(15,23,42,.055) !important;
    margin-top: 12px !important;
}

.order-table-title {
    font-size: 18px !important;
    font-weight: 950 !important;
    color: #0f172a !important;
    margin-bottom: 10px !important;
}

.order-list-header {
    background:#f8fbff !important;
    border:1px solid #e5eaf3 !important;
    border-radius:10px !important;
    padding:9px 11px !important;
    font-weight:950 !important;
    color:#475569 !important;
    font-size:12px !important;
    text-transform: uppercase;
    letter-spacing: .02em;
}

.order-cell-muted {
    color:#64748b !important;
    font-size:13.5px !important;
    font-weight:800 !important;
    line-height: 1.35 !important;
}

.order-cell-strong {
    color:#0f172a !important;
    font-size:13.5px !important;
    font-weight:900 !important;
    line-height: 1.35 !important;
}

.order-id-button button {
    min-height: 36px !important;
    font-size: 14px !important;
    font-weight: 950 !important;
    color: #1d4ed8 !important;
    background: #ffffff !important;
    border: 1px solid #dbe3ef !important;
    border-radius: 12px !important;
}

.order-id-button button:hover {
    background: #eff6ff !important;
    border-color: #93c5fd !important;
}

.order-delete-holder {
    opacity: .62;
    transition: opacity .14s ease;
}

.order-delete-holder:hover {
    opacity: 1;
}

.order-delete-holder button {
    min-height: 34px !important;
    border-radius: 10px !important;
}

/* zmenšení zbytečných mezer v tabulkovém seznamu */
.order-table-shell + div,
.order-table-shell div[data-testid="stVerticalBlock"] {
    gap: 0.25rem !important;
}

/* v18 objednávky v jasně ukončeném bloku */
.orders-page-shell {
    background: #ffffff !important;
    border: 1px solid #dbe3ef !important;
    border-radius: 24px !important;
    padding: 22px 24px 24px 24px !important;
    box-shadow: 0 14px 42px rgba(15,23,42,.07) !important;
    margin-top: 18px !important;
}

.orders-page-shell .order-table-shell {
    border: 0 !important;
    box-shadow: none !important;
    padding: 0 !important;
    margin-top: 0 !important;
    background: transparent !important;
}

.orders-page-shell .order-table-title {
    background: #f8fbff !important;
    border: 1px solid #e5eaf3 !important;
    border-radius: 14px !important;
    padding: 13px 16px !important;
    margin-bottom: 14px !important;
}

.orders-page-footer {
    border-top: 1px solid #eef2f7 !important;
    margin-top: 16px !important;
    padding-top: 12px !important;
    color: #64748b !important;
    font-size: 12px !important;
}

/* =========================================================
   v19 ULTRA PRO CRM DESIGN – PCmobilSHOP
   ========================================================= */

/* Celkový layout */
html,
body,
.stApp,
[data-testid="stAppViewContainer"],
[data-testid="stHeader"],
[data-testid="stToolbar"],
[data-testid="stDecoration"],
.main,
.main .block-container {
    background:
        radial-gradient(circle at 22% 0%, rgba(37,99,235,.12) 0, transparent 34%),
        radial-gradient(circle at 90% 12%, rgba(14,165,233,.10) 0, transparent 30%),
        linear-gradient(180deg, #f5f8ff 0%, #eef4fb 100%) !important;
    color: #071225 !important;
}

.block-container {
    max-width: 1540px !important;
    padding-top: 1.25rem !important;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background:
        linear-gradient(180deg, #ffffff 0%, #f8fbff 100%) !important;
    border-right: 1px solid rgba(148,163,184,.28) !important;
    box-shadow: 10px 0 35px rgba(15,23,42,.045) !important;
}

[data-testid="stSidebar"] img {
    filter: drop-shadow(0 8px 18px rgba(37,99,235,.12));
}

[data-testid="stSidebar"] [role="radiogroup"] label {
    border-radius: 16px !important;
    padding: 10px 12px !important;
    margin-bottom: 7px !important;
    transition: all .18s ease !important;
    border: 1px solid transparent !important;
    font-weight: 850 !important;
}

[data-testid="stSidebar"] [role="radiogroup"] label:hover {
    background: #eff6ff !important;
    border-color: #dbeafe !important;
    transform: translateX(3px);
}

.sidebar-user {
    border: 1px solid rgba(37,99,235,.18) !important;
    border-radius: 20px !important;
    padding: 16px !important;
    background:
        linear-gradient(135deg, #eff6ff 0%, #ffffff 100%) !important;
    box-shadow: 0 10px 28px rgba(37,99,235,.08) !important;
}

.sidebar-avatar {
    background: linear-gradient(135deg, #2563eb 0%, #06b6d4 100%) !important;
    box-shadow: 0 8px 20px rgba(37,99,235,.26) !important;
}

/* Nadpisy */
h1 {
    font-size: 42px !important;
    letter-spacing: -1.4px !important;
    font-weight: 1000 !important;
    color: #071225 !important;
    margin-bottom: 4px !important;
}

h2, h3 {
    letter-spacing: -.4px !important;
    font-weight: 950 !important;
}

/* Horní refresh tlačítko */
.stButton > button {
    transition: all .16s ease !important;
}

.stButton > button:hover {
    transform: translateY(-1px);
    box-shadow: 0 8px 20px rgba(37,99,235,.12) !important;
}

/* Form inputs */
.stTextInput input,
.stTextArea textarea,
.stNumberInput input,
.stSelectbox div[data-baseweb="select"],
.stMultiSelect div[data-baseweb="select"] {
    border-radius: 14px !important;
    border: 1px solid #dbe3ef !important;
    background: rgba(255,255,255,.88) !important;
    box-shadow: 0 5px 18px rgba(15,23,42,.035) !important;
    min-height: 44px !important;
}

.stTextInput input:focus,
.stTextArea textarea:focus {
    border-color: #60a5fa !important;
    box-shadow: 0 0 0 4px rgba(37,99,235,.10) !important;
}

/* =========================================================
   OBJEDNÁVKY – hlavní profesionální panel
   ========================================================= */

.orders-page-shell {
    position: relative;
    background:
        linear-gradient(180deg, rgba(255,255,255,.96) 0%, rgba(250,253,255,.94) 100%) !important;
    border: 1px solid rgba(148,163,184,.26) !important;
    border-radius: 30px !important;
    padding: 24px 26px 26px 26px !important;
    box-shadow:
        0 22px 70px rgba(15,23,42,.09),
        inset 0 1px 0 rgba(255,255,255,.85) !important;
    margin-top: 18px !important;
    overflow: hidden !important;
}

.orders-page-shell::before {
    content: "";
    position: absolute;
    top: 0;
    left: 0;
    right: 0;
    height: 5px;
    background: linear-gradient(90deg, #2563eb 0%, #06b6d4 38%, #22c55e 68%, #9333ea 100%);
}

.orders-page-shell::after {
    content: "";
    position: absolute;
    top: -120px;
    right: -120px;
    width: 260px;
    height: 260px;
    background: radial-gradient(circle, rgba(37,99,235,.12) 0%, transparent 70%);
    pointer-events: none;
}

.order-table-shell {
    background: transparent !important;
    border: none !important;
    box-shadow: none !important;
    padding: 0 !important;
    margin-top: 18px !important;
}

.order-table-title {
    display: flex !important;
    align-items: center !important;
    gap: 10px !important;
    background:
        linear-gradient(135deg, #0f2a55 0%, #173b7a 100%) !important;
    color: #ffffff !important;
    border: 0 !important;
    border-radius: 18px !important;
    padding: 16px 18px !important;
    margin-bottom: 14px !important;
    font-size: 18px !important;
    font-weight: 1000 !important;
    box-shadow: 0 12px 26px rgba(23,59,122,.20) !important;
}

.order-table-title,
.order-table-title * {
    color: #ffffff !important;
}

/* Hlavička seznamu */
.order-list-header {
    background: #f8fbff !important;
    border: 1px solid #dde8f5 !important;
    border-radius: 12px !important;
    padding: 10px 11px !important;
    font-weight: 1000 !important;
    color: #475569 !important;
    font-size: 11.5px !important;
    text-transform: uppercase !important;
    letter-spacing: .045em !important;
}

/* Řádek objednávky */
.order-row-card {
    position: relative;
    background: #ffffff !important;
    border: 1px solid #e6edf7 !important;
    border-radius: 18px !important;
    padding: 7px 10px !important;
    margin: 8px 0 !important;
    box-shadow: 0 5px 18px rgba(15,23,42,.035) !important;
    transition: all .18s ease !important;
    overflow: hidden !important;
}

.order-row-card::before {
    content: "";
    position: absolute;
    left: 0;
    top: 10px;
    bottom: 10px;
    width: 4px;
    border-radius: 999px;
    background: linear-gradient(180deg, #2563eb 0%, #06b6d4 100%);
    opacity: .75;
}

.order-row-card:hover {
    transform: translateY(-2px);
    border-color: #bfdbfe !important;
    background: linear-gradient(180deg, #ffffff 0%, #f8fbff 100%) !important;
    box-shadow:
        0 14px 34px rgba(37,99,235,.12),
        inset 0 1px 0 rgba(255,255,255,.9) !important;
}

.order-cell-strong {
    color: #0f172a !important;
    font-size: 13.6px !important;
    font-weight: 950 !important;
    line-height: 1.35 !important;
}

.order-cell-muted {
    color: #64748b !important;
    font-size: 13px !important;
    font-weight: 800 !important;
    line-height: 1.35 !important;
}

/* Klikací objednávka */
.order-id-button button {
    min-height: 38px !important;
    font-size: 14px !important;
    font-weight: 1000 !important;
    color: #1d4ed8 !important;
    background:
        linear-gradient(180deg, #ffffff 0%, #eff6ff 100%) !important;
    border: 1px solid #bfdbfe !important;
    border-radius: 14px !important;
    box-shadow: 0 5px 14px rgba(37,99,235,.08) !important;
}

.order-id-button button:hover {
    color: #ffffff !important;
    background: linear-gradient(135deg, #2563eb 0%, #06b6d4 100%) !important;
    border-color: transparent !important;
}

/* Koš */
.order-delete-holder {
    opacity: .42 !important;
    transition: opacity .14s ease, transform .14s ease !important;
}

.order-row-card:hover .order-delete-holder,
.order-delete-holder:hover {
    opacity: 1 !important;
    transform: scale(1.02);
}

.order-delete-holder button {
    min-height: 36px !important;
    border-radius: 13px !important;
    background: #ffffff !important;
}

.order-delete-holder button:hover {
    background: #fef2f2 !important;
    border-color: #fecaca !important;
}

/* Status badge – větší a luxusnější */
.pill,
span[style*="border-radius:999px"] {
    box-shadow: 0 7px 16px rgba(15,23,42,.12);
}

/* Footer seznamu */
.orders-page-footer {
    border-top: 1px solid #edf2f7 !important;
    margin-top: 18px !important;
    padding-top: 14px !important;
    color: #64748b !important;
    font-size: 12px !important;
    font-weight: 750 !important;
}

/* Karty detailu */
.mod-card,
.hero,
.compact-note-box,
.packeta-panel {
    border-radius: 24px !important;
    box-shadow: 0 14px 36px rgba(15,23,42,.07) !important;
    border-color: rgba(148,163,184,.28) !important;
}

.hero {
    background:
        linear-gradient(135deg, #ffffff 0%, #f8fbff 100%) !important;
}

/* Metriky ve statistikách */
div[data-testid="stMetric"] {
    border-radius: 22px !important;
    background:
        linear-gradient(135deg, #0f2a55 0%, #173b7a 58%, #2563eb 100%) !important;
    box-shadow: 0 18px 42px rgba(37,99,235,.18) !important;
}

/* Responsive jemné */
@media (max-width: 1100px) {
    .orders-page-shell {
        padding: 18px !important;
    }
    h1 {
        font-size: 34px !important;
    }
}
</style>
""", unsafe_allow_html=True)


def html_escape(value: object) -> str:
    return escape(str(value or ""))


def card(title: str, icon: str = ""):
    st.markdown(
        f"<div class='mod-card'><div class='mod-card-title'><span class='mod-icon'>{html_escape(icon)}</span>{html_escape(title)}</div>",
        unsafe_allow_html=True,
    )


def end_card():
    st.markdown("</div>", unsafe_allow_html=True)


def kv(rows: list[tuple[str, object]]) -> str:
    html = ["<div class='mod-kv'>"]
    for label, value in rows:
        html.append(f"<div class='mod-k'>{html_escape(label)}</div><div class='mod-v'>{html_escape(value)}</div>")
    html.append("</div>")
    return "".join(html)

def render_card_html(title: str, icon: str, inner_html: str) -> None:
    st.markdown(
        f"""
        <div class='mod-card'>
            <div class='mod-card-title'>
                <span class='mod-icon'>{html_escape(icon)}</span>{html_escape(title)}
            </div>
            {inner_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def status_pill(status: str) -> str:
    return f"<span class='pill' style='background:{status_color(status)}'>{html_escape(status)}</span>"


def progress_html(status: str, created: str, changed: str) -> str:
    steps = [
        ("Nová objednávka", "✓", created),
        ("Telefon přijat", "✓", ""),
        ("V opravě", "🔧", changed if status == "V opravě" else ""),
        ("Odesláno zpět", "🚚", changed if status == "Odesláno zpět" else ""),
        ("Dokončeno", "✓", changed if status == "Dokončeno" else ""),
    ]
    order = ["Nová objednávka", "Čekáme na doručení telefonu", "Telefon přijat", "Čekáme na díl", "V opravě", "Hotovo", "Odesláno zpět", "Dokončeno"]
    # map detailed statuses to progress index
    if status in ["Nová objednávka", "Čekáme na doručení telefonu"]:
        idx = 0
    elif status == "Telefon přijat":
        idx = 1
    elif status in ["Čekáme na díl", "V opravě", "Hotovo"]:
        idx = 2
    elif status == "Odesláno zpět":
        idx = 3
    elif status == "Dokončeno":
        idx = 3
    else:
        idx = 0

    out = ["<div class='progress'>"]
    for i, (title, icon, date) in enumerate(steps):
        active = i <= idx
        current = i == idx
        bg = "#2563eb" if current else ("#bfdbfe" if active else "#e5e7eb")
        color = "#ffffff" if current else ("#2563eb" if active else "#94a3b8")
        out.append(
            f"<div class='step'><div class='dot' style='background:{bg};color:{color};'>{icon}</div>"
            f"<div class='step-title'>{html_escape(title)}</div><div class='step-date'>{html_escape(date)}</div></div>"
        )
    out.append("</div>")
    return "".join(out)



def _get_top_action_context(order_id: str, selected_order: Optional[pd.Series] = None) -> tuple[str, str, bool]:
    """Vrátí hodnoty pro rychlá stavová tlačítka.

    Tracking se už neupravuje ručně přes text input. Bere se výhradně z Google Sheets
    ze sloupce TrackingCislo, aby se omylem nesmazal při změně stavu.
    """
    tracking = ""
    if selected_order is not None:
        tracking = str(selected_order.get("TrackingCislo", "") or "").strip()

    internal_note = st.session_state.get(f"internal_note_{order_id}", "")
    if not str(internal_note or "").strip() and selected_order is not None:
        internal_note = str(selected_order.get("InterniPoznamka", "") or "")

    send_email = st.session_state.get(f"send_email_top_{order_id}", True)
    return str(tracking or ""), str(internal_note or ""), bool(send_email)

def hero_order(selected_order: pd.Series):
    order_id = safe_text(selected_order.get("ID", ""))
    status = normalize_status_label(safe_text(selected_order.get("Stav", "")))
    created = f"{safe_text(selected_order.get('Datum', ''))} {safe_text(selected_order.get('Cas', ''))}".strip()
    changed = safe_text(selected_order.get("DatumPosledniZmenyStavu", ""))
    last_email = safe_text(selected_order.get("PosledniEmailTyp", ""))

    html = f"""
    <div class='hero'>
      <div class='hero-grid'>
        <div>
          <div class='hero-order'>{html_escape(order_id)}</div>
          {status_pill(status)}
          <div class='hero-meta'>
            📅 Přijato: <strong>{html_escape(created)}</strong><br>
            🕒 Poslední změna: <strong>{html_escape(changed)}</strong><br>
            ✉️ Poslední email: <strong>{html_escape(last_email)}</strong>
          </div>
        </div>

      </div>
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)

    top_quick_statuses = [
        ("Telefon přijat", "📱"),
        ("Čekáme na díl", "🔴"),
        ("V opravě", "🟣"),
        ("Odesláno zpět", "📦"),
    ]

    email_toggle_col, email_help_col = st.columns([1.1, 2.9])
    with email_toggle_col:
        st.checkbox(
            "Odeslat e-mail zákazníkovi",
            value=st.session_state.get(f"send_email_top_{order_id}", True),
            key=f"send_email_top_{order_id}",
        )
    with email_help_col:
        st.caption("Platí pro rychlá tlačítka změny stavu níže.")

    top_tracking_number, top_internal_note, top_send_email = _get_top_action_context(order_id, selected_order)

    top_cols = st.columns(len(top_quick_statuses))
    for idx, (quick_status, icon) in enumerate(top_quick_statuses):
        with top_cols[idx]:
            if st.button(
                f"{icon} {quick_status}",
                use_container_width=True,
                key=f"hero_quick_{quick_status}_{order_id}",
            ):
                ok, message = save_order_changes_modern(
                    selected_order=selected_order,
                    selected_order_id=order_id,
                    new_status=quick_status,
                    tracking_number=top_tracking_number,
                    internal_note=top_internal_note,
                    send_email=top_send_email,
                )
                if ok:
                    st.success(message)
                    time.sleep(0.8)
                    st.rerun()
                else:
                    st.error(message)


def save_order_changes_modern(
    selected_order: pd.Series,
    selected_order_id: str,
    new_status: str,
    tracking_number: str,
    internal_note: str,
    send_email: bool,
) -> tuple[bool, str]:
    ws = get_worksheet(ORDERS_SHEET_NAME)
    row_number = find_order_row_number_from_df(orders_df, selected_order_id)
    if not row_number:
        return False, "Objednávku se nepodařilo najít v listu."

    previous_status = str(selected_order.get("Stav", "") or "")

    # Důležité: rychlá tlačítka v horní části detailu se vykreslují dříve než
    # textové pole pro tracking. Pokud session_state ještě tracking neobsahuje,
    # nesmíme přepsat existující TrackingCislo prázdnou hodnotou.
    existing_tracking = str(selected_order.get("TrackingCislo", "") or "").strip()
    final_tracking = str(tracking_number or "").strip() or existing_tracking

    existing_internal_note = str(selected_order.get("InterniPoznamka", "") or "")
    final_internal_note = str(internal_note) if internal_note is not None else existing_internal_note

    save_cell_by_header(ws, row_number, "Stav", str(new_status))
    save_cell_by_header(ws, row_number, "InterniPoznamka", final_internal_note)
    save_cell_by_header(ws, row_number, "TrackingCislo", final_tracking)
    save_cell_by_header(ws, row_number, "DatumPosledniZmenyStavu", now_cz_string())

    if send_email and new_status != previous_status:
        if new_status in STATUS_TO_WORKER_STATUS:
            ok, message = send_status_email_via_worker(selected_order, new_status, final_tracking)
            if ok:
                save_cell_by_header(ws, row_number, "PosledniEmailTyp", f"Stav: {new_status}")
                save_cell_by_header(ws, row_number, "EmailZakaznikOdeslan", "ANO")
            else:
                save_cell_by_header(ws, row_number, "EmailZakaznikOdeslan", "NE")
                clear_data_cache()
                return False, message
        else:
            save_cell_by_header(ws, row_number, "EmailZakaznikOdeslan", "NE")
            clear_data_cache()
            return True, f"Objednávka byla uložena. Pro stav '{new_status}' není nastaven automatický e-mail."

    clear_data_cache()
    return True, "Objednávka byla aktualizována."


# =============================
# SIDEBAR
# =============================
with st.sidebar:
    st.image(LOGO_URL, width=190)
    st.markdown("###")
    page_label = st.radio(
        "Navigace",
        ["📋 Objednávky", "📦 Sklad", "📊 Statistiky", "🧾 Faktury", "🚚 Zásilkovna", "⚙️ Nastavení"],
        label_visibility="collapsed",
        index=0,
    )
    page = page_label.split(" ", 1)[1]
    st.markdown(
        """
        <div class='sidebar-user'>
          <div><span class='sidebar-avatar'>A</span><strong>Admin - Lukáš</strong></div>
          <div style='font-size:12px;color:#64748b;margin-top:6px;'>servis@pcmobilshop.cz</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =============================
# OBJEDNÁVKY – SEZNAM
# =============================
if page == "Objednávky":
    if st.session_state.selected_order_id is None:
        top1, top2 = st.columns([2.8, 1])
        with top1:
            st.title("Objednávky")
            st.caption("Servisní zakázky, stav opravy, logistika a fakturace na jednom místě.")
        with top2:
            if st.button("🔄 Obnovit data", use_container_width=True):
                clear_data_cache()
                st.rerun()

        st.markdown("<div class='orders-page-shell'>", unsafe_allow_html=True)

        toolbar1, toolbar2, toolbar3 = st.columns([1.15, 1.4, 0.8])
        with toolbar1:
            status_filter = st.multiselect("Stav", ORDER_STATUSES, default=[], placeholder="Vyberte stav")
        with toolbar2:
            search_text = st.text_input("Hledat", placeholder="Jméno, firma, telefon, email, ID, ServiceID")
        with toolbar3:
            sort_mode = st.selectbox("Řazení", ["Nejnovější nahoře", "Nejstarší nahoře"])

        filtered_df = orders_df.copy()

        if status_filter:
            filter_values = list(status_filter)
            if "Objednávka přijata" in filter_values:
                filter_values.append("Nová objednávka")
            if "Odesláno zpět" in filter_values:
                filter_values.append("Dokončeno")
            filtered_df = filtered_df[filtered_df["Stav"].isin(filter_values)]

        if search_text.strip():
            needle = search_text.strip().lower()
            mask = (
                filtered_df["ID"].astype(str).str.lower().str.contains(needle, na=False)
                | filtered_df["ServiceID"].astype(str).str.lower().str.contains(needle, na=False)
                | filtered_df["Jmeno"].astype(str).str.lower().str.contains(needle, na=False)
                | filtered_df["Firma"].astype(str).str.lower().str.contains(needle, na=False)
                | filtered_df["ICO"].astype(str).str.lower().str.contains(needle, na=False)
                | filtered_df["DIC"].astype(str).str.lower().str.contains(needle, na=False)
                | filtered_df["Telefon"].astype(str).str.lower().str.contains(needle, na=False)
                | filtered_df["Email"].astype(str).str.lower().str.contains(needle, na=False)
                | filtered_df["Produkt"].astype(str).str.lower().str.contains(needle, na=False)
                | filtered_df["Oprava"].astype(str).str.lower().str.contains(needle, na=False)
            )
            filtered_df = filtered_df[mask]

        if not filtered_df.empty:
            filtered_df = filtered_df.copy()
            filtered_df["_sort_dt"] = pd.to_datetime(
                filtered_df["Datum"].astype(str) + " " + filtered_df["Cas"].astype(str),
                errors="coerce"
            )
            filtered_df = filtered_df.sort_values("_sort_dt", ascending=(sort_mode == "Nejstarší nahoře"))

            display_df = filtered_df.copy()
            if "Stav" in display_df.columns:
                display_df["Stav"] = display_df["Stav"].apply(normalize_status_label)

            st.markdown(f"<div class='order-table-shell'><div class='order-table-title'>📋 Seznam objednávek <span style='margin-left:auto;background:rgba(255,255,255,.16);padding:6px 10px;border-radius:999px;font-size:12px;'>Zobrazeno: {len(display_df)}</span></div>", unsafe_allow_html=True)

            # Hlavička tabulky
            h_cols = st.columns([1.0, 1.25, 1.35, 1.15, 1.25, 1.2, 0.9, 1.15, 0.55])
            headers = ["Objednávka", "Stav", "Zákazník", "Telefon", "Zařízení", "Oprava", "Cena", "Tracking", ""]
            for col, header in zip(h_cols, headers):
                with col:
                    st.markdown(f"<div class='order-list-header'>{header}</div>", unsafe_allow_html=True)

            # Řádky objednávek
            for _, row in display_df.iterrows():
                order_id = str(row.get("ID", "") or "").strip()
                status = str(row.get("Stav", "") or "").strip()
                tracking = str(row.get("TrackingCislo", "") or "").strip()

                st.markdown("<div class='order-row-card'>", unsafe_allow_html=True)
                r_cols = st.columns([1.0, 1.25, 1.35, 1.15, 1.25, 1.2, 0.9, 1.15, 0.55])

                with r_cols[0]:
                    if st.button(order_id, key=f"open_order_{order_id}", use_container_width=True):
                        st.session_state.selected_order_id = order_id
                        st.session_state.delete_confirm_order_id = None
                        st.rerun()

                with r_cols[1]:
                    st.markdown(status_badge_html(status), unsafe_allow_html=True)

                with r_cols[2]:
                    st.markdown(f"<div class='order-cell-strong'>{html_escape(row.get('Jmeno', ''))}</div>", unsafe_allow_html=True)
                    firma = str(row.get("Firma", "") or "").strip()
                    if firma:
                        st.markdown(f"<div class='order-cell-muted'>{html_escape(firma)}</div>", unsafe_allow_html=True)

                with r_cols[3]:
                    st.markdown(f"<div class='order-cell-strong'>{html_escape(row.get('Telefon', ''))}</div>", unsafe_allow_html=True)

                with r_cols[4]:
                    st.markdown(f"<div class='order-cell-strong'>{html_escape(row.get('Produkt', ''))}</div>", unsafe_allow_html=True)

                with r_cols[5]:
                    st.markdown(f"<div class='order-cell-strong'>{html_escape(row.get('Oprava', ''))}</div>", unsafe_allow_html=True)

                with r_cols[6]:
                    st.markdown(f"<div class='order-cell-strong'>{html_escape(row.get('Cena', ''))}</div>", unsafe_allow_html=True)

                with r_cols[7]:
                    tracking_html = html_escape(tracking) if tracking else "&nbsp;"
                    st.markdown(f"<div class='order-cell-muted'>{tracking_html}</div>", unsafe_allow_html=True)

                with r_cols[8]:
                    if st.session_state.delete_confirm_order_id == order_id:
                        if st.button("✅", key=f"delete_confirm_{order_id}", help="Potvrdit smazání", use_container_width=True):
                            ok, msg = delete_order_by_id(orders_df, order_id)
                            if ok:
                                st.session_state.delete_confirm_order_id = None
                                st.success(msg)
                                time.sleep(0.8)
                                st.rerun()
                            else:
                                st.error(msg)

                        if st.button("↩️", key=f"delete_cancel_{order_id}", help="Zrušit mazání", use_container_width=True):
                            st.session_state.delete_confirm_order_id = None
                            st.rerun()
                    else:
                        if st.button("🗑️", key=f"delete_order_{order_id}", help="Smazat objednávku", use_container_width=True):
                            st.session_state.delete_confirm_order_id = order_id
                            st.warning(f"Potvrď smazání objednávky {order_id}.")
                            st.rerun()

                st.markdown("</div>", unsafe_allow_html=True)

            st.markdown("</div>", unsafe_allow_html=True)
            st.markdown("<div class='orders-page-footer'>Kliknutím na číslo objednávky otevřeš detail. Koš vpravo objednávku smaže až po potvrzení.</div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)
        else:
            st.info("Žádné objednávky neodpovídají filtru.")
            st.markdown("</div>", unsafe_allow_html=True)

    # =============================
    # OBJEDNÁVKY – DETAIL
    # =============================
    else:
        selected_order_id = str(st.session_state.selected_order_id)
        selected_rows = orders_df[orders_df["ID"].astype(str) == selected_order_id]

        if selected_rows.empty:
            st.warning("Objednávka už nebyla nalezena.")
            if st.button("← Zpět na seznam"):
                st.session_state.selected_order_id = None
                st.rerun()
        else:
            selected_order = selected_rows.iloc[0]

            header_left, header_title, header_actions = st.columns([1.2, 3.4, 2])
            with header_left:
                if st.button("← Zpět na seznam", use_container_width=True):
                    st.session_state.selected_order_id = None
                    st.rerun()
            with header_title:
                st.markdown(f"## Detail objednávky {selected_order_id}")
            with header_actions:
                if st.button("🔄 Obnovit detail", use_container_width=True, key=f"refresh_detail_{selected_order_id}"):
                    clear_data_cache()
                    st.session_state.selected_order_id = selected_order_id
                    st.rerun()

            hero_order(selected_order)

            # info cards
            cc1, cc2, cc3, cc4 = st.columns([1.1, 1.05, 1, 1.25])

            with cc1:
                render_card_html("Zákazník", "👤", kv([
                    ("Jméno", safe_text(selected_order["Jmeno"])),
                    ("Telefon", safe_text(selected_order["Telefon"])),
                    ("Email", safe_text(selected_order["Email"])),
                ]))

            with cc2:
                render_card_html("Zakázka", "📋", kv([
                    ("ServiceID", safe_text(selected_order["ServiceID"])),
                    ("Produkt", safe_text(selected_order["Produkt"])),
                    ("Oprava", safe_text(selected_order["Oprava"])),
                    ("Cena", safe_text(selected_order["Cena"])),
                    ("Sklad", safe_text(selected_order["SkladPriObjednani"])),
                ]))

            with cc3:
                render_card_html("Firemní údaje", "🏢", kv([
                    ("Firma", safe_text(selected_order.get("Firma", ""))),
                    ("IČO", safe_text(selected_order.get("ICO", ""))),
                    ("DIČ", safe_text(selected_order.get("DIC", ""))),
                ]))

            with cc4:
                tracking_value = str(selected_order["TrackingCislo"] or "").strip()
                render_card_html("Logistika", "🚚", kv([
                    ("Výdejní místo", safe_text(selected_order["PacketaPointValue"])),
                    ("Tracking číslo", tracking_value if tracking_value else "-"),
                ]))

                tracking_url = build_packeta_tracking_url(tracking_value)
                if tracking_url:
                    st.link_button("🔎 Sledovat zásilku", tracking_url, use_container_width=True)

                label_url = str(selected_order.get("PacketaLabelUrl", "") or "").strip()
                packet_id_value = str(selected_order.get("PacketaPacketId", "") or "").strip()
                if label_url:
                    if st.button("🖨️ Vytisknout štítek", use_container_width=True, key=f"load_label_{selected_order_id}"):
                        ok_label, pdf_bytes, label_msg = download_packeta_label_pdf(label_url)
                        if ok_label and pdf_bytes:
                            st.session_state[f"packeta_label_pdf_{selected_order_id}"] = pdf_bytes
                            st.success("Štítek byl načten. Klikněte níže na otevření štítku a v novém okně dejte tisk.")
                        else:
                            st.error(label_msg)

                    pdf_bytes = st.session_state.get(f"packeta_label_pdf_{selected_order_id}")
                    if pdf_bytes:
                        label_b64 = base64.b64encode(pdf_bytes).decode("ascii")
                        open_label_html = f"""
                            <a href="data:application/pdf;base64,{label_b64}" target="_blank" style="
                                display:block;
                                text-align:center;
                                padding:11px 14px;
                                border:1px solid #dbe3ef;
                                border-radius:12px;
                                background:#ffffff;
                                color:#0f172a !important;
                                font-weight:800;
                                text-decoration:none;
                                min-height:42px;
                                line-height:20px;
                            ">🖨️ Otevřít štítek v novém okně</a>
                        """
                        st.markdown(open_label_html, unsafe_allow_html=True)

                        st.download_button(
                            "⬇️ Stáhnout štítek PDF",
                            data=pdf_bytes,
                            file_name=f"packeta-label-{packet_id_value or selected_order_id}.pdf",
                            mime="application/pdf",
                            use_container_width=True,
                            key=f"download_label_{selected_order_id}",
                        )


            # Poznámky – malé pod hlavními kartami
            notes_a, notes_b = st.columns([1, 1])
            with notes_a:
                st.markdown(
                    f"""
                    <div class='compact-note-box'>
                        <div class='compact-note-title'>💬 Poznámka od zákazníka</div>
                        <div class='compact-note-content'>{html_escape(selected_order['Poznamka'] or 'Bez poznámky')}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            with notes_b:
                st.markdown(
                    "<div class='compact-note-box compact-internal'><div class='compact-note-title'>📝 Interní poznámka</div>",
                    unsafe_allow_html=True,
                )
                internal_note = st.text_area(
                    "Interní poznámka",
                    value=str(selected_order["InterniPoznamka"] or ""),
                    height=54,
                    label_visibility="collapsed",
                    placeholder="Zadejte interní poznámku…",
                    key=f"internal_note_{selected_order_id}",
                )
                st.markdown("<div style='font-size:11px;color:#64748b!important;margin-top:4px;'>Tato poznámka není viditelná pro zákazníka.</div></div>", unsafe_allow_html=True)



            # Spodní akce
            # Tracking číslo se už neupravuje ručně. Zobrazuje se pouze v kartě Logistika
            # a při změně stavu se bere přímo ze sloupce TrackingCislo v Google Sheets.
            action_b1, action_b2 = st.columns(2)
            with action_b1:
                if st.button("🧾 Faktura", use_container_width=True, key=f"bottom_invoice_{selected_order_id}"):
                    st.session_state[f"show_invoice_{selected_order_id}"] = not st.session_state.get(f"show_invoice_{selected_order_id}", False)
            with action_b2:
                if st.button("📦 Přidat do Zásilkovny", use_container_width=True, key=f"bottom_packeta_{selected_order_id}"):
                    st.session_state[f"show_packeta_{selected_order_id}"] = not st.session_state.get(f"show_packeta_{selected_order_id}", False)

            if st.session_state.get(f"show_invoice_{selected_order_id}", False):
                st.markdown("<div class='packeta-panel'><strong>🧾 Faktura</strong>", unsafe_allow_html=True)
                invoice_html = build_invoice_html(selected_order)
                invoice_filename = build_invoice_filename(selected_order)
                invoice_number = build_invoice_number(selected_order.get("ID", ""))

                inv_col1, inv_col2, inv_col3 = st.columns([1, 1, 1])
                with inv_col1:
                    st.info(f"Číslo faktury: {invoice_number}")
                with inv_col2:
                    st.download_button(
                        "⬇️ Stáhnout fakturu",
                        data=invoice_html.encode("utf-8"),
                        file_name=invoice_filename,
                        mime="text/html",
                        key=f"download_invoice_bottom_{selected_order_id}",
                        use_container_width=True,
                    )
                with inv_col3:
                    st.markdown(invoice_open_button_html(selected_order), unsafe_allow_html=True)

                st.components.v1.html(invoice_html, height=760, scrolling=True)
                st.markdown("</div>", unsafe_allow_html=True)

            if st.session_state.get(f"show_packeta_{selected_order_id}", False):
                st.markdown("<div class='packeta-panel'><strong>📦 Přidat do Zásilkovny</strong>", unsafe_allow_html=True)
                packet_name = st.text_input(
                    "Jméno příjemce",
                    value=str(selected_order["Jmeno"] or ""),
                    key=f"packeta_left_name_{selected_order_id}",
                )

                packet_phone = st.text_input(
                    "Telefon",
                    value=str(selected_order["Telefon"] or ""),
                    key=f"packeta_left_phone_{selected_order_id}",
                )

                packet_email = st.text_input(
                    "Email",
                    value=str(selected_order["Email"] or ""),
                    key=f"packeta_left_email_{selected_order_id}",
                )

                packet_point_name = st.text_input(
                    "Výdejní místo",
                    value=str(selected_order["PacketaPointValue"] or selected_order["PacketaPointName"] or ""),
                    key=f"packeta_left_point_name_{selected_order_id}",
                    disabled=True,
                )

                packet_point_id = st.text_input(
                    "PacketaPointId",
                    value=str(selected_order["PacketaPointId"] or ""),
                    key=f"packeta_left_point_id_{selected_order_id}",
                    disabled=True,
                )

                packet_col1, packet_col2 = st.columns(2)
                with packet_col1:
                    packet_cod = st.text_input(
                        "Dobírka",
                        value=str(selected_order["Cena"] or ""),
                        key=f"packeta_left_cod_{selected_order_id}",
                    )
                    packet_weight = st.number_input(
                        "Hmotnost (kg)",
                        min_value=0.1,
                        step=0.1,
                        value=0.5,
                        key=f"packeta_left_weight_{selected_order_id}",
                    )
                with packet_col2:
                    packet_insurance = st.text_input(
                        "Pojištění",
                        value=str(selected_order.get("PacketaInsuranceValue", "") or ""),
                        placeholder="Vyplň ručně",
                        key=f"packeta_left_insurance_{selected_order_id}",
                    )
                    st.text_input(
                        "Model zařízení",
                        value=str(selected_order["Produkt"] or ""),
                        key=f"packeta_left_product_{selected_order_id}",
                        disabled=True,
                    )

                st.caption("Vše je předvyplněné z objednávky. Ručně doplň jen pojištění a potom vytvoř zásilku přes Worker.")

                if st.button("Vytvořit zásilku v Zásilkovně", key=f"packeta_left_create_{selected_order_id}"):
                    if not str(selected_order["PacketaPointId"] or "").strip():
                        st.error("Objednávka nemá vyplněné PacketaPointId.")
                    elif not str(packet_insurance).strip():
                        st.warning("Vyplň prosím cenu pojištění.")
                    else:
                        try:
                            order_reference = str(selected_order_id or selected_order.get("ID", "") or "").strip()
                            payload = {
                                "name": packet_name,
                                "phone": packet_phone,
                                "email": packet_email,
                                "packetaPointId": packet_point_id,
                                "cod": packet_cod,
                                "insurance": packet_insurance,
                                "weight": packet_weight,
                                # Důležité: do Packety posíláme skutečné číslo zakázky, ne timestamp.
                                "packetNumber": order_reference,
                                "reference": order_reference,
                            }

                            res = requests.post(
                                WORKER_PACKETA_ENDPOINT,
                                json=payload,
                                headers=get_worker_auth_headers(),
                                timeout=30,
                            )
                            data = res.json()

                            if res.ok and data.get("success"):
                                tracking = str(data.get("trackingNumber") or "")
                                packet_id = str(data.get("packetId") or "")
                                label_url = str(data.get("labelUrl") or "")
                                normalized_insurance = str(data.get("insurance") or packet_insurance)
                                normalized_cod = str(data.get("cod") or packet_cod)

                                ws = get_worksheet(ORDERS_SHEET_NAME)
                                row_number = find_order_row_number_from_df(orders_df, selected_order_id)

                                if row_number:
                                    save_cell_by_header(ws, row_number, "TrackingCislo", tracking)
                                    save_cell_by_header(ws, row_number, "PacketaPacketId", packet_id)
                                    save_cell_by_header(ws, row_number, "PacketaLabelUrl", label_url)
                                    save_cell_by_header(ws, row_number, "PacketaCreated", "ANO")
                                    save_cell_by_header(ws, row_number, "PacketaInsuranceValue", normalized_insurance)
                                    save_cell_by_header(ws, row_number, "PacketaStatus", "Vytvořeno")
                                    save_cell_by_header(ws, row_number, "Cena", str(selected_order["Cena"] or normalized_cod))
                                    st.session_state[f"packeta_label_pdf_{selected_order_id}"] = None

                                time.sleep(1.2)
                                clear_data_cache()
                                st.session_state.selected_order_id = selected_order_id

                                # Udrž detail otevřený a vynuceně načti čerstvá data z Google Sheets
                                st.session_state.selected_order_id = selected_order_id
                                st.session_state[f"show_packeta_{selected_order_id}"] = False
                                clear_data_cache()

                                st.success(f"Zásilka vytvořena. Tracking: {tracking}")
                                time.sleep(1.0)
                                st.rerun()
                            else:
                                st.error(f"Chyba Packeta: {data}")

                        except Exception as e:
                            st.error(f"Chyba při volání Workeru: {e}")

                existing_packet_id = str(selected_order.get("PacketaPacketId", "") or "").strip()
                existing_tracking = str(selected_order.get("TrackingCislo", "") or "").strip()
                if existing_packet_id or existing_tracking:
                    st.divider()
                    st.warning(
                        "Smazání níže zatím odstraní zásilkové údaje jen z dashboardu / Google Sheets. "
                        "Pro skutečné zrušení zásilky přímo v Packetě doplníme ještě Worker endpoint /cancel-packeta."
                    )
                    if st.button("🗑️ Smazat zásilkové údaje z dashboardu", key=f"packeta_clear_{selected_order_id}"):
                        ok_clear, msg_clear = clear_packeta_data_for_order(selected_order_id)
                        if ok_clear:
                            st.success(msg_clear)
                            time.sleep(0.8)
                            st.rerun()
                        else:
                            st.error(msg_clear)

                st.markdown("</div>", unsafe_allow_html=True)

            st.markdown(
                "<div style='text-align:center;color:#94a3b8;font-size:12px;margin-top:26px;'>PCmobilSHOP – interní správa objednávek a skladu<br>Napojení: Google Sheets • Apps Script • Cloudflare Worker • Resend</div>",
                unsafe_allow_html=True,
            )


# =============================
# SKLAD
# =============================
elif page == "Sklad":
    st.title("Sklad")
    st.caption("Správa servisních dílů a dostupnosti na webu")

    stock_left, stock_right = st.columns([1.35, 1])

    with stock_left:
        card("Přehled skladu", "📦")
        stock_display = stock_df.copy()
        if not stock_display.empty:
            stock_display["PocetSkladem"] = stock_display["PocetSkladem"].apply(to_int_safe)
        st.dataframe(stock_display, use_container_width=True, hide_index=True, height=560)
        end_card()

    with stock_right:
        card("Upravit položku skladu", "✏️")

        if stock_df.empty:
            st.info("Ve skladu zatím nejsou žádné položky.")
        else:
            stock_options = [
                f"{row['ServiceID']} | {row['Produkt']} | {row['Oprava']}"
                for _, row in stock_df.iterrows()
            ]
            selected_stock_label = st.selectbox("Vyber položku", options=stock_options)
            selected_stock_row = stock_df.iloc[stock_options.index(selected_stock_label)]

            selected_service_id = str(selected_stock_row["ServiceID"])
            current_count = to_int_safe(selected_stock_row["PocetSkladem"])

            st.markdown(f"**ServiceID:** {selected_stock_row['ServiceID']}")
            st.markdown(f"**Produkt:** {selected_stock_row['Produkt']}")
            st.markdown(f"**Oprava:** {selected_stock_row['Oprava']}")
            st.markdown(f"**Díl:** {selected_stock_row['Dil']}")

            new_stock_count = st.number_input(
                "Počet skladem",
                min_value=0,
                step=1,
                value=current_count,
                key="edit_stock_count",
            )

            active_value = st.selectbox(
                "Aktivní",
                options=["ANO", "NE"],
                index=0 if str(selected_stock_row["Aktivni"]).strip() != "NE" else 1,
                key="edit_active",
            )

            contact_text = st.text_area(
                "Text při nedostupnosti",
                value=str(selected_stock_row["KontaktKdyzNeni"] or "Pro ověření dostupnosti nás kontaktujte."),
                height=100,
                key="edit_contact_text",
            )

            if st.button("Uložit změny skladu", key="save_stock_changes", type="primary", use_container_width=True):
                try:
                    ws = get_worksheet(STOCK_SHEET_NAME)
                    row_number = find_stock_row_number_from_df(stock_df, selected_service_id)

                    if not row_number:
                        st.error("Položku skladu se nepodařilo najít.")
                    else:
                        availability, status_text = normalize_stock_row(str(new_stock_count))

                        save_cell_by_header(ws, row_number, "PocetSkladem", str(new_stock_count))
                        save_cell_by_header(ws, row_number, "DostupnostWeb", availability)
                        save_cell_by_header(ws, row_number, "StatusText", status_text)
                        save_cell_by_header(ws, row_number, "KontaktKdyzNeni", contact_text)
                        save_cell_by_header(ws, row_number, "Aktivni", active_value)

                        clear_data_cache()
                        st.success("Sklad byl aktualizován.")
                        st.rerun()

                except Exception as exc:
                    st.error(f"Nepodařilo se uložit sklad: {exc}")

        end_card()

        card("Přidat novou servisní položku", "➕")
        with st.form("add_stock_item_form"):
            new_service_id = st.text_input("ServiceID", placeholder="např. 2")
            new_product = st.text_input("Produkt", placeholder="např. Samsung Galaxy S25")
            new_repair = st.text_input("Oprava", placeholder="např. Výměna baterie")
            new_part = st.text_input("Díl", placeholder="např. Baterie Samsung S25")
            new_stock_count_form = st.number_input("Počet skladem", min_value=0, step=1, value=0)
            new_contact_text = st.text_area(
                "Text při nedostupnosti",
                value="Pro ověření dostupnosti nás kontaktujte."
            )
            new_active = st.selectbox("Aktivní", options=["ANO", "NE"], index=0)

            submitted_add_stock = st.form_submit_button("Přidat položku")

            if submitted_add_stock:
                try:
                    service_id_clean = str(new_service_id).strip()
                    product_clean = str(new_product).strip()
                    repair_clean = str(new_repair).strip()
                    part_clean = str(new_part).strip()

                    if not service_id_clean:
                        st.error("Vyplň ServiceID.")
                    elif not product_clean:
                        st.error("Vyplň Produkt.")
                    elif not repair_clean:
                        st.error("Vyplň Oprava.")
                    elif service_id_exists_in_stock(stock_df, service_id_clean):
                        st.error("Tento ServiceID už ve skladu existuje.")
                    else:
                        ws = get_worksheet(STOCK_SHEET_NAME)
                        append_stock_row(
                            worksheet=ws,
                            service_id=service_id_clean,
                            produkt=product_clean,
                            oprava=repair_clean,
                            dil=part_clean,
                            pocet_skladem=int(new_stock_count_form),
                            kontakt_kdyz_neni=str(new_contact_text).strip(),
                            aktivni=str(new_active).strip(),
                        )

                        clear_data_cache()
                        st.success("Nová servisní položka byla přidána.")
                        st.rerun()

                except Exception as exc:
                    st.error(f"Nepodařilo se přidat skladovou položku: {exc}")
        end_card()


# =============================
# STATISTIKY
# =============================
elif page == "Statistiky":
    st.title("Statistiky")
    st.caption("Přehled výkonu servisních objednávek")

    stats_df = orders_df.copy()
    if not stats_df.empty:
        stats_df["Cena_num"] = stats_df["Cena"].apply(parse_price)

    today = datetime.now().strftime("%Y-%m-%d")
    today_df = stats_df[stats_df["Datum"].astype(str) == today] if not stats_df.empty else pd.DataFrame()

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.metric("Objednávky dnes", len(today_df))
    with k2:
        st.metric("Obrat dnes", f"{int(today_df['Cena_num'].sum()) if not today_df.empty else 0:,} Kč".replace(",", " "))
    with k3:
        active_count = len(stats_df[~stats_df["Stav"].astype(str).isin(["Dokončeno"])]) if not stats_df.empty else 0
        st.metric("Aktivní zakázky", active_count)
    with k4:
        repair_count = len(stats_df[stats_df["Stav"].astype(str) == "V opravě"]) if not stats_df.empty else 0
        st.metric("V opravě", repair_count)

    st.markdown("### Stav zakázek")

    if stats_df.empty:
        st.info("Zatím nejsou dostupná žádná data.")
    else:
        status_counts = (
            stats_df["Stav"]
            .astype(str)
            .apply(normalize_status_label)
            .replace("", "Bez stavu")
            .value_counts()
            .reset_index()
        )
        status_counts.columns = ["Stav", "Počet"]

        c1, c2 = st.columns([1, 1])
        with c1:
            st.dataframe(status_counts, use_container_width=True, hide_index=True)
        with c2:
            st.bar_chart(status_counts.set_index("Stav"))

        st.markdown("### Obrat podle zařízení")
        product_revenue = (
            stats_df.groupby("Produkt", dropna=False)["Cena_num"]
            .sum()
            .sort_values(ascending=False)
            .reset_index()
        )
        product_revenue["Obrat"] = product_revenue["Cena_num"].apply(lambda x: f"{int(x):,} Kč".replace(",", " "))
        st.dataframe(product_revenue[["Produkt", "Obrat"]], use_container_width=True, hide_index=True)

# =============================
# OSTATNÍ SEKCE
# =============================
else:
    st.title(page)
    st.info("Tato sekce je připravená jako navigační položka. Funkce můžeme doplnit v dalším kroku.")
