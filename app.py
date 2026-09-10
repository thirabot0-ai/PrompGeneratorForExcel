import json
import re
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import streamlit as st


ROOT = Path(__file__).parent
PROMPT_PATH = ROOT / "output" / "gemini_prompt.json"


def payload_for(chats: list[dict], order_date: str) -> dict:
    return {"date": order_date, "messages": chats}


def build_prompt(payload: dict, has_template: bool) -> str:
    template_rule = (
        "Use the attached Excel workbook as the single source of truth. Preserve its sheet names, "
        "layout, merged cells, formulas, formatting, widths, number formats, and column order exactly."
        if has_template
        else "No Excel template was attached. If you generate a workbook, use a simple order table and say that the original template was not provided."
    )
    return f"""Convert the WhatsApp order messages below into the attached Excel workbook.

{template_rule}

Use every message in the batch. Do not invent missing values. Leave unknown
template cells blank. Keep orders separated by date when the template uses date
sections. Interpret `DT` as delivery time/details and `A.n.` or `a.n.` as atas
nama (the recipient/order name).

For this specific workbook, write the data into the existing table as follows:
- Date section headers are in column F and look like `Rabu,1 Juli 2026`.
- The table headers are on the row immediately below each date header.
- Column F is normally the restaurant, but it can also contain PO codes,
  delivery text, addresses, or `DT... a.n. ...` metadata. Never create a new
  restaurant from a PO/DT/address value.
- Column H is `Nama Item`; column I is item type; J is quantity; K is ready
  stock; L is unit; M is unit price; N is item total; O is order total;
  P is delivery/ongkir detail; Q is pickup code such as T1/T2.
- A non-empty order number in column D starts a new order. Blank order-number
  cells continue the previous order until the next order number.
- Put each chat item into the matching existing order section. Put delivery
  time, address, PO code, atas nama, and ongkir in the existing row/field used
  by the template; do not move them into the restaurant field.
- If the requested date section does not exist, add a new date section by
  copying the existing date section's formatting and formulas, then fill it.

Do not merely rename the workbook. The output is invalid unless the new chat
values are visibly written into the cells and the original example values are
replaced or extended for the requested date.

Critical file rule:
- Produce exactly one workbook named `rekap_pesanan.xlsx`.
- Do not create files named V1, V2, Final, New, timestamped, or duplicate files.
- Do not add worksheets, columns, helper files, or redesign the workbook.
- Return only the completed workbook and a short note about ambiguous values.

Before returning, verify that the workbook contains the requested date,
restaurant names, item names, quantities, prices, DT times, and A.n. names.

Structured input:
{json.dumps(payload, ensure_ascii=False, indent=2)}
"""


def _date_matches(text: str, selected: str) -> bool:
    match = re.search(r"(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", text)
    if not match:
        return False
    months = {"januari": 1, "februari": 2, "maret": 3, "april": 4, "mei": 5, "juni": 6, "juli": 7, "agustus": 8, "september": 9, "oktober": 10, "november": 11, "desember": 12}
    try:
        wanted = datetime.strptime(selected, "%Y-%m-%d").date()
        return (int(match.group(3)), months.get(match.group(2).lower()), int(match.group(1))) == (wanted.year, wanted.month, wanted.day)
    except (ValueError, TypeError):
        return selected.lower() in text.lower()


def order_text_from_excel(uploaded, selected_date: str) -> str:
    try:
        import openpyxl
    except ImportError as exc:
        raise RuntimeError("Install dependencies first: pip install -r requirements.txt") from exc

    workbook = openpyxl.load_workbook(BytesIO(uploaded.getvalue()), data_only=True)
    sheet = workbook.active
    date_pattern = re.compile(r"(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday|Senin|Selasa|Rabu|Kamis|Jumat|Sabtu|Minggu)", re.I)
    active_date = ""
    orders = []
    current = None
    for row in sheet.iter_rows():
        values = [cell.value for cell in row]
        text = " ".join(str(value) for value in values if value not in (None, ""))
        if date_pattern.search(text) and re.search(r"\d{1,2}.*\d{4}", text):
            active_date = text.strip()
            continue
        if not active_date or (selected_date and not _date_matches(active_date, selected_date)):
            continue
        if not text:
            continue
        # In the example, D starts a new order. F contains either the restaurant
        # or metadata such as PO/DT/address/atas-nama for that same order.
        cells = values[5:18]
        restaurant, _, item, _, quantity, _, unit, price, item_amount, sales, delivery, pickup, buy_price = (cells + [None] * 13)[:13]
        if restaurant == "Resto" or item == "Nama Item":
            continue
        if values[3] not in (None, ""):
            current = {"restaurant": "", "items": [], "metadata": [], "total": None}
            orders.append(current)
        if current is None:
            continue
        restaurant_text = str(restaurant or "").strip()
        if restaurant_text:
            if re.match(r"^(PO\w+|DT\s*)", restaurant_text, re.I):
                current["metadata"].append(restaurant_text)
            elif not current["restaurant"]:
                current["restaurant"] = restaurant_text
        if item and str(item).strip().lower() != "ongkir":
            current["items"].append({"item": item, "quantity": quantity, "unit": unit, "price": price, "amount": item_amount})
        if item and str(item).strip().lower() == "ongkir":
            current["metadata"].append(f"ongkir: {price or delivery or pickup}")
        if sales not in (None, ""):
            current["total"] = sales
    if not orders:
        return f"No order rows found for date containing `{selected_date}`."

    lines = [f"ORDER LIST - {selected_date}"]
    for number, order in enumerate(orders, 1):
        details = [f"{number}. {order['restaurant'] or '(restaurant missing)'}"]
        items = "; ".join(f"{item['item']} x {item['quantity'] or '?'} {item['unit'] or ''} ({item['amount'] or item['price'] or 'price missing'})" for item in order["items"])
        if items:
            details.append(f"items: {items}")
        if order["total"]:
            details.append(f"total: {order['total']}")
        if order["metadata"]:
            delivery = " | ".join(order["metadata"])
            details.append(f"details: {delivery}")
            delivery_time = re.search(r"DT\s*([0-9:.]+)", delivery, re.I)
            recipient = re.search(r"(?:A\.n\.?|a\.n\.?)\s*(.+)$", delivery)
            if delivery_time:
                details.append(f"delivery time: {delivery_time.group(1)}")
            if recipient:
                details.append(f"atas nama: {recipient.group(1).strip()}")
        lines.append(" | ".join(details))
    return "\n".join(lines)


st.set_page_config(page_title="Gemini Excel Prompt Builder", layout="wide")
st.title("Gemini Excel Prompt Builder")
st.caption("Batch WhatsApp chats into one strict prompt for Gemini Web.")

with st.sidebar:
    st.header("Optional Excel template")
    template = st.file_uploader("Upload the example workbook", type=["xlsx"])
    st.info("Images are not needed. Upload the Excel only when you need Gemini to match its exact design.")

left, right = st.columns(2)
with left:
    st.subheader("1. Add chats in bulk")
    order_date = st.date_input("Order date", value=date.today())
    chat_files = st.file_uploader("Upload multiple .txt chat files", type=["txt"], accept_multiple_files=True)
    pasted = st.text_area("Or paste multiple chats", height=260, placeholder="Paste all messages here, or separate batches with headings.")
    chats = [{"name": file.name, "text": file.getvalue().decode("utf-8", errors="replace")} for file in chat_files]
    if pasted.strip():
        chats.append({"name": "pasted-chat", "text": pasted.strip()})
    payload = payload_for(chats, str(order_date))
    prompt = build_prompt(payload, template is not None)

    st.subheader("2. Generate order list from the Excel")
    summary_date = st.text_input("Date to find in the workbook", value=order_date.strftime("%Y-%m-%d"))
    if template and st.button("Generate dated order list"):
        try:
            st.session_state.order_list = order_text_from_excel(template, summary_date)
        except (OSError, RuntimeError, ValueError) as exc:
            st.error(str(exc))
    if "order_list" in st.session_state:
        st.text_area("Copyable order list", st.session_state.order_list, height=300)
        st.download_button("Download order_list.txt", st.session_state.order_list, "order_list.txt", "text/plain")

with right:
    st.subheader("3. Prompt for Gemini Web")
    st.code(prompt, language="text")
    artifact = {"prompt": prompt, "payload": payload}
    st.download_button("Download gemini_prompt.json", json.dumps(artifact, ensure_ascii=False, indent=2), "gemini_prompt.json", "application/json")
    st.download_button("Download order_payload.json", json.dumps(payload, ensure_ascii=False, indent=2), "order_payload.json", "application/json")
    if st.button("Save prompt files"):
        PROMPT_PATH.parent.mkdir(exist_ok=True)
        PROMPT_PATH.write_text(json.dumps(artifact, ensure_ascii=False, indent=2), encoding="utf-8")
        st.success("Saved output/gemini_prompt.json")
