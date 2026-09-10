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

Critical file rule:
- Produce exactly one workbook named `rekap_pesanan.xlsx`.
- Do not create files named V1, V2, Final, New, timestamped, or duplicate files.
- Do not add worksheets, columns, helper files, or redesign the workbook.
- Return only the completed workbook and a short note about ambiguous values.

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
    rows = []
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
        # The example uses F:R for its order table: restaurant, item, quantity,
        # unit, item amount, delivery details, and other recap fields.
        cells = values[5:18]
        restaurant, _, item, _, quantity, _, unit, price, item_amount, sales, delivery, pickup, buy_price = (cells + [None] * 13)[:13]
        if item or delivery or pickup or restaurant:
            rows.append({"restaurant": restaurant, "item": item, "quantity": quantity, "unit": unit, "amount": item_amount or price or sales, "delivery": delivery or pickup, "buy_price": buy_price})
    if not rows:
        return f"No order rows found for date containing `{selected_date}`."

    lines = [f"ORDER LIST — {selected_date}"]
    for number, row in enumerate(rows, 1):
        details = [f"{number}. {row['restaurant'] or '(restaurant missing)'}"]
        if row["item"]:
            details.append(f"item: {row['item']}")
        if row["quantity"] not in (None, ""):
            details.append(f"amount: {row['quantity']} {row['unit'] or ''}".strip())
        if row["amount"] not in (None, ""):
            details.append(f"price/total: {row['amount']}")
        if row["delivery"]:
            delivery = str(row["delivery"])
            delivery_time = re.search(r"DT\s*([0-9:.]+)", delivery, re.I)
            recipient = re.search(r"(?:A\.n\.?|a\.n\.?)\s*(.+)$", delivery)
            details.append(f"delivery: {delivery}")
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
