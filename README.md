# Gemini Excel Prompt Builder

```powershell
pip install -r requirements.txt
streamlit run app.py
```

Upload the Excel template only when exact formatting is needed. The Streamlit
upload is not automatically visible to Gemini Web, so download it from the app
and attach it manually in the same Gemini conversation. Add multiple
`.txt` chat files or paste chats in bulk. The app does not call Gemini or need
the client screenshots. It generates one strict prompt and JSON payload that
tells Gemini to create one fresh workbook named `rekap_pesanan.xlsx` using the
uploaded workbook only as a structure/style template, without copying its
sample order data or creating V1/V2/Final duplicates. With an uploaded
workbook, it can also create a dated order list containing items, quantities,
prices, `DT` delivery details, and `A.n.` recipient names.
