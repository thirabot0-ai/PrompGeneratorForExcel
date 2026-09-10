# Gemini Excel Prompt Builder

```powershell
pip install -r requirements.txt
streamlit run app.py
```

Upload the Excel template only when exact formatting is needed. Add multiple
`.txt` chat files or paste chats in bulk. The app does not call Gemini or need
the client screenshots. It generates one strict prompt and JSON payload that
tells Gemini to return exactly one workbook named `rekap_pesanan.xlsx`, without
V1/V2/Final duplicates. With an uploaded workbook, it can also create a dated
order list containing items, quantities, prices, `DT` delivery details, and
`A.n.` recipient names.
