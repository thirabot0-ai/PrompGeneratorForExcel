# Gemini Excel Prompt Builder

```powershell
pip install -r requirements.txt
streamlit run app.py
```

For the first run, upload the example template. For later runs, upload the
latest `rekap_pesanan.xlsx` as the current workbook to update. The Streamlit
upload is not automatically visible to Gemini Web, so download it and attach it
manually in the same Gemini conversation. Add multiple `.txt` chat files or
paste chats in bulk. The app does not call Gemini or need the client screenshots.
It generates a strict prompt for either initial creation or non-destructive
updates, without V1/V2/Final duplicates. With an uploaded workbook, it can also
create a dated order list containing items, quantities, prices, `DT` delivery
details, and `A.n.` recipient names.

The bundled `Pricelist TFF26.pdf` is loaded automatically when present. Prices
can be edited in the app; matching chat lines such as `carrot 3kg` add a
calculated `3 × price/kg` total to the Gemini payload.
