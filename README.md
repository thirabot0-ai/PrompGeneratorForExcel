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

The bundled `Thira_Fresh_Farm_Pricelist_updated.xlsx` is loaded automatically.
It reads all sheets using the `Product`, `Pricing Unit`, and `Price (Rp)`
columns. Prices can be changed with the product selector and price field; a chat line such as `carrot 3kg`
matches Baby Carrots and adds a calculated `3 × price/kg` total to the Gemini
payload.
