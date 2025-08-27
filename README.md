# EPL Excel → Streamlit App

This app loads your `EPL macro.xlsm` and lets you browse every sheet with filters, quick stats, and optional charts.

## How to run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```
Then open the URL shown in your terminal.

## Using your file
- By default, the app looks for a file named **EPL macro.xlsm** in the same folder. 
- Or you can **upload** any Excel workbook in the sidebar.
- Note: VBA/macros won't run in a web app. The app displays your data and lets you filter, sort, and chart it.

## Deploy
- Streamlit Community Cloud: upload these 3 files (`app.py`, `requirements.txt`, `README.md`) and your workbook if you want a fixed dataset, or rely on the uploader.
- Any server: `streamlit run app.py` behind your favorite process manager.
