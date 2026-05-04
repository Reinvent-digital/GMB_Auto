from io import BytesIO
from pathlib import Path

import pandas as pd
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, StreamingResponse
from gmb_utils import (METRIC_COLUMNS, build_comparison_table,
                       infer_month_label, normalize_columns,
                       save_comparison_excel)

app = FastAPI(title="GMB Report Comparator")

HOME_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <title>GMB Report Comparator</title>
  <style>
    body { font-family: Arial, sans-serif; max-width: 760px; margin: 32px auto; line-height: 1.6; color: #111; }
    h1 { margin-bottom: 0.25em; }
    .card { padding: 24px; border: 1px solid #ddd; border-radius: 14px; box-shadow: 0 12px 32px rgba(0,0,0,.08); }
    label { display: block; margin-top: 16px; font-weight: 600; }
    input[type=file], input[type=text] { width: 100%; padding: 10px 12px; margin-top: 8px; border: 1px solid #bbb; border-radius: 8px; }
    button { margin-top: 24px; padding: 12px 20px; border: none; border-radius: 10px; background: #0066cc; color: white; cursor: pointer; font-size: 16px; }
    button:hover { background: #004f9f; }
    p.note { margin-top: 16px; color: #555; }
    .footer { margin-top: 24px; padding-top: 16px; border-top: 1px solid #eee; font-size: 0.95rem; color: #666; }
  </style>
</head>
<body>
  <div class="card">
    <h1>GMB Report Comparator</h1>
    <p>Upload two GMB export files and download a combined side-by-side comparison CSV.</p>

    <form id="compare-form" action="/compare" method="post" enctype="multipart/form-data">
      <label for="previous_file">Previous month file</label>
      <input type="file" id="previous_file" name="previous_file" accept=".csv,.xlsx,.xls" required />

      <label for="current_file">Current month file</label>
      <input type="file" id="current_file" name="current_file" accept=".csv,.xlsx,.xls" required />

      <label for="previous_label">Previous month label (optional)</label>
      <input type="text" id="previous_label" name="previous_label" placeholder="February" />

      <label for="current_label">Current month label (optional)</label>
      <input type="text" id="current_label" name="current_label" placeholder="March" />

      <label for="output_format">Output format</label>
      <select id="output_format" name="output_format">
        <option value="csv">CSV</option>
        <option value="xlsx" selected>Excel (.xlsx)</option>
      </select>

      <button type="submit">Generate comparison file</button>
    </form>

    <p class="note">If labels are left blank, the app will infer month names from the uploaded filenames.</p>
  </div>

  <div class="footer">
    <p>Supports CSV exports from Google Business Profile monthly reports.</p>
  </div>
</body>
</html>"""


@app.get("/", response_class=HTMLResponse)
async def home():
    return HOME_HTML


def read_gmb_file_upload(file: UploadFile) -> pd.DataFrame:
    suffix = Path(file.filename).suffix.lower()
    file.file.seek(0)
    if suffix in (".xlsx", ".xls"):
        raw = pd.read_excel(file.file, header=[0, 1], dtype=str)
    else:
        raw = pd.read_csv(file.file, header=[0, 1], dtype=str)

    cols = normalize_columns(raw.columns)
    raw.columns = cols

    if "Business name" not in raw.columns:
        file.file.seek(0)
        if suffix in (".xlsx", ".xls"):
            raw = pd.read_excel(file.file, header=0, dtype=str)
        else:
            raw = pd.read_csv(file.file, header=0, dtype=str)
        raw.columns = normalize_columns(raw.columns)

    keep = [c for c in raw.columns if c in ("Business name", "Address") + tuple(METRIC_COLUMNS)]
    if "Business name" not in keep:
        raise ValueError("Could not find Business name column in uploaded file")

    df = raw[keep].copy()
    df["Business name"] = df["Business name"].astype(str).str.strip()
    for metric in METRIC_COLUMNS:
        if metric in df.columns:
            df[metric] = pd.to_numeric(df[metric].fillna(0).replace("", 0), errors="coerce").fillna(0).astype(int)
    return df


@app.post("/compare")
async def compare(
    previous_file: UploadFile = File(...),
    current_file: UploadFile = File(...),
    previous_label: str = Form(None),
    current_label: str = Form(None),
    output_format: str = Form("xlsx"),
):
    try:
        datasets = []
        labels = []
        
        def add_file(f, custom_label):
            label = custom_label.strip() if custom_label else infer_month_label(Path(f.filename))
            df = read_gmb_file_upload(f)
            datasets.append((df, label))
            labels.append(label)

        add_file(previous_file, previous_label)
        add_file(current_file, current_label)

        result = build_comparison_table(datasets)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc))

    buffer = BytesIO()
    if output_format == "xlsx":
        save_comparison_excel(result, buffer, labels)
        buffer.seek(0)
        filename = f"gmb_comparison_{labels[0]}_vs_{labels[1]}.xlsx".replace(" ", "_")

        return StreamingResponse(
            buffer,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"},
        )

    buffer = BytesIO()
    result.to_csv(buffer, index=False)
    buffer.seek(0)
    filename = f"gmb_comparison_{labels[0]}_vs_{labels[1]}.csv".replace(" ", "_")
    return StreamingResponse(
        buffer,
        media_type="text/csv",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )
