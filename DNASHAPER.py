import streamlit as st
import zipfile
import pandas as pd
from io import StringIO, BytesIO
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl import Workbook
import csv

st.title("🧬 DNA Shape Table Generator (with Sequence Info)")

uploaded_zip = st.file_uploader("Upload ZIP file containing `.txt` files (from DNAShapeR)", type=["zip"])
uploaded_fasta = st.file_uploader("Upload the corresponding `.fasta` file", type=["fasta", "fa", "txt"])

if uploaded_zip and uploaded_fasta:
    st.success("✅ Both ZIP and FASTA uploaded!")

    # --- Robust FASTA parser (handles multiline sequences) ---
    fasta_lines = uploaded_fasta.read().decode('utf-8').splitlines()
    sequence_ids, sequences = [], []
    current_id, current_seq_lines = None, []

    for line in fasta_lines:
        if line.startswith(">"):
            if current_id is not None:
                sequence_ids.append(current_id)
                sequences.append(''.join(current_seq_lines))
            current_id = line[1:].strip()
            current_seq_lines = []
        else:
            current_seq_lines.append(line.strip())

    if current_id is not None:
        sequence_ids.append(current_id)
        sequences.append(''.join(current_seq_lines))

    # --- Read ZIP files ---
    zip_bytes = BytesIO(uploaded_zip.read())
    dataframes = {}
    row_counts = []

    with zipfile.ZipFile(zip_bytes, 'r') as zip_ref:
        txt_files = [f for f in zip_ref.namelist() if f.endswith('.txt')]

        for file_name in txt_files:
            with zip_ref.open(file_name) as file:
                lines = [line.decode('utf-8').strip().replace('\t', ' ') for line in file.readlines()]
                cleaned_lines = [' '.join(line.split()) for line in lines]
                csv_ready = '\n'.join([line.replace(' ', ',') for line in cleaned_lines])

                # Detect if header exists
                first_line = csv_ready.split('\n')[0]
                has_header = any(char.isalpha() for char in first_line)

                try:
                    df = pd.read_csv(StringIO(csv_ready), header=0 if has_header else None)
                except pd.errors.ParserError:
                    reader = csv.reader(StringIO(csv_ready))
                    max_len = max(len(row) for row in reader)
                    reader = csv.reader(StringIO(csv_ready))
                    df = pd.DataFrame([row + ['0'] * (max_len - len(row)) for row in reader])

                df = df.apply(pd.to_numeric, errors='coerce').fillna(0)
                df = df.dropna(axis=1, how='all')
                base_name = file_name.split('/')[-1].replace('.txt', '')
                df[f"avg({base_name})"] = df.mean(axis=1)
                dataframes[base_name] = df
                row_counts.append(len(df))

    # --- Emergency auto-trim if mismatch is small
    min_len = min(row_counts[0], len(sequence_ids))
    if abs(len(sequence_ids) - row_counts[0]) == 1:
        st.warning("⚠️ FASTA and TXT row mismatch by 1 — auto-trimmed.")
        sequence_ids = sequence_ids[:min_len]
        sequences = sequences[:min_len]
        for key in dataframes:
            dataframes[key] = dataframes[key].iloc[:min_len]
        row_counts = [min_len]

    # --- Validation
    if len(set(row_counts)) != 1:
        st.error("❌ Not all `.txt` files have the same number of rows.")
        for fname, df in dataframes.items():
            st.text(f"{fname}: {df.shape[0]} rows")
        st.stop()
    if len(sequence_ids) != row_counts[0]:
        st.error(f"❌ FASTA sequence count ({len(sequence_ids)}) does not match data rows ({row_counts[0]}).")
        st.stop()

    # --- Create Excel
    output = BytesIO()
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Combined Data"
    ws_meta = wb.create_sheet("FASTA Meta")

    grey_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    red_font = Font(color="FF0000", bold=True)
    center_align = Alignment(horizontal="center", vertical="center")

    # --- Metadata sheet
    ws_meta.append(["Total Sequences", len(sequence_ids)])
    ws_meta.append(["Example ID", sequence_ids[0]])
    ws_meta.append(["Example Sequence", sequences[0][:100] + ("..." if len(sequences[0]) > 100 else "")])

    # --- Combined sheet: header rows
    ws1.cell(row=2, column=1, value="Sequence ID")
    ws1.cell(row=2, column=2, value="Sequence")

    start_col = 3
    for df_name, df in dataframes.items():
        col_count = df.shape[1]
        col_start = start_col
        col_end = col_start + col_count - 1
        ws1.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_end)
        header_cell = ws1.cell(row=1, column=col_start, value=df_name)
        header_cell.fill = grey_fill
        header_cell.alignment = center_align

        for j in range(df.shape[1] - 1):  # all but last
            ws1.cell(row=2, column=col_start + j, value=f"{df_name}_{j+1}")
        ws1.cell(row=2, column=col_end, value=f"avg({df_name})").font = red_font
        start_col += col_count

    # --- Fill rows
    for i in range(row_counts[0]):
        ws1.cell(row=i + 3, column=1, value=sequence_ids[i])
        ws1.cell(row=i + 3, column=2, value=sequences[i])

    start_col = 3
    for df in dataframes.values():
        for i, row in df.iterrows():
            for j, val in enumerate(row):
                cell = ws1.cell(row=i + 2, column=start_col + j, value=val)
                if j == len(row) - 1:
                    cell.font = red_font
        start_col += df.shape[1]

    wb.save(output)
    output.seek(0)

    st.success("✅ Excel file generated with Sequence ID, Sequence, and all features.")
    st.download_button("📥 Download Excel", output, file_name="DNAshape_with_sequence.xlsx")
