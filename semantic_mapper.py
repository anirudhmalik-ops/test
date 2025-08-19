import pandas as pd
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity
import logging

logger = logging.getLogger(__name__)
model = SentenceTransformer('all-MiniLM-L6-v2')

def load_excel(file_path):
    xl = pd.ExcelFile(file_path)
    data = {}
    for sheet in xl.sheet_names:
        df = xl.parse(sheet)
        df.columns = [str(col).strip() for col in df.columns]
        data[sheet] = df
    return data

def extract_template_structure(template_data):
    structure = {}
    for sheet_name, df in template_data.items():
        if df.empty:
            continue
        structure[sheet_name] = {
            "columns": list(df.columns),
            "rows": df.iloc[:, 0].dropna().astype(str).tolist()
        }
    return structure

def match_rows(template_rows, input_rows):
    template_emb = model.encode(template_rows)
    input_emb = model.encode(input_rows)

    mapping = {}
    for i, t_row in enumerate(template_rows):
        sims = cosine_similarity([template_emb[i]], input_emb)[0]
        best_match_idx = sims.argmax()
        mapping[t_row] = input_rows[best_match_idx]
    return mapping

def build_output(template_structure, input_data, row_mapping):
    output = {"sheets": {}}
    for sheet_name, meta in template_structure.items():
        output["sheets"][sheet_name] = []
        input_df = input_data.get(sheet_name)
        if input_df is None or input_df.empty:
            continue

        for row_label in meta["rows"]:
            matched_label = row_mapping.get(row_label)
            if matched_label is None:
                logger.warning(f"No match found for row: {row_label}")
                continue

            matched_row = input_df[input_df.iloc[:, 0].astype(str) == matched_label]
            if matched_row.empty:
                logger.warning(f"No data found for matched row: {matched_label}")
                continue

            row_dict = {}
            for col in meta["columns"]:
                if col in matched_row.columns:
                    row_dict[col] = matched_row[col].values[0]
                else:
                    row_dict[col] = None
            output["sheets"][sheet_name].append(row_dict)
    return output
