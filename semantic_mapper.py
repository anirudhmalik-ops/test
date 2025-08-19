import pandas as pd
import logging
import difflib
import math

try:
    from sentence_transformers import SentenceTransformer
    _st_model = SentenceTransformer('all-MiniLM-L6-v2')
    _have_embed = True
except Exception:
    _st_model = None
    _have_embed = False

logger = logging.getLogger(__name__)

def _cosine_similarity_vector(vec_a, vec_b):
    dot = sum(a * b for a, b in zip(vec_a, vec_b))
    norm_a = math.sqrt(sum(a * a for a in vec_a))
    norm_b = math.sqrt(sum(b * b for b in vec_b))
    if norm_a == 0 or norm_b == 0:
        return 0.0
    return dot / (norm_a * norm_b)

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
    if not template_rows or not input_rows:
        return {}
    mapping = {}
    if _have_embed:
        template_emb = _st_model.encode(template_rows)
        input_emb = _st_model.encode(input_rows)
        for i, t_row in enumerate(template_rows):
            similarities = [
                _cosine_similarity_vector(template_emb[i], input_vec)
                for input_vec in input_emb
            ]
            best_match_idx = max(range(len(similarities)), key=lambda idx: similarities[idx])
            mapping[t_row] = input_rows[best_match_idx]
    else:
        for t_row in template_rows:
            best = difflib.get_close_matches(t_row, input_rows, n=1)
            mapping[t_row] = best[0] if best else None
    return mapping

def match_columns(template_columns, input_columns):
    if not template_columns or not input_columns:
        return {}
    template_columns_str = [str(col) for col in template_columns]
    input_columns_str = [str(col) for col in input_columns]

    column_mapping = {}
    if _have_embed:
        template_emb = _st_model.encode(template_columns_str)
        input_emb = _st_model.encode(input_columns_str)
        for i, t_col in enumerate(template_columns_str):
            similarities = [
                _cosine_similarity_vector(template_emb[i], input_vec)
                for input_vec in input_emb
            ]
            best_match_idx = max(range(len(similarities)), key=lambda idx: similarities[idx])
            column_mapping[t_col] = input_columns_str[best_match_idx]
    else:
        for t_col in template_columns_str:
            best = difflib.get_close_matches(t_col, input_columns_str, n=1)
            column_mapping[t_col] = best[0] if best else input_columns_str[0]
    return column_mapping

def build_output(template_structure, input_data, row_mapping_by_sheet):
    output = {"sheets": {}}
    for sheet_name, meta in template_structure.items():
        output["sheets"][sheet_name] = []
        input_df = input_data.get(sheet_name)
        if input_df is None or input_df.empty:
            continue
        # Compute column mapping between template and input
        template_columns = meta.get("columns", [])
        input_columns = list(input_df.columns)
        column_mapping = match_columns(template_columns, input_columns)

        sheet_row_mapping = row_mapping_by_sheet.get(sheet_name, {})
        for row_label in meta["rows"]:
            matched_label = sheet_row_mapping.get(row_label)
            if matched_label is None:
                logger.warning(f"No match found for row: {row_label}")
                continue

            matched_row = input_df[input_df.iloc[:, 0].astype(str) == matched_label]
            if matched_row.empty:
                logger.warning(f"No data found for matched row: {matched_label}")
                continue

            row_dict = {}
            for template_col in template_columns:
                input_col = column_mapping.get(str(template_col))
                if input_col in matched_row.columns:
                    row_dict[template_col] = matched_row[input_col].values[0]
                else:
                    row_dict[template_col] = None
            output["sheets"][sheet_name].append(row_dict)
    return output
