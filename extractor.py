import re
import pandas as pd
from docx import Document

def clean_data(text):
    """Removes non-printable characters and whitespace."""
    if not text:
        return ""
    cleaned = "".join(ch for ch in text if ord(ch) >= 32 or ch in "\n\r\t")
    return cleaned.strip()

def parse_number(text, format_type):
    """
    STRICTLY parses a string into a float based on format_type.
    """
    if not text:
        return None
    text = text.strip()
    is_negative = False
    inner_text = text
    
    if inner_text.startswith('(') and inner_text.endswith(')'):
        is_negative = True
        inner_text = inner_text[1:-1].strip()
    elif inner_text.startswith('-'):
        is_negative = True
        inner_text = inner_text[1:].strip()
        
    if not inner_text or not re.match(r'^[0-9,.]+$', inner_text):
        return None
        
    try:
        if format_type == 'Vietnam':
            val_str = inner_text.replace('.', '').replace(',', '.')
        else:
            val_str = inner_text.replace(',', '')
        val = float(val_str)
        return -val if is_negative else val
    except (ValueError, TypeError):
        return None

def extract_table_data(file_path, format_type, mode='Number'):
    """
    Extracts data from all tables in a Word document, categorized by Table Index.
    """
    doc = Document(file_path)
    data = []
    
    for t_idx, table in enumerate(doc.tables, 1):
        for r_idx, row in enumerate(table.rows, 1):
            for c_idx, cell in enumerate(row.cells, 1):
                raw_text = clean_data(cell.text)
                if not raw_text:
                    continue
                    
                num_val = parse_number(raw_text, format_type)
                
                is_num = num_val is not None
                
                # Filter based on mode
                if (mode == 'Number' and is_num) or (mode == 'Non-Number' and not is_num):
                    data.append({
                        'Table': t_idx,
                        'Address': f"Table {t_idx}_R{r_idx}C{c_idx}",
                        'Value': num_val,
                        'Raw': raw_text
                    })
                    
    return pd.DataFrame(data)
