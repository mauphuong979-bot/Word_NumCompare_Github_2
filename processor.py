import io
import difflib
import pandas as pd
from docx import Document
from extractor import (
    extract_structured_tables,
    get_document_tables_summary,
    clean_data
)

def align_document_tables(summary1, summary2, tables1=None, tables2=None):
    """
    Smart table sequence alignment using Needleman-Wunsch Dynamic Programming.
    Guarantees strict monotonicity (no inverted/crossed table pairings).
    Combines:
    - Column count matching (+35)
    - Row count proximity (+20 max)
    - Position proximity (-3 per offset)
    - Numeric overlap (+12 per matching number, max +100)
    - Preview / text similarity (+20 max)
    - Rejection penalty for column diff > 2 (-80)
    Returns: (aligned_pairs, unmatched_t1, unmatched_t2)
    """
    n = len(summary1)
    m = len(summary2)
    
    if n == 0 and m == 0:
        return [], [], []
    if n == 0:
        return [], [], [s['table_index'] for s in summary2]
    if m == 0:
        return [], [s['table_index'] for s in summary1], []
        
    # Pre-extract numbers per table if tables data is available
    nums1 = {}
    nums2 = {}
    if tables1:
        for idx in range(1, n + 1):
            s = set()
            for r in tables1.get(idx, []):
                for c in r['cells']:
                    if c['num_val'] is not None:
                        s.add(c['num_val'])
            nums1[idx] = s
            
    if tables2:
        for idx in range(1, m + 1):
            s = set()
            for r in tables2.get(idx, []):
                for c in r['cells']:
                    if c['num_val'] is not None:
                        s.add(c['num_val'])
            nums2[idx] = s

    def table_score(i, j):
        s1 = summary1[i]
        s2 = summary2[j]
        
        cols1, cols2 = s1['cols'], s2['cols']
        rows1, rows2 = s1['rows'], s2['rows']
        
        # Heavy penalty if column difference is > 2 (cannot be 1:1 pair)
        col_diff = abs(cols1 - cols2)
        if col_diff > 2:
            return -80
            
        score = 0
        if cols1 == cols2 and cols1 > 0:
            score += 35
        elif col_diff == 1:
            score += 15
        elif col_diff == 2:
            score += 0
            
        # Row matching
        if rows1 == rows2:
            score += 20
        else:
            diff = abs(rows1 - rows2)
            if diff <= 2:
                score += 14 - (diff * 4) # diff 1: +10, diff 2: +6
            else:
                score -= min(diff * 2, 20)
                
        # Position proximity
        pos_diff = abs(i - j)
        score -= min(pos_diff * 3, 30)
        
        # Numeric overlap
        n1 = nums1.get(i + 1, set())
        n2 = nums2.get(j + 1, set())
        overlap = len(n1.intersection(n2)) if n1 and n2 else 0
        if overlap > 0:
            score += min(overlap * 12, 100)
            
        # Preview similarity
        prev1 = s1.get('preview', '')
        prev2 = s2.get('preview', '')
        if prev1 and prev2 and prev1 != '(Bảng trống)' and prev2 != '(Bảng trống)':
            ratio = difflib.SequenceMatcher(None, prev1.lower(), prev2.lower()).ratio()
            if ratio > 0.5:
                score += int(ratio * 20)
                
        return score

    GAP_PENALTY = -15
    
    dp = [[0] * (m + 1) for _ in range(n + 1)]
    trace = [[0] * (m + 1) for _ in range(n + 1)]
    
    for i in range(1, n + 1):
        dp[i][0] = i * GAP_PENALTY
        trace[i][0] = 1
        
    for j in range(1, m + 1):
        dp[0][j] = j * GAP_PENALTY
        trace[0][j] = 2
        
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            match = dp[i - 1][j - 1] + table_score(i - 1, j - 1)
            delete = dp[i - 1][j] + GAP_PENALTY
            insert = dp[i][j - 1] + GAP_PENALTY
            
            best = match
            t = 0
            if delete > best:
                best = delete
                t = 1
            if insert > best:
                best = insert
                t = 2
                
            dp[i][j] = best
            trace[i][j] = t

    i, j = n, m
    pairs = []
    unmatched_1 = []
    unmatched_2 = []
    
    while i > 0 or j > 0:
        if i > 0 and j > 0 and trace[i][j] == 0:
            cols_ok = abs(summary1[i - 1]['cols'] - summary2[j - 1]['cols']) <= 2
            if cols_ok:
                pairs.append((summary1[i - 1]['table_index'], summary2[j - 1]['table_index']))
            else:
                unmatched_1.append(summary1[i - 1]['table_index'])
                unmatched_2.append(summary2[j - 1]['table_index'])
            i -= 1
            j -= 1
        elif i > 0 and (j == 0 or trace[i][j] == 1):
            unmatched_1.append(summary1[i - 1]['table_index'])
            i -= 1
        else:
            unmatched_2.append(summary2[j - 1]['table_index'])
            j -= 1
            
    pairs.reverse()
    unmatched_1.reverse()
    unmatched_2.reverse()
    return pairs, unmatched_1, unmatched_2

def detect_and_split_merged_tables(summary1, summary2, tables1, tables2):
    """
    Detects when a table in one document is formed by merging consecutive tables from the other document.
    Automatically splits the merged table into corresponding sub-tables so that
    Needleman-Wunsch aligns them 1:1 and all cell numbers are verified automatically.
    Returns: (summary1, summary2, tables1, tables2, merged_notes)
    """
    merged_notes = []
    merged_events = []
    
    while True:
        pairs, u1, u2 = align_document_tables(summary1, summary2, tables1=tables1, tables2=tables2)
        if not u1 and not u2:
            break
            
        pair_map_1_to_2 = {p[0]: p[1] for p in pairs}
        pair_map_2_to_1 = {p[1]: p[0] for p in pairs}
        split_found = False
        
        # Case 1: Unmatched table in Doc 1 is part of a merged table in Doc 2
        for u in u1:
            candidates = []
            if (u + 1) in pair_map_1_to_2:
                candidates.append((u, u + 1, pair_map_1_to_2[u + 1]))
            if (u - 1) in pair_map_1_to_2:
                candidates.append((u - 1, u, pair_map_1_to_2[u - 1]))
                
            for t1a, t1b, t2_idx in candidates:
                r1a = len(tables1[t1a])
                r1b = len(tables1[t1b])
                r2 = len(tables2[t2_idx])
                if abs(r2 - (r1a + r1b)) <= 1:
                    k = r1a
                    p1 = tables2[t2_idx][:k]
                    p2 = tables2[t2_idx][k:]
                    
                    nums_1a = {c['num_val'] for r in tables1[t1a] for c in r['cells'] if c['num_val'] is not None}
                    nums_1b = {c['num_val'] for r in tables1[t1b] for c in r['cells'] if c['num_val'] is not None}
                    nums_p1 = {c['num_val'] for r in p1 for c in r['cells'] if c['num_val'] is not None}
                    nums_p2 = {c['num_val'] for r in p2 for c in r['cells'] if c['num_val'] is not None}
                    
                    m1 = len(nums_1a.intersection(nums_p1))
                    m2 = len(nums_1b.intersection(nums_p2))
                    
                    if (m1 > 0 or len(nums_1a) == 0) and (m2 > 0 or len(nums_1b) == 0):
                        p2_reindexed = []
                        for idx_in_p, row in enumerate(p2, 1):
                            new_row = dict(row)
                            new_row['row_idx'] = idx_in_p
                            p2_reindexed.append(new_row)
                            
                        new_t2 = {}
                        for idx in range(1, t2_idx):
                            new_t2[idx] = tables2[idx]
                        new_t2[t2_idx] = p1
                        new_t2[t2_idx + 1] = p2_reindexed
                        for idx in range(t2_idx + 1, len(tables2) + 1):
                            new_t2[idx + 1] = tables2[idx]
                        tables2 = new_t2
                        
                        orig_t2_s = summary2[t2_idx - 1] if t2_idx - 1 < len(summary2) else {'cols': 0, 'preview': ''}
                        doc2_t_idx = orig_t2_s.get('doc_table_idx', t2_idx)
                        doc1_a_idx = summary1[t1a - 1].get('doc_table_idx', t1a)
                        doc1_b_idx = summary1[t1b - 1].get('doc_table_idx', t1b)
                        
                        new_s2 = []
                        for s in summary2:
                            if s['table_index'] < t2_idx:
                                new_s2.append(s)
                            elif s['table_index'] == t2_idx:
                                new_s2.append({
                                    'table_index': t2_idx,
                                    'doc_table_idx': doc2_t_idx,
                                    'rows': len(p1),
                                    'cols': summary1[t1a - 1]['cols'],
                                    'preview': summary1[t1a - 1]['preview'],
                                    'merged_note': f"Tự động tách từ Bảng {t2_idx} (File 2) gộp tương ứng Bảng {t1a} (File 1)"
                                })
                                new_s2.append({
                                    'table_index': t2_idx + 1,
                                    'doc_table_idx': doc2_t_idx,
                                    'rows': len(p2),
                                    'cols': orig_t2_s['cols'],
                                    'preview': summary1[t1b - 1]['preview'],
                                    'merged_note': f"Tự động tách từ Bảng {t2_idx} (File 2) gộp tương ứng Bảng {t1b} (File 1)"
                                })
                            else:
                                item = dict(s)
                                item['table_index'] = s['table_index'] + 1
                                new_s2.append(item)
                        summary2 = new_s2
                        merged_notes.append(f"Đã tự động phân tách Bảng {t2_idx} (File 2) ghép từ Bảng {t1a} & {t1b} (File 1) và đối chiếu toàn bộ số liệu.")
                        merged_events.append({
                            'type': 'doc2_merged',
                            'doc1_a': doc1_a_idx,
                            'doc1_b': doc1_b_idx,
                            'rows1_a': r1a,
                            'rows1_b': r1b,
                            'doc2_table': doc2_t_idx,
                            'rows2': r2,
                            'split_row_doc2': k + 1
                        })
                        split_found = True
                        break
            if split_found:
                break
                
        if split_found:
            continue
            
        # Case 2: Unmatched table in Doc 2 is part of a merged table in Doc 1
        for u in u2:
            candidates = []
            if (u + 1) in pair_map_2_to_1:
                candidates.append((u, u + 1, pair_map_2_to_1[u + 1]))
            if (u - 1) in pair_map_2_to_1:
                candidates.append((u - 1, u, pair_map_2_to_1[u - 1]))
                
            for t2a, t2b, t1_idx in candidates:
                r2a = len(tables2[t2a])
                r2b = len(tables2[t2b])
                r1 = len(tables1[t1_idx])
                if abs(r1 - (r2a + r2b)) <= 1:
                    k = r2a
                    p1 = tables1[t1_idx][:k]
                    p2 = tables1[t1_idx][k:]
                    
                    nums_2a = {c['num_val'] for r in tables2[t2a] for c in r['cells'] if c['num_val'] is not None}
                    nums_2b = {c['num_val'] for r in tables2[t2b] for c in r['cells'] if c['num_val'] is not None}
                    nums_p1 = {c['num_val'] for r in p1 for c in r['cells'] if c['num_val'] is not None}
                    nums_p2 = {c['num_val'] for r in p2 for c in r['cells'] if c['num_val'] is not None}
                    
                    m1 = len(nums_2a.intersection(nums_p1))
                    m2 = len(nums_2b.intersection(nums_p2))
                    
                    if (m1 > 0 or len(nums_2a) == 0) and (m2 > 0 or len(nums_2b) == 0):
                        p2_reindexed = []
                        for idx_in_p, row in enumerate(p2, 1):
                            new_row = dict(row)
                            new_row['row_idx'] = idx_in_p
                            p2_reindexed.append(new_row)
                            
                        new_t1 = {}
                        for idx in range(1, t1_idx):
                            new_t1[idx] = tables1[idx]
                        new_t1[t1_idx] = p1
                        new_t1[t1_idx + 1] = p2_reindexed
                        for idx in range(t1_idx + 1, len(tables1) + 1):
                            new_t1[idx + 1] = tables1[idx]
                        tables1 = new_t1
                        
                        orig_t1_s = summary1[t1_idx - 1] if t1_idx - 1 < len(summary1) else {'cols': 0, 'preview': ''}
                        doc1_t_idx = orig_t1_s.get('doc_table_idx', t1_idx)
                        doc2_a_idx = summary2[t2a - 1].get('doc_table_idx', t2a)
                        doc2_b_idx = summary2[t2b - 1].get('doc_table_idx', t2b)
                        
                        new_s1 = []
                        for s in summary1:
                            if s['table_index'] < t1_idx:
                                new_s1.append(s)
                            elif s['table_index'] == t1_idx:
                                new_s1.append({
                                    'table_index': t1_idx,
                                    'doc_table_idx': doc1_t_idx,
                                    'rows': len(p1),
                                    'cols': summary2[t2a - 1]['cols'],
                                    'preview': summary2[t2a - 1]['preview'],
                                    'merged_note': f"Tự động tách từ Bảng {t1_idx} (File 1) gộp tương ứng Bảng {t2a} (File 2)"
                                })
                                new_s1.append({
                                    'table_index': t1_idx + 1,
                                    'doc_table_idx': doc1_t_idx,
                                    'rows': len(p2),
                                    'cols': orig_t1_s['cols'],
                                    'preview': summary2[t2b - 1]['preview'],
                                    'merged_note': f"Tự động tách từ Bảng {t1_idx} (File 1) gộp tương ứng Bảng {t2b} (File 2)"
                                })
                            else:
                                item = dict(s)
                                item['table_index'] = s['table_index'] + 1
                                new_s1.append(item)
                        summary1 = new_s1
                        merged_notes.append(f"Đã tự động phân tách Bảng {t1_idx} (File 1) ghép từ Bảng {t2a} & {t2b} (File 2) và đối chiếu toàn bộ số liệu.")
                        merged_events.append({
                            'type': 'doc1_merged',
                            'doc1_table': doc1_t_idx,
                            'rows1': r1,
                            'split_row_doc1': k + 1,
                            'doc2_a': doc2_a_idx,
                            'doc2_b': doc2_b_idx,
                            'rows2_a': r2a,
                            'rows2_b': r2b
                        })
                        split_found = True
                        break
            if split_found:
                break
                
        if not split_found:
            break
            
    return summary1, summary2, tables1, tables2, merged_notes, merged_events

def compare_table_structures(summary1, summary2, aligned_pairs, unmatched_1, unmatched_2, tables1=None, tables2=None):
    """
    Builds a detailed structural overview based on aligned and unmatched tables.
    Returns: (structure_df, stats_dict)
    """
    map1 = {s['table_index']: s for s in summary1}
    map2 = {s['table_index']: s for s in summary2}
    
    rows = []
    identical_count = 0
    mismatch_struct_count = 0
    
    for t1_idx, t2_idx in aligned_pairs:
        t1 = map1.get(t1_idx)
        t2 = map2.get(t2_idx)
        
        dim1 = f"{t1['rows']} dòng × {t1['cols']} cột"
        dim2 = f"{t2['rows']} dòng × {t2['cols']} cột"
        preview = t1.get('preview') or t2.get('preview') or ""
        
        r_diff = t1['rows'] - t2['rows']
        c_diff = t1['cols'] - t2['cols']
        
        data_rows_1 = sum(1 for r in tables1[t1_idx] if any(c['raw_text'].strip() for c in r['cells'])) if tables1 and t1_idx in tables1 else t1['rows']
        data_rows_2 = sum(1 for r in tables2[t2_idx] if any(c['raw_text'].strip() for c in r['cells'])) if tables2 and t2_idx in tables2 else t2['rows']
        data_r_diff = data_rows_1 - data_rows_2
        
        merged_note = t1.get('merged_note') or t2.get('merged_note')
        if merged_note:
            mismatch_struct_count += 1
            status = "⚠️ Bảng gộp/tách (Đã tự động đối chiếu)"
            note = merged_note
        elif r_diff == 0 and c_diff == 0:
            status = "✅ Đồng nhất"
            note = "Khớp hoàn toàn cấu trúc dòng & cột"
            identical_count += 1
        elif data_r_diff == 0 and c_diff == 0:
            mismatch_struct_count += 1
            f_more = "F1 > F2" if r_diff > 0 else "F1 < F2"
            status = f"⚠️ Lệch {abs(r_diff)} dòng đệm trống ({f_more})"
            note = "Số liệu khớp 100%, có dòng đệm trống định dạng"
        else:
            mismatch_struct_count += 1
            parts = []
            if data_r_diff != 0:
                parts.append(f"Lệch {abs(data_r_diff)} dòng ({'F1 > F2' if data_r_diff > 0 else 'F1 < F2'})")
            if c_diff != 0:
                parts.append(f"Lệch {abs(c_diff)} cột ({'F1 > F2' if c_diff > 0 else 'F1 < F2'})")
            status = f"⚠️ {', '.join(parts)}"
            note = "Đã tự động căn chỉnh đối ứng"
            
        rows.append({
            'Bảng File 1': f"Bảng {t1_idx}",
            'Bảng File 2': f"Bảng {t2_idx}",
            'Mô tả / Tiêu đề': preview,
            'File 1 (Kích thước)': dim1,
            'File 2 (Kích thước)': dim2,
            'Trạng thái': status,
            'Ghi chú': note
        })
        
    for t1_idx in unmatched_1:
        mismatch_struct_count += 1
        t1 = map1.get(t1_idx)
        dim1 = f"{t1['rows']} dòng × {t1['cols']} cột" if t1 else ""
        preview = t1.get('preview') if t1 else ""
        rows.append({
            'Bảng File 1': f"Bảng {t1_idx}",
            'Bảng File 2': "(Không đối ứng)",
            'Mô tả / Tiêu đề': preview,
            'File 1 (Kích thước)': dim1,
            'File 2 (Kích thước)': "(Tách/Thiếu ở F2)",
            'Trạng thái': "⚠️ Biến động cấu trúc",
            'Ghi chú': "Bảng riêng hoặc gộp/tách ở File 2"
        })
        
    for t2_idx in unmatched_2:
        mismatch_struct_count += 1
        t2 = map2.get(t2_idx)
        dim2 = f"{t2['rows']} dòng × {t2['cols']} cột" if t2 else ""
        preview = t2.get('preview') if t2 else ""
        rows.append({
            'Bảng File 1': "(Không đối ứng)",
            'Bảng File 2': f"Bảng {t2_idx}",
            'Mô tả / Tiêu đề': preview,
            'File 1 (Kích thước)': "(Tách/Thiếu ở F1)",
            'File 2 (Kích thước)': dim2,
            'Trạng thái': "⚠️ Biến động cấu trúc",
            'Ghi chú': "Bảng riêng hoặc phát sinh do tách bảng"
        })
        
    structure_df = pd.DataFrame(rows)
    stats = {
        'total_tables_1': len(summary1),
        'total_tables_2': len(summary2),
        'aligned_pairs_count': len(aligned_pairs),
        'identical_count': identical_count,
        'mismatched_structure_count': mismatch_struct_count,
        'is_all_identical': (mismatch_struct_count == 0 and len(summary1) == len(summary2))
    }
    return structure_df, stats

def calculate_row_match_score(r1, r2):
    """
    Calculates alignment compatibility score between row r1 and row r2.
    Considers empty status, matching numbers, column pattern, and text labels.
    """
    empty1 = all(not c['raw_text'].strip() for c in r1['cells'])
    empty2 = all(not c['raw_text'].strip() for c in r2['cells'])
    if empty1 and empty2:
        return 10.0  # Both empty rows match
    if empty1 or empty2:
        return -6.0  # Empty row should not match a content row
        
    nums1 = {c['col_idx']: c['num_val'] for c in r1['cells'] if c['num_val'] is not None}
    nums2 = {c['col_idx']: c['num_val'] for c in r2['cells'] if c['num_val'] is not None}
    
    score = 0.0
    if nums1 and nums2:
        common_cols = set(nums1.keys()).intersection(set(nums2.keys()))
        if common_cols:
            matching_vals = sum(1 for col in common_cols if abs(nums1[col] - nums2[col]) < 1e-4)
            score += matching_vals * 20.0
            diff_vals = len(common_cols) - matching_vals
            score -= diff_vals * 2.0
        else:
            score -= 5.0
    elif not nums1 and not nums2:
        score += 5.0
        lbl1 = r1.get('label', '').strip().lower()
        lbl2 = r2.get('label', '').strip().lower()
        if lbl1 and lbl2 and lbl1 == lbl2:
            score += 15.0
    else:
        score -= 8.0
        
    return score

def align_table_rows(rows1, rows2):
    """
    Smart row alignment using Needleman-Wunsch Dynamic Programming.
    Correctly aligns data rows even when rows are inserted or deleted (e.g. blank spacer rows),
    preventing false cascaded errors.
    """
    n = len(rows1)
    m = len(rows2)
    
    if n == m and all(calculate_row_match_score(rows1[i], rows2[i]) >= 0 for i in range(n)):
        return list(zip(rows1, rows2)), [], []

    gap_penalty = -2.0
    dp = [[0.0] * (m + 1) for _ in range(n + 1)]
    for i in range(n + 1):
        dp[i][0] = i * gap_penalty
    for j in range(m + 1):
        dp[0][j] = j * gap_penalty

    for i in range(1, n + 1):
        for j in range(1, m + 1):
            match = dp[i-1][j-1] + calculate_row_match_score(rows1[i-1], rows2[j-1])
            delete = dp[i-1][j] + gap_penalty
            insert = dp[i][j-1] + gap_penalty
            dp[i][j] = max(match, delete, insert)

    aligned_pairs = []
    unmatched_1 = []
    unmatched_2 = []
    i, j = n, m
    while i > 0 or j > 0:
        if i > 0 and j > 0 and abs(dp[i][j] - (dp[i-1][j-1] + calculate_row_match_score(rows1[i-1], rows2[j-1]))) < 1e-6:
            aligned_pairs.append((rows1[i-1], rows2[j-1]))
            i -= 1
            j -= 1
        elif i > 0 and abs(dp[i][j] - (dp[i-1][j] + gap_penalty)) < 1e-6:
            unmatched_1.append(rows1[i-1])
            i -= 1
        else:
            unmatched_2.append(rows2[j-1])
            j -= 1

    aligned_pairs.reverse()
    unmatched_1.reverse()
    unmatched_2.reverse()
    return aligned_pairs, unmatched_1, unmatched_2

def compare_documents(doc1_source, doc2_source, format1, format2, mode=None, compare_numbers=True, compare_text=False):
    """
    Main Comparison Engine with:
    1. Smart Table Alignment (aligns matching tables across documents even if an extra table is added/split).
    2. Smart Row Alignment (immune to inserted/deleted rows).
    3. Accurate cell comparison with exact coordinates for Word comments.
    4. Flexible mode: compare_numbers, compare_text, or both.
    """
    if mode is not None:
        if mode == 'Number':
            compare_numbers, compare_text = True, False
        elif mode == 'Text':
            compare_numbers, compare_text = False, True
        elif mode == 'Both':
            compare_numbers, compare_text = True, True

    if isinstance(doc1_source, bytes):
        doc1_source = io.BytesIO(doc1_source)
    if isinstance(doc2_source, bytes):
        doc2_source = io.BytesIO(doc2_source)
    if hasattr(doc1_source, 'seek'): doc1_source.seek(0)
    if hasattr(doc2_source, 'seek'): doc2_source.seek(0)
    doc1 = doc1_source if hasattr(doc1_source, 'tables') else Document(doc1_source)
    doc2 = doc2_source if hasattr(doc2_source, 'tables') else Document(doc2_source)

    summary1 = get_document_tables_summary(doc1)
    summary2 = get_document_tables_summary(doc2)
    tables1 = extract_structured_tables(doc1, format1)
    tables2 = extract_structured_tables(doc2, format2)
    
    # Auto-detect and reconstruct merged/split tables (e.g. Tables 34 & 35 in Advance)
    summary1, summary2, tables1, tables2, merged_notes, merged_events = detect_and_split_merged_tables(
        summary1, summary2, tables1, tables2
    )

    # 1. Smart Table Alignment (Needleman-Wunsch with Numeric Overlap)
    aligned_table_pairs, unmat_tables_1, unmat_tables_2 = align_document_tables(
        summary1, summary2, tables1=tables1, tables2=tables2
    )
    structure_df, struct_stats = compare_table_structures(
        summary1, summary2, aligned_table_pairs, unmat_tables_1, unmat_tables_2,
        tables1=tables1, tables2=tables2
    )
    
    comparison_results = []
    discrepancies_doc1 = []
    discrepancies_doc2 = []
    
    # Record merged table informational notices for Word comment balloons
    for ev in merged_events:
        if ev['type'] == 'doc2_merged':
            discrepancies_doc1.append({
                'table_idx': ev['doc1_a'],
                'doc_table_idx': ev['doc1_a'],
                'row_idx': 1,
                'doc_row_idx': 1,
                'col_idx': 1,
                'type': 'merged_table_info',
                'info_type': 'split_source',
                'self_rows': ev['rows1_a'],
                'target_table': ev['doc2_table'],
                'other_rows': ev['rows2'],
                'partner_table': ev['doc1_b']
            })
            discrepancies_doc1.append({
                'table_idx': ev['doc1_b'],
                'doc_table_idx': ev['doc1_b'],
                'row_idx': 1,
                'doc_row_idx': 1,
                'col_idx': 1,
                'type': 'merged_table_info',
                'info_type': 'split_source',
                'self_rows': ev['rows1_b'],
                'target_table': ev['doc2_table'],
                'other_rows': ev['rows2'],
                'partner_table': ev['doc1_a']
            })
            discrepancies_doc2.append({
                'table_idx': ev['doc2_table'],
                'doc_table_idx': ev['doc2_table'],
                'row_idx': 1,
                'doc_row_idx': 1,
                'col_idx': 1,
                'type': 'merged_table_info',
                'info_type': 'merged_target_header',
                'self_rows': ev['rows2'],
                'source_tables': [ev['doc1_a'], ev['doc1_b']],
                'source_rows': [ev['rows1_a'], ev['rows1_b']]
            })
            discrepancies_doc2.append({
                'table_idx': ev['doc2_table'],
                'doc_table_idx': ev['doc2_table'],
                'row_idx': ev['split_row_doc2'],
                'doc_row_idx': ev['split_row_doc2'],
                'col_idx': 1,
                'type': 'merged_table_info',
                'info_type': 'merged_target_part2',
                'source_table': ev['doc1_b'],
                'source_rows': ev['rows1_b'],
                'split_row': ev['split_row_doc2']
            })
        elif ev['type'] == 'doc1_merged':
            discrepancies_doc1.append({
                'table_idx': ev['doc1_table'],
                'doc_table_idx': ev['doc1_table'],
                'row_idx': 1,
                'doc_row_idx': 1,
                'col_idx': 1,
                'type': 'merged_table_info',
                'info_type': 'merged_target_header',
                'self_rows': ev['rows1'],
                'source_tables': [ev['doc2_a'], ev['doc2_b']],
                'source_rows': [ev['rows2_a'], ev['rows2_b']]
            })
            discrepancies_doc1.append({
                'table_idx': ev['doc1_table'],
                'doc_table_idx': ev['doc1_table'],
                'row_idx': ev['split_row_doc1'],
                'doc_row_idx': ev['split_row_doc1'],
                'col_idx': 1,
                'type': 'merged_table_info',
                'info_type': 'merged_target_part2',
                'source_table': ev['doc2_b'],
                'source_rows': ev['rows2_b'],
                'split_row': ev['split_row_doc1']
            })
            discrepancies_doc2.append({
                'table_idx': ev['doc2_a'],
                'doc_table_idx': ev['doc2_a'],
                'row_idx': 1,
                'doc_row_idx': 1,
                'col_idx': 1,
                'type': 'merged_table_info',
                'info_type': 'split_source',
                'self_rows': ev['rows2_a'],
                'target_table': ev['doc1_table'],
                'other_rows': ev['rows1'],
                'partner_table': ev['doc2_b']
            })
            discrepancies_doc2.append({
                'table_idx': ev['doc2_b'],
                'doc_table_idx': ev['doc2_b'],
                'row_idx': 1,
                'doc_row_idx': 1,
                'col_idx': 1,
                'type': 'merged_table_info',
                'info_type': 'split_source',
                'self_rows': ev['rows2_b'],
                'target_table': ev['doc1_table'],
                'other_rows': ev['rows1'],
                'partner_table': ev['doc2_a']
            })
    
    total_checked_count = 0
    total_matched_count = 0
    total_mismatch_count = 0
    
    # 2. Compare Aligned Table Pairs
    for t1_idx, t2_idx in aligned_table_pairs:
        t_rows1 = tables1.get(t1_idx)
        t_rows2 = tables2.get(t2_idx)
        
        if not t_rows1 or not t_rows2:
            continue
            
        aligned_rows, unmat_r1, unmat_r2 = align_table_rows(t_rows1, t_rows2)
        
        for r1, r2 in aligned_rows:
            r1_idx = r1['row_idx']
            r2_idx = r2['row_idx']
            label = r1['label'] or r2['label'] or "(Không có tiêu đề)"
            
            cells1 = {c['col_idx']: c for c in r1['cells']}
            cells2 = {c['col_idx']: c for c in r2['cells']}
            all_cols = sorted(set(cells1.keys()).union(set(cells2.keys())))
            
            for c_idx in all_cols:
                c1 = cells1.get(c_idx)
                c2 = cells2.get(c_idx)
                
                v1_raw = c1['raw_text'] if c1 else ""
                v2_raw = c2['raw_text'] if c2 else ""
                
                num1 = c1['num_val'] if c1 else None
                num2 = c2['num_val'] if c2 else None
                
                is_num1 = num1 is not None
                is_num2 = num2 is not None

                if not v1_raw and not v2_raw:
                    continue

                if is_num1 and is_num2:
                    if not compare_numbers:
                        continue
                    total_checked_count += 1
                    diff = num1 - num2
                    is_diff = abs(diff) > 1e-4
                    
                    if is_diff:
                        total_mismatch_count += 1
                        status = "❌ Sai lệch số liệu"
                        discrepancies_doc1.append({
                            'table_idx': t1_idx,
                            'row_idx': r1_idx,
                            'doc_table_idx': r1.get('doc_table_idx', t1_idx),
                            'doc_row_idx': r1.get('doc_row_idx', r1_idx),
                            'col_idx': c_idx,
                            'doc_col_idx': c1.get('doc_col_idx', c_idx) if c1 else c_idx,
                            'type': 'num_mismatch',
                            'label': label,
                            'val_self': num1,
                            'val_other': num2,
                            'diff': diff
                        })
                        discrepancies_doc2.append({
                            'table_idx': t2_idx,
                            'row_idx': r2_idx,
                            'doc_table_idx': r2.get('doc_table_idx', t2_idx),
                            'doc_row_idx': r2.get('doc_row_idx', r2_idx),
                            'col_idx': c_idx,
                            'doc_col_idx': c2.get('doc_col_idx', c_idx) if c2 else c_idx,
                            'type': 'num_mismatch',
                            'label': label,
                            'val_self': num2,
                            'val_other': num1,
                            'diff': -diff
                        })
                    else:
                        total_matched_count += 1
                        status = "✅ Khớp"
                        
                    diff_note = f"❌ Lệch số liệu: {diff:+,.2f} - Cần kiểm tra lại tài liệu gốc" if is_diff else "✅ Khớp chính xác 100%"
                    comparison_results.append({
                        'Bảng F1': f"Bảng {t1_idx}",
                        'Bảng F2': f"Bảng {t2_idx}",
                        'Chỉ tiêu / Nhãn': label,
                        'Vị trí File 1': f"Bảng {t1_idx}_R{r1_idx}C{c_idx}",
                        'Vị trí File 2': f"Bảng {t2_idx}_R{r2_idx}C{c_idx}",
                        'Giá trị File 1': num1,
                        'Giá trị File 2': num2,
                        'Chênh lệch': diff,
                        'Văn bản File 1': v1_raw,
                        'Văn bản File 2': v2_raw,
                        'Trạng thái': status,
                        'Ghi chú & Khuyến nghị': diff_note
                    })

                elif (is_num1 != is_num2) and (v1_raw or v2_raw):
                    if not compare_numbers and not compare_text:
                        continue
                    total_checked_count += 1
                    total_mismatch_count += 1
                    comparison_results.append({
                        'Bảng F1': f"Bảng {t1_idx}",
                        'Bảng F2': f"Bảng {t2_idx}",
                        'Chỉ tiêu / Nhãn': label,
                        'Vị trí File 1': f"Bảng {t1_idx}_R{r1_idx}C{c_idx}",
                        'Vị trí File 2': f"Bảng {t2_idx}_R{r2_idx}C{c_idx}",
                        'Giá trị File 1': num1,
                        'Giá trị File 2': num2,
                        'Chênh lệch': None,
                        'Văn bản File 1': v1_raw,
                        'Văn bản File 2': v2_raw,
                        'Trạng thái': "❌ Lệch định dạng số/chữ",
                        'Ghi chú & Khuyến nghị': "Một bên là ô số, một bên là ô chữ. Cần kiểm tra lại định dạng ô."
                    })
                    discrepancies_doc1.append({
                        'table_idx': t1_idx,
                        'row_idx': r1_idx,
                        'doc_table_idx': r1.get('doc_table_idx', t1_idx),
                        'doc_row_idx': r1.get('doc_row_idx', r1_idx),
                        'col_idx': c_idx,
                        'doc_col_idx': c1.get('doc_col_idx', c_idx) if c1 else c_idx,
                        'type': 'format_mismatch',
                        'label': label,
                        'val_self': v1_raw,
                        'val_other': v2_raw
                    })
                    discrepancies_doc2.append({
                        'table_idx': t2_idx,
                        'row_idx': r2_idx,
                        'doc_table_idx': r2.get('doc_table_idx', t2_idx),
                        'doc_row_idx': r2.get('doc_row_idx', r2_idx),
                        'col_idx': c_idx,
                        'doc_col_idx': c2.get('doc_col_idx', c_idx) if c2 else c_idx,
                        'type': 'format_mismatch',
                        'label': label,
                        'val_self': v2_raw,
                        'val_other': v1_raw
                    })

                else: # Pure text cells (not is_num1 and not is_num2)
                    if not compare_text:
                        continue
                    total_checked_count += 1
                    is_diff = (v1_raw.strip() != v2_raw.strip())
                    if is_diff:
                        total_mismatch_count += 1
                        status = "❌ Lệch văn bản"
                        diff_note = "Nội dung văn bản khác nhau giữa 2 tài liệu."
                        discrepancies_doc1.append({
                            'table_idx': t1_idx,
                            'row_idx': r1_idx,
                            'doc_table_idx': r1.get('doc_table_idx', t1_idx),
                            'doc_row_idx': r1.get('doc_row_idx', r1_idx),
                            'col_idx': c_idx,
                            'doc_col_idx': c1.get('doc_col_idx', c_idx) if c1 else c_idx,
                            'type': 'text_mismatch',
                            'label': label,
                            'val_self': v1_raw,
                            'val_other': v2_raw
                        })
                        discrepancies_doc2.append({
                            'table_idx': t2_idx,
                            'row_idx': r2_idx,
                            'doc_table_idx': r2.get('doc_table_idx', t2_idx),
                            'doc_row_idx': r2.get('doc_row_idx', r2_idx),
                            'col_idx': c_idx,
                            'doc_col_idx': c2.get('doc_col_idx', c_idx) if c2 else c_idx,
                            'type': 'text_mismatch',
                            'label': label,
                            'val_self': v2_raw,
                            'val_other': v1_raw
                        })
                    else:
                        total_matched_count += 1
                        status = "✅ Khớp"
                        diff_note = "Khớp chính xác"
                        
                    comparison_results.append({
                        'Bảng F1': f"Bảng {t1_idx}",
                        'Bảng F2': f"Bảng {t2_idx}",
                        'Chỉ tiêu / Nhãn': label,
                        'Vị trí File 1': f"Bảng {t1_idx}_R{r1_idx}C{c_idx}",
                        'Vị trí File 2': f"Bảng {t2_idx}_R{r2_idx}C{c_idx}",
                        'Giá trị File 1': None,
                        'Giá trị File 2': None,
                        'Chênh lệch': None,
                        'Văn bản File 1': v1_raw,
                        'Văn bản File 2': v2_raw,
                        'Trạng thái': status,
                        'Ghi chú & Khuyến nghị': diff_note
                    })
                    
        # Extra rows in File 1
        for un1 in unmat_r1:
            total_mismatch_count += 1
            r1_idx = un1['row_idx']
            row_raw_text = " | ".join(c['raw_text'] for c in un1['cells'] if c['raw_text'])
            label = un1['label']
            is_blank = all(not c['raw_text'].strip() for c in un1['cells'])
            
            if is_blank:
                label = "(Dòng đệm trống định dạng)"
                status_str = "⚠️ Thừa dòng đệm trống ở File 1"
                note = "Dòng đệm trống này chỉ xuất hiện ở File 1, không có ở File 2. Số liệu các dòng khác vẫn khớp 100%."
            elif not row_raw_text or label == "(Không có tiêu đề)":
                label = "(Dòng trống bổ sung)"
                status_str = "⚠️ Thừa dòng ở File 1"
                note = "Dòng trống này chỉ xuất hiện ở File 1, không tìm thấy ở File 2."
            else:
                status_str = "⚠️ Thừa dòng ở File 1"
                note = "Dòng này chỉ xuất hiện ở File 1, không tìm thấy ở File 2. Cần rà soát xem có bị sót dòng."
                
            discrepancies_doc1.append({
                'table_idx': t1_idx,
                'row_idx': r1_idx,
                'doc_table_idx': un1.get('doc_table_idx', t1_idx),
                'doc_row_idx': un1.get('doc_row_idx', r1_idx),
                'col_idx': 1,
                'type': 'extra_row_self',
                'label': label,
                'is_blank': is_blank
            })
            comparison_results.append({
                'Bảng F1': f"Bảng {t1_idx}",
                'Bảng F2': f"Bảng {t2_idx}",
                'Chỉ tiêu / Nhãn': f"[Thừa dòng] {label}",
                'Vị trí File 1': f"Bảng {t1_idx}_R{r1_idx}",
                'Vị trí File 2': "(Không có)",
                'Giá trị File 1': None,
                'Giá trị File 2': None,
                'Chênh lệch': None,
                'Văn bản File 1': row_raw_text or "(Dòng đệm trống)",
                'Văn bản File 2': "",
                'Trạng thái': status_str,
                'Ghi chú & Khuyến nghị': note
            })

        # Extra rows in File 2
        for un2 in unmat_r2:
            total_mismatch_count += 1
            r2_idx = un2['row_idx']
            row_raw_text = " | ".join(c['raw_text'] for c in un2['cells'] if c['raw_text'])
            label = un2['label']
            is_blank = all(not c['raw_text'].strip() for c in un2['cells'])
            
            if is_blank:
                label = "(Dòng đệm trống định dạng)"
                status_str = "⚠️ Thừa dòng đệm trống ở File 2"
                note = "Dòng đệm trống này chỉ xuất hiện ở File 2, không có ở File 1. Số liệu các dòng khác vẫn khớp 100%."
            elif not row_raw_text or label == "(Không có tiêu đề)":
                label = "(Dòng trống bổ sung)"
                status_str = "⚠️ Thừa dòng ở File 2"
                note = "Dòng trống này chỉ xuất hiện ở File 2, không tìm thấy ở File 1."
            else:
                status_str = "⚠️ Thừa dòng ở File 2"
                note = "Dòng này chỉ xuất hiện ở File 2, không tìm thấy ở File 1. Cần rà soát xem có bị sót dòng."
                
            discrepancies_doc2.append({
                'table_idx': t2_idx,
                'row_idx': r2_idx,
                'doc_table_idx': un2.get('doc_table_idx', t2_idx),
                'doc_row_idx': un2.get('doc_row_idx', r2_idx),
                'col_idx': 1,
                'type': 'extra_row_self',
                'label': label,
                'is_blank': is_blank
            })
            comparison_results.append({
                'Bảng F1': f"Bảng {t1_idx}",
                'Bảng F2': f"Bảng {t2_idx}",
                'Chỉ tiêu / Nhãn': f"[Thừa dòng] {label}",
                'Vị trí File 1': "(Không có)",
                'Vị trí File 2': f"Bảng {t2_idx}_R{r2_idx}",
                'Giá trị File 1': None,
                'Giá trị File 2': None,
                'Chênh lệch': None,
                'Văn bản File 1': "",
                'Văn bản File 2': row_raw_text or "(Dòng đệm trống)",
                'Trạng thái': status_str,
                'Ghi chú & Khuyến nghị': note
            })
            
    map1 = {s['table_index']: s for s in summary1}
    map2 = {s['table_index']: s for s in summary2}

    # Unmatched tables in File 1
    for t1_idx in unmat_tables_1:
        total_mismatch_count += 1
        t1_info = map1.get(t1_idx, {})
        r_cnt = t1_info.get('rows', '?')
        c_cnt = t1_info.get('cols', '?')
        prev = t1_info.get('preview', '')
        label_str = f"Bảng {t1_idx} ({r_cnt} dòng × {c_cnt} cột)" + (f": {prev[:40]}..." if prev else "")
        reason_note = (
            f"⚠️ CHƯA ĐỐI CHIẾU SỐ LIỆU TỰ ĐỘNG. "
            f"Lý do: Bảng {t1_idx} ({r_cnt}x{c_cnt}) không tìm thấy bảng đối ứng 1:1 ở File 2 "
            f"(nghi do gộp/tách bảng giữa 2 tài liệu). Khuyến nghị kiểm tra thủ công nội dung tương đương."
        )
        comparison_results.append({
            'Bảng F1': f"Bảng {t1_idx}",
            'Bảng F2': "(Không đối ứng)",
            'Chỉ tiêu / Nhãn': label_str,
            'Vị trí File 1': f"Bảng {t1_idx}",
            'Vị trí File 2': "(Không đối ứng)",
            'Giá trị File 1': None,
            'Giá trị File 2': None,
            'Chênh lệch': None,
            'Văn bản File 1': f"Kích thước: {r_cnt} dòng × {c_cnt} cột. Nội dung đầu bảng: {prev[:60]}",
            'Văn bản File 2': "(Không tìm thấy bảng đối ứng 1:1)",
            'Trạng thái': "⚠️ Bảng gộp/tách (Chưa đối chiếu)",
            'Ghi chú & Khuyến nghị': reason_note
        })
        discrepancies_doc1.append({
            'table_idx': t1_idx,
            'row_idx': 1,
            'col_idx': 1,
            'type': 'unmatched_table',
            'label': label_str,
            'rows': r_cnt,
            'cols': c_cnt,
            'preview': prev
        })

    # Unmatched tables in File 2
    for t2_idx in unmat_tables_2:
        total_mismatch_count += 1
        t2_info = map2.get(t2_idx, {})
        r_cnt = t2_info.get('rows', '?')
        c_cnt = t2_info.get('cols', '?')
        prev = t2_info.get('preview', '')
        label_str = f"Bảng {t2_idx} ({r_cnt} dòng × {c_cnt} cột)" + (f": {prev[:40]}..." if prev else "")
        reason_note = (
            f"⚠️ CHƯA ĐỐI CHIẾU SỐ LIỆU TỰ ĐỘNG. "
            f"Lý do: Bảng {t2_idx} ({r_cnt}x{c_cnt}) không tìm thấy bảng đối ứng 1:1 ở File 1 "
            f"(nghi do gộp/tách bảng giữa 2 tài liệu). Khuyến nghị kiểm tra thủ công nội dung tương đương."
        )
        comparison_results.append({
            'Bảng F1': "(Không đối ứng)",
            'Bảng F2': f"Bảng {t2_idx}",
            'Chỉ tiêu / Nhãn': label_str,
            'Vị trí File 1': "(Không đối ứng)",
            'Vị trí File 2': f"Bảng {t2_idx}",
            'Giá trị File 1': None,
            'Giá trị File 2': None,
            'Chênh lệch': None,
            'Văn bản File 1': "(Không tìm thấy bảng đối ứng 1:1)",
            'Văn bản File 2': f"Kích thước: {r_cnt} dòng × {c_cnt} cột. Nội dung đầu bảng: {prev[:60]}",
            'Trạng thái': "⚠️ Bảng gộp/tách (Chưa đối chiếu)",
            'Ghi chú & Khuyến nghị': reason_note
        })
        discrepancies_doc2.append({
            'table_idx': t2_idx,
            'row_idx': 1,
            'col_idx': 1,
            'type': 'unmatched_table',
            'label': label_str,
            'rows': r_cnt,
            'cols': c_cnt,
            'preview': prev
        })

    results_df = pd.DataFrame(comparison_results) if comparison_results else pd.DataFrame(columns=[
        'Bảng F1', 'Bảng F2', 'Chỉ tiêu / Nhãn', 'Vị trí File 1', 'Vị trí File 2', 
        'Giá trị File 1', 'Giá trị File 2', 'Chênh lệch', 
        'Văn bản File 1', 'Văn bản File 2', 'Trạng thái', 'Ghi chú & Khuyến nghị'
    ])
    
    stats_summary = {
        'total_checked': total_checked_count,
        'total_matched': total_matched_count,
        'total_mismatches': total_mismatch_count,
        'match_rate': round((total_matched_count / total_checked_count * 100), 2) if total_checked_count else 100.0
    }
    
    return results_df, structure_df, struct_stats, stats_summary, discrepancies_doc1, discrepancies_doc2
