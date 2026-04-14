import pandas as pd

def compare_dataframes(df1, df2, mode='Number'):
    """
    Compares two dataframes of extracted word table data.
    """
    if df1.empty and df2.empty:
        return pd.DataFrame(), "RESULT: NO DATA FOUND", 0
        
    # Rename columns for clarity
    df1_c = df1.copy().rename(columns={'Value': 'Value 1', 'Raw': 'Text 1'})
    df2_c = df2.copy().rename(columns={'Value': 'Value 2', 'Raw': 'Text 2'})
    
    # Merge on Table and Address
    merged = pd.merge(
        df1_c[['Table', 'Address', 'Value 1', 'Text 1']], 
        df2_c[['Table', 'Address', 'Value 2', 'Text 2']], 
        on=['Table', 'Address'], 
        how='outer'
    )
    
    merged = merged.sort_values(by=['Table', 'Address'])
    
    if mode == 'Number':
        # Fill NaNs with 0 for calculation
        v1 = merged['Value 1'].fillna(0)
        v2 = merged['Value 2'].fillna(0)
        merged['Diff'] = v1 - v2
        diff_mask = merged['Diff'].abs() > 1e-6
    else:
        # For Non-Number mode, we just compare the raw text
        # Fill NaNs with empty string
        t1 = merged['Text 1'].fillna("")
        t2 = merged['Text 2'].fillna("")
        # We can add a 'Diff' column as a boolean or string indicator for text
        merged['Diff'] = 0 # Dummy for consistency in UI columns if needed
        diff_mask = (t1 != t2)
    
    mismatches = diff_mask.sum()
    
    if mismatches == 0:
        result_msg = "✅ VERIFICATION SUCCESSFUL: No discrepancies found. All data points are consistent."
    elif mode == 'Number':
        result_msg = f"⚠️ ATTENTION REQUIRED: {mismatches} numerical discrepancies identified. Please review the details below."
    else:
        result_msg = f"🔍 AUDIT ALERT: {mismatches} text mismatches detected in table headers or labels."
        
    return merged, result_msg, mismatches
