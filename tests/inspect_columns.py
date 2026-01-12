import tabula
import pandas as pd
from pathlib import Path

base_dir = Path("data/results/sem2")
file = "EN1054.pdf"
path = base_dir / file

print(f"Processing {path}...")
try:
    # Use stream=True as established
    tables = tabula.read_pdf(path, pages="all", stream=True, pandas_options={'header': None})
    
    print(f"Found {len(tables)} tables.")
    
    for i, df in enumerate(tables):
        print(f"\n--- Table {i} ---")
        print(f"Shape: {df.shape}")
        
        # Simple print of first few rows
        print(df.head())
        
        # Check for non-empty columns
        df_clean = df.dropna(how='all').dropna(axis=1, how='all')
        print(f"Clean Shape: {df_clean.shape}")
        print("Clean snippet:")
        print(df_clean.head())
        
except Exception as e:
    print(f"Error processing {file}: {e}")
