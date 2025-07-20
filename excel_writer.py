# excel_writer.py

"""
Writes paper data and summary to an Excel file with multiple sheets.
Applies formatting for better readability.
"""

import pandas as pd
import os
from datetime import datetime
from typing import List, Dict
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

OUTPUT_DIR = "output"
# Define column widths for readability
COLUMN_WIDTHS = {
    'Title': 60,
    'Journal': 25,
    'Published Date': 15,
    'Authors': 40,
    'Abstract': 80,
    'URL': 15,
}

def write_to_excel(papers_data: List[Dict], start_date: str, end_date: str):
    """
    Writes the collected paper data and a summary to an Excel file.
    """
    if not papers_data:
        print("No data to write. Excel file will not be created.")
        return

    # Ensure output directory exists
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    # Prepare data for DataFrame
    df = pd.DataFrame(papers_data)
    
    # --- Create Hyperlinks ---
    # Create a display-friendly URL column before converting to formula
    df['URL_Display'] = df['URL']
    df['URL'] = df['URL'].apply(
        lambda x: f'=HYPERLINK("{x}", "Link")' if pd.notna(x) and "http" in x else ""
    )


    # --- Create Summary Data ---
    collection_period = f"{start_date.replace('/', '-')} to {end_date.replace('/', '-')}"
    total_papers = len(df)
    collection_timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    journal_counts = df["Journal"].value_counts().reset_index()
    journal_counts.columns = ["Journal Name", "Number of Papers"]

    # --- Write to Excel ---
    start_date_str = start_date.replace('/', '')[2:]
    end_date_str = end_date.replace('/', '')[2:]
    file_name = f"{start_date_str}_{end_date_str}_Papers.xlsx"
    file_path = os.path.join(OUTPUT_DIR, file_name)

    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        # --- Summary Sheet ---
        summary_df = pd.DataFrame({
            "Summary Metric": ["Collection Period", "Total Papers", "Collection Timestamp"],
            "Value": [collection_period, total_papers, collection_timestamp]
        })
        summary_df.to_excel(writer, sheet_name="Summary", index=False, header=False, startrow=1)
        journal_counts.to_excel(writer, sheet_name="Summary", index=False, startrow=len(summary_df) + 3)
        
        # --- Papers Sheet ---
        # Write the main data, keeping the original URL for display if needed
        df[['Title', 'Journal', 'Published Date', 'Authors', 'Abstract', 'URL']].to_excel(
            writer, sheet_name="Papers", index=False
        )

        # --- Apply Formatting ---
        worksheet_papers = writer.sheets["Papers"]
        worksheet_summary = writer.sheets["Summary"]

        # Set column widths and text wrapping for Papers sheet
        for i, col_name in enumerate(df[['Title', 'Journal', 'Published Date', 'Authors', 'Abstract', 'URL']].columns, 1):
            col_letter = get_column_letter(i)
            width = COLUMN_WIDTHS.get(col_name, 20) # Default width 20
            worksheet_papers.column_dimensions[col_letter].width = width
            
            # Apply text wrapping to Title and Abstract
            if col_name in ["Title", "Abstract"]:
                for cell in worksheet_papers[col_letter]:
                    cell.alignment = Alignment(wrap_text=True, vertical='top')

        # Set column widths for Summary sheet
        worksheet_summary.column_dimensions['A'].width = 25
        worksheet_summary.column_dimensions['B'].width = 40


    print("\nSuccessfully created Excel file: {}".format(file_path))