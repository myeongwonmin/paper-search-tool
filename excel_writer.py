# excel_writer.py

"""
Writes paper data and summary to an Excel file with multiple sheets.
Applies formatting for better readability.
"""

import pandas as pd
import os
import re
from datetime import datetime
from typing import List, Dict
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from openpyxl.cell.rich_text import TextBlock, CellRichText
from openpyxl.cell.text import InlineFont

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

def filter_papers_by_keyword(papers_data: List[Dict], keyword: str) -> List[Dict]:
    """
    Filters papers that contain the keyword in their title (case-insensitive).
    """
    filtered_papers = []
    for paper in papers_data:
        title = paper.get("Title", "")
        if keyword.lower() in title.lower():
            filtered_papers.append(paper.copy())
    return filtered_papers

def create_rich_text_with_keyword_highlighting(text: str, keyword: str) -> CellRichText:
    """
    Creates rich text with only the keyword highlighted in red and bold.
    Handles multiple occurrences and case-insensitive matching.
    """
    if not text or not keyword or text == "N/A":
        return CellRichText(text or "")
    
    # Create fonts
    normal_font = InlineFont()
    highlight_font = InlineFont(color="FF0000", b=True)
    
    # Find all occurrences of keyword (case-insensitive)
    rich_text = CellRichText()
    text_lower = text.lower()
    keyword_lower = keyword.lower()
    
    last_end = 0
    start = 0
    
    while True:
        # Find next occurrence of keyword
        start = text_lower.find(keyword_lower, start)
        if start == -1:
            break
            
        # Add text before keyword
        if start > last_end:
            rich_text.append(TextBlock(normal_font, text[last_end:start]))
        
        # Add highlighted keyword
        end = start + len(keyword)
        rich_text.append(TextBlock(highlight_font, text[start:end]))
        
        last_end = end
        start = end
    
    # Add remaining text after last keyword
    if last_end < len(text):
        rich_text.append(TextBlock(normal_font, text[last_end:]))
    
    return rich_text

def write_to_excel(papers_data: List[Dict], start_date: str, end_date: str, keywords: List[str] = None):
    """
    Writes the collected paper data and a summary to an Excel file.
    Creates additional sheets for each keyword if keywords are provided.
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

        # --- Keyword-specific Sheets ---
        if keywords:
            for keyword in keywords:
                filtered_papers = filter_papers_by_keyword(papers_data, keyword)
                if filtered_papers:
                    print(f"Found {len(filtered_papers)} papers containing '{keyword}' in title.")
                    
                    # Create DataFrame for filtered papers
                    keyword_df = pd.DataFrame(filtered_papers)
                    
                    # Create hyperlinks for URLs
                    keyword_df['URL'] = keyword_df['URL'].apply(
                        lambda x: f'=HYPERLINK("{x}", "Link")' if pd.notna(x) and "http" in x else ""
                    )
                    
                    # Write to sheet with keyword name
                    sheet_name = f"Keyword={keyword}"
                    keyword_df[['Title', 'Journal', 'Published Date', 'Authors', 'Abstract', 'URL']].to_excel(
                        writer, sheet_name=sheet_name, index=False
                    )
                else:
                    print(f"No papers found containing '{keyword}' in title.")

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

        # Apply formatting to keyword sheets
        if keywords:
            for keyword in keywords:
                sheet_name = f"Keyword={keyword}"
                if sheet_name in writer.sheets:
                    worksheet_keyword = writer.sheets[sheet_name]
                    
                    # Set column widths and text wrapping for keyword sheets (same as Papers sheet)
                    for i, col_name in enumerate(['Title', 'Journal', 'Published Date', 'Authors', 'Abstract', 'URL'], 1):
                        col_letter = get_column_letter(i)
                        width = COLUMN_WIDTHS.get(col_name, 20) # Default width 20
                        worksheet_keyword.column_dimensions[col_letter].width = width
                        
                        # Apply text wrapping to Title and Abstract
                        if col_name in ["Title", "Abstract"]:
                            for cell in worksheet_keyword[col_letter]:
                                cell.alignment = Alignment(wrap_text=True, vertical='top')
                    
                    # Apply keyword highlighting to titles and abstracts
                    for i, col_name in enumerate(['Title', 'Journal', 'Published Date', 'Authors', 'Abstract', 'URL'], 1):
                        if col_name in ["Title", "Abstract"]:
                            col_letter = get_column_letter(i)
                            for row_idx, cell in enumerate(worksheet_keyword[col_letter], 1):
                                if row_idx > 1 and cell.value:  # Skip header row
                                    # Create rich text with keyword highlighting
                                    rich_text = create_rich_text_with_keyword_highlighting(str(cell.value), keyword)
                                    cell.value = rich_text


    print("\nSuccessfully created Excel file: {}".format(file_path))