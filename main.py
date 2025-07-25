# main.py

"""
Main execution script for the PubMed Paper Pipeline.
Orchestrates the process of searching, fetching, extracting, and writing paper data.
"""

import time
from datetime import datetime, timedelta
from tqdm import tqdm

# Import project modules
from config import JOURNAL_LIST
import pubmed_client
import data_extractor
import excel_writer

def get_date_range():
    """Gets the date range from the user."""
    print("Select date range mode:")
    print("1. Specific date range (YYYY/MM/DD)")
    print("2. Recent N days")

    while True:
        choice = input("Enter your choice (1 or 2): ")
        if choice in ["1", "2"]:
            break
        print("Invalid choice. Please enter 1 or 2.")

    if choice == "1":
        while True:
            try:
                start_date_str = input("Enter start date (YYYY/MM/DD): ")
                datetime.strptime(start_date_str, "%Y/%m/%d")
                end_date_str = input("Enter end date (YYYY/MM/DD): ")
                datetime.strptime(end_date_str, "%Y/%m/%d")
                return start_date_str, end_date_str
            except ValueError:
                print("Invalid date format. Please use YYYY/MM/DD.")
    else: # choice == "2"
        while True:
            try:
                days = int(input("Enter number of recent days to search: "))
                if days > 0:
                    end_date = datetime.now()
                    start_date = end_date - timedelta(days=days)
                    return start_date.strftime("%Y/%m/%d"), end_date.strftime("%Y/%m/%d")
                print("Please enter a positive number.")
            except ValueError:
                print("Invalid input. Please enter a number.")

def get_keywords():
    """Gets keywords from the user for filtering papers by title."""
    print("\nEnter keywords to filter papers by title.")
    print("You can enter multiple keywords separated by commas for separate sheets (e.g., enzyme, e. coli, deep learning)")
    print("Use + to connect keywords for OR logic in a single sheet (e.g., Alphafold+ESMfold)")
    print("You can mix both: enzyme, protein+fold, ML creates 3 sheets")
    print("Leave empty to skip keyword filtering.")
    
    keywords_input = input("Enter keywords: ").strip()
    
    if not keywords_input:
        return []
    
    # Split by comma and clean up whitespace
    keywords = [keyword.strip() for keyword in keywords_input.split(",") if keyword.strip()]
    
    if keywords:
        # Display parsed keywords with their types
        display_keywords = []
        for keyword in keywords:
            if '+' in keyword:
                sub_keywords = [k.strip() for k in keyword.split('+')]
                display_keywords.append(f"{keyword} (OR: {', '.join(sub_keywords)})")
            else:
                display_keywords.append(keyword)
        print(f"Keywords to search for: {'; '.join(display_keywords)}")
    
    return keywords

def main():
    """Main function to run the pipeline."""
    print("--- PubMed Paper Pipeline ---")
    start_date, end_date = get_date_range()
    keywords = get_keywords()
    print("Searching for papers from {} to {}.".format(start_date, end_date))
    print()

    all_papers = []
    # Use tqdm for a progress bar
    for journal in tqdm(JOURNAL_LIST, desc="Processing Journals"):
        # Search for article IDs
        article_ids = pubmed_client.search_articles(journal, start_date, end_date)
        
        if article_ids:
            # Fetch details for these IDs
            xml_data = pubmed_client.fetch_article_details(article_ids)
            if xml_data:
                # Extract structured info from XML
                papers = data_extractor.extract_paper_info(xml_data)
                all_papers.extend(papers)
        
        # Be respectful to the API server - increased to 1 second for safety margin
        time.sleep(1.0)

    if all_papers:
        print("\nFound a total of {} papers.".format(len(all_papers)))
        excel_writer.write_to_excel(all_papers, start_date, end_date, keywords)
    else:
        print("\nNo papers found for the given criteria.")

if __name__ == "__main__":
    main()
