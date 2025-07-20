# pubmed_client.py

"""
Client for interacting with the PubMed API (E-utilities).
"""

import requests
import time
from typing import List, Dict, Optional

# E-utilities base URLs
EUTILS_BASE_URL = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/"
ESEARCH_URL = EUTILS_BASE_URL + "esearch.fcgi"
EFETCH_URL = EUTILS_BASE_URL + "efetch.fcgi"

# It's good practice to identify yourself to the NCBI
API_KEY = None  # Using without an API key is fine for low-frequency use
TOOL_NAME = "PaperPipeline"
ADMIN_EMAIL = "your_email@example.com"  # Replace with your email

def search_articles(journal: str, start_date: str, end_date: str) -> List[str]:
    """
    Searches for article IDs in a specific journal within a date range.
    """
    params = {
        "db": "pubmed",
        "term": f'"{journal}"[Journal] AND ("{start_date}"[Date - Publication] : "{end_date}"[Date - Publication])',
        "retmode": "json",
        "retmax": 1000,  # Max results per request
        "tool": TOOL_NAME,
        "email": ADMIN_EMAIL,
    }
    if API_KEY:
        params["api_key"] = API_KEY

    try:
        response = requests.get(ESEARCH_URL, params=params)
        response.raise_for_status()  # Raise an exception for bad status codes
        data = response.json()
        return data.get("esearchresult", {}).get("idlist", [])
    except requests.exceptions.RequestException as e:
        print(f"Error during API search for {journal}: {e}")
        return []

def fetch_article_details(article_ids: List[str]) -> Optional[Dict]:
    """
    Fetches detailed information for a list of article IDs.
    """
    if not article_ids:
        return None

    params = {
        "db": "pubmed",
        "id": ",".join(article_ids),
        "retmode": "xml",
        "tool": TOOL_NAME,
        "email": ADMIN_EMAIL,
    }
    if API_KEY:
        params["api_key"] = API_KEY

    try:
        # Use POST for a large number of IDs
        response = requests.post(EFETCH_URL, data=params)
        response.raise_for_status()
        return response.text  # Return XML content as text
    except requests.exceptions.RequestException as e:
        print(f"Error fetching article details: {e}")
        return None
