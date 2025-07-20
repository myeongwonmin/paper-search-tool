# data_extractor.py

"""
Extracts and parses paper information from PubMed's XML data.
"""

import xml.etree.ElementTree as ET
from typing import List, Dict

def get_full_text(element):
    """
    Recursively concatenates text from an element and all its children.
    This handles cases where text is split by formatting tags like <i> or <b>.
    """
    text = element.text or ""
    for child in element:
        text += get_full_text(child)
        if child.tail:
            text += child.tail
    return text.strip()

def extract_paper_info(xml_data: str) -> List[Dict]:
    """
    Parses the XML response from PubMed and extracts relevant paper details.
    """
    papers = []
    if not xml_data:
        return papers

    try:
        root = ET.fromstring(xml_data)
        for article in root.findall(".//PubmedArticle"):
            paper_info = {}

            # Title
            title_element = article.find(".//ArticleTitle")
            paper_info["Title"] = get_full_text(title_element) if title_element is not None else "N/A"

            # Journal
            journal_element = article.find(".//ISOAbbreviation")
            paper_info["Journal"] = journal_element.text if journal_element is not None else "N/A"

            # Published Date
            pub_date = article.find(".//PubDate")
            year = pub_date.findtext("Year", "N/A")
            month = pub_date.findtext("Month", "01").zfill(2)
            day = pub_date.findtext("Day", "01").zfill(2)
            paper_info["Published Date"] = f"{year}-{month}-{day}"

            # Authors
            author_list = []
            for author in article.findall(".//Author"):
                lastname = author.findtext("LastName", "")
                forename = author.findtext("ForeName", "")
                if lastname:
                    author_list.append(f"{lastname}, {forename}")
            paper_info["Authors"] = ", ".join(author_list) if author_list else "N/A"

            # Abstract
            abstract_element = article.find(".//Abstract")
            # Abstract can also have complex structures
            if abstract_element is not None:
                abstract_parts = []
                for part in abstract_element.findall(".//AbstractText"):
                    abstract_parts.append(get_full_text(part))
                paper_info["Abstract"] = " ".join(abstract_parts)
            else:
                paper_info["Abstract"] = "N/A"


            # URL
            pmid_element = article.find(".//PMID")
            if pmid_element is not None:
                pmid = pmid_element.text
                paper_info["URL"] = f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/"
            else:
                paper_info["URL"] = "N/A"

            papers.append(paper_info)
    except ET.ParseError as e:
        print(f"Error parsing XML data: {e}")

    return papers