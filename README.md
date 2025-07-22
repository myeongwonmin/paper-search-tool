# PubMed Paper Search Pipeline

A Python tool for automatically searching and collecting research papers from multiple scientific journals using the PubMed API.

## Features

- 🔍 Search papers from 39 major scientific journals
- 📅 Flexible date range selection (specific dates or recent N days)
- 🔑 **NEW**: Keyword filtering with automatic highlighting
- 📊 Export results to Excel format with multiple sheets
- 🚀 Progress tracking with visual progress bar
- 🛡️ Rate-limited API calls to respect PubMed servers

## Supported Journals

The tool searches across 39 leading journals in life sciences, biotechnology, and related fields:

- **Nature Family**: Nature, Nature Biotechnology, Nature Methods, Nature Communications, etc.
- **Cell Family**: Cell, Cell Systems, Cell Reports, Cell Chemical Biology
- **Science**: Science, Science Advances
- **Specialized Journals**: Bioinformatics, PNAS, Protein Science, Chemical Science
- **Review Journals**: Trends in Biotechnology, Annual Review of Microbiology

*See [config.py](config.py) for the complete list.*

## Installation

1. **Clone the repository:**
   ```bash
   git clone <repository-url>
   cd paper_search
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

## ⚠️ Required Configuration Before Use

**IMPORTANT**: Before running the tool, you must configure your unique identifiers to avoid conflicts with other users:

1. **Edit `pubmed_client.py`:**
   ```python
   TOOL_NAME = "YourName_PaperSearch"    # Change to your unique name
   ADMIN_EMAIL = "your.email@domain.com" # Change to your email
   ```

2. **Examples:**
   ```python
   TOOL_NAME = "John_PaperSearch"
   ADMIN_EMAIL = "john.doe@university.edu"
   ```
   ```python
   TOOL_NAME = "Lab2024_Papers"
   ADMIN_EMAIL = "researcher@lab.org"
   ```

**Why this is important:**
- Each user needs a unique tool identifier to avoid NCBI usage conflicts
- Provides proper identification to NCBI as required by their usage guidelines
- Ensures your usage is independent from other users of this tool

## Usage

Run the main script:
```bash
python main.py
```

### Input Options

#### 1. Date Range Selection
- **Specific Date Range**: Enter start and end dates in YYYY/MM/DD format
- **Recent Days**: Enter number of recent days to search (e.g., 7 for last week)

#### 2. Keyword Filtering (New Feature!)
After selecting dates, you can enter keywords to create filtered sheets:
- Enter multiple keywords separated by commas
- Case-insensitive matching ("enzyme" matches "Enzyme", "ENZYME")
- Partial word matching ("enzyme" matches "enzymes", "enzymatic")
- Leave empty to skip keyword filtering

### Example Usage
```
--- PubMed Paper Pipeline ---
Select date range mode:
1. Specific date range (YYYY/MM/DD)
2. Recent N days
Enter your choice (1 or 2): 2
Enter number of recent days to search: 7

Enter keywords to filter papers by title.
You can enter multiple keywords separated by commas (e.g., enzyme, e. coli, deep learning)
Leave empty to skip keyword filtering.
Enter keywords: enzyme, machine learning, CRISPR
Keywords to search for: enzyme, machine learning, CRISPR
```

## Output

Results are saved as Excel files in the `output/` directory with the naming format:
```
YYMMDD_YYMMDD_Papers.xlsx
```

### Excel File Structure

Each Excel file contains multiple sheets:

1. **Summary Sheet**: Collection statistics and journal counts
2. **Papers Sheet**: All collected papers with complete information
3. **Keyword Sheets** (if keywords provided): 
   - Sheet name format: `Keyword=enzyme`, `Keyword=machine learning`
   - Contains only papers with the keyword in their title
   - **Keyword Highlighting**: Keywords in both Title and Abstract columns are highlighted in **red and bold**
   - Supports multiple keyword occurrences in the same text

### Example Output Sheets
```
📄 250719_250722_Papers.xlsx
├── Summary              # Collection statistics
├── Papers               # All 156 papers found
├── Keyword=enzyme       # 12 papers containing "enzyme"
├── Keyword=machine learning # 8 papers containing "machine learning"
└── Keyword=CRISPR       # 5 papers containing "CRISPR"
```

### Keyword Highlighting Features
- **Smart Matching**: Case-insensitive and partial word matching
- **Visual Highlighting**: Keywords appear in **red and bold** in Excel
- **Multiple Occurrences**: All keyword instances in the same cell are highlighted
- **Both Columns**: Highlighting applied to both Title and Abstract columns

## Project Structure

```
paper_search/
├── main.py              # Main execution script
├── config.py            # Journal list configuration
├── pubmed_client.py     # PubMed API client
├── data_extractor.py    # Paper information extraction
├── excel_writer.py      # Excel file generation
├── requirements.txt     # Python dependencies
├── manual.txt          # Detailed user manual (Korean)
├── README.md           # This file
└── output/             # Generated Excel files
```

## Requirements

- Python 3.6+
- Internet connection
- Dependencies listed in `requirements.txt`:
  - requests
  - pandas
  - openpyxl
  - tqdm

## API Information

This tool uses the free PubMed E-utilities API:
- No API key required
- Rate-limited to 1.0 seconds between requests for safety margin
- Searches public database only
- No personal information collected or transmitted

## Contributing

Feel free to submit issues and enhancement requests. To add new journals:
1. Edit the `JOURNAL_LIST` in `config.py`
2. Update the journal count in this README

## License

This project is open source and available under the [MIT License](LICENSE).

## Acknowledgments

- Built using the [PubMed E-utilities API](https://www.ncbi.nlm.nih.gov/books/NBK25497/)
- Thanks to NCBI for providing free access to scientific literature data