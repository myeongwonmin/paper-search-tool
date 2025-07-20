# PubMed Paper Search Pipeline

A Python tool for automatically searching and collecting research papers from multiple scientific journals using the PubMed API.

## Features

- 🔍 Search papers from 31 major scientific journals
- 📅 Flexible date range selection (specific dates or recent N days)
- 📊 Export results to Excel format
- 🚀 Progress tracking with visual progress bar
- 🛡️ Rate-limited API calls to respect PubMed servers

## Supported Journals

The tool searches across 31 leading journals in life sciences, biotechnology, and related fields:

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

### Date Range Options

1. **Specific Date Range**: Enter start and end dates in YYYY/MM/DD format
2. **Recent Days**: Enter number of recent days to search (e.g., 7 for last week)

### Example
```
--- PubMed Paper Pipeline ---
Select date range mode:
1. Specific date range (YYYY/MM/DD)
2. Recent N days
Enter your choice (1 or 2): 2
Enter number of recent days to search: 7
```

## Output

Results are saved as Excel files in the `output/` directory with the naming format:
```
YYMMDD_YYMMDD_Papers.xlsx
```

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
- Rate-limited to 0.5 seconds between requests
- Searches public database only
- No personal information collected or transmitted

## Contributing

Feel free to submit issues and enhancement requests. To add new journals:
1. Edit the `JOURNAL_LIST` in `config.py`
2. Update the journal count in this README and `manual.txt`

## License

This project is open source and available under the [MIT License](LICENSE).

## Acknowledgments

- Built using the [PubMed E-utilities API](https://www.ncbi.nlm.nih.gov/books/NBK25497/)
- Thanks to NCBI for providing free access to scientific literature data