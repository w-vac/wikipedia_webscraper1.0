# Wikipedia Web Scraper

This Python project scrapes Wikipedia articles starting from a user-provided URL (or a random article if no URL is provided). The program recursively follows links to other Wikipedia articles, saves the titles and URLs of the visited pages, and exports the data to CSV and Excel files.

## Features

- Recursively scrapes Wikipedia articles.
- Stores the titles and URLs of visited pages.
- Saves data in both CSV and Excel formats.
- Allows the user to start from a specific Wikipedia page or a random one.
- Automatically adjusts column widths in the Excel file for better readability.

## Requirements

- Python 3.x
- `requests`
- `beautifulsoup4`
- `pandas`
- `openpyxl`

## Installation & Running

1. Clone this repository or download the project files.

```
git clone https://github.com/yourusername/wiki-web-scraper.git
```

2. Navigate to the project directory
```
cd wikipedia_webscraper1.0
```

3. Install the required Python packages by running:

```
pip install -r requirements.txt
```

4. Running the project

```
python scraper1.0py
```
## How it Works

1. The scraper starts from a Wikipedia page and scrapes the content.
2. It follows a random valid Wikipedia link from the current page.
3. The process repeats until interrupted or when no more valid links are found.
4. The visited pages' titles and URLs are saved to CSV and Excel files.
5. A completion message is printed, and the user is given the option to open the CSV and Excel files.


