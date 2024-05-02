# Hacker News Scraper

This Node.js script scrapes the top 10 articles from Hacker News and allows users to output the data in either CSV format or an Excel file with optional AI-generated summaries for each article.

## Features

- Web Scraping: Extracts the top 10 articles from Hacker News using Playwright.
  Output Formats: Choose between CSV or Excel outputs.
- AI Summaries: Optional AI-generated summaries for each article when opting for Excel output.
- Interactive CLI: Interactive command-line prompts to guide through the process.

## Prerequisites

- Node.js
- npm (Node Package Manager)

## Installation

First, clone the repository to your local machine and navigate into the project directory. Then, install the necessary dependencies.

npm install

## Usage

To run the script, use the following command in the terminal:

node index.js

Once the script is running, it will prompt you to choose whether you want a summary for each article. Type 'Y' for yes or 'N' for no, and press Enter.

### CSV Output

If you choose not to receive summaries, the script will scrape the articles and directly write the data into a CSV file named hacker-news-top-10-articles.csv.

### Excel Output

If you opt for summaries, the script will:

1. Scrape the articles.
2. Write initial data into a CSV file.
3. Convert this CSV to an Excel file.
4. Scrape AI-generated summaries for each article.
5. Update the Excel file with these summaries.

The final Excel file will be saved as hacker-news-top-10-articles.xlsx.

### Important Notes

- The script uses a headless browser to fetch and process web content. Ensure your firewall or network settings allow such operations.
- AI summaries are fetched using an external service (SummarAIse, another project of mine), its quality or availability may vary.
