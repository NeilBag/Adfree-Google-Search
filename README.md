# Google Search App

## Overview

This is a Flask web application that allows users to perform Google searches, display the results in a table, download the results as an Excel spreadsheet, and email the results with an Excel attachment.

## Features

- Perform a Google search and display up to 500 results.
- Display search results in a formatted table.
- Download search results as an Excel spreadsheet.
- Email search results with the Excel attachment.

## Project Structure

google_search_app/
├── app.py
├── templates/
│ ├── index.html
│ └── results.html
├── static/
│ └── css/
│ └── style.css
├── requirements.txt
└── README.md

## Requirements

- Python 3.7+
- Flask
- requests
- BeautifulSoup4
- pandas
- openpyxl
- xlsxwriter

## Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/neilbag/google_search_app.git
   cd google_search_app