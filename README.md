# Data Extraction and NLP - Text Analyzer

A Python tool for analyzing text content from URLs, calculating various readability metrics and sentiment scores.

## Features

- Extracts article content from URLs
- Calculates sentiment scores (positive, negative, polarity, subjectivity)
- Computes readability metrics (fog index, complex words, etc.)
- Processes multiple URLs from an Excel input file
- Outputs analysis results to Excel

## Requirements

- Python 3.6+
- Required packages:
  - pandas
  - requests
  - beautifulsoup4
  - nltk
  - syllables

## Setup

1. Clone the repository:
```
git clone https://github.com/yourusername/text-analyzer.git
cd text-analyzer
```

2. Install requirements:
```
pip install -r requirements.txt
```

3. Prepare input files:
   - Place your `Input.xlsx` file in the project directory
   - Ensure `positive-words.txt` and `negative-words.txt` are present

## Usage

Run the script:
```
python text_analysis.py
```

The script will process URLs from `Input.xlsx` and generate `Output Data Structure.xlsx` with analysis results.

## Project Structure

- `text_analysis.py`: Main script
- `requirements.txt`: List of Python dependencies
- `positive-words.txt`: Dictionary of positive words for sentiment analysis
- `negative-words.txt`: Dictionary of negative words for sentiment analysis
