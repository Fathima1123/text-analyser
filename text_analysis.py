import pandas as pd
import requests
from bs4 import BeautifulSoup
import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.corpus import stopwords
import re
import os
import syllables

# Download required NLTK data
nltk.download('punkt')
nltk.download('stopwords')
nltk.download('averaged_perceptron_tagger')

class TextAnalyzer:
    def __init__(self):
        # Load positive and negative word lists
        with open('positive-words.txt', 'r') as file:
            self.positive_words = set(file.read().splitlines())
        with open('negative-words.txt', 'r') as file:
            self.negative_words = set(file.read().splitlines())
        
        self.stop_words = set(stopwords.words('english'))
        self.personal_pronouns = set(['i', 'me', 'my', 'mine', 'we', 'us', 'our', 'ours'])
    
    def extract_article(self, url):
        try:
            response = requests.get(url)
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # Find article title and text (adjust selectors based on website structure)
            title = soup.find('h1').text.strip() if soup.find('h1') else ""
            article_body = soup.find('article')
            
            if article_body:
                paragraphs = article_body.find_all('p')
                text = ' '.join([p.text for p in paragraphs])
                return f"{title}\n\n{text}"
            return None
        except Exception as e:
            print(f"Error extracting from {url}: {str(e)}")
            return None
    
    def clean_text(self, text):
        # Remove special characters and convert to lowercase
        text = re.sub(r'[^\w\s]', '', text.lower())
        return text
    
    def get_word_count(self, text):
        words = word_tokenize(text)
        return len([word for word in words if word.isalnum()])
    
    def get_complex_word_count(self, text):
        try:
            words = word_tokenize(self.clean_text(text))
            return sum(1 for word in words if syllables.estimate(word) >= 3)
        except Exception as e:
            print(f"Error in complex word count: {e}")
            return 0
    
    def get_personal_pronoun_count(self, text):
        words = word_tokenize(self.clean_text(text))
        return sum(1 for word in words if word in self.personal_pronouns)
    
    def calculate_sentiment_scores(self, text):
        words = word_tokenize(self.clean_text(text))
        words = [word for word in words if word not in self.stop_words]
        
        positive_score = sum(1 for word in words if word in self.positive_words)
        negative_score = sum(1 for word in words if word in self.negative_words)
        
        polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
        subjectivity_score = (positive_score + negative_score) / (len(words) + 0.000001)
        
        return positive_score, negative_score, polarity_score, subjectivity_score
    
    def calculate_readability_metrics(self, text):
        try:
            sentences = sent_tokenize(text)
            words = word_tokenize(text)
            word_count = len([word for word in words if word.isalnum()])
            
            if not sentences or word_count == 0:
                return self._get_default_metrics()
            
            avg_sentence_length = word_count / len(sentences)
            complex_word_count = self.get_complex_word_count(text)
            percent_complex_words = (complex_word_count / word_count * 100) if word_count > 0 else 0
            
            fog_index = 0.4 * (avg_sentence_length + percent_complex_words)
            
            syllable_count = sum(syllables.estimate(word) for word in words if word.isalnum())
            syllable_per_word = syllable_count / word_count if word_count > 0 else 0
            
            avg_word_length = sum(len(word) for word in words if word.isalnum()) / word_count if word_count > 0 else 0
            
            return {
                'avg_sentence_length': avg_sentence_length,
                'percent_complex_words': percent_complex_words,
                'fog_index': fog_index,
                'complex_word_count': complex_word_count,
                'word_count': word_count,
                'syllable_per_word': syllable_per_word,
                'avg_word_length': avg_word_length
            }
        except Exception as e:
            print(f"Error in readability metrics: {e}")
            return self._get_default_metrics()
    
    def _get_default_metrics(self):
        return {
            'avg_sentence_length': 0,
            'percent_complex_words': 0,
            'fog_index': 0,
            'complex_word_count': 0,
            'word_count': 0,
            'syllable_per_word': 0,
            'avg_word_length': 0
        }
    
    def analyze_text(self, text):
        if not text:
            return None
        
        try:
            pos_score, neg_score, polarity, subjectivity = self.calculate_sentiment_scores(text)
            readability_metrics = self.calculate_readability_metrics(text)
            personal_pronouns = self.get_personal_pronoun_count(text)
            
            return {
                'POSITIVE SCORE': pos_score,
                'NEGATIVE SCORE': neg_score,
                'POLARITY SCORE': polarity,
                'SUBJECTIVITY SCORE': subjectivity,
                'AVG SENTENCE LENGTH': readability_metrics['avg_sentence_length'],
                'PERCENTAGE OF COMPLEX WORDS': readability_metrics['percent_complex_words'],
                'FOG INDEX': readability_metrics['fog_index'],
                'AVG NUMBER OF WORDS PER SENTENCE': readability_metrics['avg_sentence_length'],
                'COMPLEX WORD COUNT': readability_metrics['complex_word_count'],
                'WORD COUNT': readability_metrics['word_count'],
                'SYLLABLE PER WORD': readability_metrics['syllable_per_word'],
                'PERSONAL PRONOUNS': personal_pronouns,
                'AVG WORD LENGTH': readability_metrics['avg_word_length']
            }
        except Exception as e:
            print(f"Error in text analysis: {e}")
            return None

def main():
    try:
        # Load input Excel file
        input_df = pd.read_excel('Input.xlsx')
        
        analyzer = TextAnalyzer()
        results = []
        
        # Create directory for extracted articles
        os.makedirs('extracted_articles', exist_ok=True)
        
        for index, row in input_df.iterrows():
            url_id = row['URL_ID']
            url = row['URL']
            
            print(f"Processing URL {url_id}: {url}")
            
            # Extract article
            article_text = analyzer.extract_article(url)
            
            if article_text:
                # Save extracted text
                with open(f'extracted_articles/{url_id}.txt', 'w', encoding='utf-8') as f:
                    f.write(article_text)
                
                # Analyze text
                analysis_results = analyzer.analyze_text(article_text)
                
                if analysis_results:
                    result_row = row.to_dict()
                    result_row.update(analysis_results)
                    results.append(result_row)
            else:
                print(f"Failed to extract text from URL {url_id}")
        
        # Create output DataFrame and save to Excel
        if results:
            output_df = pd.DataFrame(results)
            output_df.to_excel('Output Data Structure.xlsx', index=False)
            print("Analysis completed successfully!")
        else:
            print("No results were generated.")
            
    except Exception as e:
        print(f"An error occurred in main: {e}")

if __name__ == "__main__":
    main()