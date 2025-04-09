from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from nltk.stem import PorterStemmer
from collections import Counter
import string
import nltk

nltk.download('punkt')
nltk.download('stopwords')

def preprocess_text(text):
    """Normalize and clean text"""
    text = text.lower().translate(str.maketrans('', '', string.punctuation))
    tokens = word_tokenize(text)
    stop_words = set(stopwords.words('english'))
    ps = PorterStemmer()
    return [ps.stem(w) for w in tokens if w not in stop_words]

def check_plagiarism(input_text, existing_texts):
    """Check for plagiarism with similarity score"""
    input_words = Counter(preprocess_text(input_text))
    results = []
    
    for text in existing_texts:
        text_words = Counter(preprocess_text(text))
        common = sum((input_words & text_words).values())
        total = sum((input_words | text_words).values())
        similarity = common / total if total > 0 else 0
        results.append({
            'text': text,
            'similarity': similarity
        })
    
    max_similarity = max(results, key=lambda x: x['similarity'], default={'similarity': 0})
    return {
        'is_plagiarized': max_similarity['similarity'] > 0.3,
        'score': max_similarity['similarity'],
        'matches': results
    }