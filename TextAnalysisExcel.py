# Import necessary libraries
import nltk
from nltk.corpus import stopwords
import spacy
from collections import Counter
import string
import pandas as pd  # Import the pandas library

# Download stopwords data if not already downloaded
nltk.download('stopwords')
stop_words = set(stopwords.words('english'))

# Load spaCy model for lemmatization
nlp = spacy.load('en_core_web_sm')

# Task 1: Remove Stopwords
def remove_stopwords(text):
    # Split the text into words
    words = text.split()
    # Filter out stopwords
    filtered_words = [word for word in words if word.lower() not in stop_words]
    # Reconstruct the text without stopwords
    return ' '.join(filtered_words)

# Task 2: Lemmatization (Returning Verbs to Their Root)
def lemmatize_text(text):
    # Tokenize and lemmatize the text
    doc = nlp(text)
    lemmatized_text = ' '.join([token.lemma_ if token.lemma_ != '-PRON-' else token.text for token in doc])
    return lemmatized_text

# Task 3: Calculate Word Frequencies and Perform Part-of-Speech Tagging
def calculate_word_frequencies(text):
    # Split the text into words
    words = text.split()
    # Calculate word frequencies
    word_freq = Counter(words)
    return word_freq

def perform_pos_tagging(text):
    # Tokenize and perform part-of-speech tagging
    doc = nlp(text)
    pos_tags = [(token.text, token.pos_) for token in doc]
    return pos_tags

# Task 4: Remove Punctuation
def remove_punctuation(text):
    # Remove all punctuation characters except for spaces
    return text.translate(str.maketrans('', '', string.punctuation))

# Task 5: Sort by Decreasing Frequency
def sort_by_frequency(word_freq):
    sorted_word_freq = dict(sorted(word_freq.items(), key=lambda item: item[1], reverse=True))
    return sorted_word_freq

# Task 6: Print Words List with POS and Frequency
def print_word_info(sorted_word_freq, pos_tags, title):
    data = []  # Create a list to store data

    for word, freq in sorted_word_freq.items():
        pos = next((tag for tag in pos_tags if tag[0] == word), ("N/A", "N/A"))
        data.append([word, pos[1], freq])  # Append data to the list

    # Create a pandas DataFrame from the data
    df = pd.DataFrame(data, columns=["Word", "POS", "Frequency"])

    # Sort the DataFrame by POS and then by frequency
    df = df.sort_values(by=["POS", "Frequency"], ascending=[True, False])

    # Specify the file path for the Excel file
    file_path = r""

    # Save the DataFrame to an Excel file with the story title as the file name
    file_name = f"{title}.xlsx"
    excel_file_path = f"{file_path}\\{file_name}"
    df.to_excel(excel_file_path, index=False)

# Example usage:
title = ""
text = ""
text = remove_stopwords(text)
text = lemmatize_text(text)
text = remove_punctuation(text)
word_freq = calculate_word_frequencies(text)
pos_tags = perform_pos_tagging(text)
sorted_word_freq = sort_by_frequency(word_freq)
print_word_info(sorted_word_freq, pos_tags, title)









