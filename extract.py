import requests
from bs4 import BeautifulSoup
import openpyxl
import nltk
import re
import pandas as pd
from nltk.tokenize import sent_tokenize
from nltk.sentiment.vader import SentimentIntensityAnalyzer



# Load the Excel file
input_file = "input.xlsx"
wb = openpyxl.load_workbook(input_file)
sheet = wb.active

# Create lists to store data for the output Excel sheet
data = {
    "URL_ID": [],
    "Title": [],
    "Sentimental Analysis": [],
    "Positive Score": [],
    "Negative Score": [],
    "Polarity Score": [],
    "Subjectivity Score": [],
    "Average Sentence Length": [],
    "Percentage of Complex Words": [],
    "Fog Index": [],
    "Average Number of Words Per Sentence": [],
    "Complex Word Count": [],
    "Word Count": [],
    "Syllable Count Per Word": [],
    "Personal Pronouns": [],
    "Average Word Length": [],
}

# Load custom positive and negative word dictionaries
positive_words = set()
negative_words = set()
stop_words = set()

with open("positive_words.txt", "r") as f:
    positive_words = set(f.read().splitlines())

with open("negative_words.txt", "r") as f:
    negative_words = set(f.read().splitlines())

with open("stop_words.txt", "r") as f:
    stop_words = set(f.read().splitlines())

# Function to clean and tokenize text by removing stop words
def clean_and_tokenize(text):
    tokens = nltk.word_tokenize(text)
    cleaned_tokens = [word for word in tokens if word.lower() not in stop_words]
    return cleaned_tokens

# Function to count syllables in a word
def count_syllables(word):
    # Basic syllable counting 
    word = word.lower()
    count = 0
    vowels = "aeiouy"
    if word[0] in vowels:
        count += 1
    for i in range(1, len(word)):
        if word[i] in vowels and word[i - 1] not in vowels:
            count += 1
    if word.endswith(("es", "ed")):
        count -= 1
    if count == 0:
        count = 1  # Ensure at least one syllable for short words
    return count

# Loop through rows in the Excel file to extract data
for row in sheet.iter_rows(min_row=2, values_only=True):
    url_id, article_url = row

    try:
        response = requests.get(article_url)

        if response.status_code == 200:
            soup = BeautifulSoup(response.text, "html.parser")

            # Extract title and article text
            title = soup.find("title").text.strip()
            article_text = ""
            paragraphs = soup.find_all("p")
            for paragraph in paragraphs:
                article_text += paragraph.text.strip() + "\n"

            # Custom Sentiment Analysis with stop words removal
            cleaned_tokens = clean_and_tokenize(article_text)
            positive_score = sum(1 for word in cleaned_tokens if word in positive_words)
            negative_score = sum(1 for word in cleaned_tokens if word in negative_words)
            polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
            subjectivity_score = (positive_score + negative_score) / (len(cleaned_tokens) + 0.000001)

            # Rest of the code for readability analysis
            sentences = sent_tokenize(article_text)
            tokens = nltk.word_tokenize(article_text)
            average_sentence_length = len(tokens) / len(sentences)
            complex_word_count = sum(1 for word in tokens if len(word) > 2)
            percentage_complex_words = (complex_word_count / len(tokens)) * 100
            fog_index = 0.4 * (average_sentence_length + percentage_complex_words)
            average_words_per_sentence = len(tokens) / len(sentences)
            word_count = len(tokens)

            # Count syllables per word
            syllable_count_per_word = [
                count_syllables(word) for word in tokens
            ]

            # Count personal pronouns
            personal_pronouns = len(
                re.findall(r"\b(?:I|we|my|ours|us)\b", article_text, re.IGNORECASE)
            )

            # Calculate average word length
            total_chars = sum(len(word) for word in tokens)
            average_word_length = total_chars / len(tokens)



            # Add data to the dictionary
            data["URL_ID"].append(url_id)
            data["Title"].append(title)
            data["Sentimental Analysis"].append(polarity_score)
            data["Positive Score"].append(positive_score)
            data["Negative Score"].append(negative_score)
            data["Polarity Score"].append(polarity_score)
            data["Subjectivity Score"].append(subjectivity_score)
            data["Average Sentence Length"].append(average_sentence_length)
            data["Percentage of Complex Words"].append(percentage_complex_words)
            data["Fog Index"].append(fog_index)
            data["Average Number of Words Per Sentence"].append(average_words_per_sentence)
            data["Complex Word Count"].append(complex_word_count)
            data["Word Count"].append(word_count)
            data["Syllable Count Per Word"].append(syllable_count_per_word)
            data["Personal Pronouns"].append(personal_pronouns)
            data["Average Word Length"].append(average_word_length)

            print(f"Processed {article_url}")
        else:
            print(f"Failed to retrieve data from {article_url}. Status code: {response.status_code}")

    except Exception as e:
        print(f"Error extracting data from {article_url}: {str(e)}")

# Close the Excel file
wb.close()

# Create a DataFrame from the collected data
df = pd.DataFrame(data)

# Save the DataFrame to an output Excel file
output_file = "output.xlsx"
df.to_excel(output_file, index=False)

print(f"Data analysis completed. Results saved to {output_file}")