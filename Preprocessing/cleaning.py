import pandas as pd
import re
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from nltk.stem import WordNetLemmatizer
import nltk
from contractions import (
    fix,
)

# Download required NLTK data
nltk.download("punkt")
nltk.download("stopwords")
nltk.download("wordnet")
nltk.download("omw-1.4")

# Load the dataset
data = pd.read_csv("IMDB_Review_Dataset.csv", names=["review"], encoding="utf-8")


# Define the cleaning function
def clean_text(text):
    # Expand contractions
    text = fix(text)
    # Convert to lowercase
    text = text.lower()
    # Remove HTML tags
    text = re.sub(r"<.*?>", "", text)
    # Remove special characters, numbers, and punctuation
    text = re.sub(r"[^a-zA-Z\s]", "", text)
    # Remove extra whitespace
    text = re.sub(r"\s+", " ", text).strip()
    return text


# Apply cleaning to the reviews
data["Cleaned_Review"] = data["review"].apply(clean_text)

# Tokenize the reviews
data["Tokenized_Review"] = data["Cleaned_Review"].apply(word_tokenize)

# Remove stop words
stop_words = set(stopwords.words("english"))
data["Filtered_Review"] = data["Tokenized_Review"].apply(
    lambda x: [word for word in x if word not in stop_words]
)

# Lemmatize the words
lemmatizer = WordNetLemmatizer()
data["Lemmatized_Review"] = data["Filtered_Review"].apply(
    lambda x: [lemmatizer.lemmatize(word) for word in x]
)

# Join the tokens back into a single string
data["Final_Review"] = data["Lemmatized_Review"].apply(lambda x: " ".join(x))

# Save the cleaned dataset to a new CSV file
data[["Final_Review"]].to_csv("Cleaned_IMDB_Reviews.csv", index=False)

print("Data cleaning complete. Cleaned dataset saved to  Cleaned_IMDB_Reviews.csv.")
