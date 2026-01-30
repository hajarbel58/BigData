from nltk.corpus import wordnet
import json

# Define the keywords dictionary
keywords = {
    "positive": [
        "love",
        "glad",
        "amazing",
        "happy",
        "good",
        "excellent",
        "wonderful",
        "fantastic",
        "brilliant",
        "satisfied",
        "success",
    ],
    "negative": [
        "horrible",
        "frustrating",
        "terrible",
        "bad",
        "worst",
        "disappointing",
        "awful",
        "regret",
    ],
}


# Function to find synonyms
def get_synonyms(word):
    synonyms = set()
    for syn in wordnet.synsets(word):
        for lemma in syn.lemmas():
            synonyms.add(lemma.name())
    return list(synonyms)


# Extend the keywords
for category, words in keywords.items():
    expanded = []
    for word in words:
        expanded.extend(get_synonyms(word))
    keywords[category] = list(set(expanded))  # Remove duplicates


def clean_keywords(words):
    clean_words = []
    for word in words:
        clean_words.append(word.replace("_", " "))  # Replace underscores with spaces
    return clean_words


# Apply to all categories
for category, words in keywords.items():
    keywords[category] = clean_keywords(words)

# Save the cleaned dictionary
with open("extended_sentiment_keywords.json", "w") as file:
    json.dump(keywords, file)
