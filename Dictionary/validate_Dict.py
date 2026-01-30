from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer
import json

# import keywords from extended_sentiment_keywords.json

with open("extended_sentiment_keywords.json", "r") as file:
    keywords = json.load(file)

# Initialize VADER sentiment analyzer
analyzer = SentimentIntensityAnalyzer()

# Validate words
validated_keywords = {"positive": [], "negative": []}
for category, words in keywords.items():
    for word in words:
        sentiment_score = analyzer.polarity_scores(word)["compound"]
        if (category == "positive" and sentiment_score > 0.1) or (
            category == "negative" and sentiment_score < -0.1
        ):
            validated_keywords[category].append(word)

# Save validated keywords
with open("keywords.json", "w") as file:
    json.dump(validated_keywords, file)
