import sys
import json

# Load the sentiment keywords
keywords = {
    "positive": [
        "brilliant",
        "awesome",
        "magnificent",
        "good",
        "beneficial",
        "fantastic",
        "secure",
        "fulfill",
        "astound",
        "make love",
        "happy",
        "beloved",
        "wonderful",
        "making love",
        "honest",
        "fantastical",
        "glorious",
        "satisfied",
        "winner",
        "effective",
        "bright",
        "dear",
        "glad",
        "splendid",
        "trade good",
        "marvelous",
        "amazing",
        "love",
        "sexual love",
        "smart as a whip",
        "erotic love",
        "love life",
        "enjoy",
        "well",
        "gratify",
        "succeeder",
        "dearest",
        "passion",
        "success",
        "goodness",
        "superb",
        "amaze",
        "grand",
        "satisfy",
        "terrific",
        "excellent",
        "respectable",
        "safe",
        "honorable",
    ],
    "negative": [
        "horrific",
        "thwarting",
        "frustrating",
        "fearsome",
        "ugly",
        "dread",
        "dreadful",
        "terrible",
        "terribly",
        "sorrow",
        "dire",
        "fearful",
        "disappointing",
        "dissatisfactory",
        "regret",
        "horrendous",
        "disappoint",
        "nasty",
        "worst",
        "frightful",
        "awful",
        "painful",
        "wicked",
        "frightfully",
        "frustrate",
        "tough",
        "dreaded",
        "defective",
        "horrifying",
        "badly",
        "horrible",
        "severe",
        "regretful",
        "direful",
        "frightening",
        "risky",
        "bad",
    ],
}

positive_keywords = set(keywords["positive"])
negative_keywords = set(keywords["negative"])

# Mapper function
for line in sys.stdin:
    review = line.strip().lower()
    positive_count = sum(1 for word in review.split() if word in positive_keywords)
    negative_count = sum(1 for word in review.split() if word in negative_keywords)

    if positive_count > negative_count:
        sentiment = "positive"
    elif negative_count > positive_count:
        sentiment = "negative"
    else:
        sentiment = "neutral"

    print(f"{sentiment}\t1")
