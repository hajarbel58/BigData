import sys

sentiment_counts = {}

# Reducer function
for line in sys.stdin:
    sentiment, count = line.strip().split("\t")
    count = int(count)
    if sentiment in sentiment_counts:
        sentiment_counts[sentiment] += count
    else:
        sentiment_counts[sentiment] = count

for sentiment, count in sentiment_counts.items():
    print(f"{sentiment}\t{count}")
