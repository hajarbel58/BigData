import matplotlib.pyplot as plt
import seaborn as sns

# Load the sentiment results
sentiment_counts = {}
with open("sentiment_results.txt", "r") as f:
    for line in f:
        sentiment, count = line.strip().split("\t")
        sentiment_counts[sentiment] = int(count)

# Improved Visualization
sns.set_theme(style="whitegrid")  # Use a clean grid style
plt.figure(figsize=(10, 6))  # Set figure size

# Create the bar plot
colors = sns.color_palette("pastel", len(sentiment_counts))  # Use pastel colors
bars = plt.bar(
    sentiment_counts.keys(), sentiment_counts.values(), color=colors, edgecolor="black"
)

# Add value annotations on top of the bars
for bar in bars:
    height = bar.get_height()
    plt.text(
        bar.get_x() + bar.get_width() / 2,
        height,
        f"{height}",
        ha="center",
        va="bottom",
        fontsize=12,
        color="black",
    )

# Customize labels and title
plt.xlabel("Sentiment", fontsize=14)
plt.ylabel("Count", fontsize=14)
plt.title("IMDB Reviews Sentiment Analysis", fontsize=16, fontweight="bold")
plt.xticks(fontsize=12)
plt.yticks(fontsize=12)

# Add grid lines
plt.grid(axis="y", linestyle="--", alpha=0.7)

# Show the plot
plt.tight_layout()
plt.show()
