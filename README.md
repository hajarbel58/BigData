# Big Data Sentiment Analysis using MapReduce

## Objective
Perform sentiment analysis on IMDB reviews using the MapReduce programming model to classify sentiments as Positive, Negative, or Neutral.

---

## Prerequisites
- Basic understanding of MapReduce
- Python (for Hadoop streaming) or access to a Hadoop cluster

---

## Dataset
**File**: `IMDB_Review_Dataset.csv`

- Each line is a review with optional sentiment label
- Example sentiments: positive, negative, neutral

---

## Project Structure
```
Project Directory
│
├── Dictionary/       
├── MapReduce/      
├── Preprocessing/          
├── results/
├── IMDB reviews project report.pdf            
└── Project_BigData_Sentiments_Analysis_Tasks.pdf 
```

---

## Steps

### 1. Sentiment Keyword Dictionary
- **Positive**: love, amazing, happy, good, excellent, wonderful, etc.
- **Negative**: horrible, terrible, bad, disappointing, regret, etc.

### 2. Preprocessing
- Remove stop words (e.g., the, is, and)
- Strip special characters and punctuation

### 3. MapReduce Program
- **Mapper**: Tokenizes text and tags with sentiment keywords
- **Reducer**: Aggregates and classifies sentiment counts

### 4. Upload Dataset to HDFS
```bash
hdfs dfs -put IMDB_Review_Dataset.csv /user/yourname/input
```

### 5. Run MapReduce Job
```bash
hadoop jar /path/to/hadoop-streaming.jar \
  -input /user/yourname/input \
  -output /user/yourname/output \
  -mapper Mapper.py \
  -reducer Reducer.py
```

### 6. Fetch Results
```bash
hdfs dfs -get /user/yourname/output results/
```

### 7. Visualize Results
Use any of the following:
- Python + Matplotlib
- Excel
- Tableau

---

## Final Sentiment Counts
| Sentiment | Count  |
|-----------|--------|
| Positive  | 28389  |
| Neutral   | 11988  |
| Negative  | 9623   |

### Analysis
1. **Dominant Sentiment**: Positive
2. **Implication**: Indicates a generally favorable perception of content on IMDB. This could imply higher platform trustworthiness and positive consumer sentiment.

---

## License
This project is for educational purposes only.

