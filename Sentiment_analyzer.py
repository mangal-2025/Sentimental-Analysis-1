import pandas as pd
from nltk.sentiment.vader import SentimentIntensityAnalyzer
import nltk

# --- A. Setup ---

# Download the VADER lexicon if it hasn't been downloaded yet
try:
    nltk.data.find('sentiment/vader_lexicon.zip')
except:
    nltk.download('vader_lexicon')

# Initialize the VADER analyzer
sid = SentimentIntensityAnalyzer()

# Define the name of your Excel file and the text column
INPUT_FILE = 'CRM_v2.xlsx'
TEXT_COLUMN = 'Requirement' # Match this to the column name in your Excel file
OUTPUT_FILE = 'feedback_analyzed.xlsx'

print(f"Starting analysis on {INPUT_FILE}...")


# --- B. The Core Analysis Function ---

def get_sentiment(text):
    """Calculates the VADER sentiment scores for a given text."""
    # Handle cases where the cell might be empty or non-string
    if pd.isna(text) or not isinstance(text, str):
        return 0, 'NEUTRAL 😐' # Return a default neutral score

    # Get the polarity scores
    scores = sid.polarity_scores(text)
    compound_score = scores['compound']
    
    # Determine the overall sentiment label based on the compound score
    if compound_score >= 0.05:
        sentiment = "POSITIVE 😊"
    elif compound_score <= -0.05:
        sentiment = "NEGATIVE 😠"
    else:
        sentiment = "NEUTRAL 😐"
        
    return compound_score, sentiment


# --- C. Read, Process, and Write ---

try:
    # 1. Read the Excel file into a pandas DataFrame (like a table)
    df = pd.read_excel(INPUT_FILE)

    # 2. Apply the analysis function to every row in the specified column
    # The .apply() method runs the get_sentiment function on every value in the 'Review_Text' column.
    # It returns two new lists (Compound Score and Sentiment Label).
    df[['VADER_Compound_Score', 'Sentiment_Label']] = df[TEXT_COLUMN].apply(
        lambda x: pd.Series(get_sentiment(x))
    )

    # 3. Save the resulting DataFrame (which now includes the new columns) back to a new Excel file
    df.to_excel(OUTPUT_FILE, index=False) # index=False prevents writing the internal row numbers

    print(f"\n✅ Analysis complete!")
    print(f"Results saved to {OUTPUT_FILE}")

except FileNotFoundError:
    print(f"\n❌ ERROR: Could not find the file named '{INPUT_FILE}'.")
    print("Please make sure the Excel file is in the same folder as the Python script.")