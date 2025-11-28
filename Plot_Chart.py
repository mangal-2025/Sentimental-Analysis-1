import pandas as pd
import matplotlib.pyplot as plt

# --- 1. Define file paths ---
INPUT_FILE = 'feedback_analyzed.xlsx'
CHART_FILENAME = 'sentiment_pie_chart.png'

# --- 2. Load the file and prepare data for charting ---
try:
    # Load the output file
    df = pd.read_excel(INPUT_FILE)
    
    # Count the occurrences of each sentiment label
    sentiment_counts = df['Sentiment_Label'].value_counts()
    
    # Prepare data for plotting
    labels = sentiment_counts.index
    sizes = sentiment_counts.values
    
    # Define colors for consistency (Green for Positive, Red for Negative, Grey for Neutral)
    label_to_color = {
        'POSITIVE 😊': 'tab:green',
        'NEGATIVE 😠': 'tab:red',
        'NEUTRAL 😐': 'tab:gray'
    }
    # Ensure colors match the labels in the correct order
    colors = [label_to_color[label] for label in labels]

    # --- 3. Create the Pie Chart ---
    plt.figure(figsize=(8, 8))
    plt.pie(
        sizes, 
        labels=labels, 
        colors=colors, 
        autopct='%1.1f%%', # Display percentages with one decimal place
        startangle=90,     # Start the first slice at the top
        wedgeprops={'edgecolor': 'black'}, # Adds a border to slices
        textprops={'fontsize': 12}
    )

    # Add a title and ensure the pie is a circle
    plt.title('Overall Sentiment Distribution of Feedback', fontsize=16)
    plt.axis('equal') 

    # Save the chart to a file
    plt.savefig(CHART_FILENAME)
    plt.close()

    print(f"\nChart saved successfully as: {CHART_FILENAME}")

except FileNotFoundError:
    print(f"\n❌ ERROR: Could not find the file named '{INPUT_FILE}'.")
    print("Please make sure the 'feedback_analyzed.xlsx' file is in the same folder.")
except KeyError:
    print(f"\n❌ ERROR: The column 'Sentiment_Label' was not found in the Excel file.")
    print("Ensure your analysis script ran correctly and created the output file.")