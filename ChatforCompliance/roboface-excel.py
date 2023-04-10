import pandas as pd
import re
import os
from roboface import write_answer

# Define the Excel file name
excel_filename = input('Please enter the filename of the Excel that you placed in the "Excels" folder: ')

# Construct the full file path
script_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(script_dir)
excels_dir = os.path.join(parent_dir, 'Excels')
excel_filepath = os.path.join(excels_dir, excel_filename)

# Load the Excel file
print("Loading the Excel file...")
xlsx = pd.read_excel(excel_filepath, sheet_name=None, header=None)

def is_question(sentence):
    pattern = r'\b(what|where|when|why|how|which|who|whom|whose)\b.*|\byou(r)?\b.*|\?.*|\bplease\b.*|\b(provide|describe)\b.*|Name of '
    match = re.search(pattern, sentence, re.MULTILINE)
    if match:
        return True
    else:
        return False

def cell_coordinates(row, col):
    column_letter = chr(ord('A') + col)
    return f"{column_letter}{row + 1}"

# Create a DataFrame to store questions, answers, and cell locations
qa_df = pd.DataFrame(columns=["Sheet Name", "Cell", "Question", "Answer"])

# Loop through each sheet
print("Processing sheets...")
for sheet_name, df in xlsx.items():
    print(f"Processing sheet: {sheet_name}")
    
    # Identify cells containing questions
    questions = []
    for row in range(df.shape[0]):
        for col in range(df.shape[1]):
            if pd.notna(df.iloc[row, col]) and is_question(str(df.iloc[row, col])):
                questions.append((df.iloc[row, col], row, col))

    # Process the questions and write the questions, answers, and cell locations to the DataFrame
    for question, row, col in questions:
        # Replace 'fake_write_answers' with the function you use to generate answers
        answer = write_answer(question)
        cell = cell_coordinates(row, col)
        new_row = pd.DataFrame({"Sheet Name": [sheet_name], "Cell": [cell], "Question": [question], "Answer": [answer]})
        qa_df = pd.concat([qa_df, new_row], ignore_index=True)        
    
    print(f"Finished processing sheet: {sheet_name}")

print("Writing questions, answers, and cell locations to a new sheet...")

# Write the DataFrame to a new sheet in the Excel file
with pd.ExcelWriter(excel_filepath, engine='openpyxl', mode='a') as writer:
    qa_df.to_excel(writer, sheet_name='Roboface QRA', index=False)

# Print completion message
print("All questions and answers, along with their cell locations, have been written to a new sheet in the Excel file.  Output has also be recorded to the data.csv file found in the 'config' folder.")
