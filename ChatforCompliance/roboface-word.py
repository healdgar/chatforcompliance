import pandas as pd
import re
import os
from roboface import write_answer
import docx
import sys
import os

# Define the Word file name
word_filename = input('Please enter the filename of the Word document that you placed in the "Docs" folder: ')

# Get the directory containing the script or executable
current_dir = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))

# Construct the full file path
parent_dir = os.path.dirname(current_dir)
docs_dir = os.path.join(parent_dir, 'Docs')
word_filepath = os.path.join(docs_dir, word_filename)

# Load the Word document
print("Loading the Word document...")
doc = docx.Document(word_filepath)

def is_question(sentence):
    pattern = r'\b(what|where|when|why|how|which|who|whom|whose)\b.*|\byou(r)?\b.*|\?.*|\bplease\b.*|\b(provide|describe)\b.*|Name of '
    match = re.search(pattern, sentence, re.MULTILINE)
    if match:
        return True
    else:
        return False

# Create a DataFrame to store questions, answers, and paragraph numbers
qa_df = pd.DataFrame(columns=["Paragraph Number", "Question", "Answer"])

# Loop through each paragraph
print("Processing paragraphs...")
for i, para in enumerate(doc.paragraphs):
    print(f"Processing paragraph {i}")
    
    # Identify paragraphs containing questions
    if is_question(para.text):
        print(f"Calling write_answer with question: {para.text}")  # Add this line for debugging
        # Replace 'write_answer' with the function you use to generate answers
        answer = write_answer(para.text)
        print(f"Received answer: {answer}")  # Add this line for debugging
        new_row = pd.DataFrame({"Paragraph Number": [i], "Question": [para.text], "Answer": [answer]})
        qa_df = pd.concat([qa_df, new_row], ignore_index=True)

    
    print(f"Finished processing paragraph {i}")

print("Writing questions, answers, and paragraph numbers to a new Excel file...")

# Get the base name of the Word file without the extension
word_basename = os.path.splitext(os.path.basename(word_filepath))[0]

# Create the new Excel file path
excel_filename = f"{word_basename}_qa.xlsx"
excel_dir = os.path.join(parent_dir, 'Excels')
excel_filepath = os.path.join(excel_dir, excel_filename)

# Write the DataFrame to a new sheet in the Excel file using the openpyxl engine
with pd.ExcelWriter(excel_filepath, engine='openpyxl') as writer:
    qa_df.to_excel(writer, sheet_name='Roboface QRA', index=False)

# Print completion message
print(f"All questions and answers, along with their paragraph numbers, have been written to a new Excel file '{excel_filename}'.")
