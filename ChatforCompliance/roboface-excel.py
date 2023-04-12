import pandas as pd
import re
import os
import sys
from roboface import write_answer

# Determine if the script is running as a compiled executable or as a Python script
if getattr(sys, 'frozen', False):
    # Running as an exe
    excels_dir = os.path.dirname(sys.executable)
    excels_dir = os.path.join(os.path.dirname(excels_dir), "Excels")
else:
    # Running as a script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    parent_dir = os.path.dirname(script_dir)
    excels_dir = os.path.join(parent_dir, 'Excels')


# Define the Excel file name.  Permit the user to enter the file name without the extension.  Provide appropriate error messages if the file is not found.
print(f"Looking for Excel files in: {excels_dir}")
while True:
    excel_filename = input('Please enter the filename of the Excel that you placed in the "Excels" folder: ')
    file_ext = os.path.splitext(excel_filename)[-1].lower()

    if file_ext == "":
        if os.path.isfile(os.path.join(excels_dir, f"{excel_filename}.xlsx")):
            excel_filename += ".xlsx"
            break
        elif os.path.isfile(os.path.join(excels_dir, f"{excel_filename}.xls")):
            excel_filename += ".xls"
            break
        else:
            print("File not found. Please make sure the file is in the 'Excels' folder and try again.")
    elif file_ext in ['.xlsx', '.xls']:
        if os.path.isfile(os.path.join(excels_dir, excel_filename)):
            break
        else:
            print("File not found. Please make sure the file is in the 'Excels' folder and try again.")
    else:
        print("Invalid file extension. Please enter a valid Excel filename.")

source = f"roboface-excel.py, excel_filename={excel_filename}"

excel_filepath = os.path.join(excels_dir, excel_filename)

# Load the Excel file
print("Loading the Excel file...")

try:
    xlsx = pd.read_excel(excel_filepath, sheet_name=None, header=None)
except PermissionError:
    print("Excel file is open in another program. Please close the file and hit 'Y' to continue.")
    while True:
        user_input_qdetect = input("Continue? (Y/N): ")
        if user_input_qdetect.lower() == "y":
            break
        elif user_input_qdetect.lower() == "n":
            exit()
        else:
            print("Please enter Y or N.")
            continue  
print("Excel file loaded.")

#check whether user wishes to use English language question detection regex
while True:
    user_input_qdetect = input("Use English language question detection regex? (Y/N): ")
    if user_input_qdetect.lower() == "y":
        break
    elif user_input_qdetect.lower() == "n":
        break
    else:
        print("Please enter Y or N.")
        continue

# Check if the 'Roboface QRA' sheet exists and remove it if it does
with pd.ExcelWriter(excel_filepath, engine='openpyxl', mode='a') as writer:
    if 'Roboface QRA' in writer.book:
        writer.book.remove(writer.book['Roboface QRA'])

# Define a function to identify questions
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
    
    # Identify cells containing questions and do not process blank cells as questions
    questions = []
    for row in range(df.shape[0]):
        for col in range(df.shape[1]):
            cell_value = df.iloc[row, col]
            if pd.notna(cell_value):
                if user_input_qdetect.lower() == "y":
                    if is_question(str(cell_value)):
                        questions.append((cell_value, row, col))
                else:
                    questions.append((cell_value, row, col))


    # Process the questions and write the questions, answers, and cell locations to the DataFrame
    for index, (question, row, col) in enumerate(questions):
        # write answer is passing a source parameter as well for logging purposes
        answer = write_answer(question, source)
        cell = cell_coordinates(row, col)
        new_row = pd.DataFrame({"Sheet Name": [sheet_name], "Cell": [cell], "Question": [question], "Answer": [answer]})
        
        # Append the new_row to the qa_df DataFrame
        qa_df = pd.concat([qa_df, new_row], ignore_index=True)

        # Display progress information
        remaining_questions = len(questions) - index - 1
        progress_message = f"Processing Question {index + 1} of {len(questions)} ({remaining_questions} remaining)"
        print("\033[31m" + progress_message + "\033[0m")
        print("Current Row:", row, "Column:", col)


    print(f"Finished processing sheet: {sheet_name}.  Overwriting Roboface QRA sheet in Excel file if it exists...")


#check to see if Excel file is open in another program
try:
    with pd.ExcelWriter(excel_filepath, engine='openpyxl', mode='a') as writer:
        qa_df.to_excel(writer, sheet_name='Roboface QRA', index=False)
except PermissionError:
    print("Excel file is open in another program. Please close the file and hit 'Y' to continue.")
    while True:
        user_input = input("Continue? (Y/N): ")
        if user_input.lower() == "y":
            break
        elif user_input.lower() == "n":
            exit()
        else:
            print("Please enter Y or N.")
            continue    
print("Writing questions, answers, and cell locations to a new sheet...")

# Print completion message
print("All questions and answers, along with their cell locations, have been written to a new sheet in the Excel file.  Output has also be recorded to the data.csv file found in the 'config' folder.")