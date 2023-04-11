import pandas as pd
import nltk
from QRAtoChat import write_answer
import random

# Read the Excel file and extract the text
df = pd.read_excel('ExcelTest.xlsx')
text = ' '.join(df.iloc[:, 0].tolist())

# Tokenize the text into sentences
sentences = nltk.sent_tokenize(text)

# Identify requests for information
requests = []
for sent in sentences:
    tokens = nltk.word_tokenize(sent)
    pos_tags = nltk.pos_tag(tokens)
    for i in range(len(pos_tags)-1):
        if (pos_tags[i][1].startswith('VB') or pos_tags[i][1].startswith('MD')) and pos_tags[i+1][0].lower() == 'you' or 'your':
            requests.append(sent)

# Extract the information surrounding each request and write the answer to the appropriate location in the Excel file
answered_questions = set()

for r in requests:
    if r not in answered_questions:
        # Extract the information surrounding the request
        before = text.split(r)[0].split()[-5:]
        after = text.split(r)[1].split()[:5]
        info = ' '.join(before + after)

        # Use info to generate an answer to the request using a function or API
        answer = write_answer(answers)

        # Find the row index and column index of the question in the Excel file
        # Find the row index and column index of the question in the Excel file
        question = r
        for i in range(len(df.index)):
            for j in range(len(df.columns)):
                cell_value = str(df.iloc[i, j])
                if question in cell_value:
                    question_row = i
                    question_col = j
                    break


        # Write the answer to the appropriate location in the Excel file
        # Find the first blank cell in the row
        for i in range(question_col + 1, len(df.columns)):
            if pd.isna(df.iloc[question_row, i]):
                df.iloc[question_row, i] = answer
                break

        # Save the updated Excel file
        df.to_excel('ExcelTest.xlsx', sheet_name='Sheet1', index=False)

        # Add the question to the set of answered questions
        answered_questions.add(r)
