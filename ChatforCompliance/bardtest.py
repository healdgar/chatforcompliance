import openpyxl

excel_file = 'ExcelTest.xlsx'
def extract_questions(excel_file):
    """Extracts all text that would be considered questions from an Excel file.

    Args:
        excel_file (str): The path to the Excel file.

    Returns:
        list: A list of all questions found in the Excel file.
    """
    # Open the Excel file in read-only mode.
    workbook = openpyxl.load_workbook(excel_file)

    # Get the list of all worksheets in the workbook.
    worksheets = workbook.worksheets

    # Iterate through all of the worksheets.
    for worksheet in worksheets:
        # Get the range of cells in the worksheet.
        range = worksheet.range

        # Get the text from all of the cells in the range.
        text = range.text

        # If the range is a merged cell, extract the text from all of the merged cells.
        if range.merged:
            question = text.split(",")

        # If the range is a protected worksheet, extract the text from the unlocked cells.
        if range.protected:
            question = question[question.isdigit()]

        # If the text is not a question, continue.
        if not question.startswith("What is") or not question.startswith("How do"):
            continue

        # Add the question to the list of questions.
        questions.append(question)

    return questions

extract_questions(excel_file)
print(questions)

import openpyxl

excel_file = "my_excel_file.xlsx"

workbook = openpyxl.load_workbook(excel_file)