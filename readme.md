My First RAG: "Roboface"
In 2023, I created a RAG that effectively vectorized data in a CSV and wrote answers to compliance questions into an Excel file. By use of careful prompting, grounding, and context control, its answers were acceptably accurate 97% of the time.
========

Description
-----------

Roboface is a script for finding the most relevant answers to a given query using OpenAI's GPT-4 model. It leverages semantic search and a pre-existing Question-Response-Answer (QRA) dataset to generate responses. Roboface version 3-24-2023 includes multi-lingual semantic QRA matching, similarity threshold setting, and improved logging.

Dependencies
------------

-   Python 3.7+
-   openai
-   openpyxl
-   configparser
-   requests
-   numpy
-   scikit-learn
-   tqdm
-   glob

Setup
-----

1.  Install the required Python packages using pip:

    perlCopy code

    `pip install openai openpyxl configparser requests numpy scikit-learn tqdm glob`

2.  Update `config.ini` and `secrets.ini` with the appropriate settings and API keys.
3.  Ensure the QRA dataset (in .xlsx format) is placed in the `config` directory.

Usage
-----

To run the script, execute the following command in your terminal:

Copy code

`python roboface.py`

Enter a search term when prompted. The script will then find the most relevant answers to the query using semantic search with the QRA dataset and GPT-4.

You can ask additional questions by entering 'Y' when prompted or exit the script by entering 'N'.

Key Functions
-------------

-   `generate_embeddings_file(csv_list, filename)`: Generates an embeddings file for the QRA dataset.
-   `get_parameters(...)`: Returns a dictionary of parameters used to log the response generation process.
-   `count_tokens(text)`: Counts the number of tokens in a given text string.
-   `truncate_prior_context(prior_questions, max_tokens=context_length)`: Truncates the prior questions to no more than the specified maximum tokens.
-   `write_answer(searchTerm, source="direct Roboface query")`: Writes the generated answer to a CSV file and updates the prior context.

Configuration
-------------

Settings for the script can be configured in `config.ini`, including temperature, max tokens, model, answer library prefix, number of matches, and context length.

The `secrets.ini` file should contain the API key for OpenAI.

Output
------

The script generates a CSV file named `data.csv` in the `config` directory, containing the timestamp, user, model, temperature, max response size, search term, similarity threshold, matching context, tuning context, tokens, answer, and source.

Roboface Excel Question-Answer Automation
=========================================

This Python script automates the process of identifying questions in an Excel file, answering them using the Roboface API, and saving the answers to a new sheet in the same Excel file.

Usage
-----

1.  Place your Excel file in the `Excels` folder located in the same directory as the script.
2.  Run the script.
3.  When prompted, enter the name of your Excel file (without the extension).
4.  Choose whether to use the English language question detection regex by entering "Y" or "N".
5.  The script will process each sheet in the Excel file, identify questions, and answer them using the Roboface API.
6.  The answers will be written to a new sheet called "Roboface QRA" in the same Excel file.

Important Notes
---------------

-   The script supports both `.xlsx` and `.xls` file formats.
-   The script uses the `roboface.write_answer` function to obtain answers from the Roboface API.
-   The script uses regex for English language question detection if the user opts for it. This can be disabled by entering "N" when prompted.
-   If the "Roboface QRA" sheet already exists in the Excel file, it will be overwritten.
-   Ensure that the Excel file is closed before running the script to avoid `PermissionError`.

Dependencies
------------

-   pandas
-   openpyxl
-   re
-   os
-   sys
-   roboface

Script Details
--------------

The script first checks if it's running as a compiled executable or as a Python script, and determines the correct path to the `Excels` folder accordingly. It then prompts the user for the Excel file name and checks if it exists in the `Excels` folder.

The script reads the Excel file using `pandas` and processes each sheet to identify questions based on the user's choice of using English language question detection regex. If the user opts for question detection, the script uses the `is_question` function, which contains a regex pattern to identify questions in English.

The script creates a DataFrame to store the questions, answers, and cell locations. For each question, it calls the `write_answer` function from the `roboface` module to get the answer and saves it to the DataFrame. Finally, the script writes the DataFrame to a new sheet called "Roboface QRA" in the Excel file.

The script provides progress information during the processing of the sheets, indicating the current row and column being processed, and the number of remaining questions.
