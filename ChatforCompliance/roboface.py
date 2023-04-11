import openai
import sys
import csv
import requests
import getpass
import glob
import os
import openpyxl
import configparser
import json
import pickle
import numpy as np
from openai.embeddings_utils import get_embedding
from sklearn.metrics.pairwise import cosine_similarity
import sklearn.metrics._pairwise_distances_reduction._datasets_pair
import sklearn.metrics._pairwise_distances_reduction._middle_term_computer
from tqdm import tqdm
from datetime import datetime
from itertools import islice

# Roboface version 3-24-2023: adds multi-lingual semantic QRA matching, similarity threshold setting, and better logging
# Determine if the script is running as a compiled executable or as a Python script
if getattr(sys, 'frozen', False):
    # Running as a compiled executable
    script_dir = os.path.dirname(sys.executable)
else:
    # Running as a regular Python script
    script_dir = os.path.dirname(os.path.abspath(__file__))

# Construct the paths to the config directory (one level above the script or executable)
config_dir = os.path.join(script_dir, '..', 'config')

# Read the configuration files
config = configparser.ConfigParser()
config.read(os.path.join(config_dir, 'config.ini'))
secrets = configparser.ConfigParser()
secrets.read(os.path.join(config_dir, 'secrets.ini'))

# Get the API key from the secrets.ini file
api_key = secrets.get('auth', 'api_key')  # Change 'auth' to the correct section name if needed

# Set the API key for the openai library
openai.api_key = api_key

#initialize prior questions and answers
prior_questions = []
prior_answers = []

# Get parameters from the config file
temperature = float(config.get('parameters', 'temperature'))
max_tokens = int(config.get('parameters', 'max_tokens'))
model = config.get('parameters', 'model')
answer_library_prefix = config.get('filepaths', 'answer_library_prefix')
num_matches = int(config.get('parameters', 'num_matches'))
context_length = int(config.get('parameters', 'context_length'))

# Get config directory path
config_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'config')


def generate_embeddings_file(csv_list, filename):
    print(f"Embedding text in Client Answer Library ({filename})")
    embeddings_list = []
    
    # Get the API key from the secrets.ini file
    api_key = secrets.get('auth', 'api_key')
    
    # Update the OpenAI package configuration with the API key
    openai.api_key = api_key
    
    for i, row in enumerate(csv_list):
        line = (row['Question'] or '') + ' ' + (row['Answer'] or '') + ' ' + (row['Supporting Answer'] or '')
        line = line.replace('Rimini Street', 'ACME')  # Replace "Rimini Street" with "ACME"
        line = line.replace('RSI', 'ACME')
        line_embeddings = get_embedding(line, engine='text-embedding-ada-002')
        embeddings_list.append(line_embeddings)
        print(f"Progress: {i + 1}/{len(csv_list)}")
    with open(filename, 'wb') as f:
        pickle.dump(embeddings_list, f)
    print("Embeddings file created.")


# Find the most recently modified file that matches the filename pattern
pattern = answer_library_prefix + '*'
suffix = '.xlsx'
files = glob.glob(os.path.join(config_dir, pattern) + suffix)
if len(files) == 0:
    print('Error: no files found for pattern', os.path.join(config_dir, pattern) + suffix)

filename = max(files, key=os.path.getmtime)
print('using QRA file: ', filename)

# Load the Excel file
wb = openpyxl.load_workbook(filename)
ws = wb.active

header = [cell.value for cell in ws[1]]
csv_list = []
for row_index, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
    row_dict = {'ExcelRow': row_index}  # Store the Excel row number
    for i, cell_value in enumerate(row):
        row_dict[header[i]] = cell_value
    csv_list.append(row_dict)

# Check if embeddings file exists; if not, generate it
embeddings_filename = os.path.join(config_dir, f'embeddings_{os.path.basename(filename)}.pickle')
if not os.path.exists(embeddings_filename):
    generate_embeddings_file(csv_list, embeddings_filename)

# Load the embeddings file
with open(embeddings_filename, 'rb') as f:
    embeddings_list = pickle.load(f)

def get_parameters(timestamp, user, baseGPTmodel, creativity, max_tokens, searchTerm, similarity_threshold, matching_context, tuning_context, top_n_matches, tokens, answer):
    return {
        'timestamp': timestamp,
        'user': user,
        'baseGPTmodel': baseGPTmodel,
        'creativity': creativity,
        'max_tokens': max_tokens,
        'searchTerm': searchTerm,
        'similarity_threshold': similarity_threshold,
        'matching_context': matching_context,
        'tuning_context': tuning_context,
        'top_n_matches': top_n_matches,
        'tokens': tokens,
        'answer': answer
    }

def count_tokens(text):
    return len(text.split())

def truncate_prior_context(prior_questions, max_tokens=context_length):
    truncated_prior_questions = []
    current_tokens = 0

    for question in reversed(prior_questions):
        question_tokens = count_tokens(question)
        if current_tokens + question_tokens <= max_tokens:
            truncated_prior_questions.insert(0, question)
            current_tokens += question_tokens
        else:
            break

    return truncated_prior_questions

def write_answer(searchTerm):
    #variables
    answers = ""
    timestamp = datetime.utcnow().isoformat()
    user = getpass.getuser()
    tuning_context = []
    matching_context = []
    global prior_questions
    global prior_answers

    # Truncate prior context to no more than 1000 tokens
    prior_questions = truncate_prior_context(prior_questions)
    prior_answers = truncate_prior_context(prior_answers)

    # Get embeddings for the search term
    engine = "text-embedding-ada-002"  # Specify the same engine used in generate_embeddings_file
    search_embeddings = get_embedding(searchTerm, engine=engine)

     # Load similarity threshold from the configuration file
    similarity_threshold = float(config.get('parameters', 'similarity_threshold'))
        
    # Calculate the similarity of each row to the search term using embeddings
    matches = []
    for i, row in enumerate(csv_list):
        line_embeddings = embeddings_list[i]
        if line_embeddings is not None:
            # Convert lists to NumPy arrays and then reshape
            search_embeddings_array = np.array(search_embeddings).reshape(1, -1)
            line_embeddings_array = np.array(line_embeddings).reshape(1, -1)
            
            # Calculate cosine similarity
            similarity = cosine_similarity(search_embeddings_array, line_embeddings_array)[0, 0]
            matches.append({'Row': row, 'Similarity': similarity})

    # Sort matches based on similarity score in descending order
    matches.sort(key=lambda x: x['Similarity'], reverse=True)

    # Filter matches based on similarity threshold
    filtered_matches = [match for match in matches if match['Similarity'] >= similarity_threshold]

    # Check if there are any matches that meet the threshold
    if not filtered_matches:
        # Indicate that no relevant responses were found
        matching_context = [f"No relevant content meeting the similarity threshold {similarity_threshold} were found in QRA."]
        tuning_context = [f"We have no specific information relating to this query."]
        top_n_matches = []  # Set top_n_matches to an empty list
    else:
        # Select the top n matches, where n is specified in the config file
        top_n_matches = filtered_matches[:num_matches]


        # Select the top n matches, where n is specified in the config file
    top_n_matches = filtered_matches[:num_matches]

    # Create tuning context using the top n matches and create display context

    for i, match in enumerate(top_n_matches):
        row = match['Row']
        excel_row_number = row['ExcelRow']  # Extract Excel row number from the match
        similarity_score = match['Similarity']  # Get similarity score

        # Concatenate all columns used to calculate similarity ('Question', 'Answer', 'Supporting Answer')
        full_text = (row['Question'] or '') + ' ' + (row['Answer'] or '') + ' ' + (row['Supporting Answer'] or '')
        full_text = full_text.replace('Rimini Street', 'ACME')  # Replace "Rimini Street" with "ACME"
        full_text = full_text.replace('RSI', 'ACME')

        # Construct context entry for display with Excel row number, similarity score, and full text
        display_entry = f"CAL Excel Row {excel_row_number} (Similarity: {similarity_score:.4f}): {full_text}"
        matching_context.append(display_entry)

        # Add only 'Answer' and 'Supporting Answer' to tuning context (no row number, similarity score, or 'Question')
        answer = (row['Answer'] or '') + ' ' + (row['Supporting Answer'] or '')
        answer = answer.replace('Rimini Street', 'ACME')  # Replace "Rimini Street" with "ACME"
        answer = answer.replace('RSI', 'ACME')
        tuning_context.append(answer)

    # Create a list of dictionaries with the search term, top n matches, and the messages array
    system_message = json.loads(config.get('parameters', 'system_message'))
    user_message = json.loads(config.get('parameters', 'user_message'))
    user_message['content'] = user_message['content'].format(searchTerm=searchTerm)

    # Add prior context to the messages array with alternating roles
    prior_context_messages = [{"role": "user", "content": question} for question in prior_questions] + [{"role": "assistant", "content": answer} for answer in prior_answers]
    messages = [system_message, *prior_context_messages, user_message, {"role": "user", "content": "\n".join(tuning_context)}]

    # Print the search term, top n matches, and the number of tokens in the messages array
    tuningContext = messages[:2]
    print("PRIOR CONTEXT:", prior_context_messages)
    print("CLIENT QUESTION:", searchTerm)
    print("QRA CONTEXT:", matching_context)
    # Print the number of tokens in the messages variable
    tokens = sum(len(message['content'].split()) for message in messages)
    print(f"Number of tokens in messages: {tokens}")

    # Create a dictionary with the model, messages array, temperature, and max_tokens
    data = {
        "model": model,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens,
    }

    # Convert the dictionary to JSON
    jsonData = json.dumps(data)
    # Set the API endpoint URL
    apiUrl = 'https://api.openai.com/v1/chat/completions'

    # Set the request headers
    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'Bearer {api_key}'
    }

    # Send the JSON data to the API using the requests library
    try:
        response = requests.post(apiUrl, headers=headers, data=jsonData).content
        answers = json.loads(response)['choices'][0]['message']['content']
        print('ANSWER: ', answers)
        # Get the parameters using the new function
        parameters = get_parameters(
            timestamp=timestamp,
            user=user,
            baseGPTmodel=model,
            creativity=temperature,
            max_tokens=max_tokens,
            searchTerm=searchTerm,
            similarity_threshold=similarity_threshold,
            matching_context=matching_context,
            tuning_context=tuning_context,
            top_n_matches=top_n_matches,
            tokens=tokens,
            answer=answers
        )

        # Open CSV file for writing
        csv_path = os.path.join(config_dir, 'data.csv')
        with open(csv_path, 'a', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)

            if file.tell() == 0:
                # File is empty, write header row
                writer.writerow(['timestamp', 'user', 'baseGPTmodel', 'creativity', 'max response size', 'searchTerm', 'similarity_threshold', 'matching context','tuning context', 'tokens', 'answer'])

            # Write data row
            writer.writerow([timestamp, user, model, temperature, max_tokens, searchTerm, similarity_threshold, matching_context, tuning_context, tokens, answers])

    except requests.exceptions.RequestException as e:
        with open('error.txt', 'a', encoding='utf-8') as file:
            file.write(str(e))
            file.write("\n")
        print("API call failed. Error message written to error.txt.")
    except KeyError as e:
        with open('error.txt', 'a', encoding='utf-8') as file:
            file.write(str(e))
            file.write("\n")
        print("JSON parsing error. Error message written to error.txt.")
    except Exception as e:
        print("Error content:", e)  # Print the error message directly to the console
        with open('error.txt', 'a', encoding='utf-8') as file:
            file.write(str(e))
            file.write("\n")

        print("Unexpected error. Error message written to error.txt.")

    #Append the current question and answer to the respective lists
    prior_questions.append(searchTerm)
    prior_answers.append(answers)

    return answers

# only run the script if it is called directly
# ask follow-up question and close the script if there is no additional question?

if __name__ == '__main__':
    search_term = input("Enter search term: ")
    write_answer(search_term)

    while True:
        follow_up = input("Do you have any additional questions? (Y/N): ")
        if follow_up == 'Y' or follow_up == 'y':
            search_term = input("Enter search term: ")
            write_answer(search_term)
        else:
            break





# Wait for user to hit any key before closing the script


