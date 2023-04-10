import numpy as np
from sklearn.metrics.pairwise import cosine_similarity
from openai.embeddings_utils import get_embedding

# Define a query sentence and candidate sentences
query_sentence = "What is the capital of France?"
candidate_sentences = [
    "Paris is the capital of France.",
    "The Eiffel Tower is located in Paris.",
    "France is a country in Europe.",
    "London is the capital of the United Kingdom."
]

# Generate embeddings for the query sentence and candidate sentences
query_embedding = np.array(get_embedding(query_sentence, engine="text-embedding-ada-002")).reshape(1, -1)
candidate_embeddings = [np.array(get_embedding(sentence, engine="text-embedding-ada-002")).reshape(1, -1) for sentence in candidate_sentences]

# Calculate cosine similarity between the query and each candidate sentence
similarity_scores = [cosine_similarity(query_embedding, candidate_embedding)[0][0] for candidate_embedding in candidate_embeddings]

# Rank candidate sentences based on their cosine similarity to the query sentence
ranked_indices = np.argsort(similarity_scores)[::-1]
ranked_sentences = [candidate_sentences[i] for i in ranked_indices]

# Print ranked sentences based on similarity
print("Ranked sentences based on similarity to the query:")
for i, sentence in enumerate(ranked_sentences):
    print(f"{i+1}. {sentence}")

# Print similarity scores for the ranked sentences
print("\nSimilarity scores for the ranked sentences:")
for i, score in enumerate(np.sort(similarity_scores)[::-1]):
    print(f"{i+1}. Score: {score}")


"""
import pickle
from openai.embeddings_utils import get_embedding

def generate_embeddings_file(test_string, filename):
    # Define the test string to generate embeddings
    test_string = test_string

    # Specify the model engine to be used
    engine = "text-embedding-ada-002"

    # Generate embeddings for the test string
    test_embeddings = get_embedding(test_string, engine=engine)

    # Save the embeddings to a file
    with open(filename, 'wb') as f:
        pickle.dump(test_embeddings, f)
    print("Embeddings file created.")

# Define the test string and filename
test_string = "This is a sample sentence for testing."
filename = "test_embeddings.pkl"

# Call the function with the test string and filename
generate_embeddings_file(test_string, filename)
"""