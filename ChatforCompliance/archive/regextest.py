import re

# Define the regex pattern
pattern = r'\b(what|where|when|why|how|which|who|whom|whose)\b.*|\byou(r)?\b.*|\?.*|\bplease\b.*|\b(provide|describe)\b.*|Name of|'

# Define the string to match
string = '1.1 Name of Vendor: (Single selection allowed) (Allows other) *'

# Use regular expressions to search for the pattern in the string
match = re.search(pattern, string)

# Print the matched text
if match:
    print("Matched question text:", match.group())
else:
    print("No match found.")
