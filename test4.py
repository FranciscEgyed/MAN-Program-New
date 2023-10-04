import re

import pandas as pd

# Open the input file in read mode
with open('C:/Users/Wana\Desktop/W.txt', 'r', encoding='latin-1') as file:
    # Read the content of the file
    content = file.read()

# Use regular expression to replace NUL values with the desired replacement string
content_without_null = re.sub(r'\x00', '', content)
content_without_null2 = re.sub(r'\x82', '', content_without_null)
content_without_null3 = re.sub(r'\x90', '', content_without_null2)
content_without_null4 = re.sub(r'\x83', '', content_without_null3)
content_without_null5 = re.sub(r'\x89', '', content_without_null4)
content_without_null6 = re.sub(r'\x88', '', content_without_null5)
# Open the same file in write mode to update its content
with open('input.txt', 'w') as file:
    # Write the updated content back to the file
    file.write(content_without_null6)
# Split the content into lines
lines = content_without_null6.splitlines()

# Create a list of dictionaries where each dictionary represents a row in the Excel file
data = [{'Content': line} for line in lines]

df = pd.DataFrame(data)

# Assuming your DataFrame is named 'df'
chunk_size = 100000  # You can adjust the chunk size based on your available memory

# Calculate the number of chunks
num_chunks = len(df) // chunk_size + 1

# Save DataFrame in chunks to text files
for i in range(num_chunks):
    start_idx = i * chunk_size
    end_idx = (i + 1) * chunk_size
    chunk_df = df.iloc[start_idx:end_idx]
    chunk_df.to_csv(f'output_{i}.txt', sep='\t', index=False)  # Tab-separated text file
