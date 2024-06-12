import requests
import pandas as pd
import time
import random
import os

# Paths to the input and output files
new_import_file_path = 'contacts_formatted_for_import_final.xlsx'
output_path_updated_titles_api = 'updated_contacts_with_titles_api.xlsx'

# Load the file with the names
if os.path.exists(output_path_updated_titles_api):
    # If the output file already exists, load it to resume from where it left off
    contacts_df = pd.read_excel(output_path_updated_titles_api)
else:
    # Otherwise, load the initial import file
    contacts_df = pd.read_excel(new_import_file_path)

# Function to call Genderize.io API
def get_gender(first_name):
    if pd.isna(first_name) or first_name == '':
        return ''
    while True:
        print(f"Querying API for first name: {first_name}")
        response = requests.get(f"https://api.genderize.io?name={first_name}")
        if response.status_code == 200:
            data = response.json()
            print(f"API response for {first_name}: {data}")
            if data['gender'] == 'male':
                return 'Mr'
            elif data['gender'] == 'female':
                return 'Mme'
            else:
                return 'Inconnu'
        elif response.status_code == 429:
            print(f"Rate limit exceeded for {first_name}. Waiting to retry...")
            time.sleep(60)  # Wait for 60 seconds before retrying
        else:
            print(f"Failed to get gender for {first_name}, status code: {response.status_code}")
            return 'Inconnu'

# Apply the function to the 'Prenom' column to populate the 'Qualite' column
try:
    for index, row in contacts_df.iterrows():
        if pd.isna(row['Qualite']) or row['Qualite'] == '':
            first_name = row['Prenom']
            title = get_gender(first_name)
            contacts_df.at[index, 'Qualite'] = title
            print(f"Updated row {index}: {row['Nom']} {first_name} -> {title}")
            # Wait for a random time between 1 and 10 seconds
            delay = random.randint(1, 5)
            print(f"Sleeping for {delay} seconds...")
            time.sleep(delay)
            # Save the updated dataframe to an Excel file after each update
            contacts_df.to_excel(output_path_updated_titles_api, index=False)
            print(f"Saved progress to {output_path_updated_titles_api}")
except KeyboardInterrupt:
    print("Process interrupted. Saving current progress...")
    contacts_df.to_excel(output_path_updated_titles_api, index=False)
    print(f"Progress saved to {output_path_updated_titles_api}")

print(f"Final updated file saved to {output_path_updated_titles_api}")
