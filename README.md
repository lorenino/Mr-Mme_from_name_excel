# Mr-Mme from Name Excel

This Python script uses the Genderize.io API to determine the gender (Mr or Mme) based on the first names in an Excel file. It updates the "Qualité" column in the Excel file with "Mr" or "Mme" accordingly. The script handles API rate limiting and saves progress after each update to ensure data is not lost.

## Features
- Reads an Excel file with columns "Nom", "Prenom", and "Qualité".
- Queries the Genderize.io API to determine the gender based on first names.
- Updates the "Qualité" column with "Mr" or "Mme".
- Handles rate limiting by retrying after receiving a 429 error.
- Saves progress after each update to ensure data integrity.

## Requirements
- Python 3.x
- pandas
- requests
- openpyxl

## Installation
1. Clone the repository:
    ```sh
    git clone https://github.com/lorenino/Mr-Mme_from_name_excel.git
    cd Mr-Mme_from_name_excel
    ```

2. Install the required packages:
    ```sh
    pip install pandas requests openpyxl
    ```

## Usage
1. Place your input Excel file named `contacts_formatted_for_import_final.xlsx` in the same directory as the script.

2. Run the script:
    ```sh
    python main.py
    ```

3. The script will read the input file, update the "Qualité" column, and save the results to `updated_contacts_with_titles_api.xlsx`. It will resume from where it left off if interrupted.

## Handling API Rate Limiting
The script includes logic to handle API rate limiting (HTTP 429 errors). If a rate limit error occurs, the script will wait for 60 seconds before retrying the request.

## Example Output
After running the script, the `updated_contacts_with_titles_api.xlsx` file will contain the updated "Qualité" column based on the first names in the "Prenom" column.

## Acknowledgements
- [Genderize.io](https://genderize.io) for providing the API to determine gender based on first names.
