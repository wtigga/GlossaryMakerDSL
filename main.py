import re
import pandas as pd
from datetime import datetime
import codecs
from sources import excel_file_local, output_file_dsl_local

#This is a script to create a DSL dictionary from crappy source translation files.1

def current_date():
    # Get the current date
    today = datetime.today()

    # Format the date as 'YYYY-MM-DD'
    formatted_date = today.strftime("%Y-%m-%d")
    return(str(formatted_date))

today = current_date()
source_lang = 'CHS'
column_name = 'EXTRA'
excel_file = excel_file_local
dsl_header_1 = '#NAME "GCG Glossary ' + today +'"'
dsl_header_2 = '#INDEX_LANGUAGE "Chinese"'
dsl_header_3 = '#CONTENTS_LANGUAGE "Russian"'
dsl_header = [dsl_header_1, dsl_header_2, dsl_header_3]
default_translation = 'Здесь должен быть перевод'
missing_text = 'Source text is missing'
output_file_dsl = output_file_dsl_local


def extract_key_value_pairs_v2(excel_file, source_lang_column_name, column_name):
    # Read the Excel file
    excel_data = pd.read_excel(excel_file, sheet_name=None)

    # Define the regex patterns
    key_pattern = r'(\$\[[\w\d]+\])'
    key_value_pattern = r'(\$\[[\w\d]+\])\s*→\s*([\u4e00-\u9fff]+)'

    # Initialize an empty dictionary to store key-value pairs
    key_value_dict = {}

    # Iterate through each sheet in the Excel file
    for sheet_name, sheet_data in excel_data.items():
        # First, scan the source_lang_column_name column for keys
        for cell_value in sheet_data[source_lang_column_name]:
            # Skip rows with empty or non-string cell values
            if not isinstance(cell_value, str):
                continue

            # Find all matches in the current cell
            keys = re.findall(key_pattern, cell_value)

            # Add the keys to the dictionary with empty values
            for key in keys:
                key_value_dict[key] = ""

        # Next, scan the column_name column for key-value pairs
        for cell_value in sheet_data[column_name]:
            # Skip rows with empty or non-string cell values
            if not isinstance(cell_value, str):
                continue

            # Find all matches in the current cell
            matches = re.findall(key_value_pattern, cell_value)

            # Update the key-value pairs in the dictionary
            for key, value in matches:
                key_value_dict[key] = value

    return key_value_dict

def print_key_value_pairs(key_value_dict):
    for key, value in key_value_dict.items():
        print(f"{key}: {value}")

def fill_values_with_test_translation(original_dict, test_translation):
    # Create a new dictionary with the same keys as the original dictionary
    new_dict = {key: test_translation for key in original_dict}

    return new_dict

def save_dictionaries_to_file_v4(dict1, dict2, dsl_header, file_name, missing):
    #it is important to save as UTF-16 LE BOM, otherwise GoldenDict will not recognize the dictionary.
    try:
        with open(file_name, "w", encoding="utf-16-le") as output_file:
            # Write the BOM
            output_file.write(codecs.BOM_UTF16_LE.decode("utf-16-le"))

            # Write the header
            for line in dsl_header:
                output_file.write(line + "\n")

            # Write a blank line after the header
            output_file.write("\n")

            # Write the main body of the file
            for key, value1 in dict1.items():
                try:
                    value2 = dict2[key]

                    # Check if value1 is empty and set it to 'Missing Text'
                    if not value1:
                        value1 = missing

                    output_file.write(f"{key}\n")
                    output_file.write(f"\t{value1}\n")
                    output_file.write(f"\t{value2}\n")
                    output_file.write("\n")
                except KeyError:
                    print(f"Key '{key}' not found in dict2. Skipping...")
                    continue
    except Exception as e:
        print(f"An error occurred while saving dictionaries to the file: {e}")


source_dict = extract_key_value_pairs_v2(excel_file, source_lang, column_name)
filled_dict = fill_values_with_test_translation(source_dict, default_translation)
print_key_value_pairs(filled_dict)

save_dictionaries_to_file_v4(source_dict, filled_dict, dsl_header, output_file_dsl, missing_text)
#print_key_value_pairs(result)