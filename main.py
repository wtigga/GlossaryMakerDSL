import re
import pandas as pd
from datetime import datetime
import codecs
from sources import excel_file_local, output_file_dsl_local, tm_file_local, tb_file_local, live_file_local


# This is a script to create a DSL dictionary from crappy source translation files.1

def current_date():
    # Get the current date
    today_date = datetime.today()

    # Format the date as 'YYYY-MM-DD'
    formatted_date = today_date.strftime("%Y-%m-%d")
    return str(formatted_date)


today = current_date()
source_lang_column_name = 'CHS'
comment_column_name = 'EXTRA'
target_lang_column_name = 'RU'
excel_file_path = excel_file_local
tm_file = tm_file_local
tb_file = tb_file_local
live_file = live_file_local
dsl_header_1 = '#NAME "GCG Glossary ' + today + '"'
dsl_header_2 = '#INDEX_LANGUAGE "Chinese"'
dsl_header_3 = '#CONTENTS_LANGUAGE "Russian"'
dsl_header_list = [dsl_header_1, dsl_header_2, dsl_header_3]
output_file_dsl = output_file_dsl_local
key_pattern_regex = r'\$\[\w+\d+\]'
value_pattern_regex = r'【(\$\[\w+\d+\])→((?:(?!【).)+)】'
numeric_pattern = r'^\s*\d+(\.\d+)?\s*$'
html_pattern = r'<[^>]*>'
sprite_pattern = r'\{SPRITE_PRESET#[\d]+\}'
tb_origin_column_name = 'Origin'
tb_translation_column_name = 'Translation'


# General functions

def print_key_value_pairs(key_value_dict):
    for key, value in key_value_dict.items():
        print(f"{key}: {value}")


# Create ID + Source from comments

def extract_key_value_pairs(excel_file, source_lang, comment_name, key_pattern, value_pattern):
    # This function is to extract the IDs and corresponding source language from the comment (EXTRA) column
    # Read the Excel file
    excel_data = pd.read_excel(excel_file, sheet_name=None)

    # Initialize an empty dictionary to store key-value pairs
    key_value_dict = {}

    # Iterate through each sheet in the Excel file
    for sheet_name, sheet_data in excel_data.items():
        count_keys = 0

        # First, scan the source_lang_column_name column for keys
        for cell_value in sheet_data[source_lang]:
            # Skip rows with empty or non-string cell values
            if not isinstance(cell_value, str):
                continue

            # Find all matches in the current cell
            keys = re.findall(key_pattern, cell_value)

            # Add the keys to the dictionary with empty values
            for key in keys:
                key_value_dict[key] = ""
                count_keys += 1

                # Print progress after every 100 keys
                if count_keys % 100 == 0:
                    print(f"Extracted {count_keys} keys so far from column {source_lang}")

        count_key_values = 0

        # Next, scan the comment_name column for key-value pairs
        for cell_value in sheet_data[comment_name]:
            # Skip rows with empty or non-string cell values
            if not isinstance(cell_value, str):
                continue

            # Find all matches in the current cell
            matches = re.findall(value_pattern, cell_value)

            # Update the key-value pairs in the dictionary
            for key, value in matches:
                key_value_dict[key] = value
                count_key_values += 1

                # Print progress after every 100 key-value pairs
                if count_key_values % 1000 == 0:
                    print(f"Extracted {count_key_values} key-value pairs so far from column {comment_name}")

    return key_value_dict


def clean_and_remove_numeric_values(input_dict, html_regex_pattern, sprite_regex_pattern, numeric_regex_pattern):
    cleaned_dict = {}

    html_regex = re.compile(html_regex_pattern)
    sprite_regex = re.compile(sprite_regex_pattern)
    numeric_regex = re.compile(numeric_regex_pattern)

    for key, value in input_dict.items():
        if isinstance(value, str):
            cleaned_value = html_regex.sub("", value)
            cleaned_value = sprite_regex.sub("", cleaned_value)

            if not numeric_regex.match(cleaned_value):
                cleaned_dict[key] = cleaned_value
        else:
            cleaned_dict[key] = value

    return cleaned_dict


def clean_keys_and_remove_numeric_values(input_dict, html_regex, sprite_regex):
    cleaned_dict = {}

    for key, value in input_dict.items():
        # Clean up HTML tags and sprite preset codes in the key
        cleaned_key = re.sub(html_regex, "", key)
        cleaned_key = re.sub(sprite_regex, "", cleaned_key)

        # Check if the cleaned key is a number (integer, float, or number as a string)
        try:
            float(cleaned_key)
        except ValueError:
            # If not a number, add the cleaned key-value pair to the cleaned_dict
            cleaned_dict[cleaned_key] = value

    return cleaned_dict


####
def extract_source_translation(excel_filepath, source_lang, target_lang):
    # Read the Excel file
    excel_data = pd.read_excel(excel_filepath, sheet_name=None)

    # Initialize an empty dictionary to store translations
    translation_dict = {}

    # Initialize a counter for processed rows
    processed_rows = 0

    # Iterate through each sheet in the Excel file
    for sheet_name, sheet_data in excel_data.items():
        # Iterate over the rows, skipping the header (first row)
        for index, row in sheet_data.iterrows():
            if index == 0:
                continue

            # Get the source language text and target language translation from the row
            source_text = row[source_lang]
            target_text = row[target_lang]

            # Add the translation to the dictionary
            translation_dict[source_text] = target_text

            # Increment the processed rows counter
            processed_rows += 1

            # Print progress update every 100 processed rows
            if processed_rows % 1000 == 0:
                print(f"Processed {processed_rows} rows so far")

    return translation_dict


def save_dictionaries_to_file_v4(dict1, dict2, dsl_header, file_name, missing):
    # it is important to save as UTF-16 LE BOM, otherwise GoldenDict will not recognize the dictionary.
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


def combine_columns_to_dictionary(file_path, source_lang, target_lang):
    # Read the Excel file with all sheets into a dictionary of DataFrames
    df_dict = pd.read_excel(file_path, sheet_name=None)

    result_dict = {}

    # Iterate through each sheet
    for sheet_name, df in df_dict.items():
        # Filter the DataFrame to keep only the specified columns
        filtered_df = df[[source_lang, target_lang]]

        # Remove rows with NaN values in either column
        filtered_df = filtered_df.dropna(subset=[source_lang, target_lang])

        # Convert the filtered DataFrame into a dictionary
        temp_dict = filtered_df.set_index(source_lang)[target_lang].to_dict()

        # Merge the temp_dict into the result_dict
        result_dict.update(temp_dict)

        # Print progress after every 100 rows
        current_row = len(result_dict)
        if current_row % 1000 == 0:
            print(f"Parsed {current_row} strings so far")

    return result_dict


def merge_dictionaries(dict1, dict2):
    new_dict = {}

    for key, value in dict1.items():
        if value in dict2:
            new_dict[key] = dict2[value]
        else:
            new_dict[key] = ""

    return new_dict


def merge_dictionaries_no_na(dict1, dict2):
    new_dict = {}

    for key, value in dict1.items():
        new_dict[key] = dict2.get(value, value)

    return new_dict


def remove_empty_values(dct):
    return {key: value for key, value in dct.items() if value != ""}


def merge_dicts(dict1, dict2):
    return {**dict1, **dict2}


def merge_dictionaries_overwrite_empty(dict1, dict2):
    new_dict = dict1.copy()

    for key, value in dict2.items():
        if key in new_dict and new_dict[key] == "":
            new_dict[key] = value

    return new_dict


# USAGE


code_and_source_language = extract_key_value_pairs(excel_file_path,
                                                   source_lang_column_name,
                                                   comment_column_name,
                                                   key_pattern_regex,
                                                   value_pattern_regex)
code_and_source_language = clean_and_remove_numeric_values(code_and_source_language,
                                                           html_pattern,
                                                           sprite_pattern,
                                                           numeric_pattern)
code_and_source_language_sorted = {k: v for k, v in sorted(code_and_source_language.items())}
print('Source code finished')
# print_key_value_pairs(code_and_source_language)


live_source_and_translation = extract_source_translation(live_file,
                                                         source_lang_column_name,
                                                         target_lang_column_name)
live_source_and_translation = clean_and_remove_numeric_values(live_source_and_translation,
                                                              html_pattern,
                                                              sprite_pattern,
                                                              numeric_pattern)
live_source_and_translation = clean_keys_and_remove_numeric_values(live_source_and_translation,
                                                                   html_pattern,
                                                                   sprite_pattern)
print('Live finished')

current_source_and_translation = extract_source_translation(excel_file_path,
                                                            source_lang_column_name,
                                                            target_lang_column_name)
current_source_and_translation = clean_and_remove_numeric_values(current_source_and_translation,
                                                                 html_pattern,
                                                                 sprite_pattern,
                                                                 numeric_pattern)
current_source_and_translation = clean_keys_and_remove_numeric_values(current_source_and_translation,
                                                                      html_pattern,
                                                                      sprite_pattern)

merged_current = merge_dictionaries(code_and_source_language, current_source_and_translation)
merged_live = merge_dictionaries(code_and_source_language, live_source_and_translation)

final = merge_dictionaries_overwrite_empty(merged_current, merged_live)
final_sorted = {k: v for k, v in sorted(final.items())}
print_key_value_pairs(final_sorted)

print_key_value_pairs(code_and_source_language_sorted)
