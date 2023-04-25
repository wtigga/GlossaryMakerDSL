import re
import pandas as pd
from datetime import datetime
import codecs
import tkinter as tk
from tkinter import ttk
from tkinter import Entry, filedialog, StringVar
import sys
import xlsxwriter
import glob
import os
import threading
from tkinter.messagebox import showinfo
import webbrowser

script_version = '0.1'
modification_date = '2023-04-25'
script_name_short = 'GCG Term Extractor'
script_name = str(script_name_short + ', v' + script_version + ', ' + modification_date)


# This is a script to create a DSL dictionary from crappy source translation files.

def current_date():
    # Get the current date
    today_date = datetime.today()

    # Format the date as 'YYYY-MM-DD'
    formatted_date = today_date.strftime("%Y-%m-%d")
    return str(formatted_date)


today = current_date()
source_lang_column_name = 'CHS'
target_lang_column_name = 'RU'
comment_column_name = 'EXTRA'

list_of_columns = ['CHS', 'RU', 'EXTRA', 'Origin', 'Translation', 'en', 'kr', 'cht', 'jp', 'th', 'vi', 'id', 'es', 'ru',
                   'pt', 'de', 'fr', 'CHT', 'DE', 'EN', 'ES', 'FR', 'ID', 'JP', 'KR', 'PT', 'RU', 'TH', 'VI', 'TR',
                   'IT']
excel_file_path = ""
output_excel_file_path = ""
tm_file = ""
tb_file = ""
live_file = ""
dsl_header_1 = '#NAME "GCG Glossary ' + today + '. Made with ' + script_name + '"'
dsl_header_2 = '#INDEX_LANGUAGE "' + source_lang_column_name + '"'
dsl_header_3 = '#CONTENTS_LANGUAGE "' + target_lang_column_name + '"'
dsl_header_list = [dsl_header_1, dsl_header_2, dsl_header_3]
output_file_dsl = ""
key_pattern_regex = r'\$\[\w+\d+\]'
value_pattern_regex = r'【(\$\[\w+\d+\])→((?:(?!【).)+)】'
numeric_pattern = r'^\s*\d+(\.\d+)?\s*$'
html_pattern = r'<[^>]*>'
sprite_pattern = r'\{SPRITE_PRESET#[\d]+\}'
tb_origin_column_name = 'Origin'
tb_translation_column_name = 'Translation'
main_dict_for_output = {}
code_and_source_output = {}
live_source_and_translation_output = {}
enriched_output = {}
missing_text = 'The text is missing'
folder_content_files = []


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


def save_dictionaries_to_file_v4(dict1, dict2, file_name, missing=missing_text, dsl_header=dsl_header_list):
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


def update_dict_without_empty_values(original_dict, new_dict):
    for key, value in new_dict.items():
        if value != "":
            original_dict[key] = value
    print('Update without empty values complete')
    return original_dict


def code_and_source_function(filename,
                             source,
                             comment,
                             key_regex=key_pattern_regex,
                             value_regex=value_pattern_regex,
                             html_regex=html_pattern,
                             sprite_regex=sprite_pattern,
                             numeric_regex=numeric_pattern):
    result = extract_key_value_pairs(filename,
                                     source,
                                     comment,
                                     key_regex,
                                     value_regex)
    result = clean_and_remove_numeric_values(result,
                                             html_regex,
                                             sprite_regex,
                                             numeric_regex)
    result_sorted = {k: v for k, v in sorted(result.items())}
    print('Code snippets and source language compilation complete')
    return result_sorted


def source_and_translation_function(filename,
                                    source,
                                    target,
                                    html_regex=html_pattern,
                                    sprite_regex=sprite_pattern,
                                    numeric_regex=numeric_pattern):
    result = extract_source_translation(filename, source, target)
    result = clean_and_remove_numeric_values(result, html_regex, sprite_regex, numeric_regex)
    result = clean_keys_and_remove_numeric_values(result, html_regex, sprite_regex)
    result_sorted = {k: v for k, v in sorted(result.items())}
    print('Source and translation combining complete')
    return result_sorted


# INTERFACE GUI

def browse_folder():
    folder_path = filedialog.askdirectory()
    folder_path_var.set(folder_path)

    xlsx_files = glob.glob(os.path.join(folder_path, "*.xlsx"))
    folder_content_files.extend(xlsx_files)
    print('Folder selected:' + str(folder_path))
    print('Files found: ' + str(folder_content_files))


def browse_dsl_output():
    default_name = "dictionary_" + source_lang_column_name + "_" + target_lang_column_name + "_" + str(today) + ".dsl"
    file_path = filedialog.asksaveasfilename(filetypes=[("DSL Lingvo Dict", "*.dsl")], defaultextension=".dsl", initialfile=default_name)
    if file_path:
        output_file_dsl.set(file_path)
        file_path_entry_output_dsl.config(state="normal")
        file_path_entry_output_dsl.delete(0, tk.END)
        file_path_entry_output_dsl.insert(0, file_path)
        file_path_entry_output_dsl.config(state="readonly")


def browse_file_excel_output():
    default_name = "glossary_" + source_lang_column_name + "_" + target_lang_column_name + "_" + str(today) + ".xlsx"
    file_path = filedialog.asksaveasfilename(filetypes=[("Excel files", "*.xlsx")], defaultextension=".xlsx", initialfile=default_name)
    if file_path:
        output_excel_file_path.set(file_path)
        file_path_entry_output.config(state="normal")
        file_path_entry_output.delete(0, tk.END)
        file_path_entry_output.insert(0, file_path)
        file_path_entry_output.config(state="readonly")


class TextRedirector:
    def __init__(self, widget):
        self.widget = widget

    def write(self, string):
        self.widget.insert(tk.END, string)
        self.widget.see(tk.END)

    def flush(self):
        pass


def source_lang_selection(event):
    selected_source_lang = source_lang_var.get()
    global source_lang_column_name
    source_lang_column_name = selected_source_lang
    print("Selected source language:", source_lang_column_name)


def target_lang_selection(event):
    selected_target_lang = target_lang_var.get()
    global target_lang_column_name
    target_lang_column_name = selected_target_lang
    print("Selected target language:", target_lang_column_name)


def comment_selection(event):
    selected_comment = comment_var.get()
    global comment_column_name
    comment_column_name = selected_comment
    print("Selected comment:", comment_column_name)


def batch_processing(file_list, source_lang, target_lang, comment_column):
    global code_and_source_output
    global main_dict_for_output

    code_and_source_full = {}
    source_and_translation_full = {}

    file_count = len(file_list)
    root.after(0, progress.configure, {'maximum': file_count, 'value': 0})  # Configure the progress bar

    for index, each in enumerate(file_list):
        code_and_source_language = code_and_source_function(each, source_lang, comment_column)
        source_and_translation = source_and_translation_function(each, source_lang, target_lang)
        code_and_source_full = update_dict_without_empty_values(code_and_source_full, code_and_source_language)
        source_and_translation_full = update_dict_without_empty_values(source_and_translation_full,
                                                                       source_and_translation)

        root.after(0, progress.step)  # Update the progress bar

    main_dict_for_output = merge_dictionaries(code_and_source_full, source_and_translation_full)
    code_and_source_output = code_and_source_full
    showinfo("Processing Complete", "Processing is now complete.")
    progress['value'] = 0
    return None


def batch_execute():
    batch_processing(folder_content_files, source_lang_column_name, target_lang_column_name, comment_column_name)
    print('Batch Finished')


def start_batch_processing():
    task_thread = threading.Thread(target=batch_processing, args=(folder_content_files, source_lang_column_name, target_lang_column_name, comment_column_name))
    task_thread.start()


def dicts_to_excel(dict1, dict2, source_lang, target_lang, output_excel_path):
    # Create a DataFrame from the dictionaries
    data = {'CODE': list(dict1.keys()), source_lang: list(dict1.values()), target_lang: list(dict2.values())}
    df = pd.DataFrame(data)

    # Save the DataFrame to an Excel file
    writer = pd.ExcelWriter(output_excel_path, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')

    # Set column widths and freeze the header row
    worksheet = writer.sheets['Sheet1']
    worksheet.freeze_panes(1, 0)
    worksheet.set_column(0, 0, 40)
    worksheet.set_column(1, 1, 40)
    worksheet.set_column(2, 2, 40)

    # Save and close the Excel file
    writer.close()
    print('Saved to Excel!')


def save_to_excel():
    dicts_to_excel(code_and_source_output, main_dict_for_output, source_lang_column_name, target_lang_column_name,
                   output_excel_file_path.get())


def save_to_dsl():
    save_dictionaries_to_file_v4(code_and_source_output, main_dict_for_output, output_file_dsl.get())
    print('Saved to DSL!')


# Initialize the main window
root = tk.Tk()
root.geometry("450x700")
root.title(script_name)

# Upper part with 'Preparation' title
preparation_label = ttk.Label(root, text=script_name_short, font=("Helvetica", 12))
preparation_label.grid(row=0, column=0, pady=20, padx=20, columnspan=2, sticky='W')

# Create a button to browse for a folder
browse_folder_button = ttk.Button(root, text="Browse Folder", command=browse_folder)
browse_folder_button.grid(row=1, column=0, padx=20, sticky='W')
folder_path_var = StringVar()

# Create a text field to display the selected folder path
folder_path_entry = Entry(root, textvariable=folder_path_var, width=50, state="readonly", )
folder_path_entry.grid(row=1, column=0, padx=125, sticky='W')

# Dropdown menus
# Create StringVar for each dropdown menu
source_lang_var = tk.StringVar(root, name="source_lang")
target_lang_var = tk.StringVar(root, name="target_lang")
comment_var = tk.StringVar(root, name="comment")

# Set default values for StringVars
source_lang_var.set(source_lang_column_name)
target_lang_var.set(target_lang_column_name)
comment_var.set(comment_column_name)

# Create dropdown menus
source_lang_combobox = ttk.Combobox(root, textvariable=source_lang_var, values=list_of_columns, state='readonly')
source_lang_combobox.bind("<<ComboboxSelected>>", source_lang_selection)

target_lang_combobox = ttk.Combobox(root, textvariable=target_lang_var, values=list_of_columns, state='readonly')
target_lang_combobox.bind("<<ComboboxSelected>>", target_lang_selection)

comment_combobox = ttk.Combobox(root, textvariable=comment_var, values=list_of_columns, state='readonly')
comment_combobox.bind("<<ComboboxSelected>>", comment_selection)

# Set the grid positions for dropdown menus
source_lang_label = ttk.Label(root, text="Source column:")
source_lang_label.grid(row=2, column=0, sticky="W", padx=20, pady=10)
source_lang_combobox.grid(row=2, column=0, sticky='W', padx=125)

target_lang_label = ttk.Label(root, text="Target column:")
target_lang_label.grid(row=3, column=0, sticky="W", padx=(20, 0), pady=10)
target_lang_combobox.grid(row=3, column=0, sticky='W', padx=125)

comment_label = ttk.Label(root, text="Comment column:")
comment_label.grid(row=4, column=0, sticky="W", padx=(20, 0), pady=10)
comment_combobox.grid(row=4, column=0, sticky='W', padx=125)

# Create the Execute button
execute_button = ttk.Button(root, text="Process files", command=start_batch_processing, width="30")
execute_button.grid(row=5, column=0, padx=20, pady=5, sticky='w')

# Browse output file  and file path entry
browse_button = ttk.Button(root, text="Output Excel", command=browse_file_excel_output)
browse_button.grid(row=7, column=0, padx=20, sticky='w')
output_excel_file_path = tk.StringVar()
file_path_entry_output = ttk.Entry(root, textvariable=output_excel_file_path, state="readonly", width=50)
file_path_entry_output.grid(row=7, column=0, padx=125, sticky='w')

# Browse output file DSL  and file path entry
browse_button_dsl = ttk.Button(root, text="Output DSL", command=browse_dsl_output)
browse_button_dsl.grid(row=11, column=0, padx=20, sticky='w')
output_file_dsl = tk.StringVar()
file_path_entry_output_dsl = ttk.Entry(root, textvariable=output_file_dsl, state="readonly", width=50)
file_path_entry_output_dsl.grid(row=11, column=0, padx=125, sticky='w')

# Create the DSL save button
dsl_button = ttk.Button(root, text="Save to DSL", command=save_to_dsl)
dsl_button.grid(row=13, column=0, padx=350, pady=5, sticky='w')

# Create save to Excel button
excel_button = ttk.Button(root, text="Save to Excel", command=save_to_excel)
excel_button.grid(row=13, column=0, padx=250, pady=5, sticky='w')

progress = ttk.Progressbar(root, length=400, mode='determinate')
progress.grid(row=14, column=0, pady=10, padx=20, sticky='w')

# Add a text widget for displaying print output
output_text = tk.Text(root, wrap="word", height=10, width=50)
output_text.grid(row=15, column=0, columnspan=2, padx=20, pady=(30, 10), sticky='w')

# Create a scrollbar and attach it to the text widget
scrollbar = ttk.Scrollbar(root, command=output_text.yview)
scrollbar.grid(row=15, column=0, sticky="w", padx=0)
output_text.config(yscrollcommand=scrollbar.set)

# Redirect stdout to the text widget
sys.stdout = TextRedirector(output_text)


# Text in the bottom
def open_url(url):
    webbrowser.open(url)


about_label = tk.Label(root, text="github.com/wtigga\nVladimir Zhdanov", fg="blue", cursor="hand2", justify="left")
about_text = tk.Label(root, text=script_name)
about_text.grid(row=16, column=0, sticky='w', padx=20, pady=0)
about_label.bind("<Button-1>",
                 lambda event: open_url("https://github.com/wtigga/GlossaryMakerDSL"))
about_label.grid(row=17, column=0, sticky='w', padx=20, pady=0)

label = tk.Label(root, text="This tool parse all <XLSX> files in the folder, extracts regex-defined"
                            "\ncode from the <Comments> column, finds it's source and translations,"
                            "\nand saves as a simple spreadsheet and DSL (dictionary).", justify="left")
label.grid(row=18, column=0, padx=20, pady=0, sticky='W')

# Start the main loop
root.mainloop()

'''While the logic and the architecture are products of the author's thinking capabilities,
some of the functions in this code were written with the help of OpenAi's ChatGPT 4.'''
