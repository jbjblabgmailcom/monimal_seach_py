import os
import re
import fitz  # PyMuPDF
import pandas as pd
from collections import defaultdict

def dict_to_excel(data_dict):
    files_in_folder = [f for f in os.listdir('.') if os.path.isfile(f)]

    numbers = []
    for file in files_in_folder:
    
        file = file.split('.')
        text = file[0]
        if re.search(r'output_', text):
                text = text.split('_')
                numbers.append(int(text[1]))
    if len(numbers) == 0:
        filename = 'output_1.xlsx'
    else: 
        filename = 'output_' + str(max(numbers)+1) + '.xlsx'
    
    # Create a DataFrame from the dictionary
    df = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in data_dict.items()]))
    
    # Transpose the DataFrame so keys are in the first column
    df = df.transpose()
    
    # Reset the index so the keys become the first column
    df.reset_index(inplace=True)
    
    # Rename the columns to make it clear
    df.columns = ["Key"] + [f"Value {i+1}" for i in range(df.shape[1] - 1)]
    
    # Write the DataFrame to an Excel file
    df.to_excel(filename, index=False)
    print(f"Excel file '{filename}' created successfully.")


def is_number(s):
    try:
        float(s)  # Try converting to float (works for both int and float strings)
        return True
    except ValueError:
        return False



def extract_numbers_from_raports(input_path):
    # Lista do przechowywania wyekstrahowanych danych liczbowych
    extracted_values = []

    # Przechodzenie przez wszystkie pliki PDF w folderze
    for filename in os.listdir(input_path):
        if filename.endswith(".pdf"):
            file_path = os.path.join(input_path, filename)

            # Otwarcie pliku PDF
            with fitz.open(file_path) as pdf_document:
                # Przechodzenie przez wszystkie strony w pliku PDF
                for page_num in range(pdf_document.page_count):
                    page = pdf_document[page_num]

                    # Wyciąganie tekstu z każdej strony
                    text = page.get_text()

                    # Wyszukiwanie wszystkich danych liczbowych (bez nagłówków i stopek)
                    #numbers = re.findall(r'(?<!\S)-?\d+(?:\.\d+)?(?!\S)', text)
                    values = re.findall(r'(?<!\S)-?\d+(?:\.\d+)?(?!\S)|\b\w+\b', text)
                    
                    # Dodawanie wyekstrahowanych danych do listy
                    extracted_values.extend(values)

                    
              
                            
    return extracted_values


def convert_to_rounded_float_strings(data, special_strings):
    def extract_and_format(value):
        # Convert value to a string for processing
        value_str = str(value).strip()
        
        # Check for any special string at the start of the value
        special_char = ""
        for s in special_strings:
            if value_str.startswith(s):
                special_char = s
                # Remove the special string from the beginning of the value
                value_str = value_str[len(s):].strip()
                break

        # Extract the numeric part of the value (including possible decimal points)
        match = re.search(r'[-+]?\d*\.?\d+', value_str)
        if match:
            #Convert to float and round to 3 decimal places
            rounded_value = round(float(match.group()), 3)
            # Format as a string with 3 decimal places, prefixed by the special character (if any)
            if special_char:
                return f"{special_char}_{rounded_value:.3f}"
            if not special_char:
                return f"{rounded_value:.3f}"
           
        else:
            return "NOT_FOUND"  # Return a default value if no numbers are found

    # Apply the extraction and formatting function to each item in the dictionary
    return {key: [extract_and_format(value) for value in values] for key, values in data.items()}



def filter_values(data, values_to_filter):
    # Use dictionary comprehension to filter each list
    return {key: [value for value in values if value not in values_to_filter] for key, values in data.items()}




def extract_column_values(excel_file_path, header_text):

    try:
        excel_file = pd.ExcelFile(excel_file_path)
    except FileNotFoundError:
        print(f"Error: The file {excel_file_path} does not exist.")
        return {}
    except Exception as e:
        print(f"An error occurred while opening the file: {e}")
        return {}

    result = {}


    for sheet_name in excel_file.sheet_names:
        try:
           
            df = pd.read_excel(excel_file, sheet_name=sheet_name)

        
            if header_text in df.columns:
              
                column_values = df[header_text].dropna().tolist()

                result[sheet_name] = column_values
                print(f"Found '{header_text}' column in sheet '{sheet_name}'. Extracted {len(column_values)} values.")
            else:
                print(f"'{header_text}' column not found in sheet '{sheet_name}'.")
        except Exception as e:
            print(f"An error occurred while processing sheet '{sheet_name}': {e}")

    return result



def pull_nominals_from_excell(dir_path):
    

    script_dir =os.path.dirname(os.path.abspath(__file__)) + dir_path
    list_from_excell = []
   
    excel_files = [f for f in os.listdir(script_dir) if f.endswith(('.xlsx', '.xls'))]

    if not excel_files:
        print("No Excel files found in the script directory.")
        return
    elif len(excel_files) > 1:
        print("Multiple Excel files found. Please ensure only one Excel file is present or modify the script to specify the filename.")
        return

    excel_file = excel_files[0]
    excel_file_path = os.path.join(script_dir, excel_file)
    print(f"Processing Excel file: {excel_file}")

    # Specify the header text to search for
    header_text = "XXX"

    # Extract the values
    extracted_data = extract_column_values(excel_file_path, header_text)
    #print(extracted_data)
    

    values_to_filter = {'Model or Drawing Dimension Value'}
    extracted_data = filter_values(extracted_data, values_to_filter)
    print(extracted_data)

    special_strings = ["Ø", "Rz", "R", "ISO", "M"]
    
    extracted_data = convert_to_rounded_float_strings(extracted_data, special_strings)
    print(extracted_data)



    # Display the results
    for sheet, values in extracted_data.items():
        print(f"\nSheet: {sheet}")
        for idx, value in enumerate(values, start=1):
            print(f"{idx}: {value}")
            list_from_excell.append(value)
    return list_from_excell




def search_for_nominals_in_raports(arr):
    input_folder = "./raporty"
    dane_wyjsciowe = defaultdict(list)
    res = extract_numbers_from_raports(input_folder)
    print(res)
    for wyszukiwana in arr:
        
      
        newvar = str(wyszukiwana)
        splitted = newvar.split('_')
        if len(splitted) == 1:
            splitted.insert(0, '0')
        print(splitted)
        
        if splitted[0] == 'Ø':
            
            for n_index, n in enumerate(res):

                if n == splitted[1]:
                    if not is_number(res[n_index-1]) and is_number(res[n_index+2]) and ((res[n_index+4] == 'Średnica') or res[n_index+5] == 'Średnica'):
                        dane_wyjsciowe[wyszukiwana].append(res[n_index+2])

        elif splitted[0] == 'Rz':
            dane_wyjsciowe[wyszukiwana].append(f'{splitted[0]} {splitted[1]} - nie znaleziono')

        elif splitted[0] == 'ISO':
            dane_wyjsciowe[wyszukiwana].append(f'{splitted[0]} {splitted[1]} - nie znaleziono')
                        

        elif splitted[0] == 'R':
            
            for n_index, n in enumerate(res):

                if n == splitted[1]:
                    if not is_number(res[n_index-1]) and is_number(res[n_index+2]) and ((res[n_index+4] == 'Promień') or (res[n_index+5] == 'Promień')):
                        dane_wyjsciowe[wyszukiwana].append(res[n_index+2])
                    else:
                        dane_wyjsciowe[wyszukiwana].append('promień nie znaleziono')

        elif splitted[0] == 'M':
            
            for n_index, n in enumerate(res):

                if n == splitted[1]:
                    dane_wyjsciowe[wyszukiwana].append('GWINT nie znaleziono')

       
        elif splitted[0]  == '0':
            for n_index, n in enumerate(res):

                if n == splitted[1]:

                    if (res[n_index-1] == 'X') or (res[n_index-1] == 'Y') or (res[n_index-1] == 'Z'):
                        pass

                    elif (res[n_index+5] == '-' + splitted[1]) or (res[n_index+6] == '-' + splitted[1]):
                        #print('MINUS')
                        pass

                    elif is_number(res[n_index-1]) and is_number(res[n_index+1]) and (res[n_index+3] == 'Pozycja' or res[n_index+4] == 'Pozycja'):
                        dane_wyjsciowe[wyszukiwana].append(res[n_index+2])
                        
                    elif not is_number(res[n_index-1]) and is_number(res[n_index+2]):
                        dane_wyjsciowe[wyszukiwana].append(res[n_index+2])
                        
                    elif not is_number(res[n_index-1]) and not is_number(res[n_index+2]):
                        dane_wyjsciowe[wyszukiwana].append(res[n_index+1])

    return dane_wyjsciowe