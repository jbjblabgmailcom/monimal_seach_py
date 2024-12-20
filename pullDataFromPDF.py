import os
import re
import fitz  # PyMuPDF
import pandas as pd
from collections import defaultdict
from funkcje import dict_to_excel, extract_numbers_from_raports, pull_nominals_from_excell, extract_column_values, is_number
from search import search_for_nominals_in_raports
from funkcje import create_file_list, combine_nested_dicts



if __name__ == "__main__":
    
  
    nominals_from_excell_list = pull_nominals_from_excell('/nominaly/')
    file_list_dict = create_file_list('raporty')
    print(file_list_dict)
    dane_znalezione = search_for_nominals_in_raports(nominals_from_excell_list, file_list_dict)
    #print(dane_znalezione['nikon-altera'])
    #print(dane_znalezione['mitutoyo-crysta-apex'])
    # print(dane_znalezione['mitutoyo-reczna'])
    #print(dane_znalezione['nikon-altera-rpt'])
    #print(dane_znalezione)
    new_dict = combine_nested_dicts(dane_znalezione)
    #print(new_dict)
    dict_to_excel(new_dict)

