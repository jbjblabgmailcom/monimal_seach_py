import os
import re
import fitz  # PyMuPDF
import pandas as pd
from collections import defaultdict
from funkcje import dict_to_excel, extract_numbers_from_raports, pull_nominals_from_excell, extract_column_values, is_number
from funkcje import search_for_nominals_in_raports



if __name__ == "__main__":
    
  
    nominals_from_excell = pull_nominals_from_excell('/nominaly/')
    dane_znalezione = search_for_nominals_in_raports(nominals_from_excell)
    dict_to_excel(dane_znalezione)

