import os
import re
import fitz  # PyMuPDF
import pandas as pd
from collections import defaultdict
from funkcje import extract_numbers_from_raports, is_number
from funkcje import extract_data_from_raport
from funkcje import is_opposite

def search_nikon_altera(nominals_from_excell_list, res):
    #print(res)
    dane_wyjsciowe = defaultdict(list)
    #print(res)
    
    for wyszukiwana in nominals_from_excell_list:
        newvar = str(wyszukiwana)
        splitted = newvar.split('_')
        print(splitted)
        if len(splitted) == 1:
            splitted.insert(0, '0')
        #altera-dziala    
        if splitted[0] == 'Ø':
            for n_index, n in enumerate(res):
                if ((splitted[1] == n and res[n_index - 10] == 'Nom' and res[n_index - 1] == 'ŚRED')) and (res[n_index + 1] not in dane_wyjsciowe[wyszukiwana]):
                    dane_wyjsciowe[wyszukiwana].append(res[n_index +1])
        #altera-dziala            
        elif splitted[0] == 'Rz':
            for n_index, n in enumerate(res):
                if (f'nie znaleziono' not in dane_wyjsciowe[wyszukiwana]):
                    dane_wyjsciowe[wyszukiwana].append(f'nie znaleziono')
        #altera-dziala
        elif splitted[0] == 'ISO':
            for n_index, n in enumerate(res):
                if (f'nie znaleziono' not in dane_wyjsciowe[wyszukiwana]):
                    dane_wyjsciowe[wyszukiwana].append(f'nie znaleziono')
        
        #altera-ZDOBYC RAPORT Z R
        elif splitted[0] == 'R':
            for n_index, n in enumerate(res):
                if ((splitted[1] == n and res[n_index - 10] == 'Nom' and res[n_index - 1] == 'PROM')) and (res[n_index + 1] not in dane_wyjsciowe[wyszukiwana]):
                    dane_wyjsciowe[wyszukiwana].append(res[n_index +1])

        elif splitted[0] == 'M':
                
            for n_index, n in enumerate(res):
                if (f'{splitted[0]}{splitted[1]} nie znaleziono' not in dane_wyjsciowe[wyszukiwana]):
                    dane_wyjsciowe[wyszukiwana].append(f'{splitted[0]}{splitted[1]} nie znaleziono')

        #altera - dziala - dorobić rozroznianie na pozycje rownoleglosci itp
        elif splitted[0] == '0':
            for n_index, n in enumerate(res):
                if (splitted[1] == n and res[n_index - 10] == 'Nom') and (res[n_index + 1] not in dane_wyjsciowe[wyszukiwana]):
                    dane_wyjsciowe[wyszukiwana].append(res[n_index + 1])
                elif (splitted[1] == n and res[n_index -11] =='Nom') and (res[n_index -1] not in dane_wyjsciowe[wyszukiwana]):
                    dane_wyjsciowe[wyszukiwana].append(res[n_index - 1])
                else:
                    pass


    #print(dane_wyjsciowe)
    return dane_wyjsciowe

def search_mitutoyo_reczna(nominals_from_excell_list, res):
   # print(res)
    dane_wyjsciowe = defaultdict(list)
    #print(res)
    
    for wyszukiwana in nominals_from_excell_list:
        newvar = str(wyszukiwana)
        splitted = newvar.split('_')
        print(splitted)
        if len(splitted) == 1:
            splitted.insert(0, '0')

        if splitted[0] == 'Ø':
            for n_index, n in enumerate(res):
                if ((splitted[1] == n and res[n_index - 10] == 'Nom' and res[n_index - 1] == 'ŚRED')) and (res[n_index + 1] not in dane_wyjsciowe[wyszukiwana]):
                    dane_wyjsciowe[wyszukiwana].append(res[n_index +1])

        elif splitted[0] == 'Rz':
            for n_index, n in enumerate(res):
                if (f'nie znaleziono' not in dane_wyjsciowe[wyszukiwana]):
                    dane_wyjsciowe[wyszukiwana].append(f'nie znaleziono')
        elif splitted[0] == 'ISO':
            for n_index, n in enumerate(res):
                if (f'nie znaleziono' not in dane_wyjsciowe[wyszukiwana]):
                    dane_wyjsciowe[wyszukiwana].append(f'nie znaleziono')
        
        elif splitted[0] == 'R':
            for n_index, n in enumerate(res):
                if ((splitted[1] == n and res[n_index + 9] == 'Nom' and res[n_index + 6] == 'PROM')) and (res[n_index - 5] not in dane_wyjsciowe[wyszukiwana]):
                    dane_wyjsciowe[wyszukiwana].append(res[n_index - 5])


        elif splitted[0] == '0':
            for n_index, n in enumerate(res):
                if (splitted[1] == n and res[n_index - 10] == 'Nom') and (res[n_index + 1] not in dane_wyjsciowe[wyszukiwana]):
                    dane_wyjsciowe[wyszukiwana].append(res[n_index + 1])




    return dane_wyjsciowe


def search_mitutoyo_crysta_apex(nominals_from_excell_list, res):
    newelements = [999, 999, 999, 999, 999, 999, 999, 999, 999, 999, 999, 999, 999, 999, 999, 999]

    res.extend([999, 999, 999, 999, 999, 999, 999, 999, 999, 999, 999, 999, 999, 999, 999, 999])
    res = newelements + res
    #print(res)
    dane_wyjsciowe = defaultdict(list)
    
    for wyszukiwana in nominals_from_excell_list:
        dane_wyjsciowe[wyszukiwana].append(' ') 
        newvar = str(wyszukiwana)
        splitted = newvar.split('_')
        if len(splitted) == 1:
            splitted.insert(0, '0')
        
        
        
        if splitted[0] == 'Ø':
            for n_index, n in enumerate(res):

                if n == splitted[1]:
                    if (not is_number(res[n_index-1]) and is_number(res[n_index+2]) and ((res[n_index+4] == 'Średnica') or res[n_index+5] == 'Średnica')) and (res[n_index+2] not in dane_wyjsciowe[wyszukiwana]):
                        dane_wyjsciowe[wyszukiwana].append(res[n_index+2])

        
        elif splitted[0] == 'Rz' and (f'{splitted[0]} {splitted[1]} - nie znaleziono' not in dane_wyjsciowe[wyszukiwana]):
            dane_wyjsciowe[wyszukiwana].append(f'{splitted[0]} {splitted[1]} - nie znaleziono')

        
        elif splitted[0] == 'ISO' and (f'{splitted[0]} {splitted[1]} - nie znaleziono' not in dane_wyjsciowe[wyszukiwana]):
            dane_wyjsciowe[wyszukiwana].append(f'{splitted[0]} {splitted[1]} - nie znaleziono')
                            

        elif splitted[0] == 'R':
                    
            for n_index, n in enumerate(res):

                if n == splitted[1] and (res[n_index+2] not in dane_wyjsciowe[wyszukiwana]):
                    if not is_number(res[n_index-1]) and is_number(res[n_index+2]) and ((res[n_index+4] == 'Promień') or (res[n_index+5] == 'Promień')):
                        dane_wyjsciowe[wyszukiwana].append(res[n_index+2])
                    elif 'promień nie znaleziono' not in dane_wyjsciowe[wyszukiwana]:
                        dane_wyjsciowe[wyszukiwana].append('promień nie znaleziono')

        elif splitted[0] == 'M':
                
            for n_index, n in enumerate(res):

                if n == splitted[1] and ('GWINT nie znaleziono' not in dane_wyjsciowe[wyszukiwana]):
                    dane_wyjsciowe[wyszukiwana].append('GWINT nie znaleziono')
        #dziala ale sprawdzic co jezeli jest poza tolerancja
        elif splitted[0] == '∥':
            for n_index, n in enumerate(res):
                if (splitted[1] == n and res[n_index +2] =='Równoległość') and (res[n_index +1] not in dane_wyjsciowe[wyszukiwana]):
                    dane_wyjsciowe[wyszukiwana].append(res[n_index +1])
        #dziala ale sprawdzic co jezeli jest poza tolerancja
        elif splitted[0] == '⏥':
            for n_index, n in enumerate(res):
                if (splitted[1] == n and res[n_index +3] =='Płaskość') and (res[n_index +1] not in dane_wyjsciowe[wyszukiwana]):
                    dane_wyjsciowe[wyszukiwana].append(res[n_index +1])
        #nie dziala
        elif splitted[0] == '⟂':
            for n_index, n in enumerate(res):
                if (splitted[1] == n and res[n_index -2] =='PROST_3D') and (res[n_index -1] not in dane_wyjsciowe[wyszukiwana]):
                    dane_wyjsciowe[wyszukiwana].append(res[n_index - 1])
        
        elif splitted[0] == '⌖':
            for n_index, n in enumerate(res):
                if n == splitted[1]:
                    if ((res[n_index-1] == 'X') or (res[n_index-1] == 'Y') or (res[n_index-1] == 'Z')):
                        pass

                    elif (res[n_index+5] == '-' + splitted[1]) or (res[n_index+6] == '-' + splitted[1]):
                        pass
                    
                    elif (is_number(res[n_index-1]) and is_number(res[n_index+1]) and (res[n_index+3] == 'Pozycja' or res[n_index+4] == 'Pozycja')) and (res[n_index+2] not in dane_wyjsciowe[wyszukiwana]):
                        dane_wyjsciowe[wyszukiwana].append(res[n_index+2])

           
        elif splitted[0]  == '0':
            for n_index, n in enumerate(res):
                if n == splitted[1]:
                    if (not is_number(res[n_index-1]) and is_number(res[n_index+2])) and (res[n_index+2] not in dane_wyjsciowe[wyszukiwana]):
                        dane_wyjsciowe[wyszukiwana].append(res[n_index+2])
                                    
                    elif (not is_number(res[n_index-1]) and not is_number(res[n_index+2])) and (res[n_index+1] not in dane_wyjsciowe[wyszukiwana]):
                        dane_wyjsciowe[wyszukiwana].append(res[n_index+1])
                    else:
                        pass

    return dane_wyjsciowe



def search_nikon_altera_rpt(nominals_from_excell_list, res):
    res.extend([999, 999, 999, 999, 999, 999, 999])
    dane_wyjsciowe = defaultdict(list)
    #print(res)
    
    for wyszukiwana in nominals_from_excell_list:
        dane_wyjsciowe[wyszukiwana].append(' ')
        newvar = str(wyszukiwana)
        splitted = newvar.split('_')
        #print(splitted)
        if len(splitted) == 1:
            splitted.insert(0, '0')
        #altera-dziala    
        if splitted[0] == 'Ø':
            for n_index, n in enumerate(res):
                if ((splitted[1] == n and res[n_index - 1] == 'ŚRED')) and (res[n_index + 1] not in dane_wyjsciowe[wyszukiwana]):
                    dane_wyjsciowe[wyszukiwana].append(res[n_index +1])
        #altera-dziala            
        elif splitted[0] == 'Rz':
            for n_index, n in enumerate(res):
                if (f'nie znaleziono' not in dane_wyjsciowe[wyszukiwana]):
                    dane_wyjsciowe[wyszukiwana].append(f'nie znaleziono')
        #altera-dziala
        elif splitted[0] == 'ISO':
            for n_index, n in enumerate(res):
                if (f'nie znaleziono' not in dane_wyjsciowe[wyszukiwana]):
                    dane_wyjsciowe[wyszukiwana].append(f'nie znaleziono')
         

        elif splitted[0] == 'R':
            for n_index, n in enumerate(res):
                if (splitted[1] == n and res[n_index-1] == 'PROM') and (res[n_index+1] not in dane_wyjsciowe[wyszukiwana]):
                   dane_wyjsciowe[wyszukiwana].append(res[n_index +1])

        elif splitted[0] == 'M':
                
            for n_index, n in enumerate(res):
                if (f'{wyszukiwana} nie znaleziono' not in dane_wyjsciowe[wyszukiwana]):
                    dane_wyjsciowe[wyszukiwana].append(f'{wyszukiwana} nie znaleziono')

        #altera - dziala - dorobić rozroznianie na pozycje rownoleglosci itp
        elif splitted[0] == '∥':
            for n_index, n in enumerate(res):
                if (splitted[1] == n and res[n_index -2] =='PARA_3D') and (res[n_index -1] not in dane_wyjsciowe[wyszukiwana]):
                    dane_wyjsciowe[wyszukiwana].append(res[n_index - 1])
        elif splitted[0] == '⏥':
            for n_index, n in enumerate(res):
                if (splitted[1] == n and res[n_index -2] =='PŁASK') and (res[n_index -1] not in dane_wyjsciowe[wyszukiwana]):
                    dane_wyjsciowe[wyszukiwana].append(res[n_index - 1])

        elif splitted[0] == '⟂':
            for n_index, n in enumerate(res):
                if (splitted[1] == n and res[n_index -2] =='PROST_3D') and (res[n_index -1] not in dane_wyjsciowe[wyszukiwana]):
                    dane_wyjsciowe[wyszukiwana].append(res[n_index - 1])
        elif splitted[0] == '⌖':
            for n_index, n in enumerate(res):
                if (splitted[1] == n and res[n_index - 2] == 'POZ') and (res[n_index - 1] not in dane_wyjsciowe[wyszukiwana]):
                    dane_wyjsciowe[wyszukiwana].append(res[n_index -1])


        
        #odleglosci
        elif splitted[0] == '0':
            for n_index, n in enumerate(res):

              
                if (splitted[1] == n and is_number(res[n_index+1]) and is_number(res[n_index+2]) and not is_opposite(res[n_index], res[n_index-1]) and res[n_index -3] != 'ŚRED') and (res[n_index + 1] not in dane_wyjsciowe[wyszukiwana]):
                    dane_wyjsciowe[wyszukiwana].append(res[n_index + 1])
                

                 #   pass

                    


    #print(dane_wyjsciowe)
    return dane_wyjsciowe


def search_for_nominals_in_raports(nominals_from_excell_list, file_list_dict):
    input_folder = "./raporty"
    input_path = './raporty'
    dane_wyjsciowe = defaultdict(list)
    #dane_z_jednego_pliku = defaultdict(list)
    dane_altera = []
    dane_crysta = []
    dane_mitutoyo_reczna = []
    dane_altera_rpt = []
    
    
    for key, vals in file_list_dict.items():
        if key == 'nikon-altera':
            for v in vals:
                raport_file = v
                res = extract_data_from_raport(input_path, raport_file)
                dane_altera.append(search_nikon_altera(nominals_from_excell_list, res))
            dane_wyjsciowe[key].extend(dane_altera)

        if key == 'mitutoyo-crysta-apex':
            for raport_file in vals:
                res = extract_data_from_raport(input_path, raport_file)
                dane_crysta.append(search_mitutoyo_crysta_apex(nominals_from_excell_list, res))
            dane_wyjsciowe[key].extend(dane_crysta)

        if key == 'mitutoyo-reczna':
            for raport_file in vals:
                res = extract_data_from_raport(input_path, raport_file)
                dane_mitutoyo_reczna.append(search_mitutoyo_reczna(nominals_from_excell_list, res))
            dane_wyjsciowe[key].extend(dane_mitutoyo_reczna)

        if key == 'nikon-altera-rpt':
            for v in vals:
                raport_file = v
                res = extract_data_from_raport(input_path, raport_file)
                dane_altera_rpt.append(search_nikon_altera_rpt(nominals_from_excell_list, res))
            dane_wyjsciowe[key].extend(dane_altera_rpt)

    return dane_wyjsciowe
