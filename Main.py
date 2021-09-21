from multiprocessing import Lock, get_context, set_start_method

import time
import os
import shutil

from multiprocessing import Pool, Process

import CtryDecks
from countries import countries, Centers

import HRops

selected_ctry = [str(key) for key in countries.keys()]
selected_HRCenter = [str(key) for key in Centers.keys()]

if __name__ == '__main__':  
    t1 = time.perf_counter()

    HRops.CreateReport()
    
    print("Number of selected countries: ", len(selected_ctry))

    if not os.path.isdir(f"{os.getcwd()}\\Created\\"):
        os.mkdir(f"{os.getcwd()}\\Created\\")
        os.mkdir(f"{os.getcwd()}\\Created\\GovDecks\\")

    if not os.path.isdir(f"{os.getcwd()}\\Created\\GovDecks\\"):
        os.mkdir(f"{os.getcwd()}\\Created\\GovDecks\\")

    for country in selected_ctry:
        template_file = (f"{os.getcwd()}\\templates\\Monthly_Ctry_Governance_TEMPLATE_2.0.pptx")
        dest_dir = f"{os.getcwd()}\\Created\\GovDecks\\"
        shutil.copy(template_file, dest_dir)
        os.rename(os.path.join(dest_dir,'Monthly_Ctry_Governance_TEMPLATE_2.0.pptx'), os.path.join(dest_dir,f"{countries[country]}.pptx"))
    
    for country in selected_ctry:
        CtryDecks.CreateReport(country)

    # from concurrent.futures import ThreadPoolExecutor
    # with ThreadPoolExecutor(max_workers = 3) as executor:
    #     executor.map(CtryDecks.CreateReport, selected_ctry)

    # with Pool(8) as p:
    #     p.map(CtryDecks.CreateReport, selected_ctry)

        

    t2 = time.perf_counter()
    print('-' * 50)
    print(f'Code Took:{t2 - t1} seconds. That is {round((t2 - t1), 2) / (len(selected_ctry) + 1)} seconds per presentation on average')

    #CtryDecks.cursor.close()
    