from multiprocessing import Lock, get_context, set_start_method

import time

from multiprocessing import Pool, Process

import CtryDecks
from countries import countries, Centers

import HRops
import HRCenters

selected_ctry = [str(key) for key in countries.keys()]
selected_HRCenter = [str(key) for key in Centers.keys()]

if __name__ == '__main__':  
    t1 = time.perf_counter()

    HRops.CreateReport()
    
    print("Number of selected countries: ", len(selected_ctry))
    
    for country in selected_ctry:
        CtryDecks.CreateReport(country)

    for HRCenter in selected_HRCenter:
        HRCenters.CreateReport(HRCenter)
    
        
    # from concurrent.futures import ThreadPoolExecutor
    # with ThreadPoolExecutor(max_workers = 3) as executor:
    #     executor.map(CtryDecks.CreateReport, selected_ctry)

    # with Pool(8) as p:
    #     p.map(CtryDecks.CreateReport, selected_ctry)



    t2 = time.perf_counter()
    print('-' * 50)
    print(f'Code Took:{t2 - t1} seconds. That is {round((t2 - t1), 2) / (len(selected_ctry) + 5)} seconds per presentation on average')

    #CtryDecks.cursor.close()
    