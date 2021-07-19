import pyodbc

def get_summary_data(cursor, sqldate, unit):
    arguments = f"'{sqldate}','{unit}'"
    cursor.execute('exec [HRmr_get_summarySlide]' + arguments)
    for row in cursor:
        feedbackreceived = row[1]
        rating12Nom = row[0]
        createdC = row[2]
        createdU = row[3]
        resolvedC = row[4]
        resolvedU = row[5]
        openworkload = row[12]
        ageing115 = row[13]
        ageing1530 = row[14]
        ageing3060 = row[15]
        ageing60180 = row[17]
        ageingover180 = row[18]
        otdNom = row[19]
        otdDenom = row[20]
        QCDoneNom = row[21]
        QCnaNom = row[22]
        QCpassedNom = row[23]
        QCfailedNom = row[24]
        QCDoneNomVol = row[25]
        QCnaNomVol = row[26]
        QCpassedNomVol = row[27]
        QCfailedNomVol = row[28]
    return feedbackreceived, rating12Nom, createdC, createdU, resolvedC,resolvedU, \
           openworkload, ageing115, ageing1530, ageing3060, ageing60180, \
           ageingover180, otdNom, otdDenom, QCDoneNom, QCnaNom, QCpassedNom, \
           QCfailedNom, QCDoneNomVol, QCnaNomVol, QCpassedNomVol, QCfailedNomVol


def get_query_info(name):
    if name in ('BGL', 'KRK', 'SPL', 'TLL'):
        typeN = 'pHub'
    elif name in ('CSE', 'NEU', 'NAM', 'SAM', 'NEA', 'MEA', 'SEA'):
        typeN = 'pRegion'
    elif name in ('TA_', 'TM_', 'GM_', 'PA_', 'ELC', 'CB_', 'PY_', 'LD_'):
        typeN = 'pSline'
        if name == 'ELC':
            name = 'ELCM'
        else:
            name = name[0:2]
    return name, typeN


def get_cs_chart(cursor, sqldate, name):
    name, typeN = get_query_info(name)
    cursor.execute(f"exec [HRmr_get_CS_chart] '{sqldate}','{name}','{typeN}'")

    months = []
    VeryGood = []
    Good = []
    Dissatisfied = []
    HighlyDissatisfied = []
    ResponseRate = []

    for row in cursor:
        months.append(row[0])
        VeryGood.append(row[1])
        Good.append(row[2])
        Dissatisfied.append(row[3])
        HighlyDissatisfied.append(row[4])
        if row[6] == 0:
            ResponseRate.append(0)
        else:
            ResponseRate.append(round((row[5]/row[6] * 100) ,1))
    
    VeryGood = [x if x > 0 else "" for x in VeryGood]
    Good = [x if x > 0 else "" for x in Good]
    Dissatisfied = [x if x > 0 else "" for x in Dissatisfied]
    HighlyDissatisfied = [x if x > 0 else "" for x in HighlyDissatisfied]
    return months, VeryGood, Good, Dissatisfied, HighlyDissatisfied, ResponseRate


def get_otd_chart(cursor, sqldate, name):
    name, typeN = get_query_info(name)
    cursor.execute(f"exec [HRmr_get_OTD_chart] '{sqldate}','{name}','{typeN}'")

    months = []
    data = []

    for row in cursor:
        months.append(row[0])
        if row[2] == 0:
            data.append(0)
        else:
            data.append(round(row[1] / row[2] * 100, 2))
    
    
    return months, data


def get_qc_chart(cursor, sqldate, name):
    name, typeN = get_query_info(name)
    cursor.execute(f"exec [HRmr_get_QC_chart] '{sqldate}','{name}','{typeN}'")

    months = []
    data = []

    for row in cursor:
        months.append(row[0])
        if row[3] == 0:
            data.append(0)
        else:
            data.append(row[1] * 100 / row[3])
    return months, data


def get_vl_chart(cursor, sqldate, name):
    name, typeN = get_query_info(name)
    cursor.execute(f"exec [HRmr_get_VL_chart] '{sqldate}','{name}','{typeN}'")

    months = []
    created = []
    resolved = []

    for row in cursor:
        months.append(row[0])
        created.append(row[1])
        resolved.append(row[2])
    return months, created, resolved


def get_ca_chart(cursor, sqldate, name):
    name, typeN = get_query_info(name)
    cursor.execute(f"exec [HRmr_get_CA_chart] '{sqldate}','{name}','{typeN}'")
    
    months = []
    active115 = []
    active1530 = []
    active3060 = []
    activeover60 = []
    onhold115 = []
    onhold1530 = []
    onhold3060 = []
    onhold60 = []
    longrunning = []
    
    for row in cursor:
        months.append(row[0])
        active115.append(row[1])
        active1530.append(row[2])
        active3060.append(row[3])
        activeover60.append(row[4])
        onhold115.append(row[5])
        onhold1530.append(row[6])
        onhold3060.append(row[7])
        onhold60.append(row[8])
        longrunning.append(row[9])
    
    t115 = [None]*(len(active115)+len(onhold115))
    t115[::2] = active115
    t115[1::2] = onhold115
    
    t1530 = [None]*(len(active115)+len(onhold115))
    t1530[::2] = active1530
    t1530[1::2] = onhold1530
    
    t3060 = [None]*(len(active115)+len(onhold115))
    t3060[::2] = active3060
    t3060[1::2] = onhold3060
    
    to60 = [None]*(len(active115)+len(onhold115))
    to60[::2] = activeover60
    to60[1::2] = onhold60
    
    lr = [0] * (len(longrunning) * 2)
    lr[1::2] = longrunning
    
    t115 = [x if x > 0 else "" for x in t115]
    t1530 = [x if x > 0 else "" for x in t1530]
    t3060 = [x if x > 0 else "" for x in t3060]
    to60 = [x if x > 0 else "" for x in to60]
    lr = [x if x > 0 else "" for x in lr]
    
    return months, t115, t1530, t3060, to60, lr 


def get_calt_chart(cursor, sqldate, name):
    name, typeN = get_query_info(name)
    cursor.execute(f"exec [HRmr_get_CALT_chart] '{sqldate}','{name}','{typeN}'")
    
    months = []
    ag120150 = []
    ag150180 = []
    agover180 = []
    
    for row in cursor:
        months.append(row[0])
        ag120150.append(row[1]* 100)
        ag150180.append(row[2] * 100)
        agover180.append(row[3] * 100)
    
    ag120150 = [x if x > 0 else "" for x in ag120150]
    ag150180 = [x if x > 0 else "" for x in ag150180]
    agover180 = [x if x > 0 else "" for x in agover180]  
    
    return months, ag120150, ag150180, agover180