def get_summary_data(cursor, date, ctry, unit):
        arguments = f"'{ctry}','{date}', '{unit}'"
        cursor.execute('exec [govDec_get_summarySlide]' + arguments)
        for row in cursor:
            feedbackreceived = row[1]
            rating12Nom = row[0]
            createdC = row[2]
            createdU = row[3]
            resolvedU = row[5]
            openworkload = row[6]
            ageing115 = row[7]
            ageing1530 = row[8]
            ageing3060 = row[9]
            ageing60180 = row[11]
            ageingover180 = row[12]
            otdNom = row[13]
            otdDenom = row[14]
            QCDoneNom = row[15]
            QCnaNom = row[16]
            QCpassedNom = row[17]
            QCfailedNom = row[18]
            QCDoneNomVol = row[19]
            QCnaNomVol = row[20]
            QCpassedNomVol = row[21]
            QCfailedNomVol = row[22]
        return feedbackreceived, rating12Nom, createdC, createdU, resolvedU, \
               openworkload, ageing115, ageing1530, ageing3060, ageing60180, \
               ageingover180, otdNom, otdDenom, QCDoneNom, QCnaNom, QCpassedNom, \
               QCfailedNom, QCDoneNomVol, QCnaNomVol, QCpassedNomVol, QCfailedNomVol
               
               
def get_cs_chart(cursor, sqldate, ctry, sLine, unit):
    months = []
    CSAT = []
    ResponseRate = []

    if sLine == "%":
        cursor.execute(f"exec [govDec_get_summaryCharts] '{sqldate}','{ctry}','pCSat'")

        for row in cursor:
            months.append(row[0])
            if row[2] == 0:
                CSAT.append(0)
            else:
                CSAT.append(round(row[1]/row[2] * 100, 2))
            if row[2] == 0:
                ResponseRate.append(0)
            else:
                ResponseRate.append(round((row[2]/row[3] * 100) ,1))

    else:
        cursor.execute(f"exec [govDec_get_slChartData] '{sqldate}','{ctry}','{sLine}', '{unit}', 'pCSat'")

        for row in cursor:
            months.append(row[0])
            if row[5] == 0:
                CSAT.append(0)
            else:
                CSAT.append(round((row[1] + row[2])/row[5] * 100, 2))
            if row[5] == 0:
                ResponseRate.append(0)
            else:
                ResponseRate.append(round((row[5]/row[6] * 100) ,1))

    CSAT = [x if x > 0 else "" for x in CSAT]
    return months, CSAT, ResponseRate

def get_otd_chart(cursor, sqldate, ctry, sLine, unit='%'):
    months = []
    OTD = []
    
    if sLine == "%":
        cursor.execute(f"exec [govDec_get_summaryCharts] '{sqldate}','{ctry}','pOtd'")
    else:
        cursor.execute(f"exec [govDec_get_slChartData] '{sqldate}','{ctry}','{sLine}', '{unit}','pOtd'")
        
    for row in cursor:
        months.append(row[0])
        if row[2] == 0:
            OTD.append(0)
        else:
            OTD.append(round(row[1]/row[2] * 100, 2))

    OTD = [x if x > 0 else "" for x in OTD]
    return months, OTD

def get_qc_chart(cursor, sqldate, ctry, sLine, unit='%'):
    months = []
    QC = []
    
    if sLine == "%":
        cursor.execute(f"exec [govDec_get_summaryCharts] '{sqldate}','{ctry}','pQC'")
        col = 2
    else:
        cursor.execute(f"exec [govDec_get_slChartData] '{sqldate}','{ctry}','{sLine}', '{unit}','pQC'")
        col = 3

    for row in cursor:
        months.append(row[0])
        if row[col] == 0:
            QC.append(0)
        else:
            QC.append(round(row[1]/row[col] * 100, 2))
    
    QC = [x if x > 0 else "" for x in QC]
    return months, QC

def get_ca_chart(cursor, sqldate, ctry, sLine, unit='%'):
    if sLine == "%":
        months = []
        CA = []
        cursor.execute(f"exec [govDec_get_summaryCharts] '{sqldate}','{ctry}','pCA'")
        for row in cursor:
            months.append(row[0])
            if row[2] == 0:
                CA.append(0)
            else:
                CA.append(round(row[1]/row[2] * 100, 2))
            
        CA = [x if x > 0 else "" for x in CA]
        return months, CA
    
    else:
        cursor.execute(f"exec [govDec_get_slChartData] '{sqldate}','{ctry}','{sLine}', '{unit}','pCA'")
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
            


def get_vl_chart(cursor, sqldate, ctry, sLine, unit):
    cursor.execute(f"exec [govDec_get_slChartData] '{sqldate}','{ctry}','{sLine}', '{unit}','pVL'")

    months = []
    created = []
    resolved = []

    for row in cursor:
        months.append(row[0])
        created.append(row[1])
        resolved.append(row[2])
    return months, created, resolved