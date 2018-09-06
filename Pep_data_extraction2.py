# -*- coding: utf-8 -*-

import os
import numpy as np
import openpyxl
from openpyxl import load_workbook



def parse_file(file_name):#defining where the data will be extracted from in 
              #original worksheets, this will actually be carried out later

    with open(file_name, "rb") as my_file:# Read in worksheet from workbook in one folder
             workbook=load_workbook(my_file, data_only=True) #data_only gets 
                                #openpyxl to evaluate the formaul in the cells
             wsheet=workbook.active
    
    if wsheet["E67"].value is None:
        aveSeed = "*"
        rangeSeed = "*"
    else:
        lowerBSeed = wsheet["E67"].value #read in lower bound of seed size
        upperBSeed = wsheet["G67"].value # read in upper bound of seed size
        rangeSeed = "%s_%s" % (lowerBSeed,upperBSeed) #formatting seed size ranges
        aveSeed = (lowerBSeed + upperBSeed)/2
                  
    if wsheet["E68"].value is None: #checking that there is a seed mass value 
        seedmass = "*"
    else:
        seedmass = 50000/wsheet["E68"].value     #calculating average seed mass
        
    if wsheet["E37"].value is None:
        seed_spacing = "*"
    elif wsheet["E37"].value == "#DIV/0!":
        seed_spacing = "*"
    else:
        seed_spacing = wsheet["E37"].value
    
    #stem density calculations
    eStemR1 = wsheet["D340"].value    #number of stems for first rep (first harvest)
    eStemR2 = wsheet["D341"].value     #number of stems for second rep (first harvest) 
    eStemR3 = wsheet["D342"].value     #number of stems for third rep (first harvest)
    eStemR4 = wsheet["D343"].value     #number of stems for fourth rep (first harvest)
    eStemR5 = wsheet["D344"].value     #number of stems for fifth rep (first harvest)
    eStemR6 = wsheet["D345"].value     #number of stems for sixth rep (first harvest)
    ePlantR1 = wsheet["C340"].value    #number of plants in sample (first harvest)
    ePlantR2 = wsheet["C341"].value      #number of plants in sample (first harvest)
    ePlantR3 = wsheet["C342"].value   #number of plants in sample (first harvest)
    ePlantR4 = wsheet["C343"].value   #number of plants in sample (first harvest)
    ePlantR5 = wsheet["C344"].value   #number of plants in sample (first harvest)
    ePlantR6 = wsheet["C345"].value   #number of plants in sample (first harvest)

    # checking that all values are real (and not doing the caluclation if there's nothing there)                     
    earlystem_info = [eStemR1, eStemR2, eStemR3, eStemR4, eStemR5, eStemR6, 
                      ePlantR1, ePlantR2, ePlantR3, ePlantR4, ePlantR5, ePlantR6]
    earlystem_inf = ([x for x in earlystem_info if x is not None])

    stemValues = int((len(earlystem_inf))/2)
    eStems = earlystem_inf[:stemValues]
    ePlants = earlystem_inf[stemValues:]    
    HarvestArea = wsheet["D333"].value * wsheet["G333"].value * wsheet["E25"].value/100
    ConversFac = 10 / HarvestArea #should be 10,000 but units are 000/stems or plants
    earlystems = np.mean(eStems) * ConversFac # number of 000 stems per hectare
    earlyplants = np.mean(ePlants) * ConversFac # number of 000 plants per hectare

     #late stem density calculations
    lStemR1 = wsheet["D369"].value    #number of stems for first rep (second harvest)
    lStemR2 = wsheet["D370"].value     #number of stems for second rep (second harvest) 
    lStemR3 = wsheet["D371"].value     #number of stems for third rep (second harvest)
    lStemR4 = wsheet["D372"].value     #number of stems for fourth rep (second harvest)
    lStemR5 = wsheet["D373"].value     #number of stems for fifth rep (second harvest)
    lStemR6 = wsheet["D374"].value     #number of stems for sixth rep (second harvest)
    lPlantR1 = wsheet["C369"].value    #number of plants in sample (second harvest)
    lPlantR2 = wsheet["C370"].value      #number of plants in sample (second harvest)
    lPlantR3 = wsheet["C371"].value   #number of plants in sample (second harvest)
    lPlantR4 = wsheet["C372"].value   #number of plants in sample (second harvest)
    lPlantR5 = wsheet["C373"].value   #number of plants in sample (second harvest)
    lPlantR6 = wsheet["C374"].value   #number of plants in sample (second harvest)
    
    if lStemR1 is None: # making sure that there is some data present
        latestems = "*"
        lateplants = "*"
    else:
        # checking that all values are real (and not doing the caluclation if there's nothing there)                     
        latestem_info = [lStemR1, lStemR2, lStemR3, lStemR4, lStemR5, lStemR6, 
                          lPlantR1, lPlantR2, lPlantR3, lPlantR4, lPlantR5, lPlantR6]
        latestem_inf = ([x for x in latestem_info if x is not None])
    
        stemValues = int((len(latestem_inf))/2)
        lStems = latestem_inf[:stemValues]
        lPlants = latestem_inf[stemValues:]    
        HarvestArea = wsheet["D362"].value * wsheet["G362"].value * wsheet["E25"].value/100
        ConversFac = 10 / HarvestArea #should be 10,000 but units are 000/stems or plants
        latestems = np.mean(lStems) * ConversFac # number of 000 stems per hectare
        lateplants = np.mean(lPlants) * ConversFac # number of 000 plants per hectare
    
    if wsheet["F108"].value is None: #reading in value of total intended fertilizer 
        nfertilizer = "*"
    else:
        nfertilizer = wsheet["F108"].value

    
   
    #reading julian dates from data file
    jdates = []
    
    jdcells = ["E143","E191","F191","G191","H191","I191","J191","K191",\
               "L191","M191","D200","E200","F200","G200","H200","I200",\
               "J200","K200","L200","M200"]
    NoV = 0
    for cell in jdcells:
        try: # this try stops python from trying to read in values which aren't there in the datasheet
            jdcel = wsheet[cell].value.strftime("%j")
        except AttributeError:
            continue  # sometimes the cells are empty and this is fine
        jdcell = int(jdcel) #converting to an interger
        if jdcell > 0:
            NoV +=1 
            jdates.append(jdcell)
        else:
            pass
    print('NoV',NoV)
    
    
    #reading ground cover values from data file
    reps = [1, 2, 3, 4, 5, 6]
    allReps = {}
    for rep in reps:
        #generating the lists of cell names
        rowRef1 = []
        rowRef2 = []
        for i in range(68,78):
            rowRef1.append('{letter}19{number}'.format(letter = chr(i), number = (rep+1)))
            rowRef2.append('{letter}20{number}'.format(letter = chr(i), number = (rep)))
        print(rowRef1,
              rowRef2)    
         
        gcRepCells = rowRef1 + rowRef2 #list of all possible cells with GC values in them 

        gcrep = []
        NoGC = 0
        for cell in gcRepCells:
            gccel = wsheet[cell].value
            if gccel is not None:
                gccell = int(gccel)
                repNo = "GC%s" %(rep)
                print(repNo, gccell)
                NoGC +=1
                gcrep.append(gccell)
            else:
                break
        if len(gcrep) < len(jdates):
            for i in range(len(jdates)-len(gcrep)):
                gcrep.append("*")    
        print(gcrep)
        print(NoGC)
        numberReps = len(allReps)
        if NoGC > 0:               #ensuring that empty reps aren't included
            allReps[repNo] = gcrep #and don't overwrite reps with content
    print("end print", allReps)

    print("number of REPs",numberReps)
     
    if wsheet["J33"].value is None:
        Year = wsheet["I1"].value.strftime("%Y")
    else:
        Year = wsheet["J33"].value.strftime("%Y")
    
    #checking values for planting date 
    if wsheet["J33"].value is None:
        PDate = "*"
    else:
        PDate = wsheet["J33"].value.strftime("%j")
        
    crop_info = {"cropName": wsheet["D1"].value, #crop reference name/number
                  "cropVar": wsheet["E31"].value, #potato variety 
                  "NoD": NoV,     #number of days in each indiv crop
                  "NoR": numberReps,      #number of reps 
                      #(sequential, if entered rep1+, rep2-, rep3+ would count as 1) - not sure what I meant here....
                  "year": Year, #reads in year from pdate 
                  "pdate": PDate, #reads planting date and converts to Julian Date
                  "emdate": wsheet["E143"].value.strftime("%j"),#ditto emergence 
                  "met_ref": wsheet["F59"].value, #reads met data reference
                  "stem_den1": earlystems, #including early stem density
                  "stem_den2": latestems, #inc late stem density
                  "plant_den2": lateplants,
                  "seed_mass": seedmass, #including seed mass in the dictionary
                  "seed_grade": rangeSeed, #inc raw seed garde
                  "ave_seed_grade": aveSeed, #inc middle point of seed grade
                  "seed_spacing": seed_spacing,
                  "row_width": wsheet["E25"].value,#reads in value of row width 
                  "planting_density": earlyplants,
                  "appN": nfertilizer #read in applied nitrogen
                  } 
    
       
    
    return (crop_info, jdates, allReps) # returing output as a tuple not a list
                                        # so it remains accessible

