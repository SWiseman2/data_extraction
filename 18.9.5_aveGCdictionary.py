# -*- coding: utf-8 -*-
"""
Created on Fri Aug 31 12:32:02 2018

@author: x990069
"""

import Pep_data_extraction2 as ext

import glob
dummy2 = "PEP12014040.xlsx"

path  = 'F:/PhD/Historical/RawCommData/iCrop/2014/PepsiCo-2014/*.xlsx' # locating all of the workbooks in this location
#you will need to switch between E and F drives depending on laptop vs desktop
files = glob.glob(path)

folderName = ''.join(path.split(sep='/',)[-2:-1])
Pep_details = "historical_details_%s.txt" %(folderName)
    
with open(Pep_details,"a+") as g:
     for file in files:
         crop_info, jdates, allReps = ext.parse_file(file) #looks like ext.parse_file 
         #jdates = ext.parse_file(file)
         #allReps = ext.parse_file(file) 
         #gives all the crop info, jdates and gc values therefore 
         #they all need to have structures to 'collect' them in
         #want to work out how to ask for a sbuset of the data
         #i.e. just crop_info and jdates
         type(crop_info)
         numberDays = crop_info['NoD']
         aveGCjd = []
         for GCindex in range(numberDays):
             GCvalues = []
             allGCaves = []
             for key in allReps:
                 GCvalues.append(allReps[key][GCindex])
             print(GCvalues)
             GCave = sum(GCvalues)/len(GCvalues)
             print(GCave)
             aveGCjd.append(GCave)
         print(aveGCjd)
         
g.write(GCdata)