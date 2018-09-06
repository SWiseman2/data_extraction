# -*- coding: utf-8 -*-
"""
Created on Thu Jan 25 19:14:58 2018

@author: slwis
"""
import Pep_data_extraction2 as ext
import numpy

import glob

dummy = "GRV12012180.xlsx"
dummy2 = "PEP12014040.xlsx"

path  = 'F:/PhD/Historical/RawCommData/iCrop/2014/PepsiCo-2014/*.xlsx' # locating all of the workbooks in this location
#you will need to switch between E and F drives depending on laptop vs desktop
files = glob.glob(path)



number_dates = 0  # starting a counter so that each new date file produced can be labelled uniquely
def write_datefile(data_source): #crop_info, jdates):
    folderName = ''.join(path.split(sep='/',)[-2:-1])
    Pep_dates = "historical_dates_%s.txt" %(folderName)
    with open(Pep_dates, "w") as g:  
        pass
    with open(Pep_dates, "a+") as g:
        for file in files:
            crop_info, jdates, allReps = ext.parse_file(file) #looks like ext.parse_file 
                                    #gives all the crop info, jdates and gc values therefore 
                                    #they all need to have structures to 'collect' them in
                                    #want to work out how to ask for a sbuset of the data
                                    #i.e. just crop_info and jdates
            
            
                #line = "cropname: {}".format(crop_info["cropName"])
            g.write(crop_info["cropName"]) 
            g.write("\t")
            g.write(crop_info["year"])
            g.write("\t")
            jdates21 = jdates + ["*"] * (21 - len(jdates)) # adding the 
            #appropriate number of *s to the end of the list to make it up to 21
            jdates_str = [str(x) for x in jdates21]
            data = "\t".join(jdates_str)
            g.write(data)
            g.write("\n")

#write_datefile(crop_info, jdates)
write_datefile(files)


def write_detailsfile(data_source):
    folderName = ''.join(path.split(sep='/',)[-2:-1])
    Pep_details = "historical_details_%s.txt" %(folderName)
    
    ExpNames = []
    
    with open (Pep_details,"w") as g:
        headings = "\t".join((["cropName","cropVar","NoR","year",
                    "pdate","emdate","met_ref","stem_den1",
                    "stem_den2","planting_density","plant_den2",
                    "seed_mass","seed_spacing","row_width",
                    "appN","Ground cover values"]))
        
        g.write(headings)
        g.write("\n")
        NoC = 0
    with open(Pep_details,"a+") as g:
        for file in files:
            NoC += 1 #counting number of crops
            crop_info, jdates, allReps = ext.parse_file(file) #looks like ext.parse_file 
                                    #gives all the crop info, jdates and gc values therefore 
                                    #they all need to have structures to 'collect' them in
                                    #want to work out how to ask for a sbuset of the data
                                    #i.e. just crop_info and jdates
            #averaging the values of ground cover 
            
            
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
                        
            GCdata = "\t".join("%.3f" % val for val in aveGCjd)
 
            
            CName = crop_info["cropName"]
            C_Name = CName.replace(" ", "_")
            print(CName, C_Name)
                        
            g.write(C_Name)
            g.write("\t")
            g.write(crop_info["cropVar"])
            g.write("\t")        
            g.write(str(crop_info["NoR"]))
            g.write("\t")        
            g.write(crop_info["year"])
            g.write("\t")        
            g.write(crop_info["pdate"])
            g.write("\t")        
            g.write(crop_info["emdate"])
            g.write("\t")        
            g.write(crop_info["met_ref"])
            g.write("\t")        
            g.write(str(round(crop_info["stem_den1"], 3))) #need to limit to 0dp
            g.write("\t")        
            if crop_info["stem_den2"] == "*":
                g.write(crop_info["stem_den2"])
            else:
                g.write(str(round(crop_info["stem_den2"], 3)))  #need to limit to 0dp
            g.write("\t") 
            g.write(str(crop_info["planting_density"]))
            g.write("\t")
            if crop_info["plant_den2"] == "*":
                g.write(crop_info["plant_den2"])
            else:
                g.write(str(round(crop_info["plant_den2"], 3)))  #need to limit to 0dp
            g.write("\t")       
            if crop_info["seed_mass"] == "*":
                g.write(crop_info["seed_mass"])
            else:
                g.write(str(round(crop_info["seed_mass"], 3))) #need to restrict to 1dp
            g.write("\t")
            if crop_info["seed_spacing"] == "*":
                g.write(crop_info["seed_spacing"])
            else:
                g.write(str(round(crop_info["seed_spacing"], 3))) #need to restrict to 0dp
            g.write("\t")        
            g.write(str(crop_info["row_width"]))
            g.write("\t")        
       
            g.write(str(crop_info["appN"]))
            g.write("\t")
            g.write(GCdata)
            g.write("\n")
            print(file)
            
            ExpNames.append(crop_info["cropName"])

        g.write(str(NoC))
          
    with open("ExperimentNames.txt", 'w') as EN: 
        EN.write(str(ExpNames))
write_detailsfile(dummy)




