# -*- coding: utf-8 -*-
"""
Created on Wed Mar 11 14:22:02 2020

@author: sayers
"""

from emailautosend import mailthat
import os
from cleansheet import cleansheet
import pandas as pd
from admin import newest, rehead, colclean


path = "S:\\Downloads\\"     # Give the location of the files
fname = "JOB_HIST"         # Give filename prefix
#getting the newest of these files in the directory and converting to df
#stripping out the 2 metadata columns in CJR files
#standardizing the column names
def multicampus_file(infolder,fname,outfolder):
    df=colclean(rehead(pd.read_excel(newest(infolder,fname)),2))
    multicampus = df[df.id.isin(df[~(df.unit=="YRK01")][df.pay_status.isin(['A','W','P','L'])].id.unique())]
    outfile=os.path.join(outfolder,"multicampus.xls")
    multicampus.to_excel(outfile)
    return(multicampus,outfile)
    
def communications(fname,distlist):
    subj='Multicampus report'
    body= f'\n<html>\n<head>\n<p>Good Day, </p>\n<p> </p>\n<p>Attached find the multicampus employee file.</p>\n<p> </p>\n<p>Best Regards,</p>\n<p>Shane Ayers</p>\n<p>Human Resources Information Systems Manager</p>\n<p>Office of Human Resources</p>\n<p>York College</p>\n<p>The City University of New York</p>\n</body></html>\n'
    mailthat(subj,to=distlist,html=body,atch=fname)
    
def main(infolder,outfolder):
    fname="JOB_HIST"
    file=multicampus_file(infolder,fname,outfolder)
    cleansheet(file)
    dists='some set; of peoples; emails'
    #TODO replace this with a read_json call to an existing file
    communications(file,dists)


if "__name__"=="__main__":
    main("s:\\downloads","y:\\reports")
    