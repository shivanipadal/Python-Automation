# -*- coding: utf-8 -*-
"""
Created on Tue Jun 30 16:16:15 2020


"""

import sys
import os 
import xlrd
import pandas as pd
import openpyxl as xl
import datetime
import time
import argparse

dict1={}
count=0
'''
def convert(seconds) :
    seconds = seconds % (24 * 3600)
    hour = seconds // 3600
    seconds %= 3600
    minutes = seconds // 60
    seconds %= 60
    print(type(hour))
    print(type(seconds))
    print(type(minutes))
    
    return "%d:02d:%02d" % (hour,minutes,seconds)
'''

#def convert(n):
#    return str(datetime.timedelta(seconds=n))


def sumHours(Hours):
 totalSecs = 0 
 for tm in Hours: 
   timeParts = [int(s) for s in tm.split(':')] 
   totalSecs += (timeParts[0] * 60 + timeParts[1]) * 60 + timeParts[2] 

 totalSecs, sec = divmod(totalSecs, 60) 
 hr, min = divmod(totalSecs, 60) 
 return "%d:%02d:%02d" % (hr, min, sec)

def Process(indir,outdir):
  fileformatt=time.strftime("%Y%m%d-%H%M%S")
  outfile=outdir + "output_"+fileformatt+ ".csv"  
  df=pd.DataFrame(columns=['Name of the user','Source User','TotalSeconds','First Login','Last Logout','Total Hours','Date'])
  #df.columns=['Name of the user','Source User','TotalSeconds','min1','max1','Date']
  df.to_csv(outfile,index=False)
  
  
# if os.path.isfile('./output.csv'):
#  if os.path.isfile(outfile):
#    print("Output.csv alredy exist. So exiting...")
#    sys.exit()
    
  systems=['ABC','DEF','FGH','HJO']
  for root,dirs,files in os.walk(indir,topdown=False):
    for filename in files:
        if filename.endswith('.xlsx') and filename.startswith('VPN_Users_Login_Logout_Report'):
            #print(filename)
            #data = pd.ExcelFile(filename)
            filename=indir+"/"+filename
            print(filename)
            data=pd.read_excel(filename,'Sheet1')
            datetime1 = data['Generate Time'][1].date()
            time1=datetime1.month
            print(time1)
            print(datetime1)
            
            data = data[data['Source User'].isin(systems)]
        
            data1 = data.groupby(['Name of the user','Source User']).agg(
                    TotalSeconds=('Login_duration in seconds',sum),
                    min1=('Generate Time',min),
                    max1=('Generate Time',max),
                    #Hours=('Login_duration in seconds',convert(sum))
                    totalHours=('Duration in Hours',sumHours)
                    #Date=datetime1
                    )
            #data1['Hours']=str(datetime.timedelta(seconds=data1['TotalSeconds']))
            
            #data1['Hours']=convert(data1[TotalSeconds])
            #data1['Hours']=pd.to_datetime(data1['TotalSeconds'],unit='d')
            

            print(data1)                     
            data1['Date']=datetime1
        
            df1 = pd.DataFrame(data1)
                        
            df1.to_csv(outfile,mode='a',header=False)
            
def parse_args():
   """Create the arguments"""
   parser = argparse.ArgumentParser('\nWorking_hours.py --inputdir INPUTDIR --outputdir OutPUTDIR')
   parser.add_argument("-indir", "--inputdir", help="input dir")
   parser.add_argument("-outdir", "--outputdir", help="Output dir")     
   return parser.parse_args()    

def main(args):
    if args.inputdir :
        indir=args.inputdir
    else :
        indir="."
      
    if args.outputdir :    
        outdir=args.outputdir
    else :
        outdir="./"
    
    Process(indir,outdir)
    

if __name__ == "__main__":
   main(parse_args())   
