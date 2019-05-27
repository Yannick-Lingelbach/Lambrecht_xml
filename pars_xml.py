# -*- coding: utf-8 -*-
"""
Created on Mon Apr 15 09:32:51 2019

@author: LYI9FE
"""

import pandas as pd
from pandas import ExcelWriter
from os import listdir
import datetime as DT
from xml.dom import minidom

# =============================================================================
# Set Parameters
# =============================================================================

#Set Path to Datafolder
path_xml='//bosch.com/dfsrb/DfsDE/LOC/Fe/FeP/QUER/TEF/TEF1-COS/TEF1-COS-Transfer/Transfer_RB/Lambrecht/31_E/'

path_Excel='//bosch.com/dfsrb/DfsDE/LOC/Fe/FeP/QUER/TEF/TEF1-COS/TEF1-COS-Transfer/Transfer_RB/Lambrecht/'
name_Excel = '31_E.xlsx'
    
# =============================================================================
# Define parser funtion
# =============================================================================


def read_xml(path_xml):    
    # Get the Nodes defined in the XML file
    Furn_data = minidom.parse(path_xml)
    
    # =============================================================================
    # Read the Header 
    # =============================================================================
    
    header = Furn_data.getElementsByTagName("header")[0]
    header_loc = Furn_data.getElementsByTagName("location")[0]
    
    timeStam = header.getAttribute('timeStamp')
    contType = header.getAttribute('contentType')
    eventNam = header.getAttribute('eventName')
    eventId  = header.getAttribute('eventId')
    
    locAppli = header_loc.getAttribute('application')
    procName = header_loc.getAttribute('processName')
    procNo = header_loc.getAttribute('processNo')
    lineNo = header_loc.getAttribute('lineNo')
    
    # =============================================================================
    # Read the Body -> Structs and Values 
    # =============================================================================
    
    body = Furn_data.getElementsByTagName("body")[0]
    body_struct = body.getElementsByTagName("structs")[0].getElementsByTagName("resHead")[0]
    body_values = body.getElementsByTagName("values")[0]
    
    typeNo =  body_struct.getAttribute('typeNo')
    
    TOTAL = body_values.getElementsByTagName("item")[0]
    ROT_TABLE_TURN = body_values.getElementsByTagName("item")[1]
    TOTAL_WP132 = body_values.getElementsByTagName("item")[2]
    TOTAL_WP133 = body_values.getElementsByTagName("item")[3]
    
    TOTAL =  TOTAL.getAttribute('value')
    ROT_TABLE_TURN =  ROT_TABLE_TURN.getAttribute('value')
    TOTAL_WP132 =  TOTAL_WP132.getAttribute('value')
    TOTAL_WP133 =  TOTAL_WP133.getAttribute('value')
    
    return {'timeStamp': DT.datetime.strptime(timeStam[:23], "%Y-%m-%dT%H:%M:%S.%f"),
            'contentType': int(contType),
            'eventName':eventNam,
            'eventId':int(eventId),
            'application':locAppli,
            'processName': procName,
            'processNo':int(procNo),
            'lineNo': int(lineNo),
            'typeNo': typeNo,
            'TOTAL': int(TOTAL),
            'ROT_TABLE_TURN' : int(ROT_TABLE_TURN),
            'TOTAL_WP132' : int(TOTAL_WP132),
            'TOTAL_WP133' : int(TOTAL_WP133)}

# =============================================================================
# Write to Dataframe
# =============================================================================


# Extract the Column names and the corresponding measurement Unit
col_Names=['timeStamp', 'contentType', 'eventName', 'eventId', 'application', 
           'processName', 'processNo', 'lineNo', 'typeNo', 'TOTAL', 
           'ROT_TABLE_TURN', 'TOTAL_WP132', 'TOTAL_WP133']

# Extract the timeseries data and save it in a dataframe
df = pd.DataFrame(columns=col_Names)

lst_xml_files=[]
lst_xml_files= listdir(path_xml)


for i,file in enumerate(lst_xml_files):
    try:
        df = df.append(read_xml(path_xml+file),ignore_index=True)
    except:
        pass
    
    print(str(i+1)+ ' from ' + str(len(lst_xml_files)))
    
    if i>1000:
        break
    
df.sort_values('timeStamp', inplace=True)
df.set_index('timeStamp', inplace=True)

writer = ExcelWriter(path_Excel+name_Excel)
df.to_excel(writer,'Sheet1')
writer.save()





    
    