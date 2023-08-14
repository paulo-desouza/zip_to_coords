from excel_methods import *
import plotly.graph_objects as go
from dash import Dash, dcc, html
from urllib.request import urlopen
import json
import pandas as pd
import plotly.express as px


### ZIP TO LAT // LONG COORDINATES

## USES STATES DTAA TO SEARCH MORE EFFICIENTLY
## ZIP NEEDS TO MATCH STATE

 

wb = load_workbook('all_leads.xlsx')
ws = wb.active

ld = {}

for i, row in enumerate(ws):
    
    
    ld[ws["G"+str(i+6)].value] = [ws["B"+str(i+6)].value,
                                  ws["C"+str(i+6)].value,
                                  ws["D"+str(i+6)].value,
                                  ws["F"+str(i+6)].value,
                                  ws["H"+str(i+6)].value]
    


ds = {}

cds = {}

states = get_file_names()


for i, s in enumerate(states):
    
    ds[s] = {}
    
    
for i, key in enumerate(ld.keys()):
    
    if key == None:
        
        continue
    
    for _key in ds.keys():
    
        if ld[key][3] in _key:
            
            ds[_key][key] = ld[key]
        



for k, v in ds.items():
    
    zips = load_workbook(f'.\\States\\{k}')
    wsz = zips.active

    for key in v.keys():
    
        for i, row in enumerate(wsz):
            if str(key) == f"{wsz[f'A{i+1}'].value:05}":
                
                v[key].append(wsz[f'B{i+2}'].value)
                v[key].append(wsz[f'C{i+2}'].value)
                
                cds[key] = v[key]
                
                
                
                
wb = Workbook()
ws = wb.active 

for i, row in enumerate(cds.keys()):
    cds[row].append(row)
    ws.append(cds[row])
    
wb.save("result.xlsx")

# we still need to extract the dirty data 
                
                
                
                
                

                

                



            

            