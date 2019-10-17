import xlrd
import numpy as np
import pandas as pd


#competition distances [miles]
TOTAL_SWIM = 2.4
TOTAL_BIKE = 112
TOTAL_RUN = 26.2

#estimated total workout time [hrs]
SWIM_TIME = 1.5
BIKE_TIME = 6.5
RUN_TIME = 3.5

file = 'Illini Triathlon Challenge.xlsx'
wb = xlrd.open_workbook(file)
df = pd.DataFrame(columns=['athlete', 'swim_mi', 'bike_mi', 'run_mi', 'time_left_hrs'])

for index in np.arange(wb.nsheets):
    sheet = wb.sheet_by_index(index)
    athlete = sheet.name
    swim = sheet.cell_value(27,16)/4200*2.4 #convert to a fraction of 2.4 mi (actual conversion is 1760 yd/mi)
    bike = sheet.cell_value(28,16)
    run = sheet.cell_value(29,16)
    
    time_left = (
        max((TOTAL_SWIM - swim)/TOTAL_SWIM * SWIM_TIME, 0) +
        max((TOTAL_BIKE - bike)/TOTAL_BIKE * BIKE_TIME, 0) + 
        max((TOTAL_RUN  - run )/TOTAL_RUN  * RUN_TIME, 0)
                )
    
    df = df.append({
        'athlete': athlete,
        'swim_mi': swim,
        'bike_mi': bike,
        'run_mi': run,
        'time_left_hrs': time_left
        }, ignore_index=True)

df_rank = df.sort_values('time_left_hrs', ascending=True).reset_index(drop=True)
print(df_rank.round(2))