# -*- coding: utf-8 -*-
"""
Created on Mon Feb 28 11:07:23 2022

@author: ktt672
"""

#import packages
import pandas as pd
import os
import msoffcrypto
import io
import numpy as np

#Find all files, decrypt and concat
df = None
password = input('Insert your common password here:')
for i in os.listdir():
    try:
        temp = io.BytesIO()
        with open(i, 'rb') as f:
            excel = msoffcrypto.OfficeFile(f)
            excel.load_key(password)
            excel.decrypt(temp)
            try:
                if df == None:
                    df = pd.read_excel(temp)
            except:
                df = pd.concat([df, pd.read_excel(temp)])
    except:
        continue
    del temp

#Desensitize Name
try:
    del df['Name']
except:
    pass


#####Add your own criteria here##### 
criteria = {'Specimen': input('\nInput your specimen criteria. All specimens containing the word entered will retain in calculation:\nEnter here:'), 'Specialty': input('\nInput your specialty criteria. All specialty containing the word entered will retain in calculation:\nEnter here:'), 'TestDesc': input('\nInput your Test criteria. All tests containing the word entered will retain in calculation:\nEnter here:')}
for i in criteria:
    if criteria[i] == '':
        pass
    else:
        df = df[df[i].str.contains(criteria[i], case = False, regex = False, na=False)]
        
###Customised query
try:
    df = df.query(input("\nPlease enter your customised query using Python syntax.\nEnter here:"))
except:
    print("No query/query invalid. Will skip.")



#Remove [ESBL], [MRSA]... 
df = df.replace(to_replace = r'\[.*\] (.*)', value = r'\1', regex = True)

#Sort by values
df = df.sort_values(by=['CollectDate', 'HN', 'LabNo', 'OrganismSeq'])

#Group by and select the first entry
df_enq = df.groupby(['HN', 'Organism']).head(1)



#total count
df_enq_total = pd.pivot_table(df_enq, index = 'Organism', aggfunc = 'count')
#Total count remove unrelated counts
index_no = df_enq.columns.get_loc('OrganismSeq')
remove_col_list = df_enq.columns.values.tolist()[0:index_no+1]
remove_col_list.remove('Organism')
remove_col_list.remove('LabNo')
for i in remove_col_list:
    del df_enq_total[i]

#Find values in antibiotics results, and let user select the nominator of values
value_set = set(df_enq.iloc[:, index_no+1:].values.ravel())
value_selected = input('\nEnter the single result you wish as an nominator of antibiogram (eg S or I). Default will be S.\n Your input here:') or 'S'
value_set.remove(value_selected)
#Count Sensitives
df_enq_nominator = df_enq.copy()
for i in df_enq_nominator.columns:
    df_enq_nominator.loc[df_enq_nominator[i] == value_selected, i] = 1
    df_enq_nominator.loc[df_enq_nominator[i].isin(value_set), i] = 0
df_enq_nominator.fillna(0,inplace=True)
df_enq_nominator = pd.pivot_table(df_enq_nominator, index = "Organism", aggfunc = 'sum', dropna=False)

for i in remove_col_list:
    try:
        del df_enq_nominator[i]
    except:
        continue

#putting all together and calculate the antibiogram
for i in df_enq_total.columns.values.tolist():
    for j in df_enq_total.index.tolist():
        try:
            df_enq_total[i][j] = '{}/{} ({}%)'.format(df_enq_nominator[i][j], df_enq_total[i][j], int(round(df_enq_nominator[i][j]/df_enq_total[i][j]*100, 0)))
        except:
            continue

#order by organism count
df_enq_total = df_enq_total.sort_values('LabNo', ascending=False)

#remove 0s
df_enq_total = df_enq_total.replace(to_replace = 0, value = np.nan)

#Reorder 
order_list = df_enq_total.columns.tolist()
order_list.remove('LabNo')
order_list.insert(0, 'LabNo')
df_enq_total = df_enq_total[order_list]
order_list.remove('LabNo')
order_list.insert(0, 'No of Isolates')
df_enq_total.columns = order_list

#Remove columns will all null
for i in df_enq_total.columns.values.tolist():
    if df_enq_total[i].isnull().all():
        del df_enq_total[i]

#export to Excel to "output" folder
try:
    os.mkdir("output")
except:
    pass
os.chdir("output")
df_enq_total.to_excel(input('Enter the export file name, ending with ".xlsx":\nYour input here:'))


