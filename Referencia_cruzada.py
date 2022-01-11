#pip install openpyxl
#pip install pandas

import pandas as pd #Sample data I created and saved to df
#Lectura de archvo excel, la hoja DATOS
df = pd.read_excel(io = "CERTIFICADO FLETES SEPTIEMBRE 2021.xlsm", sheet_name="DATOS")

#usando pivo_table inicio 
"""Forma de uso:  pandas.pivot_table(data, values=None, index=None, columns=None, aggfunc='mean', 
fill_value=None, margins=False, dropna=True, margins_name='All', observed=False, sort=True)
"""
cur1=pd.pivot_table(data=df, index=['CHAPA','TIPO DE MATERIAL'], values='VIAJES', 
        aggfunc='sum',margins=True)
print(cur1)
#usando pivo_table fin    

#Usando referencia cruzada inicio
"""Forma de uso: pandas.crosstab(index, columns, values=None, rownames=None, colnames=None, aggfunc=None, 
margins=False, margins_name='All', dropna=True, normalize=False)"""

cur2=pd.crosstab(index=df['CHAPA'], columns=df['TIPO DE MATERIAL'],values=df['VIAJES'],
        aggfunc='sum',margins=True)
print(cur2)
#Usando referencia cruzada fin

#Exortando a Excel las consultas
df_rrss=pd.DataFrame(cur1)
df_rrss.to_excel('viajes1.xlsx')

df_rrss=pd.DataFrame(cur2)
df_rrss.to_excel('viajes2.xlsx')
