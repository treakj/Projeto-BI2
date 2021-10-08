# # %%
# !pip install pyodbc
# !pip install pandas
# !pip install numpy
# !pip install seaborn
# !pip install matplotlib
# !pip install pyarrow
# !pip install openpyxl
# !pip install xlrd
# !pip install --upgrade pip --user
# !pip install feather.format
# !pip install XlsxWriter
# %%
import pyodbc
import pandas as pd
import seaborn as sns
import numpy as np
import matplotlib.pyplot as plt
import openpyxl
from datetime import date
import datetime
import feather
import shlex
import os
import pickle
import datetime
# import rata_util as rata
from dateutil.relativedelta import relativedelta, MO
from IPython.display import HTML




# %%
# def highlight_cells():
#     # provide your criteria for highlighting the cells here
#     return ['background-color: yellow']


# analise_cruz_final = analise_cruz_final.loc[:,
#                                             ~analise_cruz_final.columns.duplicated()]
# analise_cruz_final_excel = analise_cruz_final.reset_index(
#     drop=True).style.apply(highlight_cells, subset=['teste_percentual_receita'])



# %%
import xlsxwriter
faturado_final = pd.ExcelWriter('analise_fatura_samp_teste.xlsx', engine='xlsxwriter')
analise_cruz_final.to_excel(faturado_final, sheet_name='Análise')
max_linha=analise_cruz_final.shape[0]
max_col=analise_cruz_final.shape[1]

workbook=faturado_final.book
worksheet=faturado_final.sheets['Análise']

num = workbook.add_format({'num_format': '#,##0.00'})
perc = workbook.add_format({'num_format': '0.0%'})

num.set_bg_color('#006d77')
num.set_bg_color('#83c5be')

cor_teste_valor = workbook.add_format({'num_format': '#,##0.00'})
cor_teste_perc = workbook.add_format({'num_format': '0.0%'})

cor_teste_perc.set_bg_color('#ffddd2')
cor_teste_valor.set_bg_color('#e29578')
col_len = 22

worksheet.set_column('K:K',col_len, cor_teste_valor)
worksheet.set_column('L:L',col_len, cor_teste_perc)

worksheet.set_column('M:M',col_len, cor_teste_valor)
worksheet.set_column('N:N',col_len, cor_teste_perc)

worksheet.set_column('O:O',col_len, cor_teste_valor)
worksheet.set_column('P:P',col_len, cor_teste_perc)

worksheet.set_column('Q:Q',col_len, cor_teste_valor)
worksheet.set_column('R:R',col_len, cor_teste_perc)

worksheet.set_column('S:AC',col_len, num)

worksheet.set_column('AD:AD',col_len, cor_teste_valor)
worksheet.set_column('AE:AE',col_len, cor_teste_perc)

worksheet.set_column('AF:AI',col_len, num)

worksheet.set_column('AJ:AJ',col_len, cor_teste_valor)
worksheet.set_column('AK:AK',col_len, cor_teste_perc)

worksheet.set_column('AL:AO',col_len, num)

worksheet.set_column('AP:AP',col_len, cor_teste_valor)
worksheet.set_column('AQ:AQ',col_len, cor_teste_perc)

worksheet.set_column('AR:AT',col_len, num)


worksheet.autofilter(0,1,max_linha,max_col )

faturado_final.save()
# faturado_final.close()
# %%
from sqlalchemy import create_engine,text
import sqlalchemy
# engine= create_engine('mssql+pyodbc://BDUORGS,5468/SGT_DEV?driver=ODBC+Driver+17+for+SQL+Server')
engine= create_engine('mssql+pyodbc://BDDEV2,5455/SGT_DEV?driver=ODBC+Driver+17+for+SQL+Server')
# %%
# with engine.connect() as conn:
#   result=conn.execute("select * from pwrbi.CodigosDist")
#   print(result.all())
# %%
if 'pest_db' not in engine.dialect.get_schema_names(engine):
  engine.execute(sqlalchemy.schema.CreateSchema('pest_db'))
fatura_final.drop(columns='DthAvaliacao').to_sql('analisefatura', engine,schema='pest_db', if_exists='replace')
mercado_rel.to_sql('mercado_rel', engine, schema='pest_db', if_exists='replace')
analise_cruz_final.fillna(0).reset_index().to_sql('mercado_samp', engine, schema='pest_db', if_exists='replace')
# %%
