# -*- coding: utf-8 -*-
"""
Created on Thu Oct 21 15:36:04 2021

@author: CheahY
"""

import streamlit as st
import pandas as pd
from ipfn import ipfn
import os

st.title("IPFN Application")
file_container = st.container()
sheets_container = st.container()
table_container = st.container()
result_container = st.container()

def get_sheets_and_rows(file):
    excel_file = pd.ExcelFile(uploaded_file)
    excel_sheets = excel_file.sheet_names
    with sheets_container:
        col1, col2 = st.columns(2)

        with col1:
            sheet_A = st.selectbox("Choose the Target A sheet", excel_sheets)
            sheet_B = st.selectbox("Choose the Target B sheet", excel_sheets)
            sheet_S = st.selectbox("Choose the SEED sheet", excel_sheets)
        
        with col2:
            row_A = st.number_input("Target A table start row",min_value=1,step=1,value=1,key="A")
            row_B = st.number_input("Target B table start row",min_value=1,step=1,value=1,key="B")
            row_S = st.number_input("Seed table start row",min_value=1,step=1,value=1,key="S")
    
    return sheet_A,sheet_B,sheet_S,int(row_A),int(row_B),int(row_S)

def get_number_of_field(df):
    
    i=0
    for column in df.columns:
        try:
            if column == "Year":
                i+=1
            else:
                int(column)
                pd.to_numeric(df[column])
                if i != 0:
                    break

        except ValueError:
            i+=1
            
    return i

def rename_industry_column_name(x):
    
    if type(x) == str:
        x = x.strip()
    
    if x == "Industry":
        x = "Industry sector"
    
    return x

def validate_field_name(target_field,seed_field):
    
    with table_container:
        for field in target_field:
            if field not in seed_field:
                st.write(f"{field} not found in SEED table. Please check and try again.")
                return False
        
    return True
    
def drop_unmatch_rows(df_seed,df_A,df_B):
    
    df_seed = pd.DataFrame(df_seed.stack())
    df_seed = df_seed.reset_index()
    
    for column in df_seed.columns[:-1]:
        if column in df_A.index.names:
            df_seed = df_seed[df_seed[column].isin(df_A.index.get_level_values(column).tolist())]
        elif column in df_B.index.names:
            df_seed = df_seed[df_seed[column].isin(df_B.index.get_level_values(column).tolist())]
        elif column == "Year":
            df_seed = df_seed[df_seed[column].isin(df_A.columns.tolist())]
    
    for index in df_A.index.names:
        df_A = df_A.query(f'`{index}` in {df_seed[index].tolist()}')
    
    for index in df_B.index.names:
        df_B = df_B.query(f'`{index}` in {df_seed[index].tolist()}')
    
    df_A = df_A.stack()
    df_B = df_B.stack()
        
    return df_seed,df_A,df_B
    
def read_table(file,sheet_A,sheet_B,sheet_S,row_A,row_B,row_S):
    
    # read tables by sheetname, header set to None as the start row is different
    df_seed = pd.read_excel(file,sheet_name=sheet_S,skiprows=row_S-1).dropna(axis=1,how="all").dropna(axis=0,how="all")
    df_A = pd.read_excel(file,sheet_name=sheet_A,skiprows=row_A-1).dropna(axis=1,how="all").dropna(axis=0,how="all")
    df_B = pd.read_excel(file,sheet_name=sheet_B,skiprows=row_B-1).dropna(axis=1,how="all").dropna(axis=0,how="all")
    
    # get number of field
    field_A = get_number_of_field(df_A)
    field_B = get_number_of_field(df_B)
    field_S = get_number_of_field(df_seed)
    
    # replace "Industry" to "Industry sector" in columns
    df_seed.columns = list(map(lambda name: rename_industry_column_name(name),df_seed.columns))
    df_A.columns = list(map(lambda name: rename_industry_column_name(name),df_A.columns))
    df_B.columns = list(map(lambda name: rename_industry_column_name(name),df_B.columns))
    
    # collect all the field name
    field_name_A = df_A.columns[:field_A].tolist()
    field_name_B = df_B.columns[:field_B].tolist()
    field_name_S = df_seed.columns[:field_S].tolist()
    field_name = list(set(field_name_S + field_name_A + field_name_B + ["Year"]))
    
    # set index for each tables
    df_seed = df_seed.set_index(field_name_S)
    df_A = df_A.set_index(field_name_A)
    df_B = df_B.set_index(field_name_B)
    
    # by default the columns for target table is Year, transform the year data type to integer
    df_A.columns.name = "Year"
    df_B.columns.name = "Year"
    df_A.columns = pd.to_numeric(df_A.columns)
    df_B.columns = pd.to_numeric(df_B.columns)
    
    for field in field_name:
        if field not in field_name_S:
            df_seed.columns.name = field
            field_name_S.append(field)
            break

    # check if seed columns is year
    if df_seed.columns.name == "Year":
        df_seed.columns = pd.to_numeric(df_seed.columns)

    # validate field name
    with table_container:
        if validate_field_name(field_name_A,field_name_S) and \
            validate_field_name(field_name_B,field_name_S) and \
            df_A.columns.tolist() == df_B.columns.tolist():

            st.write(f"Target A: {' x '.join(field_name_A)} x Year")
            st.write(f"Target B: {' x '.join(field_name_B)} x Year")
            st.write(f"Seed: {' x '.join(field_name_S)}")
            
            # drop unmatch rows
            df_seed,df_A,df_B = drop_unmatch_rows(df_seed,df_A,df_B)
            
            with st.expander("Click here to see table"):
                col1, col2 = st.columns(2)
        
                with col1:
                    st.write("Target A",df_A)
                with col2:
                    st.write("Target B",df_B)
                
                st.write("SEED Table",df_seed)
            
            st.write("New tables loaded.")

            return df_seed,df_A,df_B
        
        elif df_A.columns.tolist() != df_B.columns.tolist():
            st.write("Year columns in Target A does not match with Years column in Target B")
            return None,None,None
    
        else:
            return None,None,None

def generate_results(df_seed,df_A,df_B):
    
    aggregates = [df_A,df_B]
    dimensions = [list(df_A.index.names),list(df_B.index.names)]
    
    with result_container:
        col1, col2, col3 = st.columns(3)
        
        with col1:
            convergence_rate = st.number_input("Convergence rate",value=1e-5,step=1e-5,format="%.f",key="conv")
        with col2:
            rate_tolerance = st.number_input("Tolerance rate",value=1e-8,step=1e-8,format="%.f",key="tol")
        with col3:
            max_iteration = st.number_input("Maximum iteration",step=1,value=500,key="iter")
        
        if st.button("Generate Results"):

            IPF = ipfn.ipfn(df_seed, aggregates, dimensions,weight_col=0,verbose=2,
                            convergence_rate=convergence_rate,rate_tolerance=rate_tolerance,max_iteration=max_iteration)
            df,flag,df_iteration = IPF.iteration()
            st.write(df_seed)

            iteration = max(df_iteration.index)
            conv_rate = df_iteration.iat[iteration,0]
            
            st.session_state.iteration = iteration
            st.session_state.conv_rate = conv_rate
            
            st.write(f"Number of Iteration: {iteration+1}")
            st.write(f"Convergence rate: {conv_rate}")
            st.write("Saving results...")
            
            writer = pd.ExcelWriter(uploaded_file,engine='openpyxl',mode='a',
                            if_sheet_exists='new')
            df.to_excel(writer,sheet_name = 'Results',index=False,engine='openpyxl')
            writer.close()

            with result_container:
                if st.download_button("Download Results",uploaded_file,file_name=uploaded_file.name):
                    pass

with file_container:
    st.header("Choose an input file")

    uploaded_file = st.file_uploader("Choose an excel file",type="xlsx")
    
if uploaded_file is not None:
    
    sheet_A,sheet_B,sheet_S,row_A,row_B,row_S = get_sheets_and_rows(uploaded_file)
    with table_container:
  
        df_seed,df_A,df_B = read_table(uploaded_file,sheet_A,sheet_B,sheet_S,row_A,row_B,row_S)

    if df_seed is not None and \
        df_A is not None and \
        df_B is not None:
        generate_results(df_seed,df_A,df_B)
    
    st.write(uploaded_file.name)
