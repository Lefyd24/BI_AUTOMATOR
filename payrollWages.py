from time import time
import openpyxl as xl
import os
import pandas as pd
from warnings import filterwarnings
from datetime import datetime
import time
from colorama import Fore, Style, Back
import streamlit as st
import datetime as dt

filterwarnings("ignore")


@st.cache(show_spinner=False)
def wages(files):
    try:
        dataframes_to_concat = []
        start = time.time()
        counter = 0
        for file in files:
            with st.spinner(f"Processing {counter+1} file out of {len(files)}"):
                print(file)
                print(f'Checked {counter} out of {len(files)} files...')

                df = pd.read_excel(file, sheet_name="ST01", engine='openpyxl', skiprows=5,
                                   converters={"WORKING DAYS": float, 'SOCIAL SECURITY NUMBER': str,
                                               'VAT NUMBER': str, "Adjusted\nGross\nDaily\nWage": float,
                                               "HOTEL CODE": str, "HOTEL NAME": str, "MONTH": str,
                                               'DAYS OFF/SICKNESS': float,
                                               'PAYABLE\nLEAVE/\nSICKNESS/\nOTHERS': float,
                                               'PAYABLE\nSALARY\nPYLON A': float, 'PAYABLE\nAMOUNT (B)': float,
                                               'TOTAL PAYABLE\nAMOUNT  \n(14+15=16+17)': float,
                                               'Ασφαλιστικές\nΕισφορές\nΜηνός': float,
                                               'Φ.Μ.Υ.\nΜηνός': float, 'Adjusted \nGross\nDaily \nWage': str})
                # print(df.columns)
                try:
                    df.drop(["S/N", "HOTEL NAME", 'WORKING\nPOSITION', 'wp code', 'SECTOR',
                              '% Ασφαλιστικών\nΕισφορών\nΜηνός (Α)\n(H/L 7-12)',
                              'Total Adjusted\nDaily\nCost\n(*1,27)',
                              'Total Adjusted\nDaily\nCost\n(*1,27)\n(H/L 1-6)',
                              'Total Adjusted\nDaily\nCost\n(*1,27)\nVF', 'Agreed \nNet \nDaily \nWage',
                              'Unnamed: 31', '% Ασφαλιστικών\nΕισφορών\nΜηνός (Α,Β)\n(H/L 1-6)',
                              '% Φ.Μ.Υ.\nΜηνός (Α)', 'LAST NAME', 'FIRST NAME', 'SOCIAL SECURITY NUMBER'], axis=1, inplace=True)
                except KeyError as key:
                    print(f"{Fore.MAGENTA}You're probably examining a non 3300 Payroll file. Changing parameters...{Style.RESET_ALL}")
                    #print(df.columns)
                    df.drop(['S/N', 'HOTEL NAME','LAST NAME', 'FIRST NAME', 'wp code', 'SECTOR', 'SOCIAL SECURITY NUMBER','% Ασφαλιστικών\nΕισφορών\nΜηνός (Α)\n(H/L 7-12)',
                               '% Ασφαλιστικών\nΕισφορών\nΜηνός (Α,Β)\n(H/L 1-6)','% Φ.Μ.Υ.\nΜηνός (Α)',
                               'Total Adjusted\nDaily\nCost\n(*1,27)\n(H/L 7-12)',
                               'Total Adjusted\nDaily\nCost\n(*1,27)\n(H/L 1-6)',
                               'Total Adjusted\nDaily\nCost\n(*1,27)\nVF','Agreed \nNet \nDaily \nWage', 'Unnamed: 31'], inplace=True, axis=1)
                    #break
                df["Adjusted \nGross\nDaily \nWage"].replace(to_replace="-", value=0, inplace=True)
                df["Year"] = dt.date.today().year
                #print(df.columns)
                df = df[['HOTEL CODE', 'Year', 'MONTH', 'VAT NUMBER', 'FINANCIAL AGREEMENT', 'DAYS OFF AGREED',
                         'Μ/Η/Δ/ΗΑ', 'WORKING DAYS', 'DAYS OFF/SICKNESS', 'PAYABLE\nAGREEMENT',
                         'PAYABLE\nLEAVE/\nSICKNESS/\nOTHERS', 'PAYABLE\nSALARY\nPYLON A', 'PAYABLE\nAMOUNT (B)',
                         'TOTAL PAYABLE\nAMOUNT  \n(14+15=16+17)', 'Ασφαλιστικές\nΕισφορές\nΜηνός', 'Φ.Μ.Υ.\nΜηνός',
                         'Adjusted \nGross\nDaily \nWage']]
                df["FINANCIAL AGREEMENT"] = df["FINANCIAL AGREEMENT"].replace(to_replace="ΚΣΣΕ", value=0)
                df["DAYS OFF AGREED"] = df["DAYS OFF AGREED"].replace(to_replace="ΚΣΣΕ", value=0)
                df["PAYABLE\nAGREEMENT"] = df["PAYABLE\nAGREEMENT"].replace(to_replace="ΚΣΣΕ", value=0)
                df.dropna(subset=['VAT NUMBER'], inplace=True)
                #print(df.columns)
                df.rename(columns={'HOTEL CODE':"HotelCode", 'MONTH':"Month", 'VAT NUMBER':"VAT", 'FINANCIAL AGREEMENT':"FinancialAgr",'DAYS OFF AGREED':"DaysOffAgreed", 'Μ/Η/Δ/ΗΑ':"MHDHA", 'WORKING DAYS':"WorkingDays", 'DAYS OFF/SICKNESS':"DaysOffSickness",
               'PAYABLE\nAGREEMENT':"PayableAgreement", 'PAYABLE\nLEAVE/\nSICKNESS/\nOTHERS':"PayableLSO",
               'PAYABLE\nSALARY\nPYLON A':"PayableA", 'PAYABLE\nAMOUNT (B)':"PayableB",
               'TOTAL PAYABLE\nAMOUNT  \n(14+15=16+17)':"PayableTotal",
               'Ασφαλιστικές\nΕισφορές\nΜηνός':"SocialSecurity", 'Φ.Μ.Υ.\nΜηνός':"FMY",
               'Adjusted \nGross\nDaily \nWage':"AdjustedGrossWage"}, inplace=True)
                dataframes_to_concat.append(df)
                counter += 1
                # print(df)
                #break
    #print(dataframes_to_concat)
        total_ds = pd.concat(dataframes_to_concat)
        print(total_ds)
        total_ds.to_excel("PayrollWages.xlsx",sheet_name="Payroll_Wages", index=False)
        end = time.time()
        return round(end - start, 2)
            #break
    except Exception as e:
        st.warning(f"An error with the following code has occured: {e}.\nPlease contact Lefteris Fthenos to resolve.")


