import streamlit as st
from pycel import ExcelCompiler
from openpyxl import load_workbook
import logging
import os
import sys

def pycel_logging_to_console(enable=True):
    if enable:
        logger = logging.getLogger('pycel')
        logger.setLevel('INFO')

        console = logging.StreamHandler(sys.stdout)
        console.setLevel(logging.INFO)
        logger.addHandler(console)

def main():
    pycel_logging_to_console()

    # Load the workbook and select the sheet
    excel = ExcelCompiler(filename='master.xlsx')
    
    b11_value = excel.evaluate('Dynamo all!B11')
    st.write("B11 value is: ", b11_value)

    # Get input from the user
    b11_input = st.text_input("Enter a new value for B11")

    if b11_input:
        st.write("Setting value for B11")
        excel.set_value('Dynamo all!B11', b11_input)
    
        return_values = []
        b7_value = round(excel.evaluate('Dynamo all!B7'))
        b8_value = round(excel.evaluate('Dynamo all!B8'))
        b9_value = round(excel.evaluate('Dynamo all!B9'))
        return_values.append(b7_value)
        return_values.append(b8_value)
        return_values.append(b9_value)
        
        # Load the workbook with openpyxl
        wb = load_workbook('master.xlsx')
        sheet = wb['Dynamo all']

        # Update the values
        sheet['B11'] = b11_input

        wb.save('test.xlsx')
        st.write(return_values)


if __name__ == '__main__':
    main()