import streamlit as st
from pycel import ExcelCompiler
from openpyxl import load_workbook
import logging
import os
import sys
import time

def pycel_logging_to_console(enable=True):
    if enable:
        logger = logging.getLogger('pycel')
        logger.setLevel('INFO')

        console = logging.StreamHandler(sys.stdout)
        console.setLevel(logging.INFO)
        logger.addHandler(console)

def handle_dynamo_all(excel, wb):
    
    loader()
    dynamo_types = ["Dynamo100", "Dynamo200", "Dynamo500", "Dynamo1000", "Dynamo1500"]
    images = {
    "Dynamo100": "images/dynamo/Dynamo100.png",
    "Dynamo200": "images/dynamo/Dynamo200.png",
    "Dynamo500": "images/dynamo/Dynamo500.png",
    "Dynamo1000": "images/dynamo/Dynamo1000.png",
    "Dynamo1500": "images/dynamo/Dynamo2500.png",
    }
    b11_selected = st.selectbox("Dynamo Model", dynamo_types)

    if b11_selected:
        st.image(images[b11_selected], use_column_width=True)

    with st.form(key='dynamo_form'):
        aisle_widths = [2.0, 1.8, 1.7, 1.4]
        selected_aisle_width = st.select_slider('Select Aisle Width', options=aisle_widths, value=1.8)

        carry_type = ["_LifterUp/Down", "_Conveyor transfer", "_PinHookup/Hookdown"]
        load_carry_type = st.selectbox("Select Carry Type", carry_type)

        load_unit_types = ["Pallets", "Totes"]
        load_unit_type = st.selectbox('Slide to select', load_unit_types)
        
        throughput_pallets_per_hr = st.number_input('Throughput (units/hour)', min_value=0, max_value=100000, value=54)
        number_of_load_units_carried_per_combined_cycle = st.slider('Number of Load Units Carried per Combined Cycle', min_value=1, max_value=6, value=2)
        number_of_load_units_carried_per_single_cycle = st.slider('Number of Load Units Carried per Single Cycle', min_value=1, max_value=6, value=1)
        traffic_factor = st.slider('Traffic Factor in %', min_value=0, max_value=100, value=10)
        charging_factor = st.slider('Charging Factor in %', min_value=0, max_value=100, value=15)
        
        st.markdown("### Cycle times")
        col1, col2 = st.columns(2)
        col1.image('images/dynamo/DynamoCombinedCycle.png', use_column_width=True)
        col2.image('images/dynamo/DynamoSingleCycle.png', use_column_width=True)
        near_combined_percent_length_travelled = col1.slider('Near Combined % Length Travelled', min_value=0, max_value=100, value=50)
        near_combined_weightage = col2.slider('Near Combined Weightage', min_value=0, max_value=100, value=33)

        mid_combined_percent_length_travelled = col1.slider('Mid Combined % Length Travelled', min_value=0, max_value=100, value=80)
        mid_combined_weightage = col2.slider('Mid Combined Weightage', min_value=0, max_value=100, value=33)

        far_combined_percent_length_travelled = col1.slider('Far Combined % Length Travelled', min_value=0, max_value=100, value=100)
        far_combined_weightage = col2.slider('Far Combined Weightage', min_value=0, max_value=100, value=33)
        
        
        total_weightage = near_combined_weightage + mid_combined_weightage + far_combined_weightage
        if total_weightage > 100:
            st.warning('The total weightage cannot exceed 100. Please reset the weightage values.')

        st.markdown("### Distance Patch")
        st.image('images/dynamo/DynamoDistancePatch.png', use_column_width=True)
        col1, col2 = st.columns(2)
        max_distance_a = col1.number_input('Max Distance A', min_value=0, max_value=1000, value=4)
        max_turns_a = col2.number_input('No of turns A', min_value=0, max_value=1000, value=2)

        max_distance_b = col1.number_input('Max Distance B', min_value=0, max_value=1000, value=10)
        max_turns_b = col2.number_input('No of turns B', min_value=0, max_value=1000, value=2)

        max_distance_c = col1.number_input('Max Distance C', min_value=0, max_value=1000, value=10)
        max_turns_c = col2.number_input('No of turns C', min_value=0, max_value=1000, value=2)

        max_distance_d = col1.number_input('Max Distance D', min_value=0, max_value=1000, value=10)
        max_turns_d = col2.number_input('No of turns D', min_value=0, max_value=1000, value=2)

        max_distance_single_cycle = col1.number_input('Max Distance in Single Cycle', min_value=0, max_value=1000, value=10)
        max_turns_single_cycle = col2.number_input('No of turns in Single Cycle', min_value=0, max_value=1000, value=2)
        
        submit_button = st.form_submit_button(label='Fetch Details')


    if submit_button:
        excel.evaluate('Dynamo all!B11')
        excel.set_value('Dynamo all!B11', str(b11_selected))    
        excel.evaluate('Dynamo all!B12')
        excel.set_value('Dynamo all!B12', str(selected_aisle_width))
        excel.evaluate('Dynamo all!B13')
        excel.set_value('Dynamo all!B13', str(load_carry_type))
        excel.evaluate('Dynamo all!B14')
        excel.set_value('Dynamo all!B14', str(load_unit_type))
        excel.evaluate('Dynamo all!B16')
        excel.set_value('Dynamo all!B16', str(throughput_pallets_per_hr))
        excel.evaluate('Dynamo all!B17')
        excel.set_value('Dynamo all!B17', str(number_of_load_units_carried_per_combined_cycle))
        excel.evaluate('Dynamo all!B18')
        excel.set_value('Dynamo all!B18', str(number_of_load_units_carried_per_single_cycle))
        excel.evaluate('Dynamo all!B19')
        excel.set_value('Dynamo all!B19', str(traffic_factor))
        excel.evaluate('Dynamo all!B20')
        excel.set_value('Dynamo all!B20', str(charging_factor))
        excel.evaluate('Dynamo all!B25')
        excel.set_value('Dynamo all!B25', near_combined_percent_length_travelled)
        excel.evaluate('Dynamo all!C25')
        excel.set_value('Dynamo all!C25', near_combined_weightage)    
        excel.evaluate('Dynamo all!B26')
        excel.set_value('Dynamo all!B26', mid_combined_percent_length_travelled)
        excel.evaluate('Dynamo all!C26')
        excel.set_value('Dynamo all!C26', mid_combined_weightage) 
        excel.evaluate('Dynamo all!B27')
        excel.set_value('Dynamo all!B27', far_combined_percent_length_travelled)
        excel.evaluate('Dynamo all!C27')
        excel.set_value('Dynamo all!C27', far_combined_weightage) 
        excel.evaluate('Dynamo all!B32')
        excel.set_value('Dynamo all!B32', max_distance_a)
        excel.evaluate('Dynamo all!C32')
        excel.set_value('Dynamo all!C32', max_turns_a)
        excel.evaluate('Dynamo all!B33')
        excel.set_value('Dynamo all!B33', max_distance_b)
        excel.evaluate('Dynamo all!C33')
        excel.set_value('Dynamo all!C33', max_turns_b)     
        excel.evaluate('Dynamo all!B34')
        excel.set_value('Dynamo all!B34', max_distance_c)
        excel.evaluate('Dynamo all!C34')
        excel.set_value('Dynamo all!C34', max_turns_c) 
        excel.evaluate('Dynamo all!B35')
        excel.set_value('Dynamo all!B35', max_distance_d)
        excel.evaluate('Dynamo all!C35')
        excel.set_value('Dynamo all!C35', max_turns_d)
        excel.evaluate('Dynamo all!B36')
        excel.set_value('Dynamo all!B36', max_distance_single_cycle)
        excel.evaluate('Dynamo all!C36')
        excel.set_value('Dynamo all!C36', max_turns_single_cycle)       


        st.text("Values have been updated successfully!")
        
        st.text(f"Cycles/Hr {round(excel.evaluate('Dynamo all!B7')/100)}")
        st.text(f"Pallets/Hr {round(excel.evaluate('Dynamo all!B8')/100)}")
        st.text(f"Dynamos Required {excel.evaluate('Dynamo all!B9')}")


def handle_sheet2(excel, wb):
    # Similar code for Sheet2
    pass

def loader():
    with st.spinner(text='Hold on tight! :rocket:'):
       time.sleep(0.5)


def main():
    pycel_logging_to_console()

    # Load the workbook and select the sheet
    wb = load_workbook('master.xlsx')
    sheet_names = wb.sheetnames

    selected_sheet = st.sidebar.selectbox('Select a sheet', sheet_names)

    excel = ExcelCompiler(filename='master.xlsx')
    loader()
    
    if selected_sheet == 'Dynamo all':
        handle_dynamo_all(excel, wb)
    elif selected_sheet == 'Sheet2':
        handle_sheet2(excel, wb)
    # Add more elif statements for other sheets

if __name__ == '__main__':
    main()