import cProfile
from datetime import date
import streamlit as st
from pycel import ExcelCompiler
from openpyxl import load_workbook
import logging
import os
import sys
import time
import threading
from PIL import Image

def pycel_logging_to_console(enable=True):
    if enable:
        logger = logging.getLogger('pycel')
        logger.setLevel('INFO')

        console = logging.StreamHandler(sys.stdout)
        console.setLevel(logging.INFO)
        logger.addHandler(console)

def handle_cover(excel, wb, unit):
    #print Addverb in big fat bold red
    st.markdown("<h1 style='text-align: center; color: red;'>Addverb Sales FRM Tool</h1>", unsafe_allow_html=True)

def convert_to_meters(unit, value):
    if unit == 'Feet (ft)':
        return value * 0.3048
    return value

def convert_to_feet(unit, values):
    if unit == 'Feet (ft)':
        return [x* 3.28084 for x in values]
    return values

def handle_dynamo_all(excel, wb, unit):

    dynamo_types = ["Dynamo100", "Dynamo200", "Dynamo500", "Dynamo1000", "Dynamo1500"]

    dynamo_images = {}

    def load_images():
        dynamo_images.update({
            "Dynamo100": Image.open("images/dynamo/Dynamo100.png"),
            "Dynamo200": Image.open("images/dynamo/Dynamo200.png"),
            "Dynamo500": Image.open("images/dynamo/Dynamo500.png"),
            "Dynamo1000": Image.open("images/dynamo/Dynamo1000.png"),
            "Dynamo1500": Image.open("images/dynamo/Dynamo2500.png"),
            "DynamoDistancePatch":  Image.open("images/dynamo/DynamoDistancePatch.png"),
            "DynamoCombinedCycle": Image.open("images/dynamo/DynamoCombinedCycle.png"),
            "DynamoSingleCycle": Image.open("images/dynamo/DynamoSingleCycle.png"),
        })
    load_images()
    
    dynamo_aisle_widths = {
    "Dynamo100": [1.8, 1.6, 1.5, 1.2],
    "Dynamo200": [2.0, 1.8, 1.7, 1.4],
    "Dynamo500": [2.2, 2.0, 1.9, 1.6],
    "Dynamo1000": [2.2, 2.0, 1.9, 1.6],
    "Dynamo1500": [2.5, 2.3, 2.2],
    }
    dynamo_carry_type = {
    "Dynamo100": ["_Conveyor transfer", "_PinHookup/Hookdown"],
    "Dynamo200": ["_Conveyor transfer", "_PinHookup/Hookdown"],
    "Dynamo500": ["_LifterUp/Down", "_Conveyor transfer", "_PinHookup/Hookdown"],
    "Dynamo1000": ["_LifterUp/Down", "_Conveyor transfer"],
    "Dynamo1500": ["_LifterUp/Down"],
    }
    load_unit_types = ["Pallets", "Totes"]
    b11_selected = st.selectbox("Dynamo Model", dynamo_types)
    aisle_widths = dynamo_aisle_widths[b11_selected]
    carry_type = dynamo_carry_type[b11_selected]

    if unit == 'Feet (ft)':
        aisle_widths = convert_to_feet(unit, aisle_widths)
        
    if b11_selected:
        st.image(dynamo_images[b11_selected], use_column_width=True)

    with st.form(key='dynamo_form'):

        selected_aisle_width = st.select_slider('Select Aisle Width', options=aisle_widths, value=aisle_widths[0])
        load_carry_type = st.selectbox("Select Carry Type", carry_type)
        load_unit_type = st.selectbox('Slide to select', load_unit_types)
        
        throughput_pallets_per_hr = st.number_input('Throughput (units/hour)', min_value=0, max_value=100000, value=54)
        number_of_load_units_carried_per_combined_cycle = st.slider('Number of Load Units Carried per Combined Cycle', min_value=1, max_value=6, value=2)
        number_of_load_units_carried_per_single_cycle = st.slider('Number of Load Units Carried per Single Cycle', min_value=1, max_value=6, value=1)
        traffic_factor = st.slider('Traffic Factor in %', min_value=0, max_value=100, value=10)/100
        charging_factor = st.slider('Charging Factor in %', min_value=0, max_value=100, value=15)/100
        
        st.markdown("### Cycle times")
        col1, col2 = st.columns(2)
        col1.image(dynamo_images["DynamoCombinedCycle"], use_column_width=True)
        col2.image(dynamo_images["DynamoSingleCycle"], use_column_width=True)
        near_combined_weightage = st.slider('Near Combined Weightage', min_value=0, max_value=100, value=33)/100
        mid_combined_weightage = st.slider('Mid Combined Weightage', min_value=0, max_value=100, value=33)/100
        far_combined_weightage = st.slider('Far Combined Weightage', min_value=0, max_value=100, value=33)/100
        
        
        total_weightage = near_combined_weightage + mid_combined_weightage + far_combined_weightage
        if total_weightage > 100:
            st.warning('The total weightage cannot exceed 100. Please reset the weightage values.')

        st.markdown("### Distance Patch")
        st.image(dynamo_images["DynamoDistancePatch"], use_column_width=True)
        col1, col2 = st.columns(2)
        max_distance_a = col1.number_input('Max Distance A', min_value=0, max_value=100, value=4)
        max_turns_a = col2.number_input('No of turns A', min_value=0, max_value=100, value=2)

        max_distance_b = col1.number_input('Max Distance B', min_value=0, max_value=100, value=10)
        max_turns_b = col2.number_input('No of turns B', min_value=0, max_value=100, value=2)

        max_distance_c = col1.number_input('Max Distance C', min_value=0, max_value=100, value=10)
        max_turns_c = col2.number_input('No of turns C', min_value=0, max_value=100, value=2)

        max_distance_d = col1.number_input('Max Distance D', min_value=0, max_value=100, value=10)
        max_turns_d = col2.number_input('No of turns D', min_value=0, max_value=100, value=2)

        max_distance_single_cycle = col1.number_input('Max Distance in Single Cycle', min_value=0, max_value=100, value=10)
        max_turns_single_cycle = col2.number_input('No of turns in Single Cycle', min_value=0, max_value=100, value=2)
        
        submit_button = st.form_submit_button(label='Fetch Details',use_container_width=True)


    if submit_button:
        cell_values = {
            'Dynamo all!B11': str(b11_selected),
            'Dynamo all!B12': convert_to_meters(unit,selected_aisle_width),
            'Dynamo all!B13': str(load_carry_type),
            'Dynamo all!B14': str(load_unit_type),
            'Dynamo all!B16': throughput_pallets_per_hr,
            'Dynamo all!B17': number_of_load_units_carried_per_combined_cycle,
            'Dynamo all!B18': number_of_load_units_carried_per_single_cycle,
            'Dynamo all!B19': traffic_factor,
            'Dynamo all!B20': charging_factor,
            'Dynamo all!C25': near_combined_weightage,
            'Dynamo all!C26': mid_combined_weightage,
            'Dynamo all!C27': far_combined_weightage,
            'Dynamo all!B32': convert_to_meters(unit, max_distance_a),
            'Dynamo all!C32': max_turns_a,
            'Dynamo all!B33': convert_to_meters(unit, max_distance_b),
            'Dynamo all!C33': max_turns_b,
            'Dynamo all!B34': convert_to_meters(unit, max_distance_c),
            'Dynamo all!C34': max_turns_c,
            'Dynamo all!B35': convert_to_meters(unit, max_distance_d),
            'Dynamo all!C35': max_turns_d,
            'Dynamo all!B36': convert_to_meters(unit, max_distance_single_cycle),
            'Dynamo all!C36': max_turns_single_cycle,
        }

        for cell, value in cell_values.items():
            excel.evaluate(cell)
            excel.set_value(cell, value)      


        def results(excel, wb):
            st.text("Received your preferences, hang on a sec...!")
            col1, col2, col3 = st.columns(3)
            col1.metric(":red[Cycles/Hr]", round(excel.evaluate('Dynamo all!B7')))
            time.sleep(0.1)
            col2.metric(":red[Pallets/Hr]", round(excel.evaluate('Dynamo all!B8')))
            time.sleep(0.1)
            col3.metric(":red[Dynamos Required]", excel.evaluate('Dynamo all!B9'))
            
            sheet=wb['Dynamo all']
            data = {
                'B11': b11_selected,
                'B12': selected_aisle_width,
                'B13': load_carry_type,
                'B14': load_unit_type,
                'B16': throughput_pallets_per_hr,
                'B17': number_of_load_units_carried_per_combined_cycle,
                'B18': number_of_load_units_carried_per_single_cycle,
                'B19': traffic_factor,
                'B20': charging_factor,
                'C25': near_combined_weightage,
                'C26': mid_combined_weightage,
                'C27': far_combined_weightage,
                'B32': max_distance_a,
                'C32': max_turns_a,
                'B33': max_distance_b,
                'C33': max_turns_b,
                'B34': max_distance_c,
                'C34': max_turns_c,
                'B35': max_distance_d,
                'C35': max_turns_d,
                'B36': max_distance_single_cycle,
                'C36': max_turns_single_cycle
            }
            for key, value in data.items():
                sheet[key] = value
            wb.save(f'files/{date.today()}_{b11_selected}.xlsx')
            
            try: 
                with open('temp.xlsx', 'rb') as file:
                    file_data = file.read()

                st.download_button(
                    label="Download final sheet",
                    data=file_data,
                    file_name='data.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True
                )
            except Exception as e:
                st.success("Record Stored Successfully!")
                        
                
        results(excel, wb)


def handle_sheet2(excel, wb, unit):
    # Similar code for Sheet2
    pass

def loader():
    with st.spinner(text='Hold on tight! :rocket:'):
       time.sleep(0.5)

def main(wb, sheet_names, files, excel):

    selected_sheet = st.sidebar.selectbox('Select a sheet', sheet_names)
    # selected_file = st.sidebar.selectbox('History', files, key='file_select')
    # if selected_file:
    #     with open(os.path.join('./files', selected_file), 'rb') as file:
    #         file_data = file.read()
    #     st.sidebar.download_button(
    #         label="Download selected file",
    #         data=file_data,
    #         file_name=selected_file
    #     )
    unit = st.sidebar.radio("Select Unit", ('Feet (ft)', 'Meters (mtrs)'))
    print(unit)
    if selected_sheet == 'Dynamo all':
        handle_dynamo_all(excel, wb, unit)
    elif selected_sheet == 'COVER':
        handle_cover(excel, wb, unit)
    elif selected_sheet == 'Sheet2':
        handle_sheet2(excel, wb, unit)

if __name__ == '__main__':
    st.set_page_config(page_title="Sales FRM", page_icon="images/AddverbLogo.png")
    excel = ExcelCompiler(filename='master.xlsx')
    wb = load_workbook('master.xlsx')
    sheet_names = wb.sheetnames
    files = os.listdir('./files')
    main(wb, sheet_names, files, excel)