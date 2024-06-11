from datetime import date
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

def handle_cover(excel, wb):
    #print Addverb in big fat bold red
    st.markdown("<h1 style='text-align: center; color: red;'>Addverb Sales FRM Tool</h1>", unsafe_allow_html=True)

def handle_dynamo_all(excel, wb):
    
    loader()
    dynamo_types = ["Dynamo100", "Dynamo200", "Dynamo500", "Dynamo1000", "Dynamo1500"]
    dynamo_images = {
    "Dynamo100": "images/dynamo/Dynamo100.png",
    "Dynamo200": "images/dynamo/Dynamo200.png",
    "Dynamo500": "images/dynamo/Dynamo500.png",
    "Dynamo1000": "images/dynamo/Dynamo1000.png",
    "Dynamo1500": "images/dynamo/Dynamo2500.png",
    }
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
        col1.image('images/dynamo/DynamoCombinedCycle.png', use_column_width=True)
        col2.image('images/dynamo/DynamoSingleCycle.png', use_column_width=True)
        near_combined_percent_length_travelled = col1.slider('Near Combined % Length Travelled', min_value=0, max_value=100, value=50)/100
        near_combined_weightage = col2.slider('Near Combined Weightage', min_value=0, max_value=100, value=33)/100

        mid_combined_percent_length_travelled = col1.slider('Mid Combined % Length Travelled', min_value=0, max_value=100, value=80)/100
        mid_combined_weightage = col2.slider('Mid Combined Weightage', min_value=0, max_value=100, value=33)/100

        far_combined_percent_length_travelled = col1.slider('Far Combined % Length Travelled', min_value=0, max_value=100, value=100)/100
        far_combined_weightage = col2.slider('Far Combined Weightage', min_value=0, max_value=100, value=33)/100
        
        
        total_weightage = near_combined_weightage + mid_combined_weightage + far_combined_weightage
        if total_weightage > 100:
            st.warning('The total weightage cannot exceed 100. Please reset the weightage values.')

        st.markdown("### Distance Patch")
        st.image('images/dynamo/DynamoDistancePatch.png', use_column_width=True)
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
        excel.evaluate('Dynamo all!B11')
        excel.set_value('Dynamo all!B11', str(b11_selected))    
        excel.evaluate('Dynamo all!B12')
        excel.set_value('Dynamo all!B12', selected_aisle_width)
        excel.evaluate('Dynamo all!B13')
        excel.set_value('Dynamo all!B13', str(load_carry_type))
        excel.evaluate('Dynamo all!B14')
        excel.set_value('Dynamo all!B14', str(load_unit_type))
        excel.evaluate('Dynamo all!B16')
        excel.set_value('Dynamo all!B16', throughput_pallets_per_hr)
        excel.evaluate('Dynamo all!B17')
        excel.set_value('Dynamo all!B17', number_of_load_units_carried_per_combined_cycle)
        excel.evaluate('Dynamo all!B18')
        excel.set_value('Dynamo all!B18', number_of_load_units_carried_per_single_cycle)
        excel.evaluate('Dynamo all!B19')
        excel.set_value('Dynamo all!B19', traffic_factor)
        excel.evaluate('Dynamo all!B20')
        excel.set_value('Dynamo all!B20', charging_factor)
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

        with st.spinner(text='Fetching Details...'):
            time.sleep(0.2)
            @st.experimental_dialog(f"Results 🚀")
            def results(excel, wb):
                    st.text("Received your preferences, hang on a sec...!")
                    col1, col2, col3 = st.columns(3)
                    col1.metric(":red[Cycles/Hr]", round(excel.evaluate('Dynamo all!B7')))
                    time.sleep(0.2)
                    col2.metric(":red[Pallets/Hr]", round(excel.evaluate('Dynamo all!B8')))
                    time.sleep(0.2)
                    col3.metric(":red[Dynamos Required]", excel.evaluate('Dynamo all!B9'))
                    
                    sheet=wb['Dynamo all']
                    sheet['B11'] = b11_selected
                    sheet['B12'] = selected_aisle_width
                    sheet['B13'] = load_carry_type
                    sheet['B14'] = load_unit_type
                    sheet['B16'] = throughput_pallets_per_hr
                    sheet['B17'] = number_of_load_units_carried_per_combined_cycle
                    sheet['B18'] = number_of_load_units_carried_per_single_cycle
                    sheet['B19'] = traffic_factor
                    sheet['B20'] = charging_factor
                    sheet['B25'] = near_combined_percent_length_travelled
                    sheet['C25'] = near_combined_weightage
                    sheet['B26'] = mid_combined_percent_length_travelled
                    sheet['C26'] = mid_combined_weightage
                    sheet['B27'] = far_combined_percent_length_travelled
                    sheet['C27'] = far_combined_weightage
                    sheet['B32'] = max_distance_a
                    sheet['C32'] = max_turns_a
                    sheet['B33'] = max_distance_b
                    sheet['C33'] = max_turns_b
                    sheet['B34'] = max_distance_c
                    sheet['C34'] = max_turns_c
                    sheet['B35'] = max_distance_d
                    sheet['C35'] = max_turns_d
                    sheet['B36'] = max_distance_single_cycle
                    sheet['C36'] = max_turns_single_cycle
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


def handle_sheet2(excel, wb):
    # Similar code for Sheet2
    pass

def loader():
    with st.spinner(text='Hold on tight! :rocket:'):
       time.sleep(0.5)

def main():
    # pycel_logging_to_console()

    wb = load_workbook('master.xlsx')
    sheet_names = wb.sheetnames

    selected_sheet = st.sidebar.selectbox('Select a sheet', sheet_names)
    # files = os.listdir('./files')
    # selected_file = st.sidebar.selectbox('History', files, key='file_select')
    # if selected_file:
    #     with open(os.path.join('./files', selected_file), 'rb') as file:
    #         file_data = file.read()
    #     st.sidebar.download_button(
    #         label="Download selected file",
    #         data=file_data,
    #         file_name=selected_file
    #     )

    excel = ExcelCompiler(filename='master.xlsx')
    loader()
    if selected_sheet == 'Dynamo all':
        handle_dynamo_all(excel, wb)
    elif selected_sheet == 'COVER':
        handle_cover(excel, wb)
    elif selected_sheet == 'Sheet2':
        handle_sheet2(excel, wb)

if __name__ == '__main__':
    st.set_page_config(page_title="Sales FRM", page_icon="images/AddverbLogo.png")
    main()