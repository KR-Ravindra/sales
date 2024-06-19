from PIL import Image
import streamlit as st
import time
from products.utils import save_to_file, convert_to_meters, convert_to_feet, convert_to_feet_list

@st.experimental_fragment
def handle_dynamo_all(excel, wb, unit):
    
    _SHEET_NAME = "Dynamo all"

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
    product_name = st.selectbox("Dynamo Model", dynamo_types)
    aisle_widths = dynamo_aisle_widths[product_name]
    carry_type = dynamo_carry_type[product_name]

    if unit == 'Feet (ft)':
        aisle_widths = convert_to_feet_list(unit, aisle_widths)
        
    if product_name:
        st.image(dynamo_images[product_name], use_column_width=True)
    
    def validate(total_weightage):
        if total_weightage > 1:
            st.toast("Total cycle time weightage exceeds 100%", icon='🤖')
            return False
        return True
   
 
    with st.form(key='dynamo_form'):
        col1, col2 = st.columns(2)
        selected_aisle_width = col1.select_slider('Select Aisle Width', options=aisle_widths, value=aisle_widths[0])
        load_carry_type = col2.selectbox("Select Carry Type", carry_type)
        load_unit_type = col2.selectbox('Slide to select', load_unit_types)
        
        throughput_pallets_per_hr = col2.number_input('Throughput (units/hour)', min_value=0, max_value=100000, value=54)
        number_of_load_units_carried_per_combined_cycle = col1.slider('Number of Load Units Carried per Combined Cycle', min_value=1, max_value=6, value=2)
        number_of_load_units_carried_per_single_cycle = col1.slider('Number of Load Units Carried per Single Cycle', min_value=1, max_value=6, value=1)
        col1, col2 = st.columns(2)
        traffic_factor = col1.slider('Traffic Factor in %', min_value=0, max_value=100, value=10)/100
        charging_factor = col2.slider('Charging Factor in %', min_value=0, max_value=100, value=15)/100
        
        st.markdown("### Distance Patch")
        st.image(dynamo_images["DynamoDistancePatch"], use_column_width=True)
        col1, col2, col3, col4 = st.columns(4)
        max_distance_a = col1.number_input('Max Distance A', min_value=0, max_value=100, value=4)
        max_turns_a = col2.number_input('No of turns A', min_value=0, max_value=100, value=2)

        max_distance_b = col3.number_input('Max Distance B', min_value=0, max_value=100, value=10)
        max_turns_b = col4.number_input('No of turns B', min_value=0, max_value=100, value=2)

        max_distance_c = col1.number_input('Max Distance C', min_value=0, max_value=100, value=10)
        max_turns_c = col2.number_input('No of turns C', min_value=0, max_value=100, value=2)

        max_distance_d = col3.number_input('Max Distance D', min_value=0, max_value=100, value=10)
        max_turns_d = col4.number_input('No of turns D', min_value=0, max_value=100, value=2)
        col1, col2 = st.columns(2)
        max_distance_single_cycle = col1.number_input('Max Distance in Single Cycle', min_value=0, max_value=100, value=10)
        max_turns_single_cycle = col2.number_input('No of turns in Single Cycle', min_value=0, max_value=100, value=2)

        st.markdown("### Cycle times")
        col1, col2 = st.columns(2)
        col1.image(dynamo_images["DynamoCombinedCycle"], use_column_width=True)
        col2.image(dynamo_images["DynamoSingleCycle"], use_column_width=True)
        col1,col2,col3 = st.columns(3)
        near_combined_weightage = col1.number_input('Near Combined Weightage', min_value=0, max_value=100, value=33)/100
        mid_combined_weightage = col2.number_input('Mid Combined Weightage', min_value=0, max_value=100, value=33)/100
        far_combined_weightage = col3.number_input('Far Combined Weightage', min_value=0, max_value=100, value=33)/100

        total_weightage = near_combined_weightage + mid_combined_weightage + far_combined_weightage
        if total_weightage > 1:
            st.warning("Total cycle time weightage exceeds 100%")

        submit_button = st.form_submit_button(label='Fetch Details', use_container_width=True)
    
    if submit_button:
        if not validate(total_weightage):
            st.text('Registered an error!')
            st.stop()
        cell_values = {
            f'{_SHEET_NAME}!B11': str(product_name),
            f'{_SHEET_NAME}!B12': convert_to_meters(unit,selected_aisle_width),
            f'{_SHEET_NAME}!B13': str(load_carry_type),
            f'{_SHEET_NAME}!B14': str(load_unit_type),
            f'{_SHEET_NAME}!B16': throughput_pallets_per_hr,
            f'{_SHEET_NAME}!B17': number_of_load_units_carried_per_combined_cycle,
            f'{_SHEET_NAME}!B18': number_of_load_units_carried_per_single_cycle,
            f'{_SHEET_NAME}!B19': traffic_factor,
            f'{_SHEET_NAME}!B20': charging_factor,
            f'{_SHEET_NAME}!C25': near_combined_weightage,
            f'{_SHEET_NAME}!C26': mid_combined_weightage,
            f'{_SHEET_NAME}!C27': far_combined_weightage,
            f'{_SHEET_NAME}!B32': convert_to_meters(unit, max_distance_a),
            f'{_SHEET_NAME}!C32': max_turns_a,
            f'{_SHEET_NAME}!B33': convert_to_meters(unit, max_distance_b),
            f'{_SHEET_NAME}!C33': max_turns_b,
            f'{_SHEET_NAME}!B34': convert_to_meters(unit, max_distance_c),
            f'{_SHEET_NAME}!C34': max_turns_c,
            f'{_SHEET_NAME}!B35': convert_to_meters(unit, max_distance_d),
            f'{_SHEET_NAME}!C35': max_turns_d,
            f'{_SHEET_NAME}!B36': convert_to_meters(unit, max_distance_single_cycle),
            f'{_SHEET_NAME}!C36': max_turns_single_cycle,
        }

        for cell, value in cell_values.items():
            excel.evaluate(cell)
            excel.set_value(cell, value)      


        def results(excel, wb):
            st.text("Received your preferences, hang on a sec...!")
            col1, col2, col3 = st.columns(3)
            col1.metric(":red[Cycles/Hr]", round(excel.evaluate(f'{_SHEET_NAME}!B7')))
            time.sleep(0.1)
            col2.metric(":red[Pallets/Hr]", round(excel.evaluate(f'{_SHEET_NAME}!B8')))
            time.sleep(0.1)
            col3.metric(":red[Dynamos Required]", excel.evaluate(f'{_SHEET_NAME}!B9'))
            
            sheet=wb[f'{_SHEET_NAME}']
            data = {
                'B11': str(product_name),
                'B12': convert_to_meters(unit,selected_aisle_width),
                'B13': str(load_carry_type),
                'B14': str(load_unit_type),
                'B16': throughput_pallets_per_hr,
                'B17': number_of_load_units_carried_per_combined_cycle,
                'B18': number_of_load_units_carried_per_single_cycle,
                'B19': traffic_factor,
                'B20': charging_factor,
                'C25': near_combined_weightage,
                'C26': mid_combined_weightage,
                'C27': far_combined_weightage,
                'B32': convert_to_meters(unit, max_distance_a),
                'C32': max_turns_a,
                'B33': convert_to_meters(unit, max_distance_b),
                'C33': max_turns_b,
                'B34': convert_to_meters(unit, max_distance_c),
                'C34': max_turns_c,
                'B35': convert_to_meters(unit, max_distance_d),
                'C35': max_turns_d,
                'B36': convert_to_meters(unit, max_distance_single_cycle),
                'C36': max_turns_single_cycle,
            }
            for key, value in data.items():
                sheet[key] = value
            
            saved_file_name = save_to_file(wb, product_name)
            
            with open(f'files/{saved_file_name}', 'rb') as file:
                file_data = file.read()

            st.download_button(
                label="Download final sheet",
                data=file_data,
                file_name=f'{saved_file_name}',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                use_container_width=True
            )
 
        results(excel, wb)

