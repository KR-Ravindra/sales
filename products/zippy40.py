from PIL import Image
import streamlit as st
import time
from products.utils import save_to_file, convert_to_meters, convert_to_feet

@st.experimental_fragment
def handle_zippy40(excel, wb, unit):

    _SHEET_NAME = "Cycle Time ZIppy 40"

    zippy40_images = {}

    def load_images():
        zippy40_images.update({
            "Zippy40": Image.open("images/zippy/Zippy40.png"),
            "ZippyCycleTime": Image.open("images/zippy/ZippyCycleTime.png"),
        })
    load_images()
    product_name = "Zippy40"
    
    st.image(zippy40_images[product_name], use_column_width=True)
 
    def validate(total_weightage):
        if total_weightage > 1:
            st.toast("Total cycle time weightage exceeds 100%", icon='🤖')
            return False
        return True
   
    with st.form(key='zippy40_form'):
        
        col1, col2 = st.columns(2)
        distance_bw_induction_and_first_sort = col1.number_input('Distance b/w induction station and first sort point', min_value=0, max_value=100, value=7)
        throughput_sorts_per_hr = col2.number_input('Throughput (sorts/hour)', min_value=0, max_value=5000, value=500)
        traffic_factor = col1.slider('Traffic Factor in %', min_value=0, max_value=100, value=10)/100
        charging_factor = col2.slider('Charging Factor in %', min_value=0, max_value=100, value=10)/100
        
        st.image(zippy40_images["ZippyCycleTime"], use_column_width=True)
        col1, col2 = st.columns(2)
        max_barcodes_in_l_sorting_area = col1.slider(f'Max Barcodes in L sorting area ({ convert_to_feet(unit, 200) if unit == "Feet (ft)" else 200  })', min_value=1, max_value=1000, value=70)
        max_barcodes_in_w_sorting_area = col2.slider(f'Max Barcodes in W sorting area ({ convert_to_feet(unit, 200) if unit == "Feet (ft)" else 200 })', min_value=1, max_value=1000, value=45)
        total_barcodes_in_layout = st.slider('Total Barcodes in Layout', min_value=0, max_value=2000, value=230)
    

        st.markdown("### Single Cycle times")
        col1,col2,col3,col4 = st.columns(4)
        near_combined_weightage = col1.number_input(f'Near % - { convert_to_feet(unit, 4.9) if unit == "Feet (ft)" else 4.9  } ', min_value=0, max_value=100, value=20)/100
        mid_combined_weightage = col2.number_input(f'Mid % - { convert_to_feet(unit, 9.8) if unit == "Feet (ft)" else 9.8 } ', min_value=0, max_value=100, value=35)/100
        far_combined_weightage = col3.number_input(f'Far % - { convert_to_feet(unit, 24.5) if unit == "Feet (ft)" else 24.5 } ', min_value=0, max_value=100, value=35)/100
        diagonally_farthest_weightage = col4.number_input(f'Diagonal % - { convert_to_feet(unit, 49) if unit == "Feet (ft)" else 49 } ', min_value=0, max_value=100, value=10)/100

        total_weightage = near_combined_weightage + mid_combined_weightage + far_combined_weightage + diagonally_farthest_weightage
        if total_weightage > 1:
            st.warning("Total cycle time weightage exceeds 100%")

        submit_button = st.form_submit_button(label='Fetch Details', use_container_width=True)
    
    if submit_button:
        if not validate(total_weightage):
            st.text('Registered an error!')
            st.stop()
        cell_values = {
            f'{_SHEET_NAME}!B12': max_barcodes_in_l_sorting_area,
            f'{_SHEET_NAME}!B13': max_barcodes_in_w_sorting_area,
            f'{_SHEET_NAME}!B14': convert_to_meters(unit, distance_bw_induction_and_first_sort) if unit == "Feet (ft)" else distance_bw_induction_and_first_sort,
            f'{_SHEET_NAME}!B15': throughput_sorts_per_hr,
            f'{_SHEET_NAME}!B16': traffic_factor,
            f'{_SHEET_NAME}!B17': charging_factor,
            f'{_SHEET_NAME}!B18': total_barcodes_in_layout,
            f'{_SHEET_NAME}!E21': near_combined_weightage,
            f'{_SHEET_NAME}!E22': mid_combined_weightage,
            f'{_SHEET_NAME}!E23': far_combined_weightage,
            f'{_SHEET_NAME}!E24': diagonally_farthest_weightage,
        }

        for cell, value in cell_values.items():
            excel.evaluate(cell)
            excel.set_value(cell, value)      


        def results(excel, wb):
            st.text("Received your preferences, hang on a sec...!")
            cycles_per_hr = round(excel.evaluate(f'{_SHEET_NAME}!B7'))
            zippies_required = round(excel.evaluate(f'{_SHEET_NAME}!B8'))
            zippy_barcode_ratio = total_barcodes_in_layout//zippies_required
            if zippy_barcode_ratio < 8:
                st.error('Zippy barcode ratio is too low! Try increasing the total number of barcodes')
                st.stop()
            time.sleep(0.1)
            col1, col2, col3 = st.columns(3)
            col1.metric(":red[Cycles/Hr]", cycles_per_hr)
            time.sleep(0.1)
            col2.metric(":red[Zippies Required]", zippies_required)
            time.sleep(0.1)
            col3.metric(":red[Zippy - Barcode Ratio]", zippy_barcode_ratio)

            
            sheet=wb[f'{_SHEET_NAME}']
            data = {
                'B12': max_barcodes_in_l_sorting_area,
                'B13': max_barcodes_in_w_sorting_area,
                'B14': convert_to_meters(unit, distance_bw_induction_and_first_sort) if unit == "Feet (ft)" else distance_bw_induction_and_first_sort,
                'B16': throughput_sorts_per_hr,
                'B17': traffic_factor,
                'B18': charging_factor,
                'E21': near_combined_weightage,
                'E22': mid_combined_weightage,
                'E23': far_combined_weightage,
                'E24': diagonally_farthest_weightage,
            }
            for key, value in data.items():
                sheet[key] = value

            saved_file_name = save_to_file(wb, product_name)
            
            # with open(f'files/{saved_file_name}', 'rb') as file:
            #     file_data = file.read()

            # st.download_button(
            #     label="Download final sheet",
            #     data=file_data,
            #     file_name=f'{saved_file_name}',
            #     mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            #     use_container_width=True
            # )
 
        results(excel, wb)

