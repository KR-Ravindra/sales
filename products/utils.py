from datetime import datetime

def save_to_file(wb, product_name):
    # Get the current date and time
    now = datetime.now()

    # Format the date and time as a string
    date_time = now.strftime("%Y-%m-%d_%H-%M-%S")
    wb.save(f'files/{date_time}_{product_name}.xlsx')
    
    return f'{date_time}_{product_name}.xlsx'

def convert_to_meters(unit, value):
    if unit == 'Feet (ft)':
        return value * 0.3048
    return value

def convert_to_feet_list(unit, values):
    if unit == 'Feet (ft)':
        return [round(x * 3.28084, 2) for x in values]
    return values

def convert_to_feet(unit, value):
    if unit == 'Feet (ft)':
        return round(value * 3.28084, 2)
    return value