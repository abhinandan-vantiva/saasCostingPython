import pandas as pd
from datetime import datetime

def process_data():
    # Read the data from Excel files
    node_details_df = pd.read_excel('node_details.xlsx', header=0)
    customer_details_df = pd.read_excel('customer_details.xlsx', header=0)

    # Perform the left join
    result_df = pd.merge(node_details_df, customer_details_df, left_on='FacilityId', right_on='facility_id', how='left')
    result_df = result_df.drop(columns=['facility_id'])

    # Add columns for various device types
    device_types = {
        'OWM7111': 'OWM7111-GDNT-9-dev',
        'OWM0131': 'OWM0131-GDNT-R',
        'FXA5': 'FXA5000-GCNT-K-dev',
        'ST898ZB': 'ST898ZB',
        '3157100': '3157100',
        'ThermoStats': ['3157100', 'ST898ZB'],
        'Centralized Door Sensor': '3323-G',
        'Centralite Water Sensor': '3315-G',
        'Centralite Temp & Humidity Sensor': '3310-G',
        'Centralite Motion Sensor': '3328-G',
        'R-Pi': 'linuxevb',
        'Common Area Camera': ['ggdemo_camera', 'common_area_camera']
    }

    for col, value in device_types.items():
        if isinstance(value, list):
            result_df[col] = result_df['oemModel'].apply(lambda x: 1 if x in value else 0)
        else:
            result_df[col] = result_df['oemModel'].apply(lambda x: 1 if x == value else 0)

    # Group by 'facility_name' and count the devices
    facility_counts = result_df.groupby(['customer_name', 'facility_name']).sum().reset_index()

    # Group by 'customer_name'
    grouped = facility_counts.groupby('customer_name')

    # Create a dictionary of DataFrames, one for each customer
    customer_dfs = {customer_name: df for customer_name, df in grouped}

    # Add total row to each DataFrame
    for customer_name, df in customer_dfs.items():
        df['Total'] = df.iloc[:, 2:].sum(axis=1)
        total_values = df.iloc[:, 2:].sum()
        total_values['customer_name'] = 'Total'
        df = df.append(total_values, ignore_index=True)
        customer_dfs[customer_name] = df

    # Create Excel file with the current date
    current_date = datetime.now().strftime('%Y-%m-%d')
    excel_file_path = f'{current_date}.xlsx'

    with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
        for customer_name, df in customer_dfs.items():
            df.to_excel(writer, sheet_name=customer_name, index=False)

    print(f"All customer data has been written to {excel_file_path}")

if __name__ == "__main__":
    process_data()
