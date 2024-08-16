import databricks.sql as sql
import time
import pandas as pd
import databricks.sql.exc
import os
from dotenv import load_dotenv
load_dotenv()

def save_to_excel(df, table_name):
    file_name = f'{table_name}.xlsx'
    df.to_excel(file_name, sheet_name=table_name, index=False)
    print(f"Data for {table_name} saved to {file_name}")
    
def fetch_data_from_databricks_with_retry(table_name, retries=3, delay=5):
    server_hostname = os.getenv('DATABRICKS_SERVER_HOSTNAME')
    http_path = os.getenv('DATABRICKS_HTTP_PATH')
    access_token = os.getenv('DATABRICKS_ACCESS_TOKEN')
    
    attempt = 0
    while attempt < retries:
        try:
            connection = sql.connect(
                server_hostname=server_hostname,
                http_path=http_path,
                access_token=access_token
            )
            # Your code to fetch data using the connection
            with connection.cursor() as cursor:
                cursor.execute(f"SELECT * FROM `vss-iot-customer-dev`.default.{table_name}")
                data = cursor.fetchall()
                columns = [desc[0] for desc in cursor.description]  # Get column names
                nodeDetailsdf = pd.DataFrame(data, columns=columns)
                return nodeDetailsdf
        except databricks.sql.exc.RequestError as e:
            attempt += 1
            print(f"Attempt {attempt} failed: {e}. Retrying in {delay} seconds...")
            time.sleep(delay)
    
    raise Exception(f"Failed to connect after {retries} attempts")

# Example call to the function
if __name__ == 'main':
    tables = [os.getenv("CUSTOMER_DETAILS"), os.getenv("NODE_DETAILS")]

    for table_name in tables:
        df = fetch_data_from_databricks_with_retry(table_name)
        save_to_excel(df, table_name)



