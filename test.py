import configparser
import json

import psycopg2
from psycopg2 import extras


def fetch_data_from_db(bank_value, type_value, table_name, Target_Column):
    config = configparser.ConfigParser()
    config.read(".env")
    postgres_credentials = {
        "user": config.get("DEFAULT", "USER"),
        "password": config.get("DEFAULT", "PASSWORD"),
        "host": config.get("DEFAULT", "HOST"),
        "port": config.get("DEFAULT", "PORT"),
        "database": config.get("DEFAULT", "DATABASE"),
    }
    # Construct the connection string
    database_url = f"postgresql://{postgres_credentials['user']}:{postgres_credentials['password']}@{postgres_credentials['host']}:{postgres_credentials['port']}/{postgres_credentials['database']}"
    # Connect to the PostgreSQL database
    conn = psycopg2.connect(database_url)
    # Create a cursor object to execute SQL queries
    cursor = conn.cursor(cursor_factory=extras.RealDictCursor)
    try:
        # Construct the SELECT query with specific columns and WHERE conditions using f-string
        # select_query = f"SELECT * FROM {table_name} WHERE \"Bank\" = %s AND \"Type\" = %s AND 'Target_Column' = 'Narration'"
        select_query = f"SELECT * FROM {table_name} WHERE \"Bank\" = %s AND \"Type\" = %s"
        # Execute the SELECT query with the provided values for "Bank" and "Type"
        cursor.execute(select_query, (bank_value, type_value))
        # Fetch all the rows as a list of dictionaries
        result = cursor.fetchall()
        # Convert the result to JSON format
        json_data = json.dumps(result, indent=2)
        return json_data
    finally:
        # Close the cursor and connection
        cursor.close()
        conn.close()


def compare_json(json_str1, json_str2):
    # Load JSON strings into Python objects (dictionaries)
    obj1 = json.loads(json_str1)
    obj2 = json.loads(json_str2)

    # Sort dictionaries before comparing
    sorted_obj1 = json.dumps(obj1, sort_keys=True)
    sorted_obj2 = json.dumps(obj2, sort_keys=True)

    # Compare the two sorted objects
    if sorted_obj1 == sorted_obj2:
        print(column_mapping2 = json.dumps(column_mapping2, indent=2))
    else:
        print("The JSON objects are not equal.")


if __name__ == '__main__':
    column_mapping1 = fetch_data_from_db(bank_value="indusind", type_value="type1", table_name="ksv.bank_statement_column_mapping", Target_Column="Narration")
    print(column_mapping1)
    column_mapping2 = [
        {
            "Bank": "indusind",
            "Type": "type1",
            "Source_Column": "Date",
            "Target_Column": "Transaction_Date",
        },
        {
            "Bank": "indusind",
            "Type": "type1",
            "Source_Column": "Particulars",
            "Target_Column": "Narration",
        },
        {
            "Bank": "indusind",
            "Type": "type1",
            "Source_Column": "Chq./Ref. No",
            "Target_Column": "ChequeNo_RefNo",
        },
        {
            "Bank": "indusind",
            "Type": "type1",
            "Source_Column": "WithDrawal",
            "Target_Column": "Withdrawal",
        },
        {
            "Bank": "indusind",
            "Type": "type1",
            "Source_Column": "Deposit",
            "Target_Column": "Deposit",
        },
        {
            "Bank": "indusind",
            "Type": "type1",
            "Source_Column": "Balance",
            "Target_Column": "Balance",
        },
    ]
    column_mapping1 = json.dumps(column_mapping1, indent=2)
    column_mapping1 = json.loads(column_mapping1)
    narration = set(item["Source_Column"] for item in column_mapping1 if item["Bank"] == "indusind")
    print(narration)

    # compare_json(column_mapping1, column_mapping2)

