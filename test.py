import xml.etree.ElementTree as ET
import sqlite3

def process_element(element, cursor, table_name, column_names, parent_values=None):
    values = []
    if parent_values:
        values.extend(parent_values)

    for attr, value in element.attrib.items():
        if attr not in column_names:
            column_names.append(attr)
        values.append(value)

    if element.text:
        if 'text' not in column_names:
            column_names.append('text')
        values.append(element.text)

    for child in element:
        process_element(child, cursor, table_name, column_names, values)

    if not element.findall("*"):
        insert_query = "INSERT INTO {} ({}) VALUES ({})".format(
            table_name, ', '.join(column_names), ', '.join(['?'] * len(column_names))
        )
        cursor.execute(insert_query, values)

def xml_to_sqlite(xml_file, db_file):
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()
    except (FileNotFoundError, ET.ParseError) as e:
        print("Error: Failed to parse the XML file.")
        print(e)
        return

    connection = sqlite3.connect(db_file)
    cursor = connection.cursor()

    for root_element in root:
        table_name = root_element.tag

        column_names = set()
        for element in root_element.iter():
            column_names.add(element.tag)

        create_table_query = "CREATE TABLE IF NOT EXISTS {} ({})".format(table_name, ', '.join(column_names))
        try:
            cursor.execute(create_table_query)
            connection.commit()
        except sqlite3.Error as e:
            print("Error: Failed to create the table for root element: {}".format(table_name))
            print(e)
            continue

        process_element(root_element, cursor, table_name, list(column_names))

        try:
            connection.commit()
        except sqlite3.Error as e:
            print("Error: Failed to insert data into the table for root element: {}".format(table_name))
            print(e)

    connection.close()

# Usage example
xml_file_path = "example.xml"
db_file_path = "example.db"
xml_to_sqlite(xml_file_path, db_file_path)