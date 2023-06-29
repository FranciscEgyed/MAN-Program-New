import json
import xml.etree.ElementTree as ET


# Recursive function to process the XML elements
#def process_element(element):
#    # Create a dictionary for the current tag
#    tag_dict = {}
#    # Store the attributes of the element in the dictionary
#    tag_dict.update(element.attrib)
#    # Process the child elements recursively
#    for child in element:
#        child_dict = process_element(child)
#        tag_dict.setdefault(child.tag, []).append(child_dict)
#    return tag_dict


def process_element(element):
    # Create a dictionary for the current tag
    tag_dict = {}

    # Store the attributes of the element in the dictionary
    tag_dict.update(element.attrib)

    # Check if the element has text content
    if element.text and element.text.strip():
        # Store the text content under the "__text__" key
        tag_dict["__text__"] = element.text.strip()

    # Process the child elements recursively
    for child in element:
        child_dict = process_element(child)
        tag_dict.setdefault(child.tag, []).append(child_dict)

    return tag_dict

def print_dictionary(dictionary, indent=0):
    for key, value in dictionary.items():
        if isinstance(value, dict):
            print(f"{' ' * indent}{key}:")
            print_dictionary(value, indent + 5)
        elif isinstance(value, dict):
            print(f"{' ' * indent}{key}:")
            for item in value:
                if isinstance(item, dict):
                    print_dictionary(item, indent + 5)
                else:
                    print(f"{' ' * (indent + 5)}{item}")
        else:
            print(f"{' ' * indent}{key}: {value}")



# Path to your XML file
xml_file_path = "example.xml"

# Root element to process
root_element_names = ["Modules", "LengthVariants", "Wires", "CavitySeals", "Connectors",
                      "Tapes", "Terminals", "CavityPlugs"]

# Parse the XML file
tree = ET.parse(xml_file_path)
for element in root_element_names:
    # Get the root element based on the given root_element_name
    root = tree.find(".//{}".format(element))

    if root is not None:
        # Process the root element
        root_dict = process_element(root)
        # Save the dictionary to a JSON file
        json_file_path = element + ".json"
        with open(json_file_path, "w") as json_file:
            json.dump(root_dict, json_file)

    else:
        print(f"Root element '{element}' not found in the XML file.")

with open("Connectors.json", "r") as json_file:
    loaded_dictionary = json.load(json_file)
print_dictionary(loaded_dictionary)