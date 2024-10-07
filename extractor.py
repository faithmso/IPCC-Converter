import xml.etree.ElementTree as ET
import csv

# Correct path to the XML file
xml_file = r'C:\Users\FaithMu\Desktop\Chemsherpa\IPCC\2024-09-27-11-47-19-ipc1752-dec-a3.xml'

# Parse the XML
tree = ET.parse(xml_file)
root = tree.getroot()

# Namespaces
namespaces = {'ns': 'http://webstds.ipc.org/175x/2.0'}

# Helper function to extract substances from a list
def extract_substances(category_list_id, namespace, filename):
    substances = []
    
    for substance_list in root.findall(".//ns:SubstanceCategoryList", namespaces):
        # Check if the correct SubstanceCategoryListID exists in this list
        category_list_id_elem = substance_list.find('ns:SubstanceCategoryListID', namespaces)
        if category_list_id_elem is not None and category_list_id_elem.attrib['identity'] == category_list_id:
            for substance in substance_list.findall('ns:SubstanceCategory', namespaces):
                name = substance.attrib.get('name', 'N/A')
                identity_elem = substance.find('ns:SubstanceCategoryID', namespaces)
                identity = identity_elem.attrib.get('identity', 'N/A') if identity_elem is not None else 'N/A'
                substances.append((name, identity))
    
    # Write to CSV
    with open(filename, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(['Substance Name', 'Identity'])
        writer.writerows(substances)

# Extract substances from EUREACH list
extract_substances('EUREACH-0624', namespaces, 'EUREACH_substances.csv')

# Extract substances from RoHS list
extract_substances('EUROHS-1907', namespaces, 'RoHS_substances.csv')

# Extract substances from IEC list
extract_substances('IEC_62474 D19.00', namespaces, 'IEC_substances.csv')

print("Data extraction completed and CSV files generated.")
