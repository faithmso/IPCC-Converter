import tkinter as tk
from tkinter import filedialog, messagebox
import os
import pandas as pd

# Functionality imports from your code
import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime
import csv
import zipfile
import tempfile

# Track file status
file_status_list = []


def load_svhc_list():
    svhc_file_path = os.path.join(os.path.dirname(__file__), 'svhc.csv')
    svhc_list = []
    with open(svhc_file_path, newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            if 'substance_name' in row and row['substance_name']:  # Ensure the column exists and is not empty
                svhc_list.append(row['substance_name'])
            else:
                print(f"Warning: Missing or empty 'substance_name' in row: {row}")
    return svhc_list

def load_rohs_list():
    rohs_file_path = os.path.join(os.path.dirname(__file__), 'rohs.csv')
    rohs_list = []
    with open(rohs_file_path, newline='', encoding='utf-8-sig') as csvfile:  # Use 'utf-8-sig' to handle BOM
        reader = csv.DictReader(csvfile)
        for row in reader:
            # Ensure we are using the corrected column name and strip leading/trailing spaces
            substance_name = row.get('substance_name', '').strip()
            if substance_name:
                rohs_list.append(substance_name)
            else:
                print(f"Warning: Missing or empty 'substance_name' in row: {row}")
    
    print(f"Final RoHS List: {rohs_list}")
    return rohs_list



def extract_shai_files(shai_folder):
    extracted_files = []
    for filename in os.listdir(shai_folder):
        if filename.endswith('.shai'):
            shai_path = os.path.join(shai_folder, filename)
            with zipfile.ZipFile(shai_path, 'r') as zip_ref:
                temp_dir = tempfile.mkdtemp()  # Create a temporary directory
                zip_ref.extractall(temp_dir)  # Extract contents to the temporary directory
                for file in os.listdir(temp_dir):
                    if file.endswith('.xml'):
                        extracted_file_path = os.path.join(temp_dir, file)
                        extracted_files.append(extracted_file_path)
                        print(f"Extracted XML file: {extracted_file_path}")
    return extracted_files

def get_element_text(element, attribute=None, default=''):
    if element is None:
        return default
    if attribute:
        return element.get(attribute, default)
    return element.text if element.text else default

def extract_info(xml_file):
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()

        ns = {'ns': 'http://std.iec.ch/iec62474'}

        info = {
            'request_company': get_element_text(root.find('.//ns:RequestCompany', ns), 'name'),
            'supply_company': get_element_text(root.find('.//ns:SupplyCompany', ns), 'name'),
            'authorizer_name': get_element_text(root.find('.//ns:Authorizer', ns), 'name'),
            'authorizer_title': get_element_text(root.find('.//ns:Authorizer', ns), 'title'),
            'authorizer_email': get_element_text(root.find('.//ns:Authorizer', ns), 'email'),
            'authorizer_phone': get_element_text(root.find('.//ns:Authorizer', ns), 'phone'),
            'contact_name': get_element_text(root.find('.//ns:Contact', ns), 'name'),
            'contact_title': get_element_text(root.find('.//ns:Contact', ns), 'title'),
            'contact_email': get_element_text(root.find('.//ns:Contact', ns), 'email'),
            'contact_phone': get_element_text(root.find('.//ns:Contact', ns), 'phone'),
            'product_name': get_element_text(root.find('.//ns:ProductID', ns), 'name'),
            'product_id': get_element_text(root.find('.//ns:ProductID', ns), 'identifier'),
            'product_mass': get_element_text(root.find('.//ns:ProductID/ns:Mass', ns), 'mass'),
            'product_mass_unit': get_element_text(root.find('.//ns:ProductID/ns:Mass', ns), 'unitOfMeasure'),
            'response_date': get_element_text(root.find('.//ns:Response', ns), 'date', datetime.now().strftime("%Y-%m-%d")),
            'substances': []
        }

        # Add substance extraction
        for substance in root.findall('.//ns:Substance', ns):
            #initialize default values for threshhold check
            above_threshold = False
            reporting_threshold = '0.1 mass% of article [ReportingLevel:Article]'  # Default threshold

            # Check for both Threshold and ComplianceThreshold and standardize to Threshold
            threshold_element = (
                substance.find('.//ns:aboveThreshold', ns) or 
                substance.find('.//ns:aboveComplianceThreshold', ns) or  
                substance.find('.//ns:overThreshold', ns)
            )

            if threshold_element is not None:
                # Use the appropriate attribute based on which element was found
                above_threshold = (
                    threshold_element.get('aboveThreshold') == 'true' or
                    threshold_element.get('aboveComplianceThreshold') == 'true' or
                    threshhold_element.get('overThreshhold') == 'true' or
                    threshold_element.get('aboveThresholdLevel') == 'true'
                )
                #get the reporteing threshhold value or use default
                reporting_threshold = threshold_element.get('reportingThreshold', '0.1 mass% of article [ReportingLevel:Article]')
           
            # collect substance info    
            substance_info = {
                'name': get_element_text(substance, 'name'),
                'above_threshold': above_threshold,
                'threshold': reporting_threshold
            }
            # append substance information to the info dictionary
            info['substances'].append(substance_info)  # Collect substance names and threshold info

        print(f"Parsed info from {xml_file}: {info}")  # Debugging print

        # Apply fallback logic for both contact and authorizer fields

        # If contact info is missing, use authorizer info
        if not info['contact_name'] or info['contact_name'] in ['None', '-']:
            info['contact_name'] = info['authorizer_name']  # Fallback to authorizer_name
        if not info['contact_email'] or info['contact_email'] in ['None', '-']:
            info['contact_email'] = info['authorizer_email']  # Fallback to authorizer_email
        if not info['contact_phone'] or info['contact_phone'] in ['None', '-']:
            info['contact_phone'] = info['authorizer_phone']  # Fallback to authorizer_phone

        # If authorizer info is missing, use contact info
        if not info['authorizer_name'] or info['authorizer_name'] in ['None', '-']:
            info['authorizer_name'] = info['contact_name']  # Fallback to contact_name
        if not info['authorizer_email'] or info['authorizer_email'] in ['None', '-']:
            info['authorizer_email'] = info['contact_email']  # Fallback to contact_email
        if not info['authorizer_phone'] or info['authorizer_phone'] in ['None', '-']:
            info['authorizer_phone'] = info['contact_phone']  # Fallback to contact_phone

        # Add default values
        info['legal_declaration'] = "Default Legal Declaration"
        info['supplier_acceptance'] = "true"
        info['number_of_instances'] = 1
        info['instances_unit_type'] = "Each"

        return info
    except Exception as e:
        print(f"Error parsing XML file {xml_file}: {e}")
        file_status_list.append({'filename': os.path.basename(xml_file), 'status': 'failure', 'error': str(e)})
        return None

def check_svhc(substances, svhc_list):
    """
    Check if any of the substances in the product are SVHC substances, 
    and whether they are above the threshold.
    """
    svhc_info = []
    
    # Loop through the SVHC list
    for svhc in svhc_list:
        # Find the corresponding substance in the product's substance list
        matched_substance = next((sub for sub in substances if sub['name'] == svhc), None)
        
        # If the SVHC substance is present and above threshold
        if matched_substance:
            svhc_info.append({
                'name': svhc,
                'above_threshold': matched_substance['above_threshold'],
                'threshold': matched_substance['threshold']
            })
        else:
            # If the SVHC substance is not present in the product, use a standard entry
            svhc_info.append({
                'name': svhc,
                'above_threshold': False,
                'threshold': '0.1 mass% of article [ReportingLevel:Article]'
            })
    
    return svhc_info

def check_rohs(substances, rohs_list):
    """
    Check if any of the substances in the product are RoHS substances, 
    and whether they are above the threshold.
    """
    rohs_info = []
    
    # Loop through the RoHS list
    for rohs in rohs_list:
        # Find the corresponding substance in the product's substance list
        matched_substance = next((sub for sub in substances if sub['name'] == rohs), None)
        
        # If the RoHS substance is present and above threshold
        if matched_substance:
            rohs_info.append({
                'name': rohs,
                'above_threshold': matched_substance['above_threshold'],
                'threshold': matched_substance['threshold']
            })
        else:
            # If the RoHS substance is not present in the product, use a standard entry
            rohs_info.append({
                'name': rohs,
                'above_threshold': False,
                'threshold': '0.1% by weight (1000 ppm) of homogeneous materials'
            })
    
    return rohs_info

def process_folder(shai_folder):
    svhc_list = load_svhc_list()  # Load SVHC (REACH) substances
    rohs_list = load_rohs_list()  # Load RoHS substances
    complete_entries = []

    extracted_files = extract_shai_files(shai_folder)

    for xml_path in extracted_files:
        info = extract_info(xml_path)
        if info is None:  # If extraction failed, continue to the next file
            continue

        # Check substances against both SVHC (REACH) and RoHS lists
        svhc_substances = check_svhc(info['substances'], svhc_list)
        rohs_substances = check_rohs(info['substances'], rohs_list)

        # Required fields
        required_fields = [
            'response_date', 'supply_company', 'contact_name', 'contact_email', 'contact_phone',
            'authorizer_name', 'authorizer_email', 'authorizer_phone', 'product_name', 'product_id', 
            'product_mass', 'product_mass_unit', 'legal_declaration', 'supplier_acceptance', 'number_of_instances', 'instances_unit_type'
        ]

        # Check for missing fields
        missing_fields = [field for field in required_fields if info[field] == '-' or info[field] == 'None']
        if missing_fields:
            print(f"Missing fields for {xml_path}: {missing_fields}")
            file_status_list.append({'filename': os.path.basename(xml_path), 'status': 'failure', 'error': f"Missing fields: {missing_fields}"})
        else:
            complete_entry = {field: info[field] for field in required_fields}

            # Store both SVHC and RoHS substances separately
            complete_entry['svhc_substances'] = svhc_substances  # For REACH
            complete_entry['rohs_substances'] = rohs_substances  # For RoHS

            complete_entries.append(complete_entry)
            file_status_list.append({'filename': os.path.basename(xml_path), 'status': 'success', 'error': None})

    # Convert the complete entries to a DataFrame
    df = pd.DataFrame(complete_entries)
    return df, svhc_list, rohs_list  # Ensure three values are returned

class ChemSherpaApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ChemSherpa Converter")

        # Input folder selection
        self.input_folder_label = tk.Label(root, text="Select .shai Files Folder")
        self.input_folder_label.pack(pady=5)
        self.input_folder_button = tk.Button(root, text="Browse", command=self.select_input_folder)
        self.input_folder_button.pack(pady=5)
        self.input_folder_path = tk.StringVar()
        self.input_folder_display = tk.Entry(root, textvariable=self.input_folder_path, width=50)
        self.input_folder_display.pack(pady=5)

        # Output folder selection
        self.output_folder_label = tk.Label(root, text="Select Output Folder")
        self.output_folder_label.pack(pady=5)
        self.output_folder_button = tk.Button(root, text="Browse", command=self.select_output_folder)
        self.output_folder_button.pack(pady=5)
        self.output_folder_path = tk.StringVar()
        self.output_folder_display = tk.Entry(root, textvariable=self.output_folder_path, width=50)
        self.output_folder_display.pack(pady=5)

        # Checkboxes for REACH and RoHS
        self.selection_label = tk.Label(root, text="Select Compliance Type(s)")
        self.selection_label.pack(pady=5)

        self.reach_var = tk.IntVar()
        self.rohs_var = tk.IntVar()

        self.reach_checkbox = tk.Checkbutton(root, text="REACH", variable=self.reach_var)
        self.reach_checkbox.pack(pady=5)

        self.rohs_checkbox = tk.Checkbutton(root, text="RoHS", variable=self.rohs_var)
        self.rohs_checkbox.pack(pady=5)

        # Run Button
        self.run_button = tk.Button(root, text="Run", command=self.run_conversion)
        self.run_button.pack(pady=20)

    def select_input_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.input_folder_path.set(folder_selected)

    def select_output_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.output_folder_path.set(folder_selected)

    def run_conversion(self):
        input_folder = self.input_folder_path.get()
        output_folder = self.output_folder_path.get()

        if not input_folder or not os.path.exists(input_folder):
            messagebox.showerror("Error", "Please select a valid input folder.")
            return
        if not output_folder or not os.path.exists(output_folder):
            messagebox.showerror("Error", "Please select a valid output folder.")
            return

        # Check whether the user selected either REACH or RoHS or both
        include_reach = self.reach_var.get() == 1
        include_rohs = self.rohs_var.get() == 1

        if not include_reach and not include_rohs:
            messagebox.showerror("Error", "Please select at least one compliance type (REACH or RoHS).")
            return

        # Ensure the output folder exists
        os.makedirs(output_folder, exist_ok=True)

        # Process the input folder and extract data
        try:
            df, svhc_list, rohs_list = process_folder(input_folder)

            # Save the DataFrame to a CSV file
            output_csv_path = os.path.join(output_folder, 'output_info.csv')
            df.to_csv(output_csv_path, index=False)

            # Generate XML files based on the user's selection
            for _, entry in df.iterrows():
                self.generate_xml_file(entry, svhc_list, rohs_list, output_folder, include_reach, include_rohs)

            # Save the file processing status to a CSV file
            status_df = pd.DataFrame(file_status_list)
            status_csv_path = os.path.join(output_folder, 'file_processing_status.csv')
            status_df.to_csv(status_csv_path, index=False)

            messagebox.showinfo("Success", "Conversion completed successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def generate_xml_file(self, entry, svhc_list, rohs_list, output_folder, include_reach, include_rohs):
        # Ensure necessary columns are present and preprocess data
        required_columns = [
            'response_date', 'supply_company', 'contact_name', 'contact_email', 'contact_phone',
            'authorizer_name', 'authorizer_email', 'authorizer_phone', 'product_name', 'product_id', 
            'product_mass', 'product_mass_unit', 'legal_declaration', 'supplier_acceptance', 'number_of_instances', 'instances_unit_type'
        ]
        for col in required_columns:
            if col not in entry:
                raise ValueError(f"Missing required column: {col}")

        # Root element
        root = ET.Element("MainDeclaration", {
            "xmlns": "http://webstds.ipc.org/175x/2.0",
            "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
            "version": "2.0"
        })

        # BusinessInfo Section
        business_info = ET.SubElement(root, "BusinessInfo", {"mode": "Distribute"})
        response = ET.SubElement(business_info, "Response", {
            "date": entry['response_date'],
            "fieldLock": "false",
            "comment": "Generated by Assent Custom XML Generator"
        })

        # Insert the Include section here, right after the BusinessInfo and Response
        include = ET.SubElement(root, "Include")
        sectional = ET.SubElement(include, "Sectional", {"name": "MaterialInfo"})
        ET.SubElement(sectional, "SubSectional", {"name": "C"})

        # Authorizer Info
        authorizer = ET.SubElement(response, "Authorizer", {"name": entry['authorizer_name']})
        ET.SubElement(authorizer, "Email", {"address": entry['authorizer_email']})
        ET.SubElement(authorizer, "Phone", {"number": entry['authorizer_phone']})

        # SupplyCompany Info
        supply_company = ET.SubElement(response, "SupplyCompany", {"name": entry['supply_company']})
        ET.SubElement(supply_company, "CompanyID", {"identity": entry['supply_company'], "authority": "Assent Compliance"})

        # Contact Info
        contact = ET.SubElement(response, "Contact", {"name": entry['contact_name']})
        ET.SubElement(contact, "Email", {"address": entry['contact_email']})
        ET.SubElement(contact, "Phone", {"number": entry['contact_phone']})

        # Declaration Section
        ET.SubElement(business_info, "Declaration", {
            "legalType": "Standard",
            "supplierAcceptance": entry['supplier_acceptance'],
            "legalDef": "Supplier certifies that it gathered the provided information...",
            "uncertaintyStatement": ""
        })

        # Product Section
        product = ET.SubElement(root, "Product", {
            "comment": "",
            "unitType": str(entry['instances_unit_type']),
            "numberOfInstances": str(entry['number_of_instances'])  # Convert to string
        })
        product_id = ET.SubElement(product, "ProductID", {
            "itemName": entry['product_name'],
            "itemNumber":entry['product_id'],
            "manufacturingSite":"",
            "version":"", 
            "requesterItemName":"", 
            "requesterItemNumber":""
            })
        ET.SubElement(product_id, "Amount", {
            "value": str(entry['product_mass']),  # Convert to string
            "UOM": entry['product_mass_unit']
        })

        # Single MaterialInfo Section to include both REACH and RoHS substances
        material_info_section = ET.SubElement(product, "MaterialInfo")
        

         # REACH SubstanceCategoryList (if applicable)
        if include_reach:
            reach_category_list = ET.SubElement(material_info_section, "SubstanceCategoryList")
            ET.SubElement(reach_category_list, "SubstanceCategoryListID", {
                "identity": "EUREACH-0624", 
                "authority": "IPC"
            })

            for substance in entry['svhc_substances']:
                substance_category = ET.SubElement(reach_category_list, "SubstanceCategory", {
                    "name": substance['name'],
                    "reportableApplication": "All"
                })
                ET.SubElement(substance_category, "Threshold", {
                    "overThreshold": "true" if substance['above_threshold'] else "false",
                    "threshold": substance['threshold']
                })
        # Include RoHS substances if selected
        if include_rohs:
            rohs_category_list = ET.SubElement(material_info_section, "SubstanceCategoryList")
            ET.SubElement(rohs_category_list, "SubstanceCategoryListID", {
                "identity": "EUROHS-1907", 
                "authority": "IPC"
            })

            for substance in entry['rohs_substances']:
                substance_category = ET.SubElement(rohs_category_list, "SubstanceCategory", {
                    "name": substance['name'],
                    "reportableApplication": "All"
                })
                ET.SubElement(substance_category, "Threshold", {
                    "overThreshold": "true" if substance['above_threshold'] else "false",
                    "threshold": substance['threshold']
                })


        # Save the XML to a file
        xml_string = ET.tostring(root, encoding="unicode")
        xml_pretty = minidom.parseString(xml_string).toprettyxml(indent="  ")

        file_name = f"{entry['product_id']}.xml"
        file_path = os.path.join(output_folder, file_name)
        with open(file_path, "w", encoding="utf-8") as xml_file:
            xml_file.write(xml_pretty)

        print(f"XML file saved for product_id {entry['product_id']} at {file_path}")



if __name__ == "__main__":
    root = tk.Tk()
    app = ChemSherpaApp(root)
    root.geometry("600x400")  # Set the window size
    root.mainloop()

