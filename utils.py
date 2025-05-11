from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential
import json
import os
from collections import defaultdict
import re
import pandas as pd
from rapidfuzz import process, fuzz
import math

# Replace with your Azure Form Recognizer endpoint and API key
endpoint = ""
api_key = ""

# Payed acc
# endpoint = ""
# api_key  = ""

# Initialize the DocumentAnalysisClient
client = DocumentAnalysisClient(endpoint=endpoint, credential=AzureKeyCredential(api_key))

def extract_data_from_pdf(file_path):
    # Open the PDF file in binary mode
    with open(file_path, "rb") as pdf_file:
        poller = client.begin_analyze_document("prebuilt-document", document=pdf_file)
        result = poller.result()

    # Convert the result to a dictionary
    result_dict = {"pages": []}

    for page in result.pages:
        page_data = {
            "page_number": page.page_number,
            "lines": [],
            "key_value_pairs": [],
            "tables": []  # Initialize the tables key, even if no tables are found
        }

        # Extract lines of text
        for line in page.lines:
            page_data["lines"].append({"text": line.content})

        # Extract key-value pairs (with additional check for missing key_value_pairs)
        if result.key_value_pairs:
            for kvp in result.key_value_pairs:
                key = kvp.key.content if kvp.key else None
                value = kvp.value.content if kvp.value else None
                page_data["key_value_pairs"].append({"key": key, "value": value})

        # Check if the page has tables
        if hasattr(page, 'tables') and page.tables:
            for table in page.tables:
                table_data = []
                for cell in table.cells:
                    table_data.append({
                        "row_index": cell.row_index,
                        "column_index": cell.column_index,
                        "content": cell.content
                    })
                page_data["tables"].append(table_data)

        result_dict["pages"].append(page_data)
    
    save_folder = "D:\DS_projects\OCR for PO\Json files"
    
    # Ensure the save folder exists
    if not os.path.exists(save_folder):
        os.makedirs(save_folder)

    # Define the path to save the JSON file
    output_file = os.path.join(save_folder, "extracted_data.json")

    # Save the result_dict to a JSON file
    with open(output_file, "w") as json_file:
        json.dump(result_dict, json_file, indent=4)

    return result_dict

# Load dimensions from Excel
def load_product_dimensions(excel_path):
    # Read with the correct header row
    # df = pd.read_excel(excel_path, header=1, sheet_name=2)
    df = pd.read_excel(excel_path)

    product_dimensions = {}

    for idx, row in df.iterrows():
        item_name = str(row.get('item')).strip().lower()

        # Skip if item_name is NaN or empty
        if not item_name or item_name == 'nan':
            continue

        try:
            height = float(row.get('height', 0) or 0)
            width = float(row.get('width', 0) or 0)
            length = float(row.get('length', 0) or 0)
        except ValueError:
            height = width = length = 0

        # Quantity per pack (either from 'per box' or 'per bundle')
        qty_per_pack = row.get('per box')
        if pd.isna(qty_per_pack) or qty_per_pack == '':
            qty_per_pack = row.get('per bundle')
        try:
            qty_per_pack = int(qty_per_pack)
        except (ValueError, TypeError):
            qty_per_pack = 1

        volume_per_pack = height * width * length
        volume_per_unit = volume_per_pack / qty_per_pack if qty_per_pack else 0

        product_dimensions[item_name] = {
            "Item Name": item_name,
            # "Height": height,
            # "Width": width,
            # "Length": length,
            # "QuantityPerPack": qty_per_pack,
            # "VolumePerPack": volume_per_pack,
            "VolumePerUnit": volume_per_unit
        }

    return product_dimensions

# def get_closest_match(description, product_dimensions):
#     description = description.lower().strip()
#     possible_keys = list(product_dimensions.keys())
#     match = difflib.get_close_matches(description, possible_keys, n=1, cutoff=0.6)  # adjust cutoff as needed
#     if match:
#         return product_dimensions[match[0]]["VolumePerUnit"]

def normalize(text):
    text = text.lower()
    text = re.sub(r'[^a-z0-9\s]', '', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def all_tokens_in(text_a, text_b):
    # Check if all tokens of text_a exist in text_b
    tokens_a = set(text_a.split())
    tokens_b = set(text_b.split())
    return tokens_a.issubset(tokens_b)

def get_closest_match(description, product_dimensions):
    description_norm = normalize(description)
    
    best_match = None
    best_score = 0
    
    for key in product_dimensions:
        key_norm = normalize(key)

        # First: check if all tokens from Excel exist in the PDF description
        if all_tokens_in(key_norm, description_norm):
            return product_dimensions[key]["VolumePerUnit"]
        
        # Fallback: use fuzzy match as backup
        score = fuzz.token_sort_ratio(description_norm, key_norm)
        if score > best_score and score >= 70:
            best_match = key
    
    if best_match:
        return product_dimensions[best_match]["VolumePerUnit"]
    
    return None

def create_excel(processed_data, pdf_file_name):
    df = pd.DataFrame(processed_data)
    # If you want to only include selected columns:
    # df = df[['Outlet Code', 'Total Sales']]
    
    output_directory = "D:\\DS_projects\\OCR for PO\\Excel files\\"
    os.makedirs(output_directory, exist_ok=True)
    output_file = os.path.join(output_directory, f"Total_vol_n_sales_{pdf_file_name}.xlsx")
    df.to_excel(output_file, index=False)
    return output_file

def safe_float_parse(value):
    cleaned = re.sub(r'[,:\.]+$', '', value.replace(",", "").replace(":", ""))
    return float(cleaned)

## Extract details based on pdf-----------------------------------------------------------------------------------------------------------------

def process_cargils_data(data, product_dimensions):
    outlet_sales = defaultdict(float)
    # Regex pattern for product codes
    product_code_pattern = re.compile(r"^[A-Z]{2,3}\d{3,5}$")

    outlet_code_pattern = re.compile(r'^\d{3,4}\b')

    qty_pattern = re.compile(r'^[1-9]\d{0,2}(\.00)?$')

    matched_qty_values = []

    # Initialize a dictionary to store outlet-wise product details
    outlet_products = defaultdict(lambda: {"Outlet Name": "", "Products": []})

    # Variables to track outlet details
    current_outlet_details = None

    # Iterate through the pages and extract the required details
    for page in data["pages"]:
        lines = page["lines"]
        for i in range(len(lines)):
            text = lines[i]["text"]

            if "Cargills" in text and len(text.split()) > 1:
                supplier=text

            # Identify and store outlet details
            if ("EX" in text or "FH" in text or "FC" in text) and len(text.split()) > 1:
                if outlet_code_pattern.search(text):
                    # Split into Outlet Code and Outlet Name
                    parts = text.split(" ", 1)
                    outlet_code = parts[0]
                    outlet_name = parts[1]
                    # current_outlet_details = {"Outlet Code": outlet_code, "Outlet Name": outlet_name}
                    # print(current_outlet_details)
                
                else:
                    outlet_name = text
                    outlet_code = lines[i - 1]["text"]
                    print(f"Initial Outlet Code: {outlet_code}")
                    # current_outlet_details = {"Outlet Code": outlet_code, "Outlet Name": outlet_name}
                    # print(current_outlet_details)

                # Store the current outlet details
                current_outlet_details = {"Outlet Code": outlet_code, "Outlet Name": outlet_name}
                # Update outlet name in outlet_products
                outlet_products[outlet_code]["Outlet Name"] = outlet_name


            # Identify product lines
            if product_code_pattern.match(text):
                try:
                    # Extract product details
                    product_code = text
                    if lines[i + 1]["text"] == "- -" and lines[i + 2]["text"] == "- - -":
                        product_name = lines[i + 3]["text"]
                        pack_size = lines[i + 4]["text"]
                        qty = lines[i + 6]["text"]
                        # qty = find_valid_qty(lines, i + 6)
                        net_value = lines[i + 7]["text"]
                        vat_value = lines[i + 8]["text"]
                    else:
                        product_name = lines[i + 1]["text"]
                        pack_size = lines[i + 2]["text"]
                        qty = lines[i + 4]["text"]
                        # qty = find_valid_qty(lines, i + 4)
                        net_value = lines[i + 5]["text"]
                        vat_value = lines[i + 6]["text"]

                    # IF product name is not valid, this logic works
                    # Check product name validity
                    def is_valid_product_name(name):
                        return len(name.split()) >= 2 and any(c.isalpha() for c in name)

                    if not is_valid_product_name(product_name):
                        # Fallback: Search nearby for valid product name
                        search_range = 5
                        found = False
                        for offset in range(1, search_range + 1):
                            # Look below
                            if i + offset < len(lines):
                                name_candidate = lines[i + offset]["text"]
                                if is_valid_product_name(name_candidate):
                                    product_name = name_candidate
                                    found = True
                                    break

                            # Look above
                            if i - offset >= 0:
                                name_candidate = lines[i - offset]["text"]
                                if is_valid_product_name(name_candidate):
                                    product_name = name_candidate
                                    found = True
                                    break

                        # Search forward for next outlet code line
                        for j in range(i + offset + 1, len(lines)):
                            outlet_code_pattern = re.compile(r'^\d{3,4}\s+(EX|FH|FC)\s+.+')
                            if outlet_code_pattern.search(lines[j]["text"]):
                                outlet_line_index = j
                                break

                        # Extract based on outlet line index
                        if 'outlet_line_index' in locals():
                            outlet_code = lines[outlet_line_index]["text"]
                            try:
                                vat_value = lines[outlet_line_index - 2]["text"]
                                net_value = lines[outlet_line_index - 3]["text"]
                            except IndexError:
                                # Handle case where lines are not available
                                vat_value = net_value = qty = cost = pack_size = ""

                        
                        qty = None

                        for k in range(outlet_line_index - 1, max(outlet_line_index - 21, -1), -1):
                            candidate = lines[k]["text"].strip()
                            print(f"Checking candidate for qty: '{candidate}'")  # debug
                            if qty_pattern.match(candidate):
                                print(f"Matched quantity: {candidate}")
                                matched_qty_values.append(float(candidate.replace('.00', '')))
                                lowest_qty = min(matched_qty_values)
                                qty = f"{int(lowest_qty)}.00" 

                    # Checking for qunatity value is correct 
                    if not qty_pattern.match(qty):
                        # Search forward for next outlet code line
                        for j in range(i + 1, len(lines)):
                            outlet_code_pattern = re.compile(r'^\d{3,4}\s+(EX|FH|FC)\s+.+')
                            if outlet_code_pattern.search(lines[j]["text"]):
                                outlet_line_index = j
                                break

                        # Extract based on outlet line index
                        if 'outlet_line_index' in locals():
                            outlet_code = lines[outlet_line_index]["text"]
                            try:
                                vat_value = lines[outlet_line_index - 2]["text"]
                                net_value = lines[outlet_line_index - 3]["text"]
                            except IndexError:
                                # Handle case where lines are not available
                                vat_value = net_value = qty = cost = pack_size = ""
                        
                        # To detect qty value
                        qty = None

                        for k in range(outlet_line_index - 1, max(outlet_line_index - 21, -1), -1):
                            candidate = lines[k]["text"].strip()
                            print(f"Checking candidate for qty: '{candidate}'")  # debug
                            if qty_pattern.match(candidate):
                                print(f"Matched quantity: {candidate}")
                                matched_qty_values.append(float(candidate.replace('.00', '')))
                                lowest_qty = min(matched_qty_values)
                                qty = f"{int(lowest_qty)}.00"

                    if current_outlet_details:
                        try:
                            net_val = safe_float_parse(net_value)
                        except ValueError:
                            net_val = 0

                        try:
                            vat_val = safe_float_parse(vat_value)
                        except ValueError:
                            vat_val = 0

                        
                        
                        total_sales = net_val + vat_val
                        outlet_code = current_outlet_details["Outlet Code"]
                        outlet_products[outlet_code]["Products"].append({
                            "Product Code": product_code,
                            "Product Name": product_name,
                            "Quantity": qty,
                        })
                        outlet_sales[outlet_code] += total_sales
                except IndexError:
                    continue

    # Format the output with outlet code and supplier
    output_list = []
    for outlet_code, details in outlet_products.items():
        output_list.append({
            "Outlet Name": details["Outlet Name"],
            "Outlet Code": outlet_code,
            "Supplier": supplier,
            "Products": details["Products"],
            "Total Sales": round(outlet_sales[outlet_code], 2)
        })
    return output_list

def process_country_style_data(data, product_dimensions):
        def safe_float(value):
            try:
                return float(value.replace(",", "").strip())
            except:
                return 0.0
            
        # Regex pattern for product codes
        product_code_pattern = re.compile(r"^\d{4,6}$")
        outlet_sales = defaultdict(float)

        # Initialize a dictionary to store all product details by outlet
        outlet_products = defaultdict(lambda: defaultdict(lambda: {'name': '', 'quantity': 0.0}))

        # Iterate through the pages and extract the required details
        for page in data["pages"]:
            lines = page["lines"]
            for i in range(len(lines)):
                text = lines[i]["text"]

                if "PDK" in text and len(text.split()) > 1:
                    supplier = text
                
                # Identify product lines
                if product_code_pattern.match(text):
                    try:
                        # Extract product details
                        product_code = text
                        product_name = lines[i + 1]["text"]
                        CPrice = lines[i + 2]["text"]
                        
                        # Set default price if it's not a number
                        if not CPrice.isdigit():
                            print(f"Price is not a digit")
                            price = safe_float(lines[i + 4]["text"])
                            bg = safe_float(lines[i + 5]["text"])
                            kl = safe_float(lines[i + 6]["text"])
                            kw = safe_float(lines[i + 7]["text"])
                            ne = safe_float(lines[i + 8]["text"])
                            pl = safe_float(lines[i + 9]["text"])
                            tr = safe_float(lines[i + 10]["text"])
                        else:
                            print(f"Price is a digit")
                            price = safe_float(lines[i + 3]["text"])
                            bg = safe_float(lines[i + 4]["text"])
                            kl = safe_float(lines[i + 5]["text"])
                            kw = safe_float(lines[i + 6]["text"])
                            ne = safe_float(lines[i + 7]["text"])
                            pl = safe_float(lines[i + 8]["text"])
                            tr = safe_float(lines[i + 9]["text"])
                        

                        for outlet, qty in zip(['BG', 'KL', 'KW', 'NE', 'PL', 'TR'], [bg, kl, kw, ne, pl, tr]):
                            outlet_products[outlet][product_code]['name'] = product_name
                            outlet_products[outlet][product_code]['quantity'] += qty
                            outlet_products[outlet][product_code]['price'] = price

                            # Do not multiply with C/S
                            outlet_sales[outlet] += qty * price

                        print(f"Product Code: {product_code}, Product Name: {product_name}, Price: {price}, BG: {bg}, KL: {kl}, KW: {kw}, NE: {ne}, PL: {pl}, TR: {tr}")

                    except IndexError:
                        continue

        # Create a structured output for outlet products
        outlet_summary = []
        print(outlet_products)

        for outlet, products in outlet_products.items():
            outlet_data = {
                "Outlet Code": outlet,
                "Supplier": supplier,
                "Total Sales": round(outlet_sales[outlet], 2),
                "Products": []
            }
            for product_code, details in products.items():
                outlet_data["Products"].append({
                    "Product Code": product_code,
                    "Product Name": details['name'],
                    "Quantity": details['quantity']
                })
            outlet_summary.append(outlet_data)

        # print(outlet_summary)
        # Directly return the outlet_summary list with jsonify
        return outlet_summary

def process_Softlogic_data(data, product_dimensions):
        # Initialize a dictionary to store outlet-wise product details
        outlet_products = defaultdict(list)
        outlet_sales = defaultdict(float)

        # Variables to track outlet code and outlet name
        current_outlet_code = None
        current_outlet_name = None

        # Helper function to validate and adjust line for decimal values
        def get_decimal_value(lines, start_index, max_index):
            index = start_index
            while index < max_index:
                value = lines[index]["text"].strip()
                if re.match(r"^\d+(\.\d+)?$", value):  # Check if the value is a decimal number
                    return value, index
                index += 1
            return None, start_index  # Fallback to original index if no valid value is found

        # Iterate through the pages and extract the required details
        for page in data["pages"]:
            lines = page["lines"]
            for i in range(len(lines)):
                text = lines[i]["text"].strip()

                if "Softlogic" in text and len(text.split()) > 1:
                    supplier = text

                # Identify and store outlet code and name
                if re.match(r"^\d{5}$", text):  # Matches exactly 5-digit numeric codes
                    try:
                        current_outlet_code = text  # The 5-digit outlet code
                        current_outlet_name = lines[i + 1]["text"].strip()  # The next line contains the outlet name
                    except IndexError:
                        current_outlet_code = None
                        current_outlet_name = None

                # Identify product lines
                if re.match(r"^\d{6}$", text):  # Matches 6-digit item codes
                    try:
                        # Extract product details
                        item_code = text
                        item_description = lines[i + 1]["text"].strip()
                        sale = 0

                        # Validate and adjust 'Price'
                        price, price_index = get_decimal_value(lines, i + 2, len(lines))

                        vat_value = get_decimal_value(lines, price_index + 1, len(lines))

                        # Validate and adjust 'Order in Quantity'
                        order_in_quantity, _ = get_decimal_value(lines, price_index + 5, len(lines))

                        print(f"Item Code: {item_code}, Price: {price}, Order in Quantity: {order_in_quantity}")

                        # Add product details to the outlet's list
                        if current_outlet_code and current_outlet_name:
                            order_in_quantity = safe_float_parse(order_in_quantity)
                            price = safe_float_parse(price)
                            vat_value = safe_float_parse(vat_value[0])

                            sale = order_in_quantity * price * (vat_value + 100) / 100

                            outlet_products[(current_outlet_code, current_outlet_name)].append({
                                "Item Code": item_code,
                                "Product Name": item_description,
                                "Price": price,
                                "Vat": vat_value,
                                "Order in Quantity": order_in_quantity,
                            })

                        outlet_sales[(current_outlet_code, current_outlet_name)] += sale
                    except IndexError:
                        continue

        # Prepare a structured output for outlet products
        structured_output = []

        for (outlet_code, outlet_name), products in outlet_products.items():
            structured_output.append({
                "Outlet Code": outlet_code,
                "Outlet Name": outlet_name,
                "Supplier": supplier,
                "Products": products,
                "Total Sales": round(outlet_sales.get((outlet_code, outlet_name), 0), 2)
            })

        # Output the JSON-formatted string
        # outlet_products_json = json.dumps(structured_output, indent=4, sort_keys=False)

        # return jsonify(json.loads(outlet_products_json))
        print(structured_output)
        return structured_output

def process_Laugfs_data(data, product_dimensions):

        # Initialize a dictionary to store outlet-wise product details
        outlet_products = defaultdict(list)
        outlet_sales = defaultdict(float)
        outlet_volumes = defaultdict(float)

        # Variable to track the current outlet name
        current_outlet_name = None

        column_names = {"Code", "Item Name", "Cost Price", "VAT Cost Price", "Quantity", "Qty", "CostPrice", "Vat Cost Price", "Vat"}

        # Iterate through the pages and extract the required details
        for page in data["pages"]:
            lines = page["lines"]
            for i in range(len(lines)):
                text = lines[i]["text"].strip()

                if "Laugfs" in text and len(text.split()) > 1:
                    supplier=text

                # Identify and store outlet name (assumes outlet name has first letter uppercase and others lowercase)
                if re.match(r"^[A-Z][a-z]+(\s[A-Z][a-z]+)*$", text) and text not in column_names:
                    current_outlet_name = text

                # Identify product lines (4 to 7 digit item codes)
                if re.match(r"^\d{4,7}$", text):
                    try:
                        # Extract product details with dynamic adjustment for missing data
                        item_code = text

                        # Helper function to get the next valid value
                        def get_valid_value(start_index, max_index):
                            index = start_index
                            while index < max_index and lines[index]["text"].strip() == "-":
                                index += 1
                            return lines[index]["text"].strip(), index

                        item_name, item_name_index = get_valid_value(i + 1, len(lines))
                        cost_price, cost_price_index = get_valid_value(item_name_index + 1, len(lines))
                        vat_cost_price, vat_cost_price_index = get_valid_value(cost_price_index + 2, len(lines))
                        quantity, _ = get_valid_value(vat_cost_price_index + 1, len(lines))

                        print(f"Item Code: {item_code}, VAT Cost Price: {vat_cost_price}, Quantity: {quantity}")

                        # Normalize description for matching
                        description_key = item_name.lower()
                        volume_per_unit = get_closest_match(description_key, product_dimensions)
                        if volume_per_unit is None or (isinstance(volume_per_unit, float) and math.isnan(volume_per_unit)):
                            volume_per_unit = 0
                        print(f"Volume per unit for {description_key}: {volume_per_unit}")

                        try:
                            quantity = safe_float_parse(quantity)
                        except ValueError:
                            quantity = 0

                        try:
                            vat_val = safe_float_parse(vat_cost_price)
                        except ValueError:
                            vat_val = 0

                        if current_outlet_name:
                            total_sales = quantity * vat_val
                            # total_volume = float(quantity.replace(",", "")) * (volume_per_unit or 0)
                            outlet_products[current_outlet_name].append({
                                "Code": item_code,
                                "Product Name": item_name,
                                # "Cost Price": cost_price,
                                # "VAT Cost Price": vat_cost_price,
                                "Quantity": quantity,
                            })
                        else:
                            total_sales = quantity * vat_val
                            # total_volume = float(quantity.replace(",", "")) * (volume_per_unit or 0)
                            # If no outlet name is found, add product to the last known outlet
                            outlet_products["Unknown Outlet"].append({
                                "Code": item_code,
                                "Product Name": item_name,
                                # "Cost Price": cost_price,
                                # "VAT Cost Price": vat_cost_price,
                                "Quantity": quantity,
                            })
                        
                        outlet_sales[current_outlet_name] += total_sales
                        # outlet_volumes[current_outlet_name] += total_volume
                    except IndexError:
                        continue

        structured_output = []

        for (current_outlet_name), products in outlet_products.items():
            structured_output.append({
                "Outlet Name": current_outlet_name,
                "Supplier": supplier,
                "Products": products,
                "Total Sales": round(outlet_sales.get(current_outlet_name, 0), 2),
                # "Total Volume (cm3)": round(outlet_volumes.get(current_outlet_name, 0), 2)
            })

        return structured_output

def process_Arpico_data(data, product_dimensions):

            # Initialize a dictionary to store outlet-wise product details
            outlet_products = defaultdict(list)
            outlet_totals = defaultdict(float)
            outlet_volumes = defaultdict(float)

            # Variable to track the current outlet name
            current_outlet_name = None

            # Function to check if a name ends with SS, SC, or Daily
            def is_valid_outlet_name(name):
                return name.endswith("SS") or name.endswith("SC") or name.endswith("Daily")

            # Regular expression to check for valid item codes (starting with 4 capital letters)
            def is_valid_item_code(code):
                return bool(re.match(r"^[A-Z]{4}\d{5,10}$", code))  # 4 letters followed by 5-10 digits

            # Initialize a set to store unique outlet names
            unique_outlet_names = set()
            next_outlet_code = None
            outlet_codes = {}

            # Iterate through the pages and extract the required details
            for page in data["pages"]:
                lines = page["lines"]
                for i in range(len(lines)):
                    text = lines[i]["text"].strip()

                    if "Arpico" in text and len(text.split()) > 1:
                        supplier=text

                    # Capture code after 'Supply' and store for next outlet
                    if text.lower() == "supply" and i + 1 < len(lines):
                        next_outlet_code = lines[i + 1]["text"].strip()

                    # When outlet name appears, map the code stored from earlier
                    if is_valid_outlet_name(text):
                        current_outlet_name = text
                        unique_outlet_names.add(current_outlet_name)
                        if next_outlet_code:
                            outlet_codes[current_outlet_name] = next_outlet_code
                            next_outlet_code = None  # reset after use


                    # Identify and store outlet name if it ends with SS, SC, or Daily
                    if is_valid_outlet_name(text):
                        current_outlet_name = text
                        unique_outlet_names.add(current_outlet_name)
                        # print(f"Current Outlet Name: {current_outlet_name}")


                    # Extract total sales amount from line after "Total"
                    if text.strip().lower() == "total" and i + 1 < len(lines):
                        amount_text = lines[i + 1]["text"].strip().replace(",", "")
                        try:
                            amount = float(amount_text)
                            if current_outlet_name:
                                outlet_totals[current_outlet_name] = amount
                            else:
                                outlet_totals["Unknown Outlet"] = amount
                        except ValueError:
                            pass  # skip if not a valid float

                    # Identify product lines (using Order No, PLU, Item Code, Description, Rate, and Ordered)
                    if re.match(r"^\d{8,11}$|^\d{8,11} \d{5,6}$|^\d{8,11} \d{5,6} [A-Z]{4}\d{5,10}$", text):  # Checking for a numeric Order No
                        try:
                            order_no = text.strip()
                            order_no_parts = order_no.split()

                            if len(order_no_parts)==3:
                                order_no= order_no_parts[0]
                                plu = order_no_parts[1]
                                item_code = order_no_parts[2]
                                description = lines[i + 1]["text"].strip()
                                rate = lines[i + 4]["text"].strip()
                                vat = lines[i + 6]["text"].strip()
                                ordered = lines[i + 7]["text"].strip()

                            elif len(order_no_parts)==2:
                                order_no= order_no_parts[0]
                                plu = order_no_parts[1]
                                item_code = lines[i + 1]["text"].strip()
                                description = lines[i + 2]["text"].strip()
                                rate = lines[i + 5]["text"].strip()
                                vat = lines[i + 7]["text"].strip()
                                ordered = lines[i + 8]["text"].strip()
                            
                            else:
                                plu_item_code = lines[i + 1]["text"].strip()
                            
                                # Split the PLU and Item Code if they are together
                                plu_parts = plu_item_code.split()
                                if len(plu_parts) == 2:  # PLU and Item Code are separated by space
                                    plu = plu_parts[0]
                                    item_code = plu_parts[1]
                                    description = lines[i + 2]["text"].strip()
                                    rate = lines[i + 5]["text"].strip()
                                    vat = lines[i + 7]["text"].strip()
                                    ordered = lines[i + 8]["text"].strip()

                                else:
                                    # If PLU and Item Code are in separate fields, assign them accordingly
                                    plu = plu_item_code
                                    item_code = lines[i + 2]["text"].strip()
                                    description = lines[i + 3]["text"].strip()
                                    rate = lines[i + 6]["text"].strip()
                                    vat = lines[i + 8]["text"].strip()
                                    ordered = lines[i + 9]["text"].strip()
                            

                            # Normalize description for matching
                            description_key = description.lower()
                            volume_per_unit = get_closest_match(description_key, product_dimensions)
                            if volume_per_unit is None or (isinstance(volume_per_unit, float) and math.isnan(volume_per_unit)):
                                volume_per_unit = 0
                            # print(f"Volume per unit for {description_key}: {volume_per_unit}")

                            # Validate item code (starting with 4 capital letters followed by digits)
                            if not is_valid_item_code(item_code):
                                continue  # Skip this item if it doesn't match the pattern

                            try:
                                quantity = safe_float_parse(ordered)
                            except ValueError:
                                quantity = 0

                            if current_outlet_name:
                                total_volume = quantity * (volume_per_unit or 0)
                                try:
                                    rate_val = safe_float_parse(rate)
                                    vat_val = safe_float_parse(vat)
                                    total_sales = quantity * rate_val * (1 + vat_val / 100)
                                except ValueError:
                                    total_sales = 0  # or `continue` if you want to skip the entry
                                # Add product details to the correct outlet
                                outlet_products[current_outlet_name].append({
                                    # "Order No": order_no,
                                    # "PLU": plu,
                                    "Item Code": item_code,
                                    "Product Name": description,
                                    "Rate": rate,
                                    "Quantity": ordered,
                                })
                            else:
                                total_volume = quantity * (volume_per_unit or 0)
                                try:
                                    rate_val = safe_float_parse(rate)
                                    vat_val = safe_float_parse(vat)
                                    total_sales = quantity * rate_val * (1 + vat_val / 100)
                                except ValueError:
                                    total_sales = 0  # or `continue` if you want to skip the entry

                                # If no valid outlet name is found, assign to "Unknown Outlet"
                                outlet_products["Unknown Outlet"].append({
                                    # "Order No": order_no,
                                    # "PLU": plu,
                                    "Item Code": item_code,
                                    "Product Name": description,
                                    "Rate": rate,
                                    "Quantity": ordered,
                                })
                            
                            outlet_volumes[current_outlet_name] += total_volume
                            outlet_totals[current_outlet_name] += total_sales
                        except IndexError:
                            continue

            structured_output = []

            for (current_outlet_name), products in outlet_products.items():
                structured_output.append({
                    "Outlet Code": outlet_codes.get(current_outlet_name, "Unknown"),
                    "Outlet Name": current_outlet_name,
                    "Supplier": supplier,
                    "Products": products,
                    "Total Sales": round(outlet_totals.get(current_outlet_name, 0), 2)
                    # "Total Volume (cm3)": round(total_volume,2)
                })

            # Convert the extracted data to JSON format
            # outlet_products_json = json.dumps(structured_output, indent=4, sort_keys=False)

            # return jsonify(json.loads(outlet_products_json))

            return structured_output

def process_summary_order_data(data):
    outlet_sales = defaultdict(float)
    outlet_info = {}

    # Regex pattern for 4-digit outlet code
    outlet_code_pattern = re.compile(r"^\d{4}$")

    # Flatten all lines into a single list of texts
    all_texts = []
    for page in data["pages"]:
        for line in page["lines"]:
            text = line.get("text", "").strip()
            if text:
                all_texts.append(text)

    i = 0
    while i < len(all_texts) - 6:  # Minimum 7 fields per order
        if outlet_code_pattern.fullmatch(all_texts[i]):
            try:
                outlet_code = all_texts[i]
                outlet_name = all_texts[i + 1]
                order_no = all_texts[i + 2]
                order_date = all_texts[i + 3]
                net_value = all_texts[i + 4]
                vat_value = all_texts[i + 5]
                gross_value = all_texts[i + 6]

                net_val = safe_float_parse(net_value)
                vat_val = safe_float_parse(vat_value)
                gross_val = safe_float_parse(gross_value)

                outlet_sales[outlet_code] += gross_val
                outlet_info[outlet_code] = outlet_name

                i += 7
            except IndexError:
                break
        else:
            i += 1

    # Prepare final formatted output
    output_list = []
    for outlet_code, total in outlet_sales.items():
        output_list.append({
            "Outlet Code": outlet_code,
            "Outlet Name": outlet_info.get(outlet_code, ""),
            "Order No": order_no,
            "Order Date": order_date,
            "Net Value": net_val,
            "VAT Value": vat_val,
            "Gross Sales": gross_val
        })

    return output_list

def process_other_data(data, product_dimensions):
    outlet_products = defaultdict(list)
    product_sales = defaultdict(float)
    outlet_sales = defaultdict(float)
    outlet_volumes = defaultdict(float)

    for page in data["pages"]:
        lines = page["lines"]
        outlet_names = []

        for i in range(len(lines)):
            text = lines[i]["text"]
            if re.match(r"^[A-Z]{3,4}\d?$", text):
                outlet_names.append(text)

        for i in range(len(lines)):
            text = lines[i]["text"]

            if "Country" in text and len(text.split()) > 1:
                supplier = text.replace(":", "").strip()

            if text.startswith("*"):
                try:
                    item_code = text.replace('*', '').strip()
                    item_description = lines[i + 1]["text"].strip()
                    price_text = lines[i + 2]["text"]
                    price = float(price_text.replace(",", "").strip())
                    case = int(lines[i + 3]["text"])
                    outlet_cases = [int(lines[i + j]["text"]) * case for j in range(4, 14, 2)]

                    # Normalize description for matching
                    description_key = item_description.lower()
                    volume_per_unit = get_closest_match(description_key, product_dimensions)
                    if volume_per_unit is None or (isinstance(volume_per_unit, float) and math.isnan(volume_per_unit)):
                        volume_per_unit = 0
                    # print(f"Volume per unit for {description_key}: {volume_per_unit}")

                    for j, outlet_name in enumerate(outlet_names):
                        quantity = outlet_cases[j]
                        total_sales = quantity * price
                        total_volume = quantity * (volume_per_unit or 0)

                        product_sales[item_code] += total_sales
                        outlet_sales[outlet_name] += total_sales
                        outlet_volumes[outlet_name] += total_volume

                        outlet_products[outlet_name].append({
                            "Code": item_code,
                            "Product Name": item_description,
                            "Quantity": quantity,
                            "Price": price
                            # "Total Sales": total_sales,
                            # "Volume per Unit": volume_per_unit,
                            # "Total Volume": total_volume
                        })

                except (IndexError, ValueError):
                    continue

    output_list = []
    for outlet, products in outlet_products.items():
        outlet_details = {
            "Outlet Code": outlet,
            "Supplier": supplier,
            "Total Outlet Sales": round(outlet_sales[outlet],2),
            # "Total Outlet Volume (cm3)": round(outlet_volumes[outlet],2)
            "Products": products
        }
        output_list.append(outlet_details)

    return {
        "Outlet Details": output_list
    }
