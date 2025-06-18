import sys, locale
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

import certifi 
import os
import tempfile # For creating temporary files
import usb.core
import usb.backend.libusb1
import pandas as pd
import json
import requests
from datetime import datetime, timedelta
from flask import Flask, request, jsonify, session, send_file, render_template_string
from PIL import Image, ImageDraw, ImageFont

# Brother QL imports (ensure these are installed: brother_ql, pyusb, libusb1)
try:
    from brother_ql.conversion import convert
    from brother_ql.backends.helpers import send
    from brother_ql.raster import BrotherQLRaster
except ImportError:
    print("Warning: brother_ql, pyusb, or libusb1 not installed.")
    print("Printing functionality will be disabled.")
    print("Install with: pip install brother_ql pyusb libusb1")
    # Define dummy functions if imports fail, so the app can still run
    # (though printing won't work)
    def convert(*args, **kwargs):
        print("brother_ql not imported. Printing disabled.")
        return []
    def send(*args, **kwargs):
         print("brother_ql not imported. Printing disabled.")
    class BrotherQLRaster:
         def __init__(self, model):
             self.exception_on_warning = True


app = Flask(__name__)
app.secret_key = os.urandom(24) # Keep this unique and secret

# ==========================================
# CONFIGURATION & STATE

CONFIG_FILE = "app_config.json" # File to store configuration
CONFIG = {} # Dictionary to hold current configuration

ADMIN_PASSWORD = "admin" # Update the admin password as needed

# Use a Windows-compatible temporary directory
TEMP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'temp')
os.makedirs(TEMP_DIR, exist_ok=True)

# Shopify Configuration (These could also go into config.json if preferred)
STORE_DOMAIN = "homestead-gristmill.myshopify.com"
ACCESS_TOKEN = "REPLACE_ME" # Masked for security before pushing to GitHub
API_VERSION = "2024-01"
LOCATION_ID_GID = "gid://shopify/Location/79621390578" # Verified GID for the location
GRAPHQL_URL = f"https://{STORE_DOMAIN}/admin/api/{API_VERSION}/graphql.json"
HEADERS = {
    "Content-Type": "application/json",
    "X-Shopify-Access-Token": ACCESS_TOKEN
}



# Print Job State
cancel_print_flag = False
printing_in_progress = False
current_progress = 0
total_to_print = 0

# We'll keep a global reference to the spreadsheet data once loaded
SPREADSHEET_DATA = pd.DataFrame()

# ---------------- DEBUG BUS -------------------------------------
LAST_JOB_LOG: list[str] = []          # keeps the last ~150 lines


def dbg(msg: str):
    """Collect + echo debug lines with timestamp."""
    from datetime import datetime
    ts = datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)                       # still shows up in journalctl
    LAST_JOB_LOG.append(line)
    if len(LAST_JOB_LOG) > 150:
        LAST_JOB_LOG.pop(0)
# ----------------------------------------------------------------

# ==========================================
# CONFIG FILE MANAGEMENT

def load_config():
    """Loads configuration from config.json."""
    global CONFIG
    default_config = {
        'spreadsheet_file': os.path.abspath('data.xlsx'), # Default to data.xlsx in the script directory
        'logo_path': '', # Default empty, user must set
        'barcode_folder': os.path.abspath('barcodes'), # Default to 'barcodes' folder in script directory
        'front_label_folder': os.path.abspath('frontlabels'), # Default to 'frontlabels' folder in script directory
        'back_label_folder': os.path.abspath('backlabels'), # Default to 'backlabels' folder in script directory
        'font_path_regular': '', # Default empty, user must set
        'font_path_price': '', # Default empty, user must set
    }

    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as f:
                CONFIG = json.load(f)
            # Merge with defaults in case new keys are added later
            CONFIG = {**default_config, **CONFIG}
            print(f"Loaded configuration from {CONFIG_FILE}")
        else:
            CONFIG = default_config
            print(f"No {CONFIG_FILE} found. Using default configuration.")
            save_config(CONFIG) # Save default config on first run

        # Ensure barcode folder exists
        if CONFIG.get('barcode_folder') and not os.path.exists(CONFIG['barcode_folder']):
             try:
                 os.makedirs(CONFIG['barcode_folder'])
                 print(f"Created barcode folder: {CONFIG['barcode_folder']}")
             except OSError as e:
                 print(f"Error creating barcode folder {CONFIG['barcode_folder']}: {e}")
                 # Optionally set to empty if creation fails critically?
                 # CONFIG['barcode_folder'] = '' # Consider how to handle this error gracefully
                 
        # Ensure front label folder exists
        if CONFIG.get('front_label_folder') and not os.path.exists(CONFIG['front_label_folder']):
             try:
                 os.makedirs(CONFIG['front_label_folder'])
                 print(f"Created front label folder: {CONFIG['front_label_folder']}")
             except OSError as e:
                 print(f"Error creating front label folder {CONFIG['front_label_folder']}: {e}")
                 
        # Ensure back label folder exists
        if CONFIG.get('back_label_folder') and not os.path.exists(CONFIG['back_label_folder']):
             try:
                 os.makedirs(CONFIG['back_label_folder'])
                 print(f"Created back label folder: {CONFIG['back_label_folder']}")
             except OSError as e:
                 print(f"Error creating back label folder {CONFIG['back_label_folder']}: {e}")


    except (IOError, json.JSONDecodeError) as e:
        print(f"Error loading configuration: {e}")
        CONFIG = default_config # Fallback to defaults on error

def save_config(config_dict):
    """Saves configuration to config.json."""
    try:
        with open(CONFIG_FILE, 'w') as f:
            json.dump(config_dict, f, indent=4)
        print(f"Saved configuration to {CONFIG_FILE}")
    except IOError as e:
        print(f"Error saving configuration: {e}")


# ==========================================
# SHOPIFY INVENTORY FUNCTIONS (Unchanged - kept for context)
def execute_graphql_query(query, variables):
    """Sends a GraphQL query or mutation to the Shopify API."""
    data = {"query": query, "variables": variables}
    try:
        response = requests.post(
            GRAPHQL_URL,
            headers=HEADERS,
            # --- TEMPORARY: Disable SSL Verification ---
            # This is INSECURE and for debugging the certificate issue ONLY.
            # Revert to verify=certifi.where() or a proper solution ASAP.
            verify=False,
            # --- End TEMPORARY ---
            timeout=10,
            data=json.dumps(data)
        )
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"HTTP Request failed: {e}")
        try:
            # Access response text defensively
            error_details = e.response.text if e.response else "No response details available."
        except Exception:
            error_details = "No response details available (error accessing response)."
        return {"errors": [{"message": f"Request failed: {e}", "details": error_details}]}
    except json.JSONDecodeError:
        # Add defensive check for response existence and text
        response_text = response.text if 'response' in locals() and response else "No response text available."
        return {"errors": [{"message": "Failed to decode JSON response", "details": response_text}]}
    except Exception as e:
        # Catch any other unexpected errors
        print(f"An unexpected error occurred during GraphQL query: {e}")
        return {"errors": [{"message": f"An unexpected error occurred: {str(e)}"}]}

def get_inventory_item_id_by_sku(sku):
    """Finds the Inventory Item GID for a given SKU."""
    query = """
    query getInventoryItemBySku($skuQuery: String!) {
        productVariants(first: 1, query: $skuQuery) {
            edges {
                node {
                    inventoryItem {
                        id
                    }
                }
            }
        }
    }
    """
    sku_query_string = f"sku:'{sku}'"
    variables = {"skuQuery": sku_query_string}
    result = execute_graphql_query(query, variables)

    if "errors" in result:
        print(f"GraphQL Error fetching inventory item ID: {result['errors']}")
        return None
    try:
        edges = result.get("data", {}).get("productVariants", {}).get("edges", [])
        if not edges:
            print(f"No product variant found for SKU: {sku}")
            return None
        inventory_item_id = edges[0].get("node", {}).get("inventoryItem", {}).get("id")
        if not inventory_item_id:
            print(f"Variant found for SKU {sku}, but it has no inventory item associated.")
            return None
        return inventory_item_id
    except (AttributeError, IndexError, TypeError) as e:
        print(f"Error parsing GraphQL response for SKU {sku}: {e}")
        print(f"Full response data: {result.get('data')}")
        return None

def get_current_quantity(inventory_item_id, location_id_gid):
    """Fetches the current 'available' quantity for an item at a specific location."""
    query = """
    query getCurrentQuantity($inventoryItemId: ID!) {
        inventoryItem(id: $inventoryItemId) {
            id
            inventoryLevels(first: 10) {
                edges {
                    node {
                        location {
                            id
                        }
                        quantities(names: ["available"]) {
                            name
                            quantity
                        }
                    }
                }
            }
        }
    }
    """
    variables = {"inventoryItemId": inventory_item_id}
    result = execute_graphql_query(query, variables)

    if "errors" in result:
        print(f"GraphQL Error fetching current quantity: {result['errors']}")
        return None

    try:
        inventory_item_data = result.get("data", {}).get("inventoryItem")
        if not inventory_item_data:
            print(f"No inventory item data found for ID: {inventory_item_id}")
            return None

        inventory_levels = inventory_item_data.get("inventoryLevels", {}).get("edges", [])
        for edge in inventory_levels:
            level_node = edge.get("node", {})
            level_location_id = level_node.get("location", {}).get("id")

            # Check if this inventory level matches the desired location
            if level_location_id == location_id_gid:
                quantities = level_node.get("quantities", [])
                for qty_info in quantities:
                    if qty_info.get("name") == "available":
                        return int(qty_info.get("quantity", 0)) # Return the available quantity

        # If loop finishes without finding the location
        print(f"Inventory level for location {location_id_gid} not found for item {inventory_item_id}.")
        return None # Indicate location not found for this item

    except (AttributeError, IndexError, TypeError, ValueError) as e:
        print(f"Error parsing quantity response for item {inventory_item_id}: {e}")
        print(f"Full response data: {result.get('data')}")
        return None

def update_inventory_quantity(inventory_item_id, location_id_gid, new_quantity, current_quantity):
    """Adjusts the available inventory quantity for an item at a location."""
    adjustment_quantity = new_quantity - current_quantity

    # Using inventoryAdjustQuantities mutation
    mutation = """
    mutation inventoryAdjustQuantities($input: InventoryAdjustQuantitiesInput!) {
        inventoryAdjustQuantities(input: $input) {
            inventoryAdjustmentGroup {
                createdAt
                reason
                changes {
                    name
                    delta
                }
            }
            userErrors {
                field
                message
            }
        }
    }
    """

    variables = {
        "input": {
            "name": "available",
            "reason": "correction",
            "changes": [
                {
                    "delta": adjustment_quantity,
                    "inventoryItemId": inventory_item_id,
                    "locationId": location_id_gid
                }
            ]
        }
    }

    result = execute_graphql_query(mutation, variables)

    if "errors" in result:
        print(f"GraphQL Execution Error updating inventory: {result['errors']}")
        return False

    mutation_result = result.get("data", {}).get("inventoryAdjustQuantities", {})
    user_errors = mutation_result.get("userErrors", [])

    if user_errors:
        print("Shopify User Errors occurred during update:")
        for error in user_errors:
            print(f"- Field: {error.get('field')}, Message: {error.get('message')}")
        return False

    adjustment_group = mutation_result.get("inventoryAdjustmentGroup")
    if adjustment_group:
        changes = adjustment_group.get("changes", [])
        print("\n--- Inventory Update Successful ---")
        print(f"Location: {location_id_gid}")

        for change in changes:
            if change.get("name") == "available":
                delta = change.get("delta", 0)
                print(f"Change Applied: {delta:+d} units")
                # The mutation returns the *change* applied, not the final quantity
                # We could fetch the current quantity again, but for simplicity,
                # report based on the requested change.
                # print(f"New Available Quantity (estimated): {current_quantity + delta}")
        return True
    else:
        print("Update might have failed or response format unexpected.")
        print(f"Full response data: {result.get('data')}")
        return False

def update_shopify_inventory_after_printing(sku, quantity_printed):
    """Update Shopify inventory after printing labels - ADDS the printed quantity to inventory."""
    if not sku:
        print("Cannot update inventory: No SKU provided")
        return False, "No SKU provided"

    # Get the inventory item ID for this SKU
    inventory_item_id = get_inventory_item_id_by_sku(sku)
    if not inventory_item_id:
        return False, f"SKU {sku} not found in Shopify"

    # Get current quantity
    current_quantity = get_current_quantity(inventory_item_id, LOCATION_ID_GID)
    if current_quantity is None:
        # get_current_quantity already printed error message
        return False, f"Could not retrieve current quantity for SKU {sku}"

    # Calculate new quantity (ADD printed quantity)
    new_quantity = current_quantity + quantity_printed

    # Update the inventory
    # Pass current_quantity to the update function so it can calculate delta
    success = update_inventory_quantity(inventory_item_id, LOCATION_ID_GID, new_quantity, current_quantity)

    if success:
        # Message generated by update_inventory_quantity is more detailed
        return True, f"Updated inventory for SKU {sku}: added {quantity_printed}" # Simpler success message
    else:
        # Error message printed by update_inventory_quantity
        return False, f"Failed to update inventory for SKU {sku}" # Simpler failure message


# ==========================================
# HELPER FUNCTIONS

def apply_offset(img: Image.Image, dx_px: int, dy_px: int,
                 canvas_w: int, canvas_h: int) -> Image.Image:
    """
    Paste *img* onto a blank canvas of *canvas_w × canvas_h* pixels,
    shifted by (dx_px, dy_px).  Anything that would fall outside the canvas
    is clipped.
    """
    canvas = Image.new("RGB", (canvas_w, canvas_h), "white")
    canvas.paste(img, (dx_px, dy_px))
    return canvas

def load_spreadsheet():
    """Load and validate spreadsheet data with robust type checking and cleanup."""
    global CONFIG # Use the global config
    spreadsheet_file = CONFIG.get('spreadsheet_file')

    if not spreadsheet_file:
         print("Error: Spreadsheet file path not configured.")
         return pd.DataFrame()

    try:
        if not os.path.exists(spreadsheet_file):
            print(f"Error: Spreadsheet file not found at {spreadsheet_file}")
            return pd.DataFrame()

        print(f"Attempting to load spreadsheet from {spreadsheet_file}")
        df = pd.read_excel(
            spreadsheet_file,
            dtype={
                'Product': str,
                'Variant': str,
                'SKU': str,
                'BarcodePath': str, # Expecting filename now
                'FrontLabels': str, # Add this line
                'FrontLabelFiles': str # Add this line
            }
        )
        print("Spreadsheet loaded.")

        # Validate required columns
        required = ['Product', 'Variant', 'Price', 'Timeframe', 'BarcodePath', 'SKU']
        if not all(col in df.columns for col in required):
            missing = [col for col in required if col not in df.columns]
            print(f"Error: Missing required columns in spreadsheet: {missing}")
            return pd.DataFrame()

        # Clean string columns except FrontLabels which needs special handling
        string_cols = ['Product', 'Variant', 'BarcodePath', 'SKU', 'FrontLabelFiles']
        for col in string_cols:
             if col in df.columns: # Check if column exists before processing
                df[col] = df[col].astype(str).str.strip().replace(r'\s+', ' ', regex=True)
             else:
                df[col] = '' # Add missing column with default empty strings
                
        # Handle FrontLabels column specially to preserve True/False values
        if 'FrontLabels' in df.columns:
            # Convert to proper string 'True'/'False' values
            df['FrontLabels'] = df['FrontLabels'].astype(str)
            # Normalize to ensure consistent case - 'True' or 'False'
            df['FrontLabels'] = df['FrontLabels'].apply(
                lambda x: 'True' if x.strip().lower() == 'true' else 'False'
            )
        else:
            df['FrontLabels'] = 'False'  # Default value if column missing

        # Convert numeric columns with error handling
        df['Price'] = pd.to_numeric(df['Price'], errors='coerce')
        df['Timeframe'] = pd.to_numeric(df['Timeframe'], errors='coerce').fillna(0).astype(int)

        # Drop rows where essential data is missing (Product, Variant, Price, Barcode filename)
        initial_count = len(df)
        df = df.dropna(subset=['Product', 'Variant', 'Price', 'BarcodePath'])
        # Also drop if BarcodePath is an empty string after stripping
        df = df[df['BarcodePath'] != ''].reset_index(drop=True)

        cleaned_count = len(df)

        # Check for duplicates (optional but good practice)
        duplicates = df.duplicated(subset=['Product', 'Variant'], keep=False)
        if duplicates.any():
            dup_count = duplicates.sum()
            print(f"Warning: Found {dup_count} duplicate Product/Variant combinations.")


        if df.empty:
            print("No valid data remaining after cleanup.")
            return pd.DataFrame()

        print(f"Loaded {cleaned_count}/{initial_count} valid items from spreadsheet.")
        # Include all columns, not just required ones, to ensure front label columns are preserved
        return df.reset_index(drop=True)

    except Exception as e:
        print(f"Critical error loading spreadsheet: {str(e)}")
        return pd.DataFrame()


def calculate_dates(timeframe):
    """Calculates Best By and Julian dates based on timeframe in years."""
    try:
        timeframe_days = int(timeframe) * 365
    except (ValueError, TypeError):
        print(f"Warning: Invalid timeframe '{timeframe}'. Using 0 years.")
        timeframe_days = 0

    current_date = datetime.now()
    best_by_date = current_date + timedelta(days=timeframe_days)

    # Julian date is 1 year prior to the best-by date
    julian_year_back = best_by_date - timedelta(days=365)
    # Format is DayOfYearYear (e.g., 00124 for Jan 1, 2024)
    julian_date = julian_year_back.strftime("%j") + julian_year_back.strftime("%y")[-2:]

    return best_by_date.strftime("%y-%m-%d"), julian_date


# Modified to accept barcode_filename and use CONFIG['barcode_folder']
def create_label_image(best_by, price, julian, barcode_filename):
    """Creates the label image."""
    global CONFIG # Use global config

    label_width = 991
    label_height = 306
    image = Image.new("RGB", (label_width, label_height), "white")
    draw = ImageDraw.Draw(image)

    logo_path = CONFIG.get('logo_path')
    if logo_path and os.path.exists(logo_path):
        try:
            logo = Image.open(logo_path).convert("L")
            logo = logo.resize((500, 245))
            image.paste(logo, (0, 0))
        except Exception as e:
            print(f"Error loading or pasting logo from {logo_path}: {e}")

    # Load fonts from configured paths
    font_regular_path = CONFIG.get('font_path_regular')
    font_price_path = CONFIG.get('font_path_price')
    font_regular = ImageFont.load_default() # Fallback
    font_price = ImageFont.load_default() # Fallback

    try:
        if font_regular_path and os.path.exists(font_regular_path):
             font_regular = ImageFont.truetype(font_regular_path, 35)
        else:
             print(f"Warning: Regular font not found at {font_regular_path}. Using default.")
             # Attempt to load common system fonts if config fails
             try:
                 # Check common Linux paths
                 font_regular = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 35)
                 print("Using DejavuSans as fallback regular font.")
             except OSError:
                  try:
                     # Check common Windows paths (less likely for server but good fallback)
                     font_regular = ImageFont.truetype("C:/Windows/Fonts/arial.ttf", 35)
                     print("Using Arial as fallback regular font.")
                  except OSError:
                     pass # Stick to default if system fonts fail


    except OSError:
        print(f"Warning: Could not load TrueType font from {font_regular_path}. Using default.")
        font_regular = ImageFont.load_default()


    try:
        if font_price_path and os.path.exists(font_price_path):
             font_price = ImageFont.truetype(font_price_path, 56)
        else:
            print(f"Warning: Price font not found at {font_price_path}. Using default.")
            # Attempt to load common system fonts if config fails
            try:
                font_price = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 56)
                print("Using DejavuSans-Bold as fallback price font.")
            except OSError:
                 try:
                    font_price = ImageFont.truetype("C:/Windows/Fonts/arialbd.ttf", 56) # Arial Bold
                    print("Using Arial Bold as fallback price font.")
                 except OSError:
                    pass # Stick to default if system fonts fail

    except OSError:
        print(f"Warning: Could not load TrueType font from {font_price_path}. Using default.")
        font_price = ImageFont.load_default()


    best_by_text = best_by
    price_text = f"${float(price):.2f}" # Ensure price is treated as float
    julian_text = julian

    # Calculate text widths
    bw = draw.textlength(best_by_text, font=font_regular)
    pw = draw.textlength(price_text, font=font_price)
    jw = draw.textlength(julian_text, font=font_regular)

    spacing = 20
    text_area_width = 450 # Area next to logo

    # Position text block
    # Calculate Y position - bottom aligned
    best_by_bbox = draw.textbbox((0, 0), best_by_text, font=font_regular)
    price_bbox = draw.textbbox((0, 0), price_text, font=font_price)
    julian_bbox = draw.textbbox((0, 0), julian_text, font=font_regular)

    # Find the lowest point among the text elements to align bottoms
    max_text_height = max(best_by_bbox[3] - best_by_bbox[1],
                          price_bbox[3] - price_bbox[1],
                          julian_bbox[3] - julian_bbox[1])

    # Base line for the text block - slightly above the bottom edge
    text_baseline_y = label_height - 25 # Adjust padding from bottom

    # Y positions adjusted for baseline alignment (approximate)
    best_by_y = text_baseline_y - (best_by_bbox[3] - best_by_bbox[1]) # Align top based on height difference
    price_y = text_baseline_y - (price_bbox[3] - price_bbox[1])
    julian_y = text_baseline_y - (julian_bbox[3] - julian_bbox[1])

    # Calculate total width if placed side-by-side with spacing
    total_combined_width = bw + pw + jw + (2 * spacing)

    # Determine starting X position for the text block
    left_margin = 10
    if total_combined_width < text_area_width:
        # Center within the text area if it fits
        start_x = left_margin + (text_area_width - total_combined_width) // 2
    else:
        # Start from the left margin if it's too wide
        start_x = left_margin

    # Draw text
    current_x = start_x
    draw.text((current_x, best_by_y), best_by_text, fill="black", font=font_regular)
    current_x += bw + spacing

    draw.text((current_x, price_y), price_text, fill="black", font=font_price)
    current_x += pw + spacing

    draw.text((current_x, julian_y), julian_text, fill="black", font=font_regular)


    # Barcode image
    barcode_folder = CONFIG.get('barcode_folder')
    full_barcode_path = None
    if barcode_folder and barcode_filename:
        full_barcode_path = os.path.join(barcode_folder, barcode_filename)

    if full_barcode_path and os.path.exists(full_barcode_path):
        try:
            barcode_img = Image.open(full_barcode_path).convert("L")
            # Resize barcode to fit right side
            barcode_desired_width = 500
            barcode_desired_height = 285 # Leave some space at the bottom

            # Maintain aspect ratio while fitting within desired area
            original_width, original_height = barcode_img.size
            aspect_ratio = original_width / original_height

            # Calculate new size trying to fit width first
            new_width = barcode_desired_width
            new_height = int(new_width / aspect_ratio)

            # If fitting width makes height too large, fit height instead
            if new_height > barcode_desired_height:
                 new_height = barcode_desired_height
                 new_width = int(new_height * aspect_ratio)

            # Ensure dimensions are at least 1x1
            new_width = max(1, new_width)
            new_height = max(1, new_height)

            barcode_img = barcode_img.resize((new_width, new_height))

            # Position barcode - centered vertically on the right side
            barcode_x = label_width - new_width - 10 # 10px padding from right
            barcode_y = (label_height - new_height) // 2

            image.paste(barcode_img, (barcode_x, barcode_y))
        except Exception as e:
             print(f"Error loading or pasting barcode from {full_barcode_path}: {e}")
    else:
        if barcode_filename: # Only warn if a filename was expected
             print(f"Warning: Barcode file not found at {full_barcode_path or 'None'}. Barcode area will be blank.")

    # Save the image to our Windows-compatible temp directory
    temp_path = os.path.join(TEMP_DIR, "temp_label.png")
    try:
        image.save(temp_path)
        print(f"Label image saved to {temp_path}")
        return temp_path
    except Exception as e:
        print(f"Error saving temporary label image: {e}")
        return None # Return None if image saving fails


def send_to_printer(image_path):
    """Send the label image to the Brother QL-800 printer."""
    if not image_path or not os.path.exists(image_path):
        print(f"Error: Cannot print. Image file not found at {image_path}")
        return False # Return False on failure

    # --- START: Find Printer ---
    try:
        # Ensure libusb1 backend is available on Windows
        backend = usb.backend.libusb1.get_backend()
        if backend is None:
            print("libusb1 backend not found. Ensure libusb is installed and configured.")
            return False
            
        # Use the backend object when finding the device
        printer_handle = usb.core.find(idVendor=0x04f9, idProduct=0x209b, backend=backend)

        if printer_handle is None:
            print("Printer not found over USB (VID: 0x04f9, PID: 0x209b).")
            print("Make sure the printer is connected and the driver was replaced using Zadig.")
            return False

        # Detach kernel driver if needed (often required on Linux, might not be on Windows)
        # You might need to experiment with this line on Windows
        # try:
        #     printer_handle.detach_kernel_driver(0)
        # except usb.core.USBError:
        #     pass # Ignore if no kernel driver attached
            
        # Set configuration (often needed after finding device)
        # try:
        #     printer_handle.set_configuration()
        # except usb.core.USBError as e:
        #      print(f"Failed to set USB configuration: {e}")
        #      # Depending on error, might need to re-attach driver? Or it's already claimed.
        #      # This is a complex USB detail. If printing fails, this might be a place to investigate.
        #      pass # Continue and hope it works
             
    except Exception as e:
         print(f"Error finding or configuring USB printer: {e}")
         return False # Return False on failure
    # --- END: Find Printer ---


    try:
        im = Image.open(image_path).resize((991, 306))
        im = apply_offset(im, -15, -24, 991, 306)  # 35 px left, 24 px up (moved right by 12px)
        qlr = BrotherQLRaster('QL-800')
        qlr.exception_on_warning = True

        instructions = convert(
            qlr=qlr,
            images=[im],
            label='29x90', # Or your actual label size, e.g., '62' for 62mm
            rotate='90',
            threshold=70.0,
            dither=False,
            compress=False,
            cut=True
        )

        send(instructions, printer_handle, 'pyusb')
        print("Print job sent to Brother QL-800.")
        return True 

    except Exception as e:
        print(f"Print error for Brother QL-800: {e}")

        return False 


import win32print, win32ui, win32con
from PIL import Image, ImageWin

def send_to_epson(image_path, copies=1, extra_left_mm=0):
    # Configure printer name - use the exact printer name from Windows
    printer_name = "EPSON CW-C6500Au"
    
    # Open the image file
    img = Image.open(image_path)
    img = img.convert("RGB")  # Ensure no alpha channel or palette issues

    # Process for each copy requested
    for _ in range(max(1, copies)):
        # Open a handle to the printer
        hPrinter = win32print.OpenPrinter(printer_name)
        try:
            # Get printer settings (DEVMODE) – can be modified if needed
            properties = win32print.GetPrinter(hPrinter, 2)
            devmode = properties["pDevMode"]
            # e.g., to force a specific paper or orientation (optional):
            # devmode.Orientation = 2   # 1 = Portrait, 2 = Landscape
            # devmode.PaperSize   = 1   # e.g., 1 = Letter, etc. (driver-defined for labels)
            # devmode fields can be changed before creating the DC

            # Create a Device Context (DC) for the printer using the driver settings
            hDC = win32ui.CreateDC()
            hDC.CreatePrinterDC(printer_name)  # (Alternatively, win32gui.CreateDC with devmode)

            # Get printable area (HORZRES/VERTRES) and total paper size (PHYSICALWIDTH/HEIGHT)
            printable_area = (hDC.GetDeviceCaps(win32con.HORZRES), 
                              hDC.GetDeviceCaps(win32con.VERTRES))
            physical_size = (hDC.GetDeviceCaps(win32con.PHYSICALWIDTH), 
                             hDC.GetDeviceCaps(win32con.PHYSICALHEIGHT))
            # Get printer margins (PHYSICALOFFSETX/Y). Typically small or zero for label printers.
            offsets = (hDC.GetDeviceCaps(win32con.PHYSICALOFFSETX), 
                       hDC.GetDeviceCaps(win32con.PHYSICALOFFSETY))

            # Apply extra left margin if specified (converted from mm to pixels)
            dpi = hDC.GetDeviceCaps(win32con.LOGPIXELSX)
            extra_px = int((extra_left_mm / 25.4) * dpi)

            # Rotate image if it better fits the orientation of the label
            if img.width > img.height and printable_area[0] < printable_area[1]:
                img = img.rotate(90, expand=True)
            elif img.width < img.height and printable_area[0] > printable_area[1]:
                img = img.rotate(90, expand=True)

            # Scale the image to max size within printable area
            scale_x = printable_area[0] / img.width
            scale_y = printable_area[1] / img.height
            scale = min(scale_x, scale_y)
            new_width  = int(img.width * scale)
            new_height = int(img.height * scale)

            # Center the image within the total physical page (so margins are accounted for)
            # Start coordinates (top-left) for centering:
            x = (physical_size[0] - new_width) // 2 + extra_px  # Add extra_px to shift left if needed
            y = (physical_size[1] - new_height) // 2

            # Start the print job and draw the image
            hDC.StartDoc("Python Label Print")
            hDC.StartPage()
            dib = ImageWin.Dib(img)
            dib.draw(hDC.GetHandleOutput(), (x, y, x + new_width, y + new_height))
            hDC.EndPage()
            hDC.EndDoc()
            hDC.DeleteDC()
        finally:
            win32print.ClosePrinter(hPrinter)

def create_front_label(front_label_filename):
    """Creates a front label image."""
    global CONFIG
    
    front_label_folder = CONFIG.get('front_label_folder')
    
    # Enhanced debugging output
    print(f"Front label function called with filename: '{front_label_filename}'")
    print(f"Front label folder from config: '{front_label_folder}'")
    
    # Check for empty values
    if not front_label_folder:
        print("ERROR: front_label_folder not set in configuration")
        return None
        
    if not front_label_filename or pd.isna(front_label_filename) or front_label_filename.strip() == '':
        print("ERROR: front_label_filename is empty or NaN")
        return None
    
    # Ensure filename is a string and clean it
    front_label_filename = str(front_label_filename).strip()
    
    # Construct full path
    full_front_label_path = os.path.join(front_label_folder, front_label_filename)
    print(f"Constructed full front label path: '{full_front_label_path}'")
    
    # Check if file exists
    if not os.path.exists(full_front_label_path):
        print(f"ERROR: Front label file not found at '{full_front_label_path}'")
        
        # List files in directory to help with debugging
        try:
            files = os.listdir(front_label_folder)
            print(f"Files in front label folder: {files}")
        except Exception as list_err:
            print(f"Could not list directory contents: {str(list_err)}")
        
        return None
    
    try:
        # Attempt to open and process the image
        print(f"Attempting to open front label file: '{full_front_label_path}'")
        front_image = Image.open(full_front_label_path)
        print(f"Original image size: {front_image.size}, format: {front_image.format}")
        
        # Resize the image
        ROLL_DPI      = 300        
        ROLL_WIDTH_IN = 3.5        
        ROLL_LEN_IN   = 5.0        

        max_w = int(ROLL_WIDTH_IN * ROLL_DPI)         
        max_h = int(ROLL_LEN_IN   * ROLL_DPI)         

        w, h = front_image.size
        scale = min(max_w / w, max_h / h)             
        new_size = (int(w * scale), int(h * scale))
        front_image = front_image.resize(new_size)
        
        # Save temporary file to the Windows-compatible temp directory
        temp_path = os.path.join(TEMP_DIR, "temp_front_label.png")
        front_image.save(temp_path)
        print(f"Front label image saved to temporary file: '{temp_path}'")
        
        return temp_path
    except Exception as e:
        print(f"ERROR processing front label image: {str(e)}")
        import traceback
        traceback.print_exc()  # Print full stack trace for debugging
        return None


def create_back_label(back_label_filename):
    """Creates a back label image for Epson printing."""
    global CONFIG
    
    back_label_folder = CONFIG.get('back_label_folder')
    
    # Enhanced debugging output
    print(f"Back label function called with filename: '{back_label_filename}'")
    print(f"Back label folder from config: '{back_label_folder}'")
    
    # Check for empty values
    if not back_label_folder:
        print("ERROR: back_label_folder not set in configuration")
        return None
        
    if not back_label_filename or pd.isna(back_label_filename) or back_label_filename.strip() == '':
        print("ERROR: back_label_filename is empty or NaN")
        return None
    
    # Ensure filename is a string and clean it
    back_label_filename = str(back_label_filename).strip()
    
    # Construct full path
    full_back_label_path = os.path.join(back_label_folder, back_label_filename)
    print(f"Constructed full back label path: '{full_back_label_path}'")
    
    # Check if file exists
    if not os.path.exists(full_back_label_path):
        print(f"ERROR: Back label file not found at '{full_back_label_path}'")
        
        # List files in directory to help with debugging
        try:
            files = os.listdir(back_label_folder)
            print(f"Files in back label folder: {files}")
        except Exception as list_err:
            print(f"Could not list directory contents: {str(list_err)}")
        
        return None
    
    try:
        # Attempt to open and process the image
        print(f"Attempting to open back label file: '{full_back_label_path}'")
        back_image = Image.open(full_back_label_path)
        print(f"Original image size: {back_image.size}, format: {back_image.format}")
        
        # Resize the image
        ROLL_DPI      = 300        
        ROLL_WIDTH_IN = 3.5        
        ROLL_LEN_IN   = 5.0        

        max_w = int(ROLL_WIDTH_IN * ROLL_DPI)         
        max_h = int(ROLL_LEN_IN   * ROLL_DPI)         

        w, h = back_image.size
        scale = min(max_w / w, max_h / h)             
        new_size = (int(w * scale), int(h * scale))
        back_image = back_image.resize(new_size)
        
        # Save temporary file to the Windows-compatible temp directory
        temp_path = os.path.join(TEMP_DIR, "temp_back_label.png")
        back_image.save(temp_path)
        print(f"Back label image saved to temporary file: '{temp_path}'")
        
        return temp_path
    except Exception as e:
        print(f"ERROR processing back label image: {str(e)}")
        import traceback
        traceback.print_exc()  # Print full stack trace for debugging
        return None


def check_front_label_configuration():
    """Checks the configuration and spreadsheet for front label issues."""
    global CONFIG, SPREADSHEET_DATA
    
    print("\n--- Front Label Configuration Check ---")
    
    # Check if front_label_folder is configured
    front_label_folder = CONFIG.get('front_label_folder', '')
    if not front_label_folder:
        print("WARNING: front_label_folder is not configured in config.json")
        print("Front label printing will not work until this is set.")
        return
    
    # Check if the folder exists
    if not os.path.exists(front_label_folder):
        print(f"WARNING: Configured front_label_folder does not exist: '{front_label_folder}'")
        try:
            os.makedirs(front_label_folder)
            print(f"Created front label folder: '{front_label_folder}'")
        except OSError as e:
            print(f"ERROR: Failed to create front label folder: {e}")
            print("Front label printing will not work until this folder is created.")
            return
    
    if SPREADSHEET_DATA.empty:
        print("WARNING: Spreadsheet data not loaded. Cannot check front label columns.")
        return
    
    # Check for FrontLabels column
    if 'FrontLabels' not in SPREADSHEET_DATA.columns:
        print("WARNING: 'FrontLabels' column not found in spreadsheet data.")
        print("Adding 'FrontLabels' column with default value 'False'...")
        SPREADSHEET_DATA['FrontLabels'] = 'False'
    
    # Check for FrontLabelFiles column
    if 'FrontLabelFiles' not in SPREADSHEET_DATA.columns:
        print("WARNING: 'FrontLabelFiles' column not found in spreadsheet data.")
        print("Adding 'FrontLabelFiles' column with empty values...")
        SPREADSHEET_DATA['FrontLabelFiles'] = ''
    
    # Check for products with FrontLabels=True but no filename
    if 'FrontLabels' in SPREADSHEET_DATA.columns and 'FrontLabelFiles' in SPREADSHEET_DATA.columns:
        front_label_enabled = SPREADSHEET_DATA[SPREADSHEET_DATA['FrontLabels'].str.lower() == 'true']
        missing_filenames = front_label_enabled[
            (front_label_enabled['FrontLabelFiles'].isnull()) | 
            (front_label_enabled['FrontLabelFiles'] == '')
        ]
        
        if not missing_filenames.empty:
            print(f"WARNING: Found {len(missing_filenames)} products with FrontLabels=True but missing filename.")
            print("These products will show the front label option but printing will fail:")
            for idx, row in missing_filenames.iterrows():
                print(f"- {row['Product']} / {row['Variant']}")
    
    # Check if any front label files are missing
    if 'FrontLabels' in SPREADSHEET_DATA.columns and 'FrontLabelFiles' in SPREADSHEET_DATA.columns:
        front_label_enabled = SPREADSHEET_DATA[SPREADSHEET_DATA['FrontLabels'].str.lower() == 'true']
        front_label_enabled = front_label_enabled[front_label_enabled['FrontLabelFiles'] != '']
        
        missing_files = 0
        for idx, row in front_label_enabled.iterrows():
            filename = row['FrontLabelFiles']
            if not os.path.exists(os.path.join(front_label_folder, filename)):
                missing_files += 1
                print(f"WARNING: Front label file not found: '{filename}' for {row['Product']} / {row['Variant']}")
        
        if missing_files > 0:
            print(f"WARNING: {missing_files} front label files are missing from the front label folder.")
            print(f"Upload these files to: '{front_label_folder}'")
        else:
            print(f"All front label files found in '{front_label_folder}'.")
    
    # Save any changes made to spreadsheet
    if ('FrontLabels' not in SPREADSHEET_DATA.columns or 'FrontLabelFiles' not in SPREADSHEET_DATA.columns):
        try:
            spreadsheet_path = CONFIG.get('spreadsheet_file')
            if spreadsheet_path:
                SPREADSHEET_DATA.to_excel(spreadsheet_path, index=False)
                print(f"Saved updated spreadsheet with front label columns to '{spreadsheet_path}'")
        except Exception as e:
            print(f"ERROR: Failed to save updated spreadsheet: {e}")
    
    print("--- Front Label Configuration Check Complete ---\n")


# ==========================================
# FLASK ROUTES

@app.route('/')
def index():
    # The redesigned HTML template with Admin Config tab
    html_content = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Homestead Gristmill - Label Printing</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
<style>
:root {
    --primary: #2C5F2D;
    --primary-light: #97BC62;
    --secondary: #FFB30F;
    --secondary-light: #ffd175;
    --accent: #FFD700;
    --text: #2E2E2E;
    --text-light: #6E6E6E;
    --background: #F8F7F4;
    --white: #FFFFFF;
    --gray-100: #F7F7F7;
    --gray-200: #E9ECEF;
    --gray-300: #DEE2E6;
    --gray-400: #CED4DA;
    --danger: #DC3545;
    --success: #28A745;
    --warning: #FFC107;
    --info: #17A2B8;
    --shadow-sm: 0 1px 2px rgba(0,0,0,0.05);
    --shadow: 0 4px 6px rgba(0,0,0,0.05), 0 1px 3px rgba(0,0,0,0.1);
    --shadow-lg: 0 10px 15px rgba(0,0,0,0.05), 0 44px 6px rgba(0,0,0,0.05); /* Fixed typo */
    --radius-sm: 0.375rem;
    --radius: 0.5rem;
    --radius-lg: 0.75rem;
    --transition: all 0.2s ease;
}

* {
    box-sizing: border-box;
    margin: 0;
    padding: 0;
}

body {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    background: var(--background);
    color: var(--text);
    line-height: 1.6;
    min-height: 100vh;
    display: flex;
    flex-direction: column;
}

/* Header Styles */
.header {
    background: var(--white);
    box-shadow: var(--shadow);
    padding: 1.5rem 0;
    margin-bottom: 2rem;
    position: relative;
    overflow: hidden;
}

.header-content {
    position: relative;
    z-index: 2;
    width: 90%;
    max-width: 1200px;
    margin: 0 auto;
    display: flex;
    align-items: center;
    justify-content: space-between;
}

.header h1 {
    font-weight: 700;
    font-size: 1.75rem;
    color: var(--primary);
    margin: 0;
}

.header-subtitle {
    color: var(--text-light);
    font-weight: 400;
    font-size: 1rem;
}

.header-decoration {
    position: absolute;
    top: 0;
    right: 0;
    width: 30%;
    height: 100%;
    background: linear-gradient(135deg, transparent, var(--primary-light) 80%);
    opacity: 0.1;
    z-index: 1;
}

/* Container */
.container {
    width: 90%;
    max-width: 1200px;
    margin: 0 auto 2rem;
}

/* Card Component */
.card {
    background: var(--white);
    border-radius: var(--radius);
    box-shadow: var(--shadow);
    margin-bottom: 1.5rem;
    overflow: hidden;
    border: 1px solid var(--gray-200);
}

.card-header {
    padding: 1.25rem 1.5rem;
    background: var(--gray-100);
    border-bottom: 1px solid var(--gray-200);
    display: flex;
    align-items: center;
    justify-content: space-between;
    cursor: pointer; /* Make headers clickable to expand/collapse */
}

.card-header h2 {
    font-size: 1.25rem;
    font-weight: 600;
    color: var(--primary);
    margin: 0;
    display: flex;
    align-items: center;
    gap: 0.5rem;
}

.card-header h2 i {
    color: var(--secondary);
}

.card-header .toggle-icon {
     transition: transform 0.3s ease;
}
.card-header.collapsed .toggle-icon {
    transform: rotate(180deg);
}


.card-body {
    padding: 1.5rem;
}

.card-body.collapsed {
    display: none;
}


/* Search & Filters */
.search-container {
    display: flex;
    gap: 0.75rem;
    margin-bottom: 1.25rem;
}

.search-input {
    position: relative;
    flex-grow: 1;
}

.search-input i {
    position: absolute;
    left: 1rem;
    top: 50%;
    transform: translateY(-50%);
    color: var(--text-light);
}

.search-input input {
    width: 100%;
    padding: 0.75rem 1rem 0.75rem 2.5rem;
    border: 1px solid var(--gray-300);
    border-radius: var(--radius);
    font-size: 0.95rem;
    transition: var(--transition);
    box-shadow: var(--shadow-sm);
}

.search-input input:focus {
    outline: none;
    border-color: var(--primary-light);
    box-shadow: 0 0 0 3px rgba(44, 95, 45, 0.15);
}

.button {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    gap: 0.5rem;
    background: var(--primary);
    color: var(--white);
    border: none;
    border-radius: var(--radius);
    padding: 0.75rem 1.25rem;
    font-weight: 500;
    font-size: 0.95rem;
    cursor: pointer;
    transition: var(--transition);
    box-shadow: var(--shadow-sm);
}

.button:hover {
    background: #234d24;
    transform: translateY(-1px);
    box-shadow: var(--shadow);
}

.button:active {
    transform: translateY(0);
    box-shadow: var(--shadow-sm);
}

.button.secondary {
    background: var(--secondary);
}

.button.secondary:hover {
    background: #e09800;
}

.button.outline {
    background: var(--white);
    color: var(--primary);
    border: 1px solid var(--gray-300);
}

.button.outline:hover {
    border-color: var(--primary);
    background: var(--gray-100);
}

.button.danger {
    background: var(--danger);
}

.button.danger:hover {
    background: #bd2130;
}

.button.small {
    padding: 0.5rem 0.75rem;
    font-size: 0.85rem;
}

.button:disabled {
    opacity: 0.6;
    cursor: not-allowed;
    transform: none;
    box-shadow: var(--shadow-sm);
}


/* Table Styles */
.table-container {
    overflow-x: auto;
    margin-bottom: 1rem;
    border-radius: var(--radius);
    border: 1px solid var(--gray-300);
}

table {
    width: 100%;
    border-collapse: collapse;
    font-size: 0.95rem;
}

thead {
    background: var(--gray-100);
}

th, td {
    padding: 1rem;
    text-align: left;
    border-bottom: 1px solid var(--gray-200);
}

th {
    font-weight: 600;
    color: var(--primary);
}

tbody tr {
    transition: var(--transition);
}

tbody tr:hover {
    background: var(--gray-100);
}

tbody tr:last-child td {
    border-bottom: none;
}

/* Form Controls */
.form-group {
    margin-bottom: 1.25rem;
}

.form-label {
    display: block;
    margin-bottom: 0.5rem;
    font-weight: 500;
    color: var(--text);
    font-size: 0.95rem;
}

.form-control {
    width: 100%;
    padding: 0.75rem 1rem;
    border: 1px solid var(--gray-300);
    border-radius: var(--radius);
    font-size: 0.95rem;
    transition: var(--transition);
    background: var(--white);
    box-shadow: var(--shadow-sm);
}

.form-control:focus {
    outline: none;
    border-color: var(--primary-light);
    box-shadow: 0 0 0 3px rgba(44, 95, 45, 0.15);
}

.form-control-sm {
    padding: 0.5rem 0.75rem;
    font-size: 0.85rem;
}

input[type="number"].quantity-input {
    width: 80px;
    text-align: center;
    -moz-appearance: textfield;
}

input[type="number"].quantity-input::-webkit-outer-spin-button,
input[type="number"].quantity-input::-webkit-inner-spin-button {
    -webkit-appearance: none;
    margin: 0;
}

/* Alert/Message Styles */
.alert {
    padding: 1rem 1.25rem;
    border-radius: var(--radius);
    margin-bottom: 1.25rem;
    display: flex;
    align-items: center;
    gap: 0.75rem;
    border: 1px solid transparent;
}

.alert i {
    font-size: 1.25rem;
    flex-shrink: 0; /* Prevent icon from shrinking */
}

.alert-success {
    background: rgba(40, 167, 69, 0.1);
    border-color: var(--success);
    color: #155724;
}

.alert-danger {
    background: rgba(220, 53, 69, 0.1);
    border-color: var(--danger);
    color: #721c24;
}

.alert-warning {
    background: rgba(255, 193, 7, 0.1);
    border-color: var(--warning);
    color: #856404;
}

.alert-info {
    background: rgba(23, 162, 184, 0.1);
    border-color: var(--info);
    color: #0c5460;
}

/* Progress Bar */
.progress-container {
    margin-top: 1rem;
    display: none; /* Initially hidden */
}

.progress-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 0.5rem;
}

.progress-label {
    font-weight: 500;
    font-size: 0.9rem;
    color: var(--text);
}

.progress-status {
    font-size: 0.85rem;
    color: var(--text-light);
}

.progress-bar-container {
    height: 0.8rem;
    background: var(--gray-200);
    border-radius: var(--radius-sm);
    overflow: hidden;
}

.progress-bar {
    height: 100%;
    background: var(--secondary);
    width: 0;
    transition: width 0.3s ease;
    border-radius: var(--radius-sm);
}

.progress-actions {
    display: flex;
    justify-content: flex-end;
    gap: 0.75rem;
    margin-top: 0.75rem;
}

/* Admin Section */
.hidden-section {
    display: none;
}

.admin-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(350px, 1fr));
    gap: 1.5rem;
}

/* Tabs */
.tab-header {
    display: flex;
    margin-bottom: 1.5rem;
    border-bottom: 1px solid var(--gray-300);
}

.tab-button {
    padding: 0.75rem 1.5rem;
    border: none;
    background: none;
    cursor: pointer;
    font-size: 1rem;
    font-weight: 500;
    color: var(--text-light);
    border-bottom: 2px solid transparent;
    transition: all 0.2s ease;
}

.tab-button:hover {
    color: var(--primary);
}

.tab-button.active {
    color: var(--primary);
    border-bottom-color: var(--secondary);
    font-weight: 600;
}

.tab-content > div {
    display: none;
}

.tab-content > div.active {
    display: block;
}


/* Modal Styles */
.modal-backdrop {
    position: fixed;
    top: 0;
    left: 0;
    right: 0;
    bottom: 0;
    background: rgba(0, 0, 0, 0.5);
    display: flex;
    align-items: center;
    justify-content: center;
    z-index: 1000;
    opacity: 0;
    visibility: hidden;
    transition: opacity 0.2s ease, visibility 0.2s ease;
}

.modal-backdrop.active {
    opacity: 1;
    visibility: visible;
}

.modal {
    width: 90%;
    max-width: 500px;
    background: var(--white);
    border-radius: var(--radius);
    box-shadow: var(--shadow-lg);
    transform: translateY(-20px);
    transition: transform 0.3s ease;
}

.modal-backdrop.active .modal {
    transform: translateY(0);
}

.modal-header {
    padding: 1.25rem 1.5rem;
    border-bottom: 1px solid var(--gray-200);
    display: flex;
    align-items: center;
    justify-content: space-between;
}

.modal-header h3 {
    font-size: 1.25rem;
    font-weight: 600;
    color: var(--primary);
    margin: 0;
}

.modal-close {
    background: transparent;
    border: none;
    font-size: 1.25rem;
    color: var(--text-light);
    cursor: pointer;
    transition: var(--transition);
}

.modal-close:hover {
    color: var(--text);
}

.modal-body {
    padding: 1.5rem;
}

.modal-footer {
    padding: 1rem 1.5rem;
    border-top: 1px solid var(--gray-200);
    display: flex;
    justify-content: flex-end;
    gap: 0.75rem;
}

/* Utility Classes */
.text-center { text-align: center; }
.text-right { text-align: right; }
.text-primary { color: var(--primary); }
.text-secondary { color: var(--secondary); }
.text-danger { color: var(--danger); }
.text-success { color: var(--success); }
.text-warning { color: var(--warning); }
.text-info { color: var(--info); }
.text-muted { color: var(--text-light); }

.mb-1 { margin-bottom: 0.25rem; }
.mb-2 { margin-bottom: 0.5rem; }
.mb-3 { margin-bottom: 1rem; }
.mb-4 { margin-bottom: 1.5rem; }
.mb-5 { margin-bottom: 3rem; }

.mt-1 { margin-top: 0.25rem; }
.mt-2 { margin-top: 0.5rem; }
.mt-3 { margin-top: 1rem; }
.mt-4 { margin-top: 1.5rem; }
.mt-5 { margin-top: 3rem; }

.ml-auto { margin-left: auto; }
.hidden { display: none; }
.d-flex { display: flex; }
.align-center { align-items: center; }
.justify-between { justify-content: space-between; }
.gap-2 { gap: 0.5rem; }
.gap-3 { gap: 1rem; }
.gap-4 { gap: 1.5rem; }
.flex-wrap { flex-wrap: wrap; }


/* Responsive Adjustments */
@media (max-width: 768px) {
    .header h1 {
        font-size: 1.5rem;
    }

    .header-subtitle {
        font-size: 0.9rem;
    }

    .card-header h2 {
        font-size: 1.1rem;
    }

    .admin-grid {
        grid-template-columns: 1fr;
    }

    th, td {
        padding: 0.75rem;
    }
}

@media print {
    .no-print {
        display: none !important;
    }
}
</style>
</head>
<body>
<!-- Header -->
<header class="header no-print">
    <div class="header-decoration"></div>
    <div class="header-content">
        <div>
            <h1>Homestead Gristmill</h1>
            <p class="header-subtitle">Label Printing System</p>
        </div>
    </div>
</header>

<!-- Main Content Container -->
<div class="container">
    <!-- Product List Card -->
    <div class="card">
        <div class="card-header">
            <h2><i class="fas fa-tags"></i> Product List</h2>
            <i class="fas fa-chevron-up toggle-icon"></i>
        </div>
        <div class="card-body">
            <div class="search-container">
                <div class="search-input">
                    <i class="fas fa-search"></i>
                    <input type="text" id="searchInput" placeholder="Search products, variants or SKUs..." />
                </div>
                <button class="button outline" onclick="clearSearch()">
                    <i class="fas fa-times"></i> Clear
                </button>
            </div>

            <div class="table-container">
                <table id="productTable">
                    <thead>
                        <tr>
                            <th>Product</th>
                            <th>Variant</th>
                            <th>SKU</th>
                            <th>Price</th>
                            <th>Best By</th>
                            <th>Qty</th>
                            <th>Action</th>
                        </tr>
                    </thead>
                    <tbody></tbody>
                </table>
            </div>

            <!-- Print Status -->
            <div id="printStatus" class="hidden"></div>

            <!-- Progress Container -->
            <div id="progressContainer" class="progress-container">
                <div class="progress-header">
                    <span class="progress-label">Printing Status</span>
                    <span class="progress-status" id="progressStatus">0/0 labels</span>
                </div>
                <div class="progress-bar-container">
                    <div class="progress-bar" id="progressBar"></div>
                </div>
                <div class="progress-actions">
                    <button class="button danger" id="cancelBtn" onclick="cancelPrint()">
                        <i class="fas fa-stop-circle"></i> Cancel Print
                    </button>
                </div>
            </div>

            <!-- Print Result Messages -->
            <div id="printResult" class="mt-3 hidden"></div>
            <div id="inventoryResult" class="mt-3 hidden"></div>
        </div>
    </div>

    <!-- Debug Card -->
    <div class="card no-print">
      <div class="card-header">
         <h2><i class="fas fa-bug"></i> Debug log</h2>
         <i class="fas fa-chevron-up toggle-icon"></i>
      </div>
      <div class="card-body collapsed">
         <pre id="debugPre"
              style="font-size:.8rem;white-space:pre-wrap;"></pre>
      </div>
    </div>

    <script>
    function loadDebug(){
        fetch('/debug_log')
          .then(r=>r.json())
          .then(j=>{
             document.getElementById('debugPre').textContent =
                 (j.lines||[]).join('\n');
          });
    }
    // load when card header clicked:
    document.querySelectorAll('.fa-bug')
            .forEach(i=>i.closest('.card-header')
                  .addEventListener('click',loadDebug));
    </script>

    <!-- Admin Section -->
    <div class="admin-section">
        <div class="card">
            <div class="card-header">
                <h2><i class="fas fa-lock"></i> Admin Access</h2>
                <i class="fas fa-chevron-up toggle-icon"></i>
            </div>
            <div class="card-body">
                <div class="d-flex gap-3">
                    <div class="form-group" style="flex-grow: 1;">
                        <input type="password" id="adminPassword" class="form-control" placeholder="Enter admin password" />
                    </div>
                    <button class="button" id="adminLoginBtn" onclick="adminLogin()">
                        <i class="fas fa-sign-in-alt"></i> Authenticate
                    </button>
                </div>
                <div id="adminResult" class="mt-3 hidden"></div>

                <div id="adminControls" class="hidden-section mt-4">
                    <!-- Tab Headers -->
                    <div class="tab-header">
                        <button class="tab-button active" data-tab="data-management">Data Management</button>
                        <button class="tab-button" data-tab="resource-locations">Resource Locations</button>
                    </div>

                    <!-- Tab Content -->
                    <div class="tab-content">
                        <!-- Data Management Tab -->
                        <div id="data-management" class="active">
                            <div class="admin-grid">
                                <!-- Upload Data Card -->
                                <div class="card">
                                    <div class="card-header card-nested-header">
                                        <h2><i class="fas fa-file-upload"></i> Upload Spreadsheet</h2>
                                    </div>
                                    <div class="card-body">
                                        <p class="mb-3 text-muted">Upload a new Excel (.xlsx) or CSV (.csv) file to update all product data. File must contain columns: Product, Variant, Price, Timeframe, BarcodePath (filename only), SKU.</p>
                                        <div class="form-group">
                                            <input type="file" id="spreadsheetFile" class="form-control" accept=".csv,.xlsx" />
                                        </div>
                                        <button class="button" id="uploadSpreadsheetBtn" onclick="uploadSpreadsheet()">
                                            <i class="fas fa-upload"></i> Upload Spreadsheet
                                        </button>
                                        <div id="uploadResult" class="mt-3 hidden"></div>
                                    </div>
                                </div>

                                <!-- Download Data Card -->
                                <div class="card">
                                    <div class="card-header card-nested-header">
                                        <h2><i class="fas fa-file-download"></i> Download Spreadsheet</h2>
                                    </div>
                                    <div class="card-body">
                                        <p class="mb-3 text-muted">Download the current product data as an Excel file for backup or editing.</p>
                                        <button class="button secondary" onclick="downloadSpreadsheet()">
                                            <i class="fas fa-download"></i> Download Current Data
                                        </button>
                                    </div>
                                </div>
                            </div>

                             <!-- Add/Edit Product Card -->
                            <div class="card mt-4">
                                <div class="card-header card-nested-header">
                                    <h2><i class="fas fa-plus-circle"></i> Add / Edit Product</h2>
                                </div>
                                <div class="card-body">
                                    <p class="mb-3 text-muted">Enter details to add a new product or update an existing one (matching Product and Variant). Uploading a barcode will replace the old one.</p>
                                    <div class="d-flex gap-3 flex-wrap">
                                        <div class="form-group" style="flex: 1; min-width: 200px;">
                                            <label class="form-label">Product Name</label>
                                            <input type="text" id="newProductName" class="form-control" placeholder="Product Name" required />
                                        </div>
                                        <div class="form-group" style="flex: 1; min-width: 200px;">
                                            <label class="form-label">Variant</label>
                                            <input type="text" id="newVariant" class="form-control" placeholder="Variant" required />
                                        </div>
                                    </div>

                                    <div class="d-flex gap-3 flex-wrap">
                                        <div class="form-group" style="flex: 1; min-width: 200px;">
                                            <label class="form-label">SKU</label>
                                            <input type="text" id="newSKU" class="form-control" placeholder="SKU" />
                                        </div>
                                        <div class="form-group" style="width: 150px;">
                                            <label class="form-label">Price</label>
                                            <input type="number" step="0.01" id="newPrice" class="form-control" placeholder="Price" value="0.00" />
                                        </div>
                                        <div class="form-group" style="width: 150px;">
                                            <label class="form-label">Timeframe (Years)</label>
                                            <input type="number" id="newTimeframe" class="form-control" placeholder="Timeframe" value="0" min="0" />
                                        </div>
                                    </div>

                                    <div class="form-group">
                                        <label class="form-label">Upload Barcode Image (.png, .jpg, etc.)</label>
                                        <input type="file" id="newBarcodeFile" class="form-control" accept="image/*" />
                                        <small class="text-muted">Upload a new image file for the barcode. Only the filename will be stored in the spreadsheet.</small>
                                    </div>

                                    <div class="form-group mt-3">
                                        <label class="form-label">Front Label Options</label>
                                        <div class="d-flex flex-column">
                                            <div class="mb-2">
                                                <input type="checkbox" id="newHasFrontLabel" class="mr-2">
                                                <label for="newHasFrontLabel">This product has front labels</label>
                                            </div>
                                            
                                            <div id="frontLabelFileSection" style="display: none;">
                                                <div class="form-group mt-2">
                                                    <label class="form-label">Front Label Image (.png, .jpg, etc.)</label>
                                                    <input type="file" id="newFrontLabelFile" class="form-control" accept="image/*">
                                                    <small class="text-muted">Upload a new image file for the front label. Only the filename will be stored in the spreadsheet.</small>
                                                </div>
                                            </div>
                                        </div>
                                    </div>

                                    <button class="button secondary" id="addEditProductBtn" onclick="addOrEditProduct()">
                                        <i class="fas fa-save"></i> Add / Update Product
                                    </button>

                                    <div id="addEditResult" class="mt-3 hidden"></div>
                                </div>
                            </div>
                        </div>

                        <!-- Resource Locations Tab -->
                        <div id="resource-locations">
                             <div class="card">
                                <div class="card-header card-nested-header">
                                    <h2><i class="fas fa-folder-open"></i> Resource Locations</h2>
                                </div>
                                <div class="card-body">
                                     <p class="mb-3 text-muted">Configure the file paths for the system resources. These paths are relative to the server running the application.</p>

                                    <div class="form-group">
                                        <label class="form-label">Spreadsheet File Path (.xlsx or .csv)</label>
                                        <input type="text" id="configSpreadsheetPath" class="form-control" placeholder="/path/to/data.xlsx" />
                                        <small class="text-muted">The full path to your main product data spreadsheet.</small>
                                    </div>

                                    <div class="form-group">
                                        <label class="form-label">Logo Image Path (.png, .jpg, etc.)</label>
                                        <input type="text" id="configLogoPath" class="form-control" placeholder="/path/to/logo.png" />
                                        <small class="text-muted">The full path to the logo image used on labels (optional).</small>
                                    </div>

                                    <div class="form-group">
                                        <label class="form-label">Barcode Images Folder Path</label>
                                        <input type="text" id="configBarcodeFolder" class="form-control" placeholder="/path/to/barcode/folder" />
                                        <small class="text-muted">The full path to the folder containing barcode image files.</small>
                                    </div>
                                    
                                    <div class="form-group">
                                        <label class="form-label">Front Label Images Folder Path</label>
                                        <input type="text" id="configFrontLabelFolder" class="form-control" placeholder="/path/to/frontlabel/folder" />
                                        <small class="text-muted">The full path to the folder containing front label images.</small>
                                    </div>
                                    
                                    <div class="form-group">
                                        <label class="form-label">Back Label Images Folder Path</label>
                                        <input type="text" id="configBackLabelFolder" class="form-control" placeholder="/path/to/backlabel/folder" />
                                        <small class="text-muted">The full path to the folder containing back label images.</small>
                                    </div>

                                     <div class="form-group">
                                        <label class="form-label">Regular Font File Path (.ttf)</label>
                                        <input type="text" id="configFontRegularPath" class="form-control" placeholder="/path/to/arial.ttf" />
                                        <small class="text-muted">The full path to the TrueType font file for regular text (e.g., Best By, Julian date).</small>
                                    </div>

                                    <div class="form-group">
                                        <label class="form-label">Price Font File Path (.ttf)</label>
                                        <input type="text" id="configFontPricePath" class="form-control" placeholder="/path/to/arialbd.ttf" />
                                        <small class="text-muted">The full path to the TrueType font file for price text (e.g., Arial Bold).</small>
                                    </div>


                                    <button class="button secondary" id="saveConfigBtn" onclick="saveConfig()">
                                        <i class="fas fa-save"></i> Save Settings
                                    </button>

                                    <div id="adminConfigResult" class="mt-3 hidden"></div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
</div>

<!-- Print Confirmation Modal -->


<script>
// State Management
let productData = [];
let printInterval = null;
let confirmationData = {
    product: '',
    variant: '',
    sku: '',
    quantity: 1,
    updateInventory: true,
    button: null // Reference to the clicked button
};
let isAdminLoggedIn = false; // Track admin status

// Utility Functions
function $(id) {
    return document.getElementById(id);
}

function showAlert(elementId, message, type = 'success') {
    const element = $(elementId);
    if (!element) {
        console.error(`Element with ID ${elementId} not found.`);
        return;
    }
    element.innerHTML = `
        <div class="alert alert-${type}">
            <i class="fas fa-${type === 'success' ? 'check-circle' : type === 'danger' ? 'exclamation-circle' : type === 'warning' ? 'exclamation-triangle' : 'info-circle'}"></i>
            <div>${message}</div>
        </div>
    `;
    element.classList.remove('hidden');
}

function hideAlert(elementId) {
    const element = $(elementId);
     if (element) {
        element.innerHTML = '';
        element.classList.add('hidden');
    }
}

function setLoading(button, isLoading) {
    if (!button) return;
    if (isLoading) {
        const originalText = button.innerHTML;
        button.setAttribute('data-original-text', originalText);
        button.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Loading...';
        button.disabled = true;
    } else {
        const originalText = button.getAttribute('data-original-text');
        if (originalText) {
            button.innerHTML = originalText;
        }
        button.disabled = false;
    }
}

// Card Collapse/Expand
function setupCardToggles() {
    document.querySelectorAll('.card-header').forEach(header => {
        header.addEventListener('click', function() {
            const body = this.nextElementSibling;
            if (body && body.classList.contains('card-body')) {
                const isCollapsed = body.classList.toggle('collapsed');
                this.classList.toggle('collapsed', isCollapsed);
                const icon = this.querySelector('.toggle-icon');
                if (icon) {
                     if (isCollapsed) {
                        icon.classList.remove('fa-chevron-up');
                        icon.classList.add('fa-chevron-down');
                    } else {
                         icon.classList.remove('fa-chevron-down');
                        icon.classList.add('fa-chevron-up');
                    }
                }
            }
        });
    });
}


// Modal Functions
function openModal(modalId) {
    const modal = $(modalId);
    if (modal) {
        modal.classList.add('active');
        document.body.style.overflow = 'hidden';
    }
}

function closeModal(modalId) {
     const modal = $(modalId);
     if (modal) {
        modal.classList.remove('active');
        document.body.style.overflow = '';
    }
}


// Tab Functions
function setupTabs() {
    document.querySelectorAll('.tab-button').forEach(button => {
        button.addEventListener('click', function() {
            const tabId = this.getAttribute('data-tab');
            // Remove active from all buttons and contents
            document.querySelectorAll('.tab-button').forEach(btn => btn.classList.remove('active'));
            document.querySelectorAll('.tab-content > div').forEach(content => content.classList.remove('active'));

            // Add active to clicked button and corresponding content
            this.classList.add('active');
            $(tabId).classList.add('active');
        });
    });
}


// Search & Filter Functions
function filterProducts(query) {
    const lowerQuery = query.toLowerCase().trim();
    if (!productData || productData.length === 0) return [];
    return productData.filter(item => {
        return (
            (item.Product?.toLowerCase() || '').includes(lowerQuery) ||
            (item.Variant?.toLowerCase() || '').includes(lowerQuery) ||
            (item.SKU?.toLowerCase() || '').includes(lowerQuery)
        );
    });
}

function handleSearch() {
    const query = $('searchInput').value;
    const filtered = filterProducts(query);
    renderTable(filtered);
}

function clearSearch() {
    $('searchInput').value = '';
    renderTable(productData);
}

$('searchInput').addEventListener('keyup', handleSearch);

// Data Loading Functions
async function checkAdminStatus() {
    try {
        const response = await fetch('/check_admin');
        const data = await response.json();
        isAdminLoggedIn = data.loggedIn;
        if (isAdminLoggedIn) {
            $('adminControls').classList.remove('hidden-section');
            // Automatically load config when admin is logged in
            await loadConfig();
             // Collapse admin password input section after successful login
             const adminCardBody = $('adminPassword').closest('.card-body');
             const adminCardHeader = adminCardBody.previousElementSibling;
             if (adminCardBody && adminCardHeader) {
                 adminCardBody.classList.add('collapsed');
                 adminCardHeader.classList.add('collapsed');
                  const icon = adminCardHeader.querySelector('.toggle-icon');
                    if (icon) {
                       icon.classList.remove('fa-chevron-up');
                       icon.classList.add('fa-chevron-down');
                   }
             }
        } else {
             $('adminControls').classList.add('hidden-section');
             // Ensure admin password input section is expanded if not logged in
             const adminCardBody = $('adminPassword').closest('.card-body');
              const adminCardHeader = adminCardBody.previousElementSibling;
              if (adminCardBody && adminCardHeader) {
                 adminCardBody.classList.remove('collapsed');
                 adminCardHeader.classList.remove('collapsed');
                  const icon = adminCardHeader.querySelector('.toggle-icon');
                    if (icon) {
                       icon.classList.remove('fa-chevron-down');
                       icon.classList.add('fa-chevron-up');
                   }
             }
        }
    } catch (error) {
        console.error('Admin check error:', error);
        isAdminLoggedIn = false;
        $('adminControls').classList.add('hidden-section');
    }
}

function queuePrint(btn, labelType){
    const tr  = btn.closest('tr');
    const qty = tr.querySelector('.qty').value || 1;

    // ⬇️  NEW – ask once, returns true/false
    const doInv = confirm(
        "Add the quantity you're about to print to Shopify inventory?"
    );

    fetch('/print_labels',{
        method:'POST',
        headers:{'Content-Type':'application/json'},
        body:JSON.stringify({
            product : tr.children[0].innerText.trim(),
            variant : tr.children[1].innerText.trim(),
            quantity: parseInt(qty,10),
            label_type: labelType,
            updateInventory: doInv            // ← what the user chose
        })
    })
    .then(r=>r.json())
    .then(j=>{
        // Reset progress displays
        hideAlert('printResult');
        hideAlert('inventoryResult');
        $('progressContainer').style.display = 'block';
        $('progressBar').style.width = '0%';
        $('progressStatus').textContent = '0/0 labels';
        
        // Start progress polling
        pollPrintStatus();
        if (printInterval) clearInterval(printInterval);
        printInterval = setInterval(pollPrintStatus, 1000);
        
        alert(j.message || j.error || 'Job queued');
    })
    .catch(()=>alert('Network error'));
}

async function loadProducts() {
    hideAlert('printResult'); // Clear previous print messages on load
    hideAlert('inventoryResult');
    $('progressContainer').style.display = 'none'; // Hide progress bar
    try {
        const response = await fetch('/get_products');
        const data = await response.json();
        if (data.error) {
             showAlert('printResult', `Failed to load products: ${data.error}`, 'danger');
             productData = []; // Clear old data on error
        } else {
            productData = data.products || [];
            renderTable();
        }
    } catch (error) {
        showAlert('printResult', `Failed to communicate with server to load products. Please check console for details.`, 'danger');
        console.error('Error loading products:', error);
        productData = []; // Clear old data on communication error
    } finally {
        renderTable(); // Always render, even if empty
    }
}

function renderTable(data = productData) {
    const tbody = document.querySelector("#productTable tbody");

    if (!data || data.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="7" class="text-center">
                    <div class="mt-3 mb-3 text-muted">
                        <i class="fas fa-info-circle"></i> No products found or loaded.
                        ${!productData || productData.length === 0 ? 'Check admin config and upload spreadsheet.' : ''}
                    </div>
                </td>
            </tr>
        `;
        return;
    }

    tbody.innerHTML = data.map((row) => `
    <tr>
        <td><strong>${row.Product || ''}</strong></td>
        <td>${row.Variant || ''}</td>
        <td><span class="text-muted">${row.SKU || ''}</span></td>
        <td>${(row.Price!=null)?'$'+parseFloat(row.Price).toFixed(2):'-'}</td>
        <td>${(row.Timeframe!=null)?row.Timeframe+' years':'-'}</td>
        <td><input type="number" class="form-control form-control-sm qty"
                   min="1" value="1" onfocus="this.select();"></td>
        <td class="d-flex gap-2 flex-wrap">
            <button class="button small" onclick="queuePrint(this,'back')">Back</button>
            ${
                row.FrontLabels && row.FrontLabels.toLowerCase() === "true"
                ? `
                    <button class="button small" onclick="queuePrint(this,'front')">Front</button>
                    <button class="button small" onclick="queuePrint(this,'both')">Both</button>
                  `
                : ""
            }
        </td>
    </tr>
`).join('');

     // Add event listeners for quantity inputs to select all text on focus
     document.querySelectorAll('.qty').forEach(input => {
         input.addEventListener('focus', function() {
             this.select();
         });
     });

}


// Print Functions
// Functions showPrintConfirmation and confirmPrint removed as they're no longer needed
// with the new direct print button implementation

async function pollPrintStatus() {
    try {
        const resp = await fetch('/print_status');
        const data = await resp.json();

        // Update progress bar
        const current = data.current || 0;
        const total = data.total || 0;
        const percentage = total > 0 ? (current / total) * 100 : 0;
        $('progressBar').style.width = percentage + '%';
        $('progressStatus').textContent = `${current}/${total} labels`;

        // Job completed or canceled
        if (!data.inProgress) {
            clearInterval(printInterval);
            printInterval = null; // Clear interval variable
            $('progressContainer').style.display = 'none'; // Hide progress bar

            // Determine final message
            let message = '';
            let type = 'success';

            if (data.current < data.total) {
                message = `Print job canceled at ${data.current}/${data.total} labels`;
                type = 'warning';
            } else if (data.total > 0) {
                 message = `Successfully printed ${data.total} labels`;
                 type = 'success';
            } else {
                 // Handle case where job finished but total was 0 (e.g. error before loop)
                 message = 'Print job finished with no labels printed.';
                 type = 'info';
            }

            showAlert('printResult', message, type);

            // Check inventory update status if it was requested for this job
            if (sessionStorage.getItem('updateInventoryRequested') === 'true') {
                 // Use a short delay to ensure the server side update might have started
                 setTimeout(checkInventoryStatus, 500); // Poll after 0.5 seconds
                 sessionStorage.removeItem('updateInventoryRequested'); // Clear flag
            }

             // Restore the button state
             if (confirmationData.button) {
               setLoading(confirmationData.button, false);
               confirmationData.button = null; // Clear button reference
            }
        }
    } catch (err) {
        console.error('Error polling print status:', err);
        // If polling fails continuously, stop the interval
        // Consider adding a retry counter or timeout
        // clearInterval(printInterval);
        // printInterval = null;
        // $('progressContainer').style.display = 'none';
        // showAlert('printResult', 'Lost connection to server or print status check failed.', 'danger');
         // Don't necessarily stop loading immediately, the job might still be running on the server
    }
}


async function checkInventoryStatus() {
    try {
        const resp = await fetch('/inventory_status');
        const data = await resp.json();

        if (data.message) {
            showAlert('inventoryResult', data.message, 'success');
        } else if (data.error) {
            showAlert('inventoryResult', data.error, 'danger');
        } else {
            hideAlert('inventoryResult'); // Hide if no message/error
        }
    } catch (err) {
        console.error('Error checking inventory status:', err);
         showAlert('inventoryResult', 'Failed to get inventory update status from server.', 'danger');
    }
}


async function cancelPrint() {
     // Set button to loading state (Cancel button itself)
    setLoading($('cancelBtn'), true);
    try {
        const response = await fetch('/cancel_print', { method: 'POST' });
        const data = await response.json();
        console.log(data.message);
        // Polling will detect the cancellation and update UI
    } catch (error) {
        console.error('Cancel request failed', error);
        showAlert('printResult', 'Failed to send cancel request.', 'danger');
         setLoading($('cancelBtn'), false); // Stop loading on error
    }
}

// Admin Functions

async function adminLogin() {
    const button = $('adminLoginBtn');
    const password = $('adminPassword').value;
    setLoading(button, true);
    hideAlert('adminResult'); // Clear previous messages

    try {
        const response = await fetch('/admin_login', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ password: password }) // Ensure key matches backend
        });
        const data = await response.json();

        if (data.error) {
            showAlert('adminResult', data.error, 'danger');
            isAdminLoggedIn = false;
        } else {
            showAlert('adminResult', data.message, 'success');
            $('adminControls').classList.remove('hidden-section');
            $('adminPassword').value = ''; // Clear password field
            isAdminLoggedIn = true;
             // Collapse admin password input section after successful login
             const adminCardBody = $('adminPassword').closest('.card-body');
             const adminCardHeader = adminCardBody.previousElementSibling;
             if (adminCardBody && adminCardHeader) {
                 adminCardBody.classList.add('collapsed');
                 adminCardHeader.classList.add('collapsed');
                 const icon = adminCardHeader.querySelector('.toggle-icon');
                    if (icon) {
                       icon.classList.remove('fa-chevron-up');
                       icon.classList.add('fa-chevron-down');
                   }
             }

             // Load configuration settings into the form
             await loadConfig();
        }
    } catch (err) {
         showAlert('adminResult', 'Login failed due to network or server error.', 'danger');
         console.error('Admin login error:', err);
         isAdminLoggedIn = false;
    } finally {
        setLoading(button, false);
    }
}


async function uploadSpreadsheet() {
    if (!isAdminLoggedIn) {
        showAlert('uploadResult', 'Unauthorized: Admin login required.', 'danger');
        return;
    }

    const button = $('uploadSpreadsheetBtn');
    const fileInput = $('spreadsheetFile');
    setLoading(button, true);
    hideAlert('uploadResult'); // Clear previous messages

    if (!fileInput.files || fileInput.files.length === 0) {
         showAlert('uploadResult', 'Please select a file to upload.', 'warning');
         setLoading(button, false);
         return;
    }

    try {
        const formData = new FormData();
        formData.append('spreadsheet', fileInput.files[0]);

        const response = await fetch('/upload_spreadsheet', {
            method: 'POST',
            body: formData
        });
        const data = await response.json();

        if (!response.ok) { // Check for non-2xx status codes
             throw new Error(data.error || `Upload failed with status ${response.status}`);
        }

        showAlert('uploadResult', `${data.message}`, 'success');
        // Reload products after successful upload
        await loadProducts();

    } catch (error) {
        showAlert('uploadResult', `Upload failed: ${error.message}`, 'danger');
        console.error('Upload spreadsheet error:', error);
    } finally {
        setLoading(button, false);
        fileInput.value = ''; // Clear the file input
    }
}


function downloadSpreadsheet() {
    if (!isAdminLoggedIn) {
        alert('Admin login required to download.');
        return;
    }
    // Flask handles the file sending, browser will prompt download
    window.location.href = '/download_spreadsheet';
}

// Add or Edit product
async function addOrEditProduct() {
     if (!isAdminLoggedIn) {
        showAlert('addEditResult', 'Unauthorized: Admin login required.', 'danger');
        return;
    }

    const button = $('addEditProductBtn');
    setLoading(button, true);
    hideAlert('addEditResult'); // Clear previous messages

    const product = $('newProductName').value.trim();
    const variant = $('newVariant').value.trim();
    const sku = $('newSKU').value || ''; // Use value property
    const price_str = $('newPrice').value; // Use value property
    const timeframe_str = $('newTimeframe').value; // Use value property
    const barcodeFile = $('newBarcodeFile').files[0];
    
    // New front label parameters
    const hasFrontLabel = $('newHasFrontLabel').checked;
    const frontLabelFile = $('newFrontLabelFile').files[0];

    if (!product || !variant) {
        showAlert('addEditResult', 'Product and Variant are required.', 'warning');
        setLoading(button, false);
        return;
    }

    const formData = new FormData();
    formData.append('product', product);
    formData.append('variant', variant);
    formData.append('sku', sku);
    formData.append('price', price_str);
    formData.append('timeframe', timeframe_str);
    formData.append('hasFrontLabel', hasFrontLabel ? 'True' : 'False');
    
    if (barcodeFile) {
        formData.append('barcode', barcodeFile);
    }
    
    if (frontLabelFile) {
        formData.append('frontLabelFile', frontLabelFile);
    }

    try {
        const response = await fetch('/admin_add_product', {
            method: 'POST',
            body: formData
        });
        const data = await response.json();

        if (!response.ok) { // Check for non-2xx status codes
             throw new Error(data.error || `Action failed with status ${response.status}`);
        }

        showAlert('addEditResult', data.message, 'success');
        // Reload products to see the new or updated entry
        await loadProducts();

         // Clear input fields after successful add/edit
         $('newProductName').value = '';
         $('newVariant').value = '';
         $('newSKU').value = '';
         $('newPrice').value = '0.00';
         $('newTimeframe').value = '0';
         $('newBarcodeFile').value = '';
         $('newHasFrontLabel').checked = false;
         $('newFrontLabelFile').value = '';
         $('frontLabelFileSection').style.display = 'none';

    } catch (err) {
        showAlert('addEditResult', `Error adding/updating product: ${err.message}`, 'danger');
        console.error('Add/Edit product error:', err);
    } finally {
        setLoading(button, false);
    }
}

// Front label checkbox toggle
$('newHasFrontLabel').addEventListener('change', function() {
    $('frontLabelFileSection').style.display = this.checked ? 'block' : 'none';
});


// Configuration Functions
async function loadConfig() {
     if (!isAdminLoggedIn) {
        console.warn('Attempted to load config without admin login.');
        return;
    }
    try {
        const response = await fetch('/admin_config');
         if (!response.ok) {
            // If not 200 OK, it might be 403 Unauthorized or other server error
             const errorData = await response.json().catch(() => ({ error: 'Unknown error fetching config' }));
             throw new Error(errorData.error || `Failed to fetch config with status ${response.status}`);
         }
        const config = await response.json();

        // Populate config fields
        $('configSpreadsheetPath').value = config.spreadsheet_file || '';
        $('configLogoPath').value = config.logo_path || '';
        $('configBarcodeFolder').value = config.barcode_folder || '';
        $('configFrontLabelFolder').value = config.front_label_folder || '';
        $('configBackLabelFolder').value = config.back_label_folder || '';
        $('configFontRegularPath').value = config.font_path_regular || '';
        $('configFontPricePath').value = config.font_path_price || '';

         console.log("Configuration loaded into UI.");

    } catch (error) {
        showAlert('adminConfigResult', `Failed to load configuration: ${error.message}`, 'danger');
         console.error('Load config error:', error);
    }
}

async function saveConfig() {
    if (!isAdminLoggedIn) {
        showAlert('adminConfigResult', 'Unauthorized: Admin login required.', 'danger');
        return;
    }

    const button = $('saveConfigBtn');
    setLoading(button, true);
    hideAlert('adminConfigResult'); // Clear previous messages

    const configData = {
        spreadsheet_file: $('configSpreadsheetPath').value.trim(),
        logo_path: $('configLogoPath').value.trim(),
        barcode_folder: $('configBarcodeFolder').value.trim(),
        front_label_folder: $('configFrontLabelFolder').value.trim(),
        back_label_folder: $('configBackLabelFolder').value.trim(),
        font_path_regular: $('configFontRegularPath').value.trim(),
        font_path_price: $('configFontPricePath').value.trim(),
    };

     // Basic validation
     if (!configData.spreadsheet_file) {
         showAlert('adminConfigResult', 'Spreadsheet File Path is required.', 'warning');
         setLoading(button, false);
         return;
     }
     if (!configData.barcode_folder) {
          showAlert('adminConfigResult', 'Barcode Images Folder Path is required.', 'warning');
          setLoading(button, false);
          return;
     }


    try {
        const response = await fetch('/admin_config', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(configData)
        });
        const data = await response.json();

         if (!response.ok) { // Check for non-2xx status codes
             throw new Error(data.error || `Save config failed with status ${response.status}`);
         }

        showAlert('adminConfigResult', data.message, 'success');

        // After saving config, attempt to reload product data using the new paths
        await loadProducts();

    } catch (err) {
        showAlert('adminConfigResult', `Error saving configuration: ${err.message}`, 'danger');
         console.error('Save config error:', err);
    } finally {
        setLoading(button, false);
    }
}


// Initialize App
window.addEventListener('DOMContentLoaded', async () => {
    setupCardToggles(); // Setup toggles for all cards
    setupTabs(); // Setup tab switching
    // Ensure the product list card is initially expanded
    const productCardHeader = document.querySelector('.card-header h2 i.fa-tags').closest('.card-header');
    if (productCardHeader && productCardHeader.nextElementSibling) {
        productCardHeader.nextElementSibling.classList.remove('collapsed');
         productCardHeader.classList.remove('collapsed');
         const icon = productCardHeader.querySelector('.toggle-icon');
            if (icon) {
               icon.classList.remove('fa-chevron-down');
               icon.classList.add('fa-chevron-up');
           }
    }


    await checkAdminStatus(); // Check admin status first
    // loadProducts is called within checkAdminStatus if already logged in,
    // or can be called here if not logged in but still want to show current data
    // Let's always call it here, it handles the case where config isn't fully set yet.
    await loadProducts();

    // Keep polling print status if a job might have been running on server restart or page refresh
    // Check if any job might be active by polling once immediately
    pollPrintStatus();
});

</script>
</body>
</html>"""
    return render_template_string(html_content)

@app.route('/get_products')
def get_products():
    """Returns product data loaded from the configured spreadsheet."""
    global SPREADSHEET_DATA
    SPREADSHEET_DATA = load_spreadsheet() # Always attempt to reload on this endpoint call
                                         # to pick up config changes or file updates

    if SPREADSHEET_DATA.empty:
         # Return an error message if spreadsheet didn't load
         # load_spreadsheet already prints errors, but this gives feedback to the UI
         if not CONFIG.get('spreadsheet_file') or not os.path.exists(CONFIG.get('spreadsheet_file', '')):
              return jsonify({"products": [], "error": "Spreadsheet file not found or path not configured. Check Admin > Resource Locations."}), 500
         else:
             return jsonify({"products": [], "error": "Failed to load or parse spreadsheet. Check file content and column headers (Product, Variant, Price, Timeframe, BarcodePath, SKU)."}), 500

    safe_df = SPREADSHEET_DATA.where(pd.notnull(SPREADSHEET_DATA), None)
    return jsonify({
        "products": safe_df.to_dict(orient='records')
        if not SPREADSHEET_DATA.empty else []
    })


@app.route('/print_labels', methods=['POST'])
def handle_print():
    """Handles print requests."""
    global SPREADSHEET_DATA, cancel_print_flag
    global printing_in_progress, current_progress, total_to_print, CONFIG

    if SPREADSHEET_DATA.empty:
        return jsonify({"error": "No products loaded. Check spreadsheet configuration in Admin."}), 400

    data = request.json
    # print(f"DEBUG: /print_labels request data: {data}") # Keep this for debugging if needed

    product_request = data.get('product', '').strip().lower()
    variant_request = data.get('variant', '').strip().lower()
    total_qty = int(data.get('quantity', 0)) # Ensure quantity is integer, default 0
    update_inventory = data.get('updateInventory', True) # Default to True if not specified
    label_type = data.get("label_type", "back")  # new

    if total_qty <= 0:
         return jsonify({"error": "Quantity must be 1 or more."}), 400

    # Find the item case-insensitively using helper columns
    df_copy = SPREADSHEET_DATA.copy()
    df_copy['ProductLower'] = df_copy['Product'].str.strip().str.lower()
    df_copy['VariantLower'] = df_copy['Variant'].str.strip().str.lower()

    filtered = df_copy[
        (df_copy['ProductLower'] == product_request) &
        (df_copy['VariantLower'] == variant_request)
    ]

    if filtered.empty:
        return jsonify({"error": f"Product '{data.get('product')}' / Variant '{data.get('variant')}' not found in spreadsheet."}), 404

    # Get the original case version from SPREADSHEET_DATA using the index
    original_index = filtered.index[0]
    item = SPREADSHEET_DATA.loc[original_index]


    # Get SKU for inventory update
    sku = item.get('SKU', '')
    # Ensure sku is a string and not NaN
    if pd.isna(sku):
         sku = ''
    sku = str(sku).strip() # Ensure it's string

    # Store SKU and quantity in session for inventory update after printing
    session['print_sku'] = sku
    session['print_quantity'] = total_qty # Save quantity to update inventory with after printing
    session['update_inventory'] = update_inventory # Store user preference

    # Reset inventory update message state
    session['inventory_update_message'] = ''


    total_to_print = total_qty
    current_progress = 0
    printing_in_progress = True
    cancel_print_flag = False

    # IMPORTANT: This print loop blocks the Flask server.
    # For a production app, you'd offload this to a background worker (Celery, multiprocessing).
    # For this script, it will freeze the web interface during printing.
    try:
        barcode_filename = item.get('BarcodePath', '')
        if pd.isna(barcode_filename): # Handle NaN from spreadsheet
             barcode_filename = ''
            
        # Get label filenames and flags
        front_label_filename = (item.get('FrontLabelFiles') or '').strip() or None
        back_label_filename = (item.get('BackLabelFiles') or '').strip() or None
        both_labels_epson = (item.get('BothLabelsEpson') or '').strip().lower() == "true"
        
        dbg(f"Processing {item.Product}/{item.Variant} - label type: {label_type}")
        dbg(f"BothLabelsEpson: {both_labels_epson}")
        
        if front_label_filename:
            dbg(f"Front label file: {front_label_filename}")
        else:
            dbg(f"No front label file specified for this product")
            if label_type in ('front', 'both'):
                dbg(f"Warning: Front label requested but no file defined in spreadsheet")
                
        if back_label_filename:
            dbg(f"Back label file: {back_label_filename}")
        else:
            dbg(f"No back label file specified for this product")
            if label_type in ("back", "both") and both_labels_epson:
                dbg(f"Warning: Epson back label requested but no file defined in spreadsheet")

        for i in range(total_qty):
            if cancel_print_flag:
                break

            # --- Epson for both labels case ---
            if label_type in ("back", "both") and both_labels_epson:
                if back_label_filename:
                    back_label_path = create_back_label(back_label_filename)
                    if back_label_path:
                        send_to_epson(back_label_path, copies=1)
                        os.remove(back_label_path)
                    else:
                        print(f"ERROR: Failed to create back label image for Epson")
                else:
                    print(f"ERROR: BothLabelsEpson is True but no BackLabelFiles set.")

            # --- Normal "back" label (Brother) ---
            elif label_type in ("back", "both"):
                best_by, julian = calculate_dates(item.Timeframe)
                label_path = create_label_image(best_by, item.Price, julian, barcode_filename)
                if label_path:
                    send_to_printer(label_path)      # Brother QL‑800
                    os.remove(label_path)

            # --- Front label as usual ---
            if label_type in ("front", "both"):
                if front_label_filename:
                    fp = create_front_label(front_label_filename)
                    if fp:
                        send_to_epson(fp, copies=1)  # always raw
                        os.remove(fp)
                    else:
                        print(f"ERROR: Failed to create front label image - path not returned or file doesn't exist")
            
            current_progress += 1
                
            # If we reach this point, check if we were canceled
            if cancel_print_flag:
                dbg(f"Print job canceled by user after {current_progress} labels.")
                break


        # After printing loop finishes (either completed or canceled)
        printing_in_progress = False


        # Update inventory only if explicitly requested and if any labels were successfully printed
        if current_progress > 0 and sku and update_inventory:
            # Pass the *actual* number printed, not the requested total_qty
            success, message = update_shopify_inventory_after_printing(sku, current_progress)
            session['inventory_update_message'] = message
        elif not update_inventory:
             session['inventory_update_message'] = f"Printed {current_progress} labels (inventory update skipped by user request)."
        elif sku == '':
             session['inventory_update_message'] = f"Printed {current_progress} labels (inventory not updated: SKU missing)."
        elif current_progress == 0:
             session['inventory_update_message'] = f"No labels were successfully printed (inventory not updated)."
        # If current_progress > 0 but sku is None/empty, message is handled by the sku == '' check above


        if cancel_print_flag:
             session['inventory_update_message'] = f"Print job canceled at {current_progress}/{total_qty} labels. " + session.get('inventory_update_message', '')
             return jsonify({"message": f"Print job canceled at {current_progress}/{total_qty} labels. Inventory update status: {session['inventory_update_message']}"})
        else:
             return jsonify({"message": f"Printed {total_qty} labels. Inventory update status: {session.get('inventory_update_message', 'Not requested or applicable.')}"})

    except Exception as e:
        print(f"An error occurred during the print process: {e}")
        printing_in_progress = False
        # Attempt to clean up any temporary file if it exists
        if 'label_path' in locals() and label_path and os.path.exists(label_path):
            try:
                os.remove(label_path)
            except OSError as cleanup_error:
                print(f"Warning: Could not delete temporary label file after error: {cleanup_error}")

        # Update inventory message to reflect error
        session['inventory_update_message'] = f"Print job failed after {current_progress}/{total_qty} labels due to error: {str(e)}. Inventory not updated."
        return jsonify({"error": f"Print process failed: {str(e)}."}), 500


@app.route('/inventory_status', methods=['GET'])
def inventory_status():
    """Return the status of the most recent inventory update."""
    # The message is stored in the session by handle_print
    message = session.get('inventory_update_message', '')
    if not message:
         # If no message is set, it means no print job requiring inventory update finished yet
         return jsonify({"message": "Inventory update status not available."})

    # Check if the message indicates failure to categorize it for the UI
    # This is a simple check, might need refinement
    if 'Failed' in message or 'Could not' in message or 'Cannot' in message or 'not found' in message or 'missing' in message or 'error' in message.lower():
        # Clear message after reporting failure? Depends on desired UX. Let's keep it.
        return jsonify({"error": message})
    else:
        # Clear message after reporting success/completion?
        # session.pop('inventory_update_message', None) # Uncomment if you want it to show only once
        return jsonify({"message": message})


@app.route('/cancel_print', methods=['POST'])
def cancel_print():
    """Sets the flag to cancel the current print job."""
    global cancel_print_flag
    cancel_print_flag = True
    return jsonify({"message": "Print job cancellation requested."})


@app.route('/print_status', methods=['GET'])
def print_status():
    """Returns the current state of the print job."""
    global printing_in_progress, current_progress, total_to_print
    return jsonify({
        "inProgress": printing_in_progress,
        "current": current_progress,
        "total": total_to_print
    })


@app.route('/admin_login', methods=['POST'])
def handle_admin_login():
    """Handles admin login."""
    password = request.json.get('password')
    if password and password == ADMIN_PASSWORD:
        session['admin'] = True
        # Also store config in session? No, config should be server-side global.
        return jsonify({"message": "Admin access granted."})
    # Clear admin session on failed attempt? Optional security measure.
    # session.pop('admin', None)
    return jsonify({"error": "Invalid password."}), 401

@app.route('/check_admin')
def check_admin_session():
    """Checks if the user has an active admin session."""
    return jsonify({"loggedIn": session.get('admin', False)})

@app.route('/upload_spreadsheet', methods=['POST'])
def handle_upload():
    """Handles uploading a new spreadsheet file."""
    global CONFIG # Use the global config
    if not session.get('admin'):
        return jsonify({"error": "Unauthorized"}), 403

    if 'spreadsheet' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['spreadsheet']
    filename = file.filename.lower()

    if not filename.endswith(('.csv', '.xlsx')):
        return jsonify({"error": "Invalid file type. Please upload .csv or .xlsx"}), 400

    spreadsheet_path = CONFIG.get('spreadsheet_file')
    if not spreadsheet_path:
        return jsonify({"error": "Spreadsheet file path not configured in Admin settings."}), 500

    # Ensure the directory for the spreadsheet path exists
    spreadsheet_dir = os.path.dirname(spreadsheet_path)
    if spreadsheet_dir and not os.path.exists(spreadsheet_dir):
        try:
            os.makedirs(spreadsheet_dir)
            print(f"Created spreadsheet directory: {spreadsheet_dir}")
        except OSError as e:
            return jsonify({"error": f"Failed to create directory for spreadsheet: {e}"}), 500

    # Save the uploaded file temporarily to load it
    temp_path = os.path.join(TEMP_DIR, file.filename) # Use our cross-platform temporary directory
    try:
        file.save(temp_path)
    except Exception as e:
         return jsonify({"error": f"Failed to save uploaded file temporarily: {e}"}), 500

    try:
        if filename.endswith('.xlsx'):
            df = pd.read_excel(temp_path)
        else:
            df = pd.read_csv(temp_path)

        # --- Validation matching load_spreadsheet ---
        required = ['Product', 'Variant', 'Price', 'Timeframe', 'BarcodePath', 'SKU']
        if not all(col in df.columns for col in required):
            missing = [col for col in required if col not in df.columns]
            return jsonify({"error": f"Missing required columns: {missing}"}), 400
         

        # Clean string columns
        string_cols = ['Product', 'Variant', 'BarcodePath', 'SKU']
        for col in string_cols:
             if col in df.columns:
                df[col] = df[col].astype(str).str.strip().replace(r'\s+', ' ', regex=True)
             else:
                df[col] = '' # Add missing column with default empty strings

        # Convert numeric columns with error handling
        df['Price'] = pd.to_numeric(df['Price'], errors='coerce')
        df['Timeframe'] = pd.to_numeric(df['Timeframe'], errors='coerce').fillna(0).astype(int)

        # Drop rows where essential data is missing (Product, Variant, Price, Barcode filename)
        # Note: BarcodePath here is expected to be just a filename. We only check if the cell is non-empty.
        # The existence of the actual file will be checked during printing.
        initial_count = len(df)
        df = df.dropna(subset=['Product', 'Variant', 'Price']) # Price is essential for label
        # Drop if BarcodePath column exists but is empty string after stripping
        if 'BarcodePath' in df.columns:
             df = df[df['BarcodePath'] != ''].reset_index(drop=True)
        else: # If BarcodePath wasn't even a column, something is wrong based on required check, but handle defensively
             return jsonify({"error": "Spreadsheet missing BarcodePath column."}), 400

        # Make sure front label columns exist
        if 'FrontLabels' not in df.columns:
            print("Adding missing FrontLabels column")
            df['FrontLabels'] = 'False'  # Default all to False
        else:
            # Normalize existing values
            df['FrontLabels'] = df['FrontLabels'].astype(str)
            df['FrontLabels'] = df['FrontLabels'].apply(
                lambda x: 'True' if x.strip().lower() == 'true' else 'False'
            )
            
        if 'FrontLabelFiles' not in df.columns:
            print("Adding missing FrontLabelFiles column")
            df['FrontLabelFiles'] = ''  # Default to empty

        # Process FrontLabels and FrontLabelFiles columns (ensure True/False strings)
        if 'FrontLabels' in df.columns:
            df['FrontLabels'] = df['FrontLabels'].astype(str).apply(
                lambda x: 'True' if x.lower() == 'true' else 'False'
            )
        else:
            df['FrontLabels'] = 'False'

        if 'FrontLabelFiles' in df.columns:
            df['FrontLabelFiles'] = df['FrontLabelFiles'].astype(str).str.strip()
        else:
            df['FrontLabelFiles'] = ''

        # Handle BackLabelFiles
        if 'BackLabelFiles' in df.columns:
            df['BackLabelFiles'] = df['BackLabelFiles'].astype(str).str.strip()
        else:
            df['BackLabelFiles'] = ''

        # Handle BothLabelsEpson (ensure True/False string)
        if 'BothLabelsEpson' in df.columns:
            df['BothLabelsEpson'] = df['BothLabelsEpson'].astype(str).apply(
                lambda x: 'True' if x.strip().lower() == 'true' else 'False'
            )
        else:
            df['BothLabelsEpson'] = 'False'  # Default to empty

        cleaned_count = len(df)

        if df.empty:
            return jsonify({"error": "No valid product data found after checking required columns and non-empty barcode paths."}), 400

        # --- End Validation ---

        # Save the validated DataFrame to the configured spreadsheet path (as .xlsx)
        # Always save as xlsx for consistency with download_spreadsheet
        df.to_excel(spreadsheet_path, index=False)

        # Update the global SPREADSHEET_DATA with the newly loaded data
        global SPREADSHEET_DATA
        # Reload using the standard load_spreadsheet function to ensure consistency (optional but safer)
        SPREADSHEET_DATA = load_spreadsheet()
        # Or simply use the validated df:
        # SPREADSHEET_DATA = df.copy() # Use .copy() to avoid potential SettingWithCopyWarning later


        return jsonify({
            "message": f"Spreadsheet uploaded and validated successfully. Loaded {len(SPREADSHEET_DATA)} valid products.",
            "count": len(SPREADSHEET_DATA)
        })

    except Exception as e:
        return jsonify({"error": f"Processing uploaded spreadsheet failed: {str(e)}"}), 500
    finally:
        # Clean up the temporary file
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError as cleanup_error:
                print(f"Warning: Could not delete temporary uploaded file {temp_path}: {cleanup_error}")


@app.route('/download_spreadsheet')
def handle_download():
    """Allows downloading the current spreadsheet file."""
    global CONFIG # Use the global config
    if not session.get('admin'):
        return jsonify({"error": "Unauthorized"}), 403

    spreadsheet_path = CONFIG.get('spreadsheet_file')
    if not spreadsheet_path or not os.path.exists(spreadsheet_path):
         # Attempt to create a minimal file if it doesn't exist
         if not spreadsheet_path:
             return jsonify({"error": "Spreadsheet file path not configured in Admin settings."}), 500
         else:
              # Create a dummy file if missing to avoid crashing send_file
             print(f"Warning: Spreadsheet file not found at {spreadsheet_path}. Creating an empty one.")
             try:
                 pd.DataFrame(columns=['Product','Variant','Price','Timeframe','BarcodePath','SKU']).to_excel(spreadsheet_path, index=False)
             except Exception as e:
                  print(f"Error creating dummy spreadsheet file: {e}")
                  return jsonify({"error": f"Spreadsheet file not found at {spreadsheet_path} and failed to create a dummy file."}), 500

    try:
        # Ensure the file exists before trying to send it
        if not os.path.exists(spreadsheet_path):
             return jsonify({"error": f"Spreadsheet file not found at {spreadsheet_path}"}), 500
             
        # Create a temporary copy for download without modifying the original file
        # This ensures we don't change the actual data values in the spreadsheet
        if not SPREADSHEET_DATA.empty and 'FrontLabels' in SPREADSHEET_DATA.columns:
            # We'll preserve the original data by using a local copy for download
            # No need to write back to the file
            print("Preserving original FrontLabels values during download")

        return send_file(spreadsheet_path, as_attachment=True, download_name='current_data.xlsx') # Suggest a download name
    except FileNotFoundError:
        return jsonify({"error": f"Spreadsheet file not found at {spreadsheet_path}"}), 404
    except Exception as e:
         return jsonify({"error": f"Error sending spreadsheet file: {str(e)}"}), 500


@app.route('/admin_add_product', methods=['POST'])
def admin_add_product():
    """Adds a new product or updates an existing one."""
    global SPREADSHEET_DATA, CONFIG # Use global data and config
    if not session.get('admin'):
        return jsonify({"error": "Unauthorized"}), 403

    product = request.form.get('product', '').strip()
    variant = request.form.get('variant', '').strip()
    sku = request.form.get('sku', '').strip()
    price_str = request.form.get('price', '0').strip()
    timeframe_str = request.form.get('timeframe', '0').strip()
    barcode_file = request.files.get('barcode')
    
    # New front label parameters
    has_front_label = request.form.get('hasFrontLabel', 'False').strip()
    front_label_file = request.files.get('frontLabelFile')

    if not product or not variant:
        return jsonify({"error": "Product and Variant are required"}), 400

    # Validate and convert price
    try:
        price = float(price_str)
    except ValueError:
        price = 0.0 # Default to 0.0 if conversion fails

    # Validate and convert timeframe
    try:
        timeframe = int(timeframe_str)
    except ValueError:
        timeframe = 0 # Default to 0 if conversion fails

    # Ensure SPREADSHEET_DATA is loaded
    if SPREADSHEET_DATA.empty:
        # Attempt to load if not already loaded. This might fail if config is bad.
        SPREADSHEET_DATA = load_spreadsheet()
        # If still empty, initialize a new DataFrame with required columns
        if SPREADSHEET_DATA.empty:
             print("Warning: SPREADSHEET_DATA was empty and load_spreadsheet failed. Initializing empty DataFrame.")
             SPREADSHEET_DATA = pd.DataFrame(columns=['Product','Variant','Price','Timeframe',
                                                     'BarcodePath','SKU','FrontLabels','FrontLabelFiles'])

    # Find existing row based on Product and Variant (case-insensitive for lookup)
    df_copy_lookup = SPREADSHEET_DATA.copy()
    df_copy_lookup['ProductLower'] = df_copy_lookup['Product'].str.lower()
    df_copy_lookup['VariantLower'] = df_copy_lookup['Variant'].str.lower()

    mask = (df_copy_lookup['ProductLower'] == product.lower()) & \
           (df_copy_lookup['VariantLower'] == variant.lower())

    existing_index = SPREADSHEET_DATA[mask].index[0] if mask.any() else None

    barcode_filename = '' # Will store just the filename
    front_label_filename = '' # Will store just the filename

    # Process barcode file if provided
    if barcode_file:
        barcode_folder = CONFIG.get('barcode_folder')
        if not barcode_folder:
             return jsonify({"error": "Barcode folder path not configured in Admin settings."}), 500

        # Ensure barcode folder exists
        if not os.path.exists(barcode_folder):
             try:
                 os.makedirs(barcode_folder)
                 print(f"Created barcode folder: {barcode_folder}")
             except OSError as e:
                 return jsonify({"error": f"Failed to create barcode folder {barcode_folder}: {e}"}), 500

        filename = barcode_file.filename
        save_path = os.path.join(barcode_folder, filename)

        try:
            barcode_file.save(save_path)
            barcode_filename = filename # Store only the filename in the spreadsheet
            print(f"Saved barcode file: {save_path}")
        except Exception as e:
             return jsonify({"error": f"Failed to save barcode file {filename}: {e}"}), 500

    # Process front label file if provided
    if front_label_file:
        front_label_folder = CONFIG.get('front_label_folder')
        if not front_label_folder:
             return jsonify({"error": "Front label folder path not configured in Admin settings."}), 500

        # Ensure front label folder exists
        if not os.path.exists(front_label_folder):
             try:
                 os.makedirs(front_label_folder)
                 print(f"Created front label folder: {front_label_folder}")
             except OSError as e:
                 return jsonify({"error": f"Failed to create front label folder {front_label_folder}: {e}"}), 500

        filename = front_label_file.filename
        save_path = os.path.join(front_label_folder, filename)

        try:
            front_label_file.save(save_path)
            front_label_filename = filename # Store only the filename in the spreadsheet
            print(f"Saved front label file: {save_path}")
        except Exception as e:
             return jsonify({"error": f"Failed to save front label file {filename}: {e}"}), 500

    # Ensure all required columns exist in the spreadsheet data
    required_columns = ['Product', 'Variant', 'Price', 'Timeframe', 'BarcodePath', 'SKU', 
                        'FrontLabels', 'FrontLabelFiles']
    for col in required_columns:
        if col not in SPREADSHEET_DATA.columns:
            print(f"Adding missing column '{col}' to spreadsheet data")
            SPREADSHEET_DATA[col] = '' # Add missing column with empty default

    # Prepare data for the row
    new_row_data = {
        'Product': product, 
        'Variant': variant, 
        'SKU': sku,
        'Price': price,
        'Timeframe': timeframe,
        'FrontLabels': has_front_label,
        'BarcodePath': barcode_filename if barcode_filename else (SPREADSHEET_DATA.loc[existing_index, 'BarcodePath'] if existing_index is not None and 'BarcodePath' in SPREADSHEET_DATA.columns else '')
    }
    
    # Handle front label filename
    if front_label_filename:
        # New file was uploaded
        new_row_data['FrontLabelFiles'] = front_label_filename
    elif existing_index is not None and 'FrontLabelFiles' in SPREADSHEET_DATA.columns:
        # Keep existing filename if one exists
        new_row_data['FrontLabelFiles'] = SPREADSHEET_DATA.loc[existing_index, 'FrontLabelFiles']
    else:
        # Default to empty
        new_row_data['FrontLabelFiles'] = ''

    if existing_index is not None:
        # Update existing row
        for key, value in new_row_data.items():
            if key in SPREADSHEET_DATA.columns:
                SPREADSHEET_DATA.at[existing_index, key] = value

        message = "Product/Variant updated successfully."
    else:
        # Add new row using pd.concat
        new_row_df = pd.DataFrame([new_row_data])
        SPREADSHEET_DATA = pd.concat([SPREADSHEET_DATA, new_row_df], ignore_index=True)

        message = "Product/Variant added successfully."

    # Save the updated DataFrame back to the spreadsheet file
    spreadsheet_path = CONFIG.get('spreadsheet_file')
    if not spreadsheet_path:
        return jsonify({"error": "Spreadsheet file path not configured. Cannot save data."}), 500

    try:
         # Ensure the directory exists before saving
         spreadsheet_dir = os.path.dirname(spreadsheet_path)
         if spreadsheet_dir and not os.path.exists(spreadsheet_dir):
             os.makedirs(spreadsheet_dir)

         SPREADSHEET_DATA.to_excel(spreadsheet_path, index=False)
         print(f"Saved updated spreadsheet to {spreadsheet_path}")
    except Exception as e:
         return jsonify({"error": f"Failed to save spreadsheet data: {e}"}), 500

    return jsonify({"message": message})


@app.route('/admin_config', methods=['GET'])
def get_admin_config():
    """Returns the current configuration settings."""
    if not session.get('admin'):
        return jsonify({"error": "Unauthorized"}), 403
    # Return a copy to prevent accidental modification of the global CONFIG
    return jsonify(CONFIG.copy())


@app.route('/admin_config', methods=['POST'])
def save_admin_config():
    """Saves the updated configuration settings."""
    global CONFIG # Modify the global config
    if not session.get('admin'):
        return jsonify({"error": "Unauthorized"}), 403

    data = request.json

    # Update CONFIG dictionary with received data
    # Only update keys we expect
    updated = {}
    for key in ['spreadsheet_file', 'logo_path', 'barcode_folder', 'front_label_folder', 'font_path_regular', 'font_path_price']:
        if key in data:
            # Normalize paths (optional but good practice)
            path = data[key].strip()
            if path:
                 # Convert relative paths to absolute if needed? os.path.abspath(path)
                 # Or leave as is and rely on user input
                 # Let's assume user inputs absolute paths or paths relative to where the script is run
                 CONFIG[key] = path
                 updated[key] = path # Track what was updated
            else:
                 CONFIG[key] = '' # Allow emptying paths

    # Basic validation: Check if barcode folder exists/can be created
    barcode_folder = CONFIG.get('barcode_folder')
    if barcode_folder:
         if not os.path.isdir(barcode_folder):
             print(f"Configured barcode folder does not exist: {barcode_folder}")
             try:
                 os.makedirs(barcode_folder)
                 print(f"Attempted to create barcode folder: {barcode_folder}")
             except OSError as e:
                 print(f"Failed to create barcode folder: {e}")
                 # Do we fail the save, or save config but warn the user?
                 # Let's save but return a warning message
                 save_config(CONFIG) # Save even if folder creation failed
                 return jsonify({"message": f"Configuration saved, but warning: Failed to create barcode folder '{barcode_folder}': {e}"}), 200 # Use 200 with warning

         else:
             print(f"Configured barcode folder exists: {barcode_folder}")
    elif 'barcode_folder' in data and data['barcode_folder'].strip() == '':
         print("Barcode folder path was cleared in config.")
         # This is valid, but print won't work if barcodes are needed.
         
    # Basic validation: Check if front label folder exists/can be created
    front_label_folder = CONFIG.get('front_label_folder')
    if front_label_folder:
         if not os.path.isdir(front_label_folder):
             print(f"Configured front label folder does not exist: {front_label_folder}")
             try:
                 os.makedirs(front_label_folder)
                 print(f"Attempted to create front label folder: {front_label_folder}")
             except OSError as e:
                 print(f"Failed to create front label folder: {e}")
                 # Save but return a warning message
                 save_config(CONFIG) # Save even if folder creation failed
                 return jsonify({"message": f"Configuration saved, but warning: Failed to create front label folder '{front_label_folder}': {e}"}), 200 # Use 200 with warning
         else:
             print(f"Configured front label folder exists: {front_label_folder}")
    elif 'front_label_folder' in data and data['front_label_folder'].strip() == '':
         print("Front label folder path was cleared in config.")


    # Save the updated configuration to file
    try:
        save_config(CONFIG)
    except Exception as e:
         # If saving fails, return error
         return jsonify({"error": f"Failed to save configuration to file: {e}"}), 500

    # Optionally reload spreadsheet data immediately after saving config
    # This ensures the product list reflects the new spreadsheet path if it was changed
    global SPREADSHEET_DATA
    SPREADSHEET_DATA = load_spreadsheet()


    return jsonify({"message": "Configuration saved successfully."})


def check_front_label_configuration():
    """Checks the configuration and spreadsheet for front label issues."""
    global CONFIG, SPREADSHEET_DATA
    
    print("\n--- Front Label Configuration Check ---")
    
    # Check if front_label_folder is configured
    front_label_folder = CONFIG.get('front_label_folder', '')
    if not front_label_folder:
        print("WARNING: front_label_folder is not configured in config.json")
        print("Front label printing will not work until this is set.")
        return
    
    # Check if the folder exists
    if not os.path.exists(front_label_folder):
        print(f"WARNING: Configured front_label_folder does not exist: '{front_label_folder}'")
        try:
            os.makedirs(front_label_folder)
            print(f"Created front label folder: '{front_label_folder}'")
        except OSError as e:
            print(f"ERROR: Failed to create front label folder: {e}")
            print("Front label printing will not work until this folder is created.")
            return
    
    if SPREADSHEET_DATA.empty:
        print("WARNING: Spreadsheet data not loaded. Cannot check front label columns.")
        return
    
    # Check for FrontLabels column
    if 'FrontLabels' not in SPREADSHEET_DATA.columns:
        print("WARNING: 'FrontLabels' column not found in spreadsheet data.")
        print("Adding 'FrontLabels' column with default value 'False'...")
        SPREADSHEET_DATA['FrontLabels'] = 'False'
    
    # Check for FrontLabelFiles column
    if 'FrontLabelFiles' not in SPREADSHEET_DATA.columns:
        print("WARNING: 'FrontLabelFiles' column not found in spreadsheet data.")
        print("Adding 'FrontLabelFiles' column with empty values...")
        SPREADSHEET_DATA['FrontLabelFiles'] = ''
    
    # Check for products with FrontLabels=True but no filename
    if 'FrontLabels' in SPREADSHEET_DATA.columns and 'FrontLabelFiles' in SPREADSHEET_DATA.columns:
        front_label_enabled = SPREADSHEET_DATA[SPREADSHEET_DATA['FrontLabels'].str.lower() == 'true']
        missing_filenames = front_label_enabled[
            (front_label_enabled['FrontLabelFiles'].isnull()) | 
            (front_label_enabled['FrontLabelFiles'] == '')
        ]
        
        if not missing_filenames.empty:
            print(f"WARNING: Found {len(missing_filenames)} products with FrontLabels=True but missing filename.")
            print("These products will show the front label option but printing will fail:")
            for idx, row in missing_filenames.iterrows():
                print(f"- {row['Product']} / {row['Variant']}")
    
    # Check if any front label files are missing
    if 'FrontLabels' in SPREADSHEET_DATA.columns and 'FrontLabelFiles' in SPREADSHEET_DATA.columns:
        front_label_enabled = SPREADSHEET_DATA[SPREADSHEET_DATA['FrontLabels'].str.lower() == 'true']
        front_label_enabled = front_label_enabled[front_label_enabled['FrontLabelFiles'] != '']
        
        missing_files = 0
        for idx, row in front_label_enabled.iterrows():
            filename = row['FrontLabelFiles']
            if not os.path.exists(os.path.join(front_label_folder, filename)):
                missing_files += 1
                print(f"WARNING: Front label file not found: '{filename}' for {row['Product']} / {row['Variant']}")
        
        if missing_files > 0:
            print(f"WARNING: {missing_files} front label files are missing from the front label folder.")
            print(f"Upload these files to: '{front_label_folder}'")
        else:
            print(f"All front label files found in '{front_label_folder}'.")
    
    # Save any changes made to spreadsheet
    if ('FrontLabels' not in SPREADSHEET_DATA.columns or 'FrontLabelFiles' not in SPREADSHEET_DATA.columns):
        try:
            spreadsheet_path = CONFIG.get('spreadsheet_file')
            if spreadsheet_path:
                SPREADSHEET_DATA.to_excel(spreadsheet_path, index=False)
                print(f"Saved updated spreadsheet with front label columns to '{spreadsheet_path}'")
        except Exception as e:
            print(f"ERROR: Failed to save updated spreadsheet: {e}")
    
    print("--- Front Label Configuration Check Complete ---\n")


@app.route("/debug_log")
def debug_log():
    # (optional) protect with admin check:
    # if not session.get("admin"): return jsonify({"error":"unauthorised"}),403
    return jsonify({"lines": LAST_JOB_LOG[-100:]})

# ==========================================
# STARTUP

if __name__ == "__main__":
    # 1. Load configuration from file
    load_config()

    # 2. Load initial spreadsheet data using the configured path
    SPREADSHEET_DATA = load_spreadsheet()
    
    # 3. Run front label configuration check
    check_front_label_configuration()

    # 4. Run the Flask app
    # debug=False in production! Set debug=True only for development.
    # Use a production-ready server like Gunicorn or uWSGI for deployment.
    # Example (with Gunicorn): gunicorn -w 4 -b 0.0.0.0:5050 your_script_name:app
    app.run(host="0.0.0.0", port=5050, threaded=True, debug=False)  