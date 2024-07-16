import pytesseract
from PIL import Image
import os
from openpyxl import Workbook
from tqdm import tqdm

# Define the path to the tesseract executable
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  

# Define the directory containing screenshots
screenshot_dir = 'C:\\Users\\dnai7\\PycharmProjects\\pythonProject1\\Taulukko'  

# Define the output Excel file
output_excel = 'Rise_of_Castles_Units.xlsx'


# Function to extract data from a single screenshot
def extract_data_from_screenshot(image_path):
    image = Image.open(image_path)
    text = pytesseract.image_to_string(image)

    lines = text.split('\n')
    data = {}
    current_header = None

    headers = [
        'Name', 'Season points', 'Declaration points', 'Stone donations', 'Total war joined',
        'Total war demolition value', 'Demolition value in Eden', 'Merit',
        'Guild contribution exp', 'Wonder cp', 'Occupy enemy territory'
    ]

    for line in lines:
        if line.strip():  # Skip empty lines
            parts = line.split(':')
            if len(parts) == 2:
                key, value = parts
                key = key.strip().lower()
                value = value.strip().replace(',', '').replace('.', '')  # Remove commas and periods for number parsing
                if 'name' in key:
                    current_header = 'Name'
                elif 'season points' in key:
                    current_header = 'Season points'
                elif 'declaration point' in key:
                    current_header = 'Declaration points'
                elif 'stone donation' in key:
                    current_header = 'Stone donations'
                elif 'total war joined' in key:
                    current_header = 'Total war joined'
                elif 'total war demolition value' in key:
                    current_header = 'Total war demolition value'
                elif 'demolition value in eden' in key:
                    current_header = 'Demolition value in Eden'
                elif 'merit' in key:
                    current_header = 'Merit'
                elif 'guild contribution exp' in key:
                    current_header = 'Guild contribution exp'
                elif 'wonder cp' in key:
                    current_header = 'Wonder cp'
                elif 'occupy enemy territory' in key:
                    current_header = 'Occupy enemy territory'

                if current_header:
                    data[current_header] = value

    return [data.get(header, '') for header in headers]


# Create a new Excel workbook and sheet
wb = Workbook()
ws = wb.active
ws.title = "Units Data"

# Add headers
headers = [
    'Name', 'Season points', 'Declaration points', 'Stone donations', 'Total war joined',
    'Total war demolition value', 'Demolition value in Eden', 'Merit',
    'Guild contribution exp', 'Wonder cp', 'Occupy enemy territory'
]
ws.append(headers)

# Get list of screenshot files
files = [f for f in os.listdir(screenshot_dir) if f.endswith('.jpg') or f.endswith('.png')]

# Process each screenshot in the directory with a progress bar
for filename in tqdm(files, desc="Processing screenshots"):
    image_path = os.path.join(screenshot_dir, filename)
    data = extract_data_from_screenshot(image_path)
    ws.append(data)

# Save the workbook
try:
    wb.save(output_excel)
    print(f'Data extracted and saved to {output_excel}')
except Exception as e:
    print(f'Error saving the workbook: {e}')
