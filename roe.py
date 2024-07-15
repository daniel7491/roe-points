import os
from PIL import Image
import pytesseract
import pandas as pd
import re

# Configure pytesseract to use the correct language data
pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'

# The lang parameter is set to handle a variety of languages.
ocr_lang = 'eng+spa+deu+fra+ita+por+rus+chi_sim+jpn'

# Define the directory path that contains the image files
directory_path = 'C:\\Users\\dnai7\\PycharmProjects\\pythonProject1\\pics'

# Define the function to extract values from a single image
def extract_values_from_image(image_path, lang=ocr_lang):
    # Open the image file
    img = Image.open(image_path)

    # Use tesseract to do OCR on the image with the specified language
    text = pytesseract.image_to_string(img, lang=lang)

    # Print the OCR result for debugging
    print(f"OCR result for {os.path.basename(image_path)}:\n{text}")

    # Extract the name and values
    name_pattern = re.compile(r'\[Cnt\]\s(.*?)[\r\n]+', re.DOTALL)
    name_match = name_pattern.search(text)
    name = name_match.group(1).strip() if name_match else "Name not found"

    # Prepare dictionary for extracted values
    data = {
        'name': name,
        'Season points': 0,
        'Demolition Value in Eden': 0,
        'Merit': 0,
        'Wonder CP': 0,
        'Occupy Enemy Territory': 0
    }

    # Match labels to values using proximity based matching
    labels = ['Season points', 'Demolition Value in Eden', 'Merit', 'Wonder CP:', 'Occupy Enemy Territory']
    for label in labels:
        pattern = re.compile(rf"{re.escape(label)}\s*(?:\:)?\s*(\d[\d,]*)")
        match = pattern.search(text)
        if match:
            # Replace commas and cast to int
            data[label] = int(match.group(1).replace(',', ''))

    return data

# List to hold data for each image
data_list = []

# Loop over all the image files in the directory
for filename in os.listdir(directory_path):
    if filename.endswith('.png'):
        # Construct the full file path
        file_path = os.path.join(directory_path, filename)
        # Extract the data from the image
        data = extract_values_from_image(file_path)
        # Append the data to the list
        data_list.append(data)
        print(f"Data extracted for {filename}: {data}")  # Log extracted data

# Create a DataFrame with all the extracted data
df = pd.DataFrame(data_list)

# Define the Excel file path you want to save to
excel_file_path = 'C:\\Users\\dnai7\\PycharmProjects\\pythonProject1\\extracted_game_data_all.xlsx'

# Save the DataFrame to an Excel file with the correct data
df.to_excel(excel_file_path, index=False)

print(f"All data extracted and saved to {excel_file_path}")
