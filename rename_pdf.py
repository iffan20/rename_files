import os
import sys
import PyPDF2
import pandas as pd

def read_pdf_data(file_path, limit_lines=10):
    """
    Extracts data from a PDF file.

    Args:
        file_path (str): The path to the PDF file.
        limit_lines (int): Maximum number of lines to extract.

    Returns:
        list: Extracted data from the PDF file.
        str: First 10 lines from the PDF file.
    """
    extracted_data = []
    first_lines = ""
    lines_processed = 0

    with open(file_path, 'rb') as file:
        reader = PyPDF2.PdfFileReader(file)
        for page_num in range(reader.numPages):
            page = reader.getPage(page_num)
            lines = page.extractText().strip().split('\n')
            for line in lines:
                if lines_processed < limit_lines:
                    first_lines += line.strip() + "\n"
                    extracted_data.extend([data for data in line.strip().split() if len(data) == 8 and data.isdigit()])
                    lines_processed += 1
                    if extracted_data:
                        break  # Break out of the loop once the condition is met
                else:
                    break
            if extracted_data:
                break  # Break out of the outer loop if the condition was met

    return extracted_data, first_lines

def rename_files(folder_path, excel_path):
    """
    Processes PDF files in a folder, extracting data and renaming files based on project names from an Excel file.

    Args:
        folder_path (str): The path to the folder containing PDF files.
        excel_path (str): The path to the Excel file containing project data.
    """

    sys.stdout.reconfigure(encoding='utf-8')  # Set UTF-8 encoding for console output

    # Read the Excel file using pandas
    df = pd.read_excel(excel_path)

    for filename in os.listdir(folder_path):
        if filename.endswith(".pdf"):
            file_path = os.path.join(folder_path, filename)
            extracted_data, _ = read_pdf_data(file_path)

            if extracted_data:
                extracted_string = ''.join(extracted_data)
                try:
                    extracted_integers = int(extracted_string[:8])
                except ValueError:
                    print("Error: Extracted data contains non-numeric characters.")
                    continue  # Skip processing this file

                # Check if extracted data matches any values in the "รหัสนิสิต" column
                matched_projects = df[df['รหัสนิสิต'] == extracted_integers]['ชื่อโปรเจ็ค'].tolist()
                projects_name = ''.join(matched_projects)

                if projects_name:
                    new_filename = f"{extracted_integers}_{projects_name}.pdf"
                    new_path = os.path.join(folder_path, new_filename)

                    # Check if the file with the new name already exists
                    counter = 1
                    while os.path.exists(new_path):
                        new_filename = f"{extracted_integers}_{projects_name}_{counter}.pdf"
                        new_path = os.path.join(folder_path, new_filename)
                        counter += 1

                    # Rename the file
                    old_path = os.path.join(folder_path, filename)
                    os.rename(old_path, new_path)
                    print(f"File '{filename}' renamed to '{new_filename}'")
                else:
                    print(f"No project name found")
            else:
                print(f"No valid ID found ")
        else:
            print(f"Not a PDF file.")

# Example usage for a folder containing PDF files and the Excel file path
folder_path = r"C:\Users\user\Desktop\...."
excel_path = r"C:\Users\user\Downloads\......xlsx"
rename_files(folder_path, excel_path)
