import docx
import openpyxl
import sys
import os
import pandas as pd

def read_docx_data(file_path, limit_lines=10):
    """
    Extracts data from a DOCX file.

    Args:
        file_path (str): The path to the DOCX file.
        limit_lines (int): Maximum number of lines to extract.

    Returns:
        list: Extracted data from the DOCX file.
        str: First 10 lines from the DOCX file.
    """
    doc = docx.Document(file_path)
    extracted_data = []
    first_lines = ""
    lines_processed = 0

    for paragraph in doc.paragraphs:
        lines = paragraph.text.strip().split('\n')
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
    Processes DOCX files in a folder, extracting data and renaming files based on project names from an Excel file.

    Args:
        folder_path (str): The path to the folder containing DOCX files.
        excel_path (str): The path to the Excel file containing project data.
    """

    sys.stdout.reconfigure(encoding='utf-8')  # Set UTF-8 encoding for console output

    # Read the Excel file using pandas
    df = pd.read_excel(excel_path)

    for filename in os.listdir(folder_path):
        if filename.endswith(".docx"):
            file_path = os.path.join(folder_path, filename)
            extracted_data, _ = read_docx_data(file_path)

            if extracted_data:
                
                extracted_string = ''.join(extracted_data)
                print (extracted_string)
                try:
                    extracted_integers = int(extracted_string[:8])
                except ValueError:
                    print("Error: Extracted data contains non-numeric characters.")
                    continue  # Skip processing this file
                
                    
                # Check if extracted data matches any values in the "รหัสนิสิต" column
                matched_projects = df[df['รหัสนิสิต'] == extracted_integers]['ชื่อโปรเจ็ค'].tolist()
                projects_name = ''.join(matched_projects)

                if projects_name:
                    new_filename = f"{extracted_integers}_{projects_name}.docx"
                    new_path = os.path.join(folder_path, new_filename)

                    # Check if the file with the new name already exists
                    counter = 1
                    while os.path.exists(new_path):
                        digit = len(extracted_string)
                        i = 0
                        while i < digit:
                            x = 1  # Set x to 1 initially
                            # Check if the segment of 8 digits starting at index i is already in the filename
                            Index_Id = extracted_string[i:i+8]
                            if Index_Id in new_filename:
                                i += 8    
                                if x == 1:
                                    new_filename = f"{Index_Id}_{projects_name}.docx"
                                    new_path = os.path.join(folder_path, new_filename)
                                    counter -= 1
                                    x += 1
                                else:
                                    new_filename = f"{Index_Id}_{projects_name}_{counter}.docx"
                                    new_path = os.path.join(folder_path, new_filename)
                                    counter += 1
                            else:
                                break


                     # Rename the file        
                    old_path = os.path.join(folder_path, filename)
                    os.rename(old_path, new_path)
                    print(f"File '{filename}' renamed to '{new_filename}'")
                else:
                    print(f"No project name found")
            else:
                print(f"No valid ID found ")
        else:
            print(f"Not a DOCX file.")

# Example usage for a folder containing DOCX files and the Excel file path
folder_path = r"C:\Users\user\Desktop\....."  #folder path contain files 
excel_path = r"C:\Users\user\Downloads\.........xlsx" # excel path to matching value
rename_files(folder_path, excel_path)
