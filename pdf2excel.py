#add each page to seperate sheet eureka!
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import tabula

# Step 1: Welcome message
print("Welcome to the PDF to Excel Converter!")

# Step 2: Ask the user to confirm if they wish to proceed.
response = input("Do you wish to proceed? (y/n): ")

# Proceed only if the user confirms
if response.lower() == "y":
    # Step 3: Ask the user to select a PDF file
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    
    # Proceed only if the user selected a file
    if file_path:
    # Step 4: Convert the PDF to Excel and notify the user upon successful conversion
        try:
            output_file = os.path.splitext(file_path)[0] + ".xlsx"
            writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
            for i, page in enumerate(tabula.read_pdf(file_path, pages='all')):
                sheet_name = 'Page' + str(i+1)
                page.to_excel(writer, index=False, sheet_name=sheet_name)
            writer.close()
            print("PDF conversion to Excel successful!")
            
            # Step 5: Ask the user if they want to save the converted file
            response = input("Do you want to save the converted file? (y/n): ")
            if response.lower() == "y":
                # Prompt the user to enter a filename for the Excel file
                root = tk.Tk()
                root.withdraw()
                output_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
                if output_path:
                    # Save the Excel file
                    os.replace(output_file, output_path)
                    print("Excel file saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", str(e))