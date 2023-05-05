import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import tabula
from tqdm import tqdm
import time

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
            print("Initializing. . .")
            time.sleep(5)
            print("This may take a while depending on the size of the file")
            pages = tabula.read_pdf(file_path, pages='all')
            progress_bar = tqdm(total=len(pages), desc="Converting pages")

            output_file = os.path.splitext(file_path)[0] + ".xlsx"
            writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

            start_time = time.time() #initiate timer for progress bar
            for i, page in enumerate(pages):
                print("Generating your Excel file . . .")
                sheet_name = 'Page' + str(i+1)
                page.to_excel(writer, index=False, sheet_name=sheet_name)
                progress_bar.update(1)
                elapsed_time = time.time() - start_time
                progress_bar.set_description(f"Converting pages - Elapsed time: {elapsed_time:.2f} seconds")

            writer.close()
            progress_bar.close()
            total_time = time.time() - start_time
            print(f"PDF conversion to Excel successful! Total elapsed time: {total_time:.2f} seconds")

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
