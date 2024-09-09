import customtkinter as ctk
from tkinter import filedialog
import qrcode
import openpyxl
import os
import tkinter.simpledialog

class QRCodeGeneratorApp:
    def __init__(self, master):
        self.master = master
        self.master.title("QR Code Generator")
        self.master.geometry("400x300")

        # Set appearance mode and theme
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # Create GUI elements
        self.file_button = ctk.CTkButton(self.master, text="Browse", command=self.load_file, font=("Arial", 16, "bold"), text_color="white")
        self.file_button.pack(pady=20)

        self.file_label = ctk.CTkLabel(self.master, text="No file selected", font=("Arial", 16), text_color="white")
        self.file_label.pack(pady=10)

        self.generate_button = ctk.CTkButton(self.master, text="Generate QR Codes", command=self.generate_qr_codes, font=("Arial", 16, "bold"), text_color="white")
        self.generate_button.pack(pady=20)

    def load_file(self):
        # Open a file dialog to select an Excel file
        self.excel_file = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx")],
            title="Select Excel File"
        )
        if self.excel_file:
            self.file_label.configure(text=f"Selected File: {os.path.basename(self.excel_file)}")

    def generate_qr_codes(self):
        if not hasattr(self, 'excel_file') or not self.excel_file:
            self.file_label.configure(text="No file selected")
            return

        # Ask the user to input the folder name for saving QR codes
        qr_code_folder = tkinter.simpledialog.askstring("Folder Name", "Enter the name for the QR codes folder:")
        if not qr_code_folder:
            self.file_label.configure(text="Folder name not provided")
            return

        # Create the folder if it doesn't exist
        if not os.path.exists(qr_code_folder):
            os.makedirs(qr_code_folder)

        # Load the Excel workbook and worksheet
        wb = openpyxl.load_workbook(self.excel_file)
        worksheet = wb.active

        # Iterate through the rows in the Excel sheet, starting from the second row
        for row in worksheet.iter_rows(min_row=2, values_only=True):
            student_id, student_name = row
            # Combine the student name and ID
            student_info = f"{student_id} - {student_name}"
            
            # Create a unique QR code for each student
            qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_L,
                box_size=10,
                border=4,
            )
            qr.add_data(student_info)
            qr.make(fit=True)

            # Save the QR code as an image file with both student name and ID
            img = qr.make_image(fill='black', back_color='white')
            img.save(f'{qr_code_folder}/{student_id} - {student_name}.png')

        self.file_label.configure(text="QR codes generated and saved!")

if __name__ == "__main__":
    root = ctk.CTk()
    app = QRCodeGeneratorApp(root)
    root.mainloop()