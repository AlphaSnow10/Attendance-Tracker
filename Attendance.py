import cv2
import pandas as pd
import customtkinter as ctk
from PIL import Image, ImageTk
import threading
import queue
import tkinter as tk
from tkinter import filedialog, simpledialog
import datetime

class AttendanceTracker:
    def __init__(self, master):
        self.master = master
        self.master.title("Attendance Tracker")
        self.master.geometry("1000x800")

        # Set appearance mode and theme
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # Create GUI elements
        self.file_button = ctk.CTkButton(self.master, text="Browse for Excel File", command=self.load_excel)
        self.file_button.place(x=500, y = 60, anchor = ctk.CENTER)

        self.file_label = ctk.CTkLabel(self.master, text="No file selected", font=("Arial", 14))
        self.file_label.place(x=500, y = 90, anchor = ctk.CENTER)

        self.video_frame = ctk.CTkFrame(self.master, width=640, height=480)
        self.video_frame.pack(side=ctk.LEFT, padx=20, pady=20)

        self.video_title_label = ctk.CTkLabel(self.master, text="QR Code Scanner", 
                                              font=("Arial", 18, "bold"), 
                                              text_color="white")
        self.video_title_label.place(x=260, y = 180, anchor="center")

        self.attendance_frame = ctk.CTkFrame(self.master, width=320, height=480)
        self.attendance_frame.pack(side=ctk.RIGHT, padx=20, pady=20)

        self.video_label = ctk.CTkLabel(self.video_frame, text="  ", font=("Arial", 16))
        self.video_label.pack(pady=10)

        self.attendance_label = ctk.CTkLabel(self.attendance_frame, text="Attendance", font=("Arial", 20, "bold"), text_color="white")
        self.attendance_label.pack(pady=10)

        self.attendance_table = ctk.CTkScrollableFrame(self.attendance_frame, width=300, height=400)
        self.attendance_table.pack(pady=10)

        self.attendance_table_heading = ctk.CTkLabel(self.attendance_table, text="ID | Name | Status", font=("Arial", 16, "bold"), text_color="white")
        self.attendance_table_heading.grid(row=0, column=0, pady=5, sticky="ew")

        self.save_button = ctk.CTkButton(self.attendance_frame, 
                                 text="Save Attendance", 
                                 command=self.save_and_close, 
                                 font=("Arial", 16, "bold"), text_color="snow",  
                                 hover_color="midnight blue",  
                                 width=200, height=50)  
        self.save_button.pack(pady=20)

        # Initialize the camera
        self.cap = cv2.VideoCapture(0)
        self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
        self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)

        # Create a dictionary to store the scanned QR codes
        self.scanned_qr_codes = {}

        # Create queues for frame and QR code data
        self.frame_queue = queue.Queue(maxsize=10)
        self.qr_code_queue = queue.Queue()

        # Start threads
        self.video_thread = threading.Thread(target=self.capture_video)
        self.video_thread.daemon = True
        self.video_thread.start()

        self.qr_thread = threading.Thread(target=self.process_qr_codes)
        self.qr_thread.daemon = True
        self.qr_thread.start()

        # Schedule GUI updates
        self.master.after(100, self.update_gui)

    def load_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.student_df = pd.read_excel(file_path)
            self.file_label.configure(text=f"Loaded file: {file_path}")

    def capture_video(self):
        while True:
            ret, frame = self.cap.read()
            if not ret:
                break

            # Push frame to the queue
            if not self.frame_queue.full():
                self.frame_queue.put(frame)

    def process_qr_codes(self):
        qr_code_detector = cv2.QRCodeDetector()
        while True:
            if not self.frame_queue.empty():
                frame = self.frame_queue.get()

                # QR code detection
                data, bbox, rectifiedImage = qr_code_detector.detectAndDecode(frame)

                if data and data not in self.scanned_qr_codes:
                    student_info = data.split(' - ')
                    student_name = student_info[1]
                    student_id = student_info[0]

                    self.scanned_qr_codes[student_id] = True
                    self.qr_code_queue.put((student_id, student_name))

    def update_gui(self):
        if not self.frame_queue.empty():
            frame = self.frame_queue.get()
            frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(frame_rgb)
            imgtk = ImageTk.PhotoImage(image=img)
            self.video_label.imgtk = imgtk
            self.video_label.configure(image=imgtk)

        try:
            student_name, student_id = self.qr_code_queue.get_nowait()
            self.update_attendance_table(student_id, student_name)
        except queue.Empty:
            pass

        # Schedule the next update
        self.master.after(1, self.update_gui)

    def update_attendance_table(self, student_id, student_name):
        for child in self.attendance_table.winfo_children():
            if isinstance(child, ctk.CTkLabel) and student_id in child.cget("text"):
                return

        new_entry = ctk.CTkLabel(self.attendance_table, text=f"{student_name} | {student_id} | Present", font=("Arial", 12))
        row_index = len(self.attendance_table.winfo_children())
        new_entry.grid(row=row_index, column=0, pady=5)

    def save_and_close(self):
        attendance_df = pd.DataFrame(columns=['Student ID', 'Student Name', 'Status', 'Date'])

        # Get the current date
        current_date = datetime.date.today().strftime('%Y-%m-%d')

        # Create a set of scanned student IDs
        scanned_ids = set(self.scanned_qr_codes.keys())

        for index, row in self.student_df.iterrows():
            student_id = row['Student ID']
            student_name = row['Student Name']
            
            if student_id in scanned_ids:
                # Student is present
                new_row = pd.DataFrame({'Student ID': [student_id], 'Student Name': [student_name], 'Status': ['Present'], 'Date': [current_date]})
            else:
                # Student is absent
                new_row = pd.DataFrame({'Student ID': [student_id], 'Student Name': [student_name], 'Status': ['Absent'], 'Date': [current_date]})
            
            attendance_df = pd.concat([attendance_df, new_row], ignore_index=True)

        # Get file name from user
        file_name = tk.simpledialog.askstring("Save File", "Enter the name for the attendance file:")
        if file_name:
            attendance_df.to_excel(file_name + '.xlsx', index=False)

        self.cap.release()
        cv2.destroyAllWindows()
        self.master.destroy()


if __name__ == "__main__":
    root = ctk.CTk()
    app = AttendanceTracker(root)
    root.mainloop()