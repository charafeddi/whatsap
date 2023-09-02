import sys
import time
import pyautogui
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QLabel, QLineEdit, QTextEdit, QPushButton, QFileDialog, QVBoxLayout, QListWidget, QListWidgetItem, QDesktopWidget

# Create a class for the main WhatsApp Automation GUI
class WhatsAppAutomationGUI(QMainWindow):
    def __init__(self):
        super().__init__()

        # Set window properties
        self.setWindowTitle("WhatsApp Automation")
        self.resize(700, 700)
        self.center()
      
        # Create and configure GUI elements

        # Label and entry for Excel file selection
        self.excel_label = QLabel("Select Excel File:", self)
        self.excel_label.move(20, 90)
        self.excel_label.resize(200,30)
        self.excel_file_entry = QLineEdit(self)
        self.excel_file_entry.setGeometry(20, 120, 520, 30)

        # Button to browse for Excel file
        self.browse_excel_button = QPushButton("Browse", self)
        self.browse_excel_button.setGeometry(560, 120, 80, 30)
        self.browse_excel_button.clicked.connect(self.browse_excel_file)

        # Label and text input box for message
        self.message_label = QLabel("Message:", self)
        self.message_label.move(20, 160)
        self.message_text = QTextEdit(self)
        self.message_text.setGeometry(20, 200, 620, 120)

        # List widget for selected images
        self.selected_images_label = QLabel("Selected Images:", self)
        self.selected_images_label.move(20, 360)
        self.selected_images_label.resize(150,30)
        self.selected_images_list = QListWidget(self)
        self.selected_images_list.setGeometry(20, 400, 620, 80)

        # Button to add images
        self.add_images_button = QPushButton("Add Images", self)
        self.add_images_button.setGeometry(20, 500, 150, 30)
        self.add_images_button.clicked.connect(self.add_images)

        # Button to remove images
        self.remove_images_button = QPushButton("Remove Images", self)
        self.remove_images_button.setGeometry(200, 500, 150, 30)
        self.remove_images_button.clicked.connect(self.remove_images)

        # Button to send WhatsApp messages
        self.send_button = QPushButton("Send WhatsApp Messages", self)
        self.send_button.setGeometry(200, 600, 250, 40)
        self.send_button.clicked.connect(self.send_whatsapp_messages)

    # Function to handle Excel file selection using a file dialog
    def browse_excel_file(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        self.excel_file_entry.setText(file_path)

    # Function to add images using a file dialog
    def add_images(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_paths, _ = QFileDialog.getOpenFileNames(self, "Select Image Files", "", "Image Files (*.jpg *.jpeg *.png *.gif);;All Files (*)", options=options)
        
        # Add selected images to the list widget
        for file_path in file_paths:
            image_item = QListWidgetItem(file_path)
            self.selected_images_list.addItem(image_item)

    # Function to remove selected images from the list widget
    def remove_images(self):
        selected_items = self.selected_images_list.selectedItems()
        for item in selected_items:
            self.selected_images_list.takeItem(self.selected_images_list.row(item))

    # Function to send WhatsApp messages
    def send_whatsapp_messages(self):
        excel_file_path = self.excel_file_entry.text()
        message = self.message_text.toPlainText()
        selected_images = [self.selected_images_list.item(i).text() for i in range(self.selected_images_list.count())]
        
        # Load the Excel file into a pandas DataFrame
        try:
            df = pd.read_excel(excel_file_path, engine='openpyxl')  # Use engine='openpyxl' for .xlsx files
        except Exception as e:
            print(f"Error: {e}")
            df = None

        # Check if the DataFrame was loaded successfully
        if df is not None:
           # Assuming 'Contact' is the column name in your Excel file
            contact_list = df['Contact'].tolist()

            # Print the list of names
            print(contact_list)
            print(selected_images)
            
            # Function to send WhatsApp message
            def send_whatsapp_message(contact_name, message, selected_images):
                # Adjust the coordinates based on screen resolution and browser window position
                search_box_x, search_box_y = 200, 320
                message_box_x, message_box_y = 980, 1460
        
                # Open a chat with the specified contact
                pyautogui.click(search_box_x, search_box_y)
                pyautogui.write(contact_name)
                time.sleep(2)
                pyautogui.press('enter')
                time.sleep(5)  # Increase this delay to ensure WhatsApp Web fully loads

                #send whatsapp images
                send_whatsapp_images(selected_images)
                time.sleep(4)  # Add a delay to allow WhatsApp Web to fully load for writing messages
                
                # Type and send the message
                pyautogui.click(message_box_x, message_box_y)
                pyautogui.write(message)
                pyautogui.press('enter')

            # function to send Whatsapp images
            def send_whatsapp_images(selected_images):
                # Initialize file_names as an empty string
                file_names = ''

                # Adjust the coordination based on screen resolution and browser window position
                plus_btn_x, plus_btn_y = 920, 1460
                photoVideo_btn_x, photoVideo_btn_y = 1030, 1130
                send_btn_x, send_btn_y = 2420, 1440
                
                for i in selected_images:
                    
                    file_names += '"' + i.replace('/', '\\') + '" ' 
                #open the plus button
                pyautogui.click(plus_btn_x, plus_btn_y)
                time.sleep(1)

                #open the photo and video button
                pyautogui.click(photoVideo_btn_x, photoVideo_btn_y)
                time.sleep(1)

                #Type to choose the images
                #pyautogui.click(FileName_box_x, FileName_box_y)
                pyautogui.write(file_names)
                pyautogui.press("enter")
                time.sleep(3)

                #send files
                pyautogui.click(send_btn_x, send_btn_y)
                
                    
            open_whatsapp_web()
            for contact in contact_list:
                send_whatsapp_message(contact, message, selected_images)
                time.sleep(5)  # Add a delay to allow WhatsApp Web to fully load for the next contact

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)

# Function to open WhatsApp Web in your web browser
def open_whatsapp_web():
    # Adjust the coordinates based on your screen resolution and browser window position
    pyautogui.hotkey('win', 'r')  # Opens the Run dialog
    time.sleep(1)
    pyautogui.write('https://web.whatsapp.com')
    pyautogui.press('enter')
    time.sleep(15)  # Adjust the sleep duration to allow time for WhatsApp Web to load

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = WhatsAppAutomationGUI()
    window.show()
    app.exec_()  # Start the event loop

    # Your program has finished executing, so you can quit the application
    app.quit()
