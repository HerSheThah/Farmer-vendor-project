""" This application helps in connecting the farmers with their nearby markets by prompting to enter the pincode
    This application also provides the market vendors to register their market addresses

    Go through the following passage for the initial setup
    >> One should have two files (main_page.py vendor_info.py) stored in a directory
    >> The excel file (list_address.xlsx) and the text file (abstract.txt) should also be saved in the same directory

    Functions in main.py
    >> main_window() - Initial window including the application details and step to be followed
    >> popup_page() - Window displaying the options for farmers and vendor
    >> farmer_details() - Prompting to ask for farmer's data (pin-code)
    >> check() - To check if the correct data is entered and if the data is available in the database
    >> show_message() - To display all the markets found in the given pin-code
    >> is_valid_ph() - Checks if the phone number entered is valid or not
    >> is_valid_pincode() - Checks if the pin code entered is valid or not

    Functions in vendor_info.py
    >> vendor_details() - A window to prompting for the vendor details
    >> insert_ven_details() - Storing the valid addresses  in the database
"""

# Importing the necessary packages

import tkinter
from tkinter import ttk
from tkinter import *
from tkinter import messagebox as mb
from vendor_info import vendor_details
import sys
import os.path
import xlrd
import re

# Retrieving and storing the data from excel sheet
curPath = os.getcwd()
excelFile = curPath + '\list_address.xlsx'
if os.path.exists(excelFile):
    print(excelFile, ' file exist')
else:
    print(excelFile, 'does not exist')
    sys.exit()
wb = xlrd.open_workbook(excelFile)
sheet = wb.sheet_by_index(0)
no_rows = sheet.nrows
no_columns = sheet.ncols


# Clears all the data entered
def clear_farmer():
    entry1.get()
    entry1.delete(0, END)


# To display all the markets found in the given pin-code
def show_message(address_show):
    # Destroys the previous window
    farmer_window.destroy()
    # Creating a window
    message = Tk()
    message.geometry("550x500+400+100")
    # Creating main frame
    main_frame = Frame(message)
    main_frame.pack(fill=BOTH, expand=1)
    # Creating canvas
    my_canvas = Canvas(main_frame)
    my_canvas.pack(side=LEFT, fill=BOTH, expand=1)
    # Adding scroll bar
    my_scrollbar = ttk.Scrollbar(main_frame, orient=VERTICAL, command=my_canvas.yview)
    my_scrollbar.pack(side=RIGHT, fill=Y)
    # Configure The Canvas to scrollbar
    my_canvas.configure(yscrollcommand=my_scrollbar.set)
    my_canvas.bind('<Configure>', lambda e: my_canvas.configure(scrollregion=my_canvas.bbox("all")))
    # Create another frame
    second_frame = Frame(my_canvas)
    # Add that new frame to the window in the canvas
    my_canvas.create_window((1, 1), window=second_frame, anchor="n")
    lab = Label(second_frame, text="       Here are the Market Addresses we found for you\n", font="bold")
    lab.config(pady=10)
    lab.pack()
    # Printing the addresses collected in the list
    for i, address in enumerate(address_show):
        label_frame = LabelFrame(second_frame, text='Address'+str(i+1))
        label_frame.config(labelanchor="n")
        label_frame.pack(fill="both", padx=15, pady=3)
        for j, col in enumerate(address):
            if j >= 4:
                col = int(col)
                Label(label_frame, text=col, anchor="w").pack(fill="both")
            else:
                Label(label_frame, text=str(col)+", ", anchor="w").pack(fill="both")
    Label(second_frame, text="").pack(fill="both")
    message.mainloop()


# To check if the correct data is entered and if the data is available in the database
def check():
    pin = entry1.get()
    # Checking if the data is non empty
    if len(pin) != 0:
        if is_valid_pincode(pin):
            if len(pin) == 7:
                # Removing the extra spaces
                pin = pin[0:3] + pin[4:7]
            list_address = []
            # Checking if the pincode is available in the database
            # and appending to the list if the  data exists
            for i in range(0, no_rows):
                curr_pincode = sheet.cell_value(i, 4)
                if int(pin) == curr_pincode:
                    row_address = []
                    [row_address.append(sheet.cell_value(i, k)) for k in range(0, no_columns)]
                    list_address.append(row_address)
            # Displaying the address  if the list is not empty
            if len(list_address) != 0:
                show_message(list_address)
            # if no address is found display a message
            else:
                mb.showinfo("Thank you", "Sorry, we couldn't find any markets in the given pincode. "
                                         "We will surely extend our service in real soon.")

                msgbox = mb.askquestion("Exit Application", "Do you want to quit the application?", icon="warning")
                if msgbox == "yes":
                    farmer_window.destroy()
                    sys.exit()
                else:
                    clear_farmer()
        # Showing error if the data is  invalid
        else:
            mb.showerror("Pincode", "Enter a valid Pincode")
            clear_farmer()
            # print("false")
    # Showing warning if the data is empty
    else:
        mb.showwarning("Warning", "Fill the required field")


# Function for entering former details
def farmer_details():
    new_window.destroy()
    global farmer_window
    farmer_window = Tk()
    # Creating an window
    farmer_window.title("Market Application")
    farmer_window.geometry("212x150+490+200")
    # Getting pin code
    pin_label = Label(farmer_window, text="Enter the pincode of your area", font=("calibre", 10, "bold"))
    pin_label.grid(row=0, column=1, padx=10, pady=10)
    global entry1
    entry1 = Entry(farmer_window, borderwidth=5)
    entry1.grid(row=1, column=1)
    submit_button = Button(farmer_window, text="SUBMIT", command=check, underline=0, width=12)
    submit_button.grid(row=3, column=1, padx=50, pady=15, sticky=tkinter.W)


# Phone number validation function
def is_valid_ph(ph):
    # 1) Begins with 0 or 91
    # 2) Then contains 7 or 8 or 9.
    # 3) Then contains 9 digits
    pattern = re.compile("(0/91)?[7-9][0-9]{9}")
    return pattern.match(ph)


# Pincode validation function
def is_valid_pincode(pin):
    regex = "^[1-9]{1}[0-9]{2}\\s{0,1}[0-9]{3}$"
    p = re.compile(regex)
    m = re.match(p, pin)
    if m is None:
        return False
    else:
        return True


# Window displaying the options for farmers and vendor
def popup_page():
    mainWindow.destroy()
    global new_window
    # creating a window
    new_window = tkinter.Tk()
    new_window.title("Market Application")
    new_window.geometry('480x150+400+200')
    # Configuring rows and columns
    new_window.columnconfigure(0, weight=1)
    new_window.columnconfigure(1, weight=2)
    new_window.columnconfigure(2, weight=2)
    new_window.rowconfigure(0, weight=1)
    new_window.rowconfigure(1, weight=2)
    new_window.rowconfigure(2, weight=2)
    heading = tkinter.Label(new_window, text="Are you a farmer looking for markets to sell your veggies?")
    heading.config(font=("Helvetica", 12))
    heading.grid(row=0, column=1, columnspan=2, padx=5, pady=5)
    # Farmer button
    frame = tkinter.Frame(new_window)
    frame.grid(row=1, column=1, columnspan=2)
    yes_button = tkinter.Button(frame, text="YES", width=10, height=2, activeforeground="dark gray",
                                borderwidth=3, command=farmer_details)
    yes_button.grid(row=0, column=1, padx=5)
    # Vendor button
    vendor_button = tkinter.Button(frame, text="No I'm a vendor looking forward to register my shop", height=2,
                                   activeforeground="gray", borderwidth=3,
                                   command=lambda: [new_window.destroy(), vendor_details()])
    vendor_button.grid(row=0, column=3, padx=5)
    # Button brings back to the instruction page
    back_button = Button(new_window, text="Back", width=10, height=2, underline=0,
                         command=lambda: [new_window.destroy(), main_window()])
    back_button.grid(row=2, column=2, sticky="sw", pady="10")


# Initial window including the application details and step to be followed
def main_window():
    global mainWindow
    mainWindow = Tk()
    mainWindow.title("Farmer market Application")
    mainWindow.geometry('640x480+400+80')
    # Main frame
    main_frame = Frame(mainWindow)
    main_frame.pack(fill=BOTH, expand=1)
    # Create canvas
    my_canvas = Canvas(main_frame)
    my_canvas.pack(side=LEFT, fill=BOTH, expand=1)
    # Add scroll bar
    my_scrollbar = ttk.Scrollbar(main_frame, orient=VERTICAL, command=my_canvas.yview)
    my_scrollbar.pack(side=RIGHT, fill=Y)
    # Configure The Canvas
    my_canvas.configure(yscrollcommand=my_scrollbar.set)
    my_canvas.bind('<Configure>', lambda e: my_canvas.configure(scrollregion=my_canvas.bbox("all")))
    # Create another frame
    second_frame = Frame(my_canvas)
    # Add that new frame to the window in the canvas
    my_canvas.create_window((2, 2), window=second_frame, anchor="n")
    lab = Label(second_frame, text="       ABOUT", font="bold")
    lab.config(pady=10)
    # Reading and writing the text file
    with open("abstract.txt", "r") as f:
        label = Label(second_frame, text=f.read())
        label.config(justify=LEFT, padx=100)
        label.pack()
    button = Button(second_frame, text="Next", width=15, height=2, underline=0, borderwidth=2, command=popup_page)
    button.pack(pady=10)
    mainWindow.mainloop()


# If the current window is main window run the program
if __name__ == "__main__":
    main_window()
