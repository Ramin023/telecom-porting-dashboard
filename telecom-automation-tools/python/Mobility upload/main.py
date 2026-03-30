import tkinter as tk
from tkinter import filedialog
from mobExcel import mobExcel
import os


def process_data():
    
    #locate app floder
    script_dir = os.path.dirname(__file__)
    os.chdir(script_dir)
    current_directory = os.getcwd()
    
    file_temp = "upload_template.xlsx"
    filePathRead = os.path.join(current_directory, file_temp)
    
    file_config =  "configue.txt"
    filePathCongig = os.path.join(current_directory, file_config)

    #read cofigation 
    with open(filePathCongig, 'r') as file:
        for line in file:
            d = line.split(":")
            if d[0] == "User Name":
                user_name = d[1].strip()
            elif d[0] == "Save Directory":
                if len(d) == 2:
                    pathWrite = current_directory
                else:
                    pathWrite = d[1].strip()+":"+d[2].strip()


    # Extract the checkbox values
    s = var_s.get()
    c = var_c.get()
    e = var_e.get()

    # Create an instance of mobExcel and call the respective functions based on the checkboxes
    work = mobExcel(user_name, pathWrite, filePathRead)
    if s == 1:
        work.addService()
    if c == 1:
        work.addCharge()
    if e == 1:
        work.addEquipment()
    
    print ("*Done*")
    print ("============================================================================")

if __name__ == "__main__":
    # Set up the main tkinter window
    root = tk.Tk()
    root.title("App")
    root.geometry("250x100")  # Set the initial size of the window
    root.resizable(False, False)

    # Create variables to hold the checkbox values and set them to 1 (checked) by default
    var_s = tk.IntVar(value=1)
    var_c = tk.IntVar(value=1)
    var_e = tk.IntVar(value=1)

    # Create checkboxes in one row
    tk.Label(root, text="Create upload Excel").pack()

    checkbox_frame = tk.Frame(root)
    checkbox_frame.pack()

    s_checkbox = tk.Checkbutton(checkbox_frame, text="S", variable=var_s)
    s_checkbox.pack(side=tk.LEFT)

    c_checkbox = tk.Checkbutton(checkbox_frame, text="C", variable=var_c)
    c_checkbox.pack(side=tk.LEFT)

    e_checkbox = tk.Checkbutton(checkbox_frame, text="E", variable=var_e)
    e_checkbox.pack(side=tk.LEFT)

    # Create a button to trigger processing the data
    tk.Button(root, text="Submit", command=process_data).pack()

    # Run the tkinter event loop
    root.mainloop()
