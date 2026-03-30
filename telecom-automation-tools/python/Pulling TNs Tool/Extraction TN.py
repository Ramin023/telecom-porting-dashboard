import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def extract_data():
    input_file = entry_var.get()
    mode = mode_var.get()
    
    #Error message for unselect file
    if not input_file:
        messagebox.showwarning("Warning", "Please select a text file first!")
        return

    try:
        with open(input_file, 'r', encoding='utf-8', errors='ignore') as file:
            lines = file.readlines()
            full_content = "".join(lines)

        results = []
        
        # --- Extraction Logic ---
        #=========Verizon==================
        if mode == "VZ":
            pattern = r"/TN\s+(\d{3})\s+(\d{3})-(\d{4})"
            matches = re.findall(pattern, full_content, re.IGNORECASE)
            cleaned = list(dict.fromkeys(["".join(m) for m in matches]))
            results = [{"TN Numbers": tn} for tn in cleaned]
        #=========Neustar==================
        elif mode == "Neustar":
            pattern = r"WTN\s+(\d{3})-(\d{3})-(\d{4})"
            matches = re.findall(pattern, full_content, re.IGNORECASE)
            cleaned = list(dict.fromkeys(["".join(m) for m in matches]))
            results = [{"TN Numbers": tn} for tn in cleaned]
        #=========MCI==================
        elif mode == "MCI":
            pattern = r"TN\s*:\s*(\d{10})"
            matches = re.findall(pattern, full_content, re.IGNORECASE)
            cleaned = list(dict.fromkeys(matches))
            results = [{"TN Numbers": tn} for tn in cleaned]
        #=========SBC==================
        elif mode == "SBC":
            pattern = r"Working Telephone Number \(WTN\):\s*([0-9\-]+)"
            matches = re.findall(pattern, full_content, re.IGNORECASE)
            cleaned = list(dict.fromkeys([m.replace("-", "") for m in matches]))
            results = [{"TN Numbers": tn} for tn in cleaned]
        #=========Frontier==================
        elif mode == "FT":
            current_wtn = ""
            re_wtn, re_sa = r"WTN:\s*([0-9]+)", r"SA\s+(.+)"
            for line in lines:
                w_m = re.search(re_wtn, line, re.IGNORECASE)
                if w_m: current_wtn = w_m.group(1)
                s_m = re.search(re_sa, line, re.IGNORECASE)
                if s_m and current_wtn:
                    results.append({"WTN": current_wtn, "SA": s_m.group(1).strip()})
                    current_wtn = ""
        #=========CN==================
        elif mode == "CN":
            unique_tns = {}
            re_range = r"(\d{3})-(\d{3})-(\d{4})-(\d{3})-(\d{3})-(\d{4})"
            re_single = r"\b\d{3}-\d{3}-\d{4}\b"
            for line in lines:
                range_matches = re.findall(re_range, line)
                for rm in range_matches:
                    start_val = int(f"{rm[0]}{rm[1]}{rm[2]}")
                    end_val = int(f"{rm[3]}{rm[4]}{rm[5]}")
                    for i in range(start_val, end_val + 1):
                        unique_tns[str(i)] = True
                single_matches = re.findall(re_single, line)
                for sm in single_matches:
                    unique_tns[sm.replace("-", "")] = True
            results = [{"TN Numbers": tn} for tn in unique_tns.keys()]
        
        elif mode == "WS":
            # This pattern finds 10 digits in a row OR digits separated by dashes
            pattern = r"\b(\d{3})[- ]?(\d{3})[- ]?(\d{4})\b"
            matches = re.findall(pattern, full_content)
            # Combine groups and deduplicate
            cleaned = list(dict.fromkeys(["".join(m) for m in matches]))
            results = [{"TN Numbers": tn} for tn in cleaned]

        if not results:
            messagebox.showinfo("No Results", f"No data found for {mode}.")
            return

        # Export to Excel
        df = pd.DataFrame(results)
        output_path = input_file.replace(".txt", f"_{mode}_Extracted.xlsx")
        df.to_excel(output_path, index=False)
        messagebox.showinfo("Success", f"Saved to:\n{output_path}")

    except Exception as e:
        messagebox.showerror("Error", f"Error: {str(e)}")

def browse_file():
    filename = filedialog.askopenfilename(filetypes=(("Text files", "*.txt"), ("All files", "*.*")))
    if filename: entry_var.set(filename)

# --- GUI Setup ---
root = tk.Tk()
root.title("Multi-Carrier TN Extractor")
root.geometry("500x320")

# Header
tk.Label(root, text="Select Carrier:", font=("Arial", 11, "bold")).pack(pady=10)

# Radio Buttons Frame (Grid layout for circles)
radio_frame = tk.Frame(root)
radio_frame.pack(pady=5)

mode_var = tk.StringVar(value="VZ")
carriers = [
    ("MCI", "MCI"), 
    ("Verizon", "VZ"), 
    ("Frontier", "FT"), 
    ("Neustar", "Neustar"), 
    ("CN", "CN"), 
    ("SBC", "SBC"),
    ("Windstream", "WS")
]

# Arrange radio buttons in 2 columns
for i, (text, mode_code) in enumerate(carriers):
    rb = tk.Radiobutton(radio_frame, text=text, variable=mode_var, value=mode_code, font=("Arial", 10))
    rb.grid(row=i//2, column=i%2, sticky="w", padx=20, pady=5)

# File Selection Area
file_frame = tk.Frame(root)
file_frame.pack(pady=20)

tk.Label(file_frame, text="File:").grid(row=0, column=0)
entry_var = tk.StringVar()
tk.Entry(file_frame, textvariable=entry_var, width=35).grid(row=0, column=1, padx=5)
tk.Button(file_frame, text="Browse", command=browse_file).grid(row=0, column=2)

# Run Button
tk.Button(
    root, 
    text="RUN EXTRACTION", 
    command=extract_data, 
    bg="#2B579A", 
    fg="white", 
    font=("Arial", 10, "bold"), 
    width=25, 
    height=2
).pack(pady=10)

root.mainloop()