import tkinter as tk
from tkinter import ttk, simpledialog, messagebox
import pandas as pd
import math

# Constants
INITIAL_VISIBLE_ROWS = 10
MAX_ROWS = 30

# GUI Components
root = tk.Tk()
root.title("AdyMob Factory Calculator")
root.state('zoomed')

# Global variables
total_area = 0
total_cost = 0

def calculate_row(row):
    try:
        nr_placi = int(entry_nr_placi[row].get())
        latime = float(entry_latime[row].get())
        lungime = float(entry_lungime[row].get())
        cost_mp = float(entry_cost[row].get())

        # Convert units to meters
        unit = unit_var.get()
        if unit == "cm":
            latime /= 100
            lungime /= 100
        elif unit == "mm":
            latime /= 1000
            lungime /= 1000

        # Calculate perimeter, area, and cost
        perimeter = 2 * nr_placi * (latime + lungime)
        area = nr_placi * latime * lungime
        cost = area * cost_mp

        # Display results
        label_perimeter[row].config(text=f"{perimeter:.3f} m")
        label_area[row].config(text=f"{area:.3f} m²")
        label_cost[row].config(text=f"{cost:.2f} EURO")
    except ValueError:
        label_perimeter[row].config(text="Invalid")
        label_area[row].config(text="Invalid")
        label_cost[row].config(text="Invalid")

    # Recalculate totals
    calculate_totals()

def calculate_totals():
    global total_area, total_cost
    total_area = 0
    total_perimeter = 0
    total_cost = 0

    for row in range(visible_rows):
        try:
            area = float(label_area[row].cget("text").split()[0])
            perimeter = float(label_perimeter[row].cget("text").split()[0])
            cost = float(label_cost[row].cget("text").split()[0])
            total_area += area
            total_perimeter += perimeter
            total_cost += cost
        except ValueError:
            continue

    label_total_area.config(text=f"Total Area: {total_area:.3f} m²")
    label_total_perimeter.config(text=f"Total Perimeter: {total_perimeter:.3f} m")
    label_total_cost.config(text=f"Total Cost: {total_cost:.2f} EURO")
    calculate_sheets_needed()

def calculate_sheets_needed():
    try:
        sheet_width = float(entry_sheet_width.get())
        sheet_length = float(entry_sheet_length.get())
        if sheet_width > 0 and sheet_length > 0:
            sheet_area = sheet_width * sheet_length
            sheets_needed = total_area / sheet_area
            label_sheets_needed.config(text=f"{sheets_needed:.1f} sheets ({math.ceil(sheets_needed)} round up)")
    except ValueError:
        label_sheets_needed.config(text="Invalid dimensions")

def save_to_excel():
    global total_area, total_cost
    filename = simpledialog.askstring("Save File", "Enter the filename:", initialvalue="AdyMob_Calculator_Output")
    if filename:
        data = {
            "Index": [],
            "Numar Coli": [],
            "Lățime": [],
            "Lungime": [],
            "Unitate": [],
            "Cost/m² (EURO)": [],
            "Perimetru (m)": [],
            "Suprafață (m²)": [],
            "Cost Total (EURO)": []
        }

        for i in range(visible_rows):
            if entry_nr_placi[i].get() or entry_latime[i].get() or entry_lungime[i].get():
                data["Index"].append(i + 1)
                data["Numar Coli"].append(entry_nr_placi[i].get())
                data["Lățime"].append(entry_latime[i].get())
                data["Lungime"].append(entry_lungime[i].get())
                data["Unitate"].append(unit_var.get())
                data["Cost/m² (EURO)"].append(entry_cost[i].get())
                data["Perimetru (m)"].append(label_perimeter[i].cget("text"))
                data["Suprafață (m²)"].append(label_area[i].cget("text"))
                data["Cost Total (EURO)"].append(label_cost[i].cget("text"))

        df = pd.DataFrame(data)

        try:
            sheet_width = float(entry_sheet_width.get())
            sheet_length = float(entry_sheet_length.get())
            sheet_area = sheet_width * sheet_length
            sheets_needed = total_area / sheet_area
        except ValueError:
            sheet_width = sheet_length = sheet_area = sheets_needed = 0

        total_data = {
            "Total Area": f"{total_area:.3f} m²",
            "Total Cost": f"{total_cost:.2f} EURO",
            "Sheets Needed": f"{sheets_needed:.1f} ({math.ceil(sheets_needed)} round up)"
        }

        for key, value in total_data.items():
            df.loc["Total", key] = value

        df.to_excel(f"{filename}.xlsx", index=False)
        messagebox.showinfo("Save File", "Data has been saved successfully!")

def show_more_rows():
    global visible_rows
    visible_rows = min(visible_rows + 10, MAX_ROWS)
    update_table_visibility()
    if visible_rows == MAX_ROWS:
        button_more_rows.grid_remove()

def update_table_visibility():
    for row in range(MAX_ROWS):
        if row < visible_rows:
            index_labels[row].grid()
            entry_nr_placi[row].grid()
            entry_latime[row].grid()
            entry_lungime[row].grid()
            entry_cost[row].grid()
            label_perimeter[row].grid()
            label_area[row].grid()
            label_cost[row].grid()
        else:
            index_labels[row].grid_remove()
            entry_nr_placi[row].grid_remove()
            entry_latime[row].grid_remove()
            entry_lungime[row].grid_remove()
            entry_cost[row].grid_remove()
            label_perimeter[row].grid_remove()
            label_area[row].grid_remove()
            label_cost[row].grid_remove()
    root.update_idletasks()  # Ensure the scroll region is updated
    canvas.configure(scrollregion=canvas.bbox("all"))

# Create a frame for the scrollbar and the canvas
frame = tk.Frame(root)
frame.pack(fill=tk.BOTH, expand=1)

canvas = tk.Canvas(frame, bg='#A52A2A')
canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=canvas.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

canvas.configure(yscrollcommand=scrollbar.set)
canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

content_frame = tk.Frame(canvas, bg='#A52A2A')
canvas.create_window((0, 0), window=content_frame, anchor="nw")

# Define style for ttk widgets
style = ttk.Style()
style.configure("TLabel", background='#A52A2A', foreground='white')
style.configure("TButton", padding=6, relief="flat", background='#8B4513', foreground='black')

# Header row
headers = ["Index", "Numar Coli", "Lățime", "Lungime", "Cost/m² (EURO)", "Perimetru (m)", "Suprafață (m²)", "Cost Total (EURO)"]
for col, text in enumerate(headers):
    ttk.Label(content_frame, text=text).grid(row=0, column=col, padx=5, pady=5, sticky="W")

# Unit selection
unit_var = tk.StringVar(value="cm")
ttk.Label(content_frame, text="Unitate:").grid(row=1, column=9, padx=5, pady=5, sticky="E")
unit_combobox = ttk.Combobox(content_frame, textvariable=unit_var, values=["m", "cm", "mm"], state="readonly")
unit_combobox.grid(row=2, column=9, padx=5, pady=5, sticky="E")

# Input entries and row setup
entry_nr_placi = []
entry_latime = []
entry_lungime = []
entry_cost = []
label_perimeter = []
label_area = []
label_cost = []
index_labels = []

for row in range(MAX_ROWS):
    index_labels.append(ttk.Label(content_frame, text=str(row + 1)))
    entry_nr_placi.append(ttk.Entry(content_frame))
    entry_latime.append(ttk.Entry(content_frame))
    entry_lungime.append(ttk.Entry(content_frame))
    entry_cost.append(ttk.Entry(content_frame))

    index_labels[-1].grid(row=row + 1, column=0, padx=5, pady=5, sticky="W")
    entry_nr_placi[-1].grid(row=row + 1, column=1, padx=5, pady=5, sticky="W")
    entry_latime[-1].grid(row=row + 1, column=2, padx=5, pady=5, sticky="W")
    entry_lungime[-1].grid(row=row + 1, column=3, padx=5, pady=5, sticky="W")
    entry_cost[-1].grid(row=row + 1, column=4, padx=5, pady=5, sticky="W")
    label_perimeter.append(ttk.Label(content_frame, text="0.000 m"))
    label_area.append(ttk.Label(content_frame, text="0.000 m²"))
    label_cost.append(ttk.Label(content_frame, text="0.00 EURO"))

    label_perimeter[-1].grid(row=row + 1, column=5, padx=5, pady=5, sticky="W")
    label_area[-1].grid(row=row + 1, column=6, padx=5, pady=5, sticky="W")
    label_cost[-1].grid(row=row + 1, column=7, padx=5, pady=5, sticky="W")

    # Bind events to trigger calculation automatically
    entry_nr_placi[-1].bind("<FocusOut>", lambda e, r=row: calculate_row(r))
    entry_latime[-1].bind("<FocusOut>", lambda e, r=row: calculate_row(r))
    entry_lungime[-1].bind("<FocusOut>", lambda e, r=row: calculate_row(r))
    entry_cost[-1].bind("<FocusOut>", lambda e, r=row: calculate_row(r))
    unit_combobox.bind("<<ComboboxSelected>>", lambda e: [calculate_row(r) for r in range(MAX_ROWS)])

# Totals row
label_total_area = ttk.Label(content_frame, text="Total Area: 0.000 m²")
label_total_perimeter = ttk.Label(content_frame, text="Total Perimeter: 0.000 m")
label_total_cost = ttk.Label(content_frame, text="Total Cost: 0.00 EURO")
label_total_area.grid(row=MAX_ROWS + 1, column=6, padx=5, pady=5, sticky="W")
label_total_perimeter.grid(row=MAX_ROWS + 1, column=5, padx=5, pady=5, sticky="W")
label_total_cost.grid(row=MAX_ROWS + 1, column=7, padx=5, pady=5, sticky="W")

# Store Sheet Size Inputs
ttk.Label(content_frame, text="Lungime Foaie (m):").grid(row=1, column=8, padx=5, pady=5, sticky="E")
entry_sheet_length = ttk.Entry(content_frame)
entry_sheet_length.grid(row=2, column=8, padx=5, pady=5, sticky="E")
ttk.Label(content_frame, text="Lățime Foaie (m):").grid(row=3, column=8, padx=5, pady=5, sticky="E")
entry_sheet_width = ttk.Entry(content_frame)
entry_sheet_width.grid(row=4, column=8, padx=5, pady=5, sticky="E")
label_sheets_needed = ttk.Label(content_frame, text="0.0 sheets (0 round up)")
label_sheets_needed.grid(row=5, column=8, padx=5, pady=5, sticky="E")

# Button to show more rows, placed on the right side, not to affect row positioning
visible_rows = INITIAL_VISIBLE_ROWS
button_more_rows = ttk.Button(content_frame, text="Mai Multe Placi", command=show_more_rows)
button_more_rows.grid(row=6, column=8, padx=5, pady=5, sticky="E")

# Save button with sheet calculation
save_button = ttk.Button(content_frame, text="Salvează", command=lambda: [calculate_sheets_needed(), save_to_excel()])
save_button.grid(row=7, column=8, padx=5, pady=5, sticky="E")

# Calculate totals button
calculate_button = ttk.Button(content_frame, text="Calculează Totaluri", command=calculate_totals)
calculate_button.grid(row=8, column=8, padx=5, pady=5, sticky="E")

# Initial visibility update
update_table_visibility()


# Function to move focus with arrow keys
def move_focus(event):
    widget = event.widget
    if event.keysym in ['Return', 'Tab']:
        next_widget = widget.tk_focusNext()
        next_widget.focus()
        return "break"
    elif event.keysym == 'Left':
        prev_widget = widget.tk_focusPrev()
        prev_widget.focus()
        return "break"
    elif event.keysym == 'Right':
        next_widget = widget.tk_focusNext()
        next_widget.focus()
        return "break"
    elif event.keysym == 'Up':
        widget.tk_focusPrev().focus()
        return "break"
    elif event.keysym == 'Down':
        widget.tk_focusNext().focus()
        return "break"


# Bind arrow keys and Enter key to move focus
for entry in entry_nr_placi + entry_latime + entry_lungime + entry_cost:
    entry.bind("<Return>", move_focus)
    entry.bind("<Tab>", move_focus)
    entry.bind("<Left>", move_focus)
    entry.bind("<Right>", move_focus)
    entry.bind("<Up>", move_focus)
    entry.bind("<Down>", move_focus)

# Start the GUI event loop
root.mainloop()

