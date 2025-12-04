import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import pandas as pd
from comparator import perform_lookup

files = []
file_names = []
compare_vars = []
lookup_vars = []

###landing page###
def launch_app():

    root = tk.Tk()
    root.title('General Comparator Tool')
    root.geometry('500x500')

    #window title
    title = tk.Label(root, text='General Comparator Tool', font=('Montserrat', 20, 'bold'))
    title.pack(padx=5, pady=5)
        
    instruction = tk.Label(root, text='Add Excel Files and select the designator column (E.G. Customer Reference) and which columns to compare (E.G. Manufacturing Number) using the dropdown buttons.',\
                            font=('Montserrat', 12), wraplength=475, justify='center')
    instruction.pack(padx=5, pady=5)

    #files frame
    file_list_frame = tk.Frame(root)
    file_list_frame.pack(fill='both', expand=True)  

    #buttons
    button_frame = tk.Frame(root)
    button_frame.pack(pady=20)

    addfiles_button = tk.Button(button_frame, font=('Montserrat', 10), text='Add Files', command=lambda: add_files(file_list_frame))
    compare_button = tk.Button(button_frame, font=('Montserrat', 10), text='Compare', command=lambda: compare(file_list_frame))
    addfiles_button.pack(side='left', padx=20)
    compare_button.pack(side='right', padx=20)

    root.mainloop()

def add_files(frame):
    file_paths = filedialog.askopenfilenames(filetypes=[('Excel files', '*.xlsx *.xls')])
    if not file_paths:
        return

    for file_path in file_paths:
        if file_path in files:
            continue  #avoid duplicates

        try:
            df = pd.read_excel(file_path, nrows=1)
            cols = list(df.columns)
        except Exception as e:
            messagebox.showerror('Error', f'Could not read columns from {file_path}: {e}')
            continue

        #append state
        files.append(file_path)

        compare_var = tk.StringVar()
        lookup_var = tk.StringVar()

        file_names.append(os.path.basename(file_path))
        compare_vars.append(compare_var)
        lookup_vars.append(lookup_var)

        i = len(files) - 1
        row_base = i * 3

        #file label
        tk.Label(frame, text=os.path.basename(file_path), font=('Montserrat', 10, 'bold')).grid(
            row=row_base, column=0, columnspan=2, sticky='w', padx=10, pady=(10, 0)
        )

        #lookup dropdown with label
        tk.Label(frame, text='Lookup column:', font=('Montserrat', 9)).grid(
            row=row_base + 1, column=0, sticky='w', padx=10, pady=(5, 0)
        )
        lookup_dropdown = ttk.Combobox(frame, values=cols, textvariable=lookup_var, width=50, state='readonly')
        lookup_dropdown.grid(row=row_base + 1, column=1, padx=10, pady=(5, 0))
        lookup_dropdown.set('Select lookup column')
        lookup_dropdown.configure(foreground='gray')

        # comparison dropdown with label
        tk.Label(frame, text='Comparison column:', font=('Montserrat', 9)).grid(
            row=row_base + 2, column=0, sticky='w', padx=10, pady=(5, 10)
        )
        compare_dropdown = ttk.Combobox(frame, values=cols, textvariable=compare_var, width=50, state='readonly')
        compare_dropdown.grid(row=row_base + 2, column=1, padx=10, pady=(5, 10))
        compare_dropdown.set('Select column for comparison')
        compare_dropdown.configure(foreground='gray')

        #bind events to simulate placeholder behavior
        def on_click(event, box, var, placeholder):
            if box.get() == placeholder:
                box.set('')
                box.configure(foreground='black')

        def on_select(event, box):
            box.configure(foreground='black')

        lookup_dropdown.bind('<Button-1>', lambda e, b=lookup_dropdown, v=lookup_var: on_click(e, b, v, 'Select designator column'))
        lookup_dropdown.bind('<<ComboboxSelected>>', lambda e, b=lookup_dropdown: on_select(e, b))

        compare_dropdown.bind('<Button-1>', lambda e, b=compare_dropdown, v=compare_var: on_click(e, b, v, 'Select column for comparison'))
        compare_dropdown.bind('<<ComboboxSelected>>', lambda e, b=compare_dropdown: on_select(e, b))


def compare(frame):
    if len(files) < 2:
        messagebox.showerror('Error', 'Select at least 2 Excel files.')
        return
    
    file_column_map = {}
    for i, file in enumerate(files):
        file_name = file_names[i]

        lookup = lookup_vars[i].get()        
        if not lookup:
            messagebox.showerror('Error', f'No designator column selected for {os.path.basename(file)}.')
            return
        compare = compare_vars[i].get()
        if not compare:
            messagebox.showerror('Error', f'No comparison column selected for {os.path.basename(file)}.')
            return        
        file_column_map[file] = [file_name,(lookup, compare)]

    try:
        mismatches, result_df = perform_lookup(file_column_map)
        output_path = os.path.join(os.path.dirname(files[0]), "comparison_result.xlsx")
        result_df.to_excel(output_path, index=False)
        messagebox.showinfo('Done', f'{mismatches} Mismatches. \nResult saved to:{output_path}')

    except Exception as e:
        messagebox.showerror('Error', str(e))