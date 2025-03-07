'''
update march 5, 25
category to diff sections
need to work on ui/ux

'''


import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import os
import win32print
import re
import win32api
from ttkwidgets import CheckboxTreeview
import threading

PRINT_CONFIG = {
    "master": {
        "folder": r'X:\\',
        "prefix": "",
        "suffix": "",
        "extension": [".pdf", ".jpg"]
    },
    "qc": {
        "folder": r'S:\QCLAB\QC PDF Drawing',
        "prefix": "QC",
        "suffix": "",
        "extension": ".pdf"
    },
    "quality alert": {
        "folder": r'S:\QCLAB\Quality Alert',
        "prefix": "",
        "suffix": "",
        "extension": ".pdf"
    },
    "sample approval": {
        "folder": r'S:\QCLAB\Sample Approval',
        "prefix": "",
        "suffix": "",
        "extension": ".pdf"
    },
    "packing": {
        "folder": r'S:\QCLAB\Packing',
        "prefix": "",
        "suffix": "",
        "extension": [".pdf", ".jpg"]
    },
    "blue book": {
        "folder": r'S:\QCLAB\Blue Book',
        "prefix": "",
        "suffix": "",
        "extension": ".docx"
    },
    "set-up sheets": {
        "folder": r'S:\Set-up Sheets',
        "prefix": "",
        "suffix": "",
        "extension": ".pdf"
    }
}

class PrintApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Bluebook Printer - Vision Profile Extrusions - Made by Duc Mai")
        self.geometry("800x600")
        self.configure(bg="#f0f0f0")
        
        #self.check_vars = []  # Stores (file_path, BooleanVar) tuples
        # UI Elements
        self.create_widgets()
        self.center_window()
        
    def center_window(self):
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'+{x}+{y}')

    '''def create_widgets(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        # Main container
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Input section
        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(input_frame, text="Drawing Number:", font=('Arial', 11)).pack(side=tk.LEFT)
        self.entry_number = ttk.Entry(input_frame, width=25, font=('Arial', 11))
        self.entry_number.pack(side=tk.LEFT, padx=5)
        
        # Action buttons
        self.btn_search = ttk.Button(input_frame, text="Search Files", command=self.start_search)
        self.btn_search.pack(side=tk.LEFT, padx=5)
        
        self.btn_print_selected = ttk.Button(input_frame, 
                                            text="Print Selected", 
                                            command=self.print_selected,
                                            state=tk.DISABLED)
        self.btn_print_selected.pack(side=tk.LEFT, padx=5)

        # Initialize selection frame (fix for AttributeError)
        self.selection_frame = ttk.LabelFrame(main_frame, text="Found Files", padding=10)
        self.selection_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # Add "Select All" checkbox to the top-right corner of selection frame
        self.select_all_var = tk.BooleanVar(value=False)  # Variable to track "Select All" state
        self.select_all_checkbox = ttk.Checkbutton(
            self.selection_frame,
            text="Select All",
            variable=self.select_all_var,
            command=self.toggle_select_all  # Function to toggle all checkboxes
        )
        self.select_all_checkbox.pack(side=tk.TOP, anchor=tk.NE, padx=5, pady=5)

        # Canvas and scrollbar for scrollable area
        self.canvas = tk.Canvas(self.selection_frame, borderwidth=0)
        self.scrollbar = ttk.Scrollbar(self.selection_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(
            scrollregion=self.canvas.bbox("all")
        ))
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # Progress and log sections (existing components)
        self.progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, mode='determinate')
        self.progress.pack(fill=tk.X, pady=10)
        
        log_frame = ttk.Frame(main_frame)
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.log_console = scrolledtext.ScrolledText(
            log_frame, wrap=tk.WORD, font=('Consolas', 10),
            bg='#ffffff', padx=10, pady=10
        )
        self.log_console.pack(fill=tk.BOTH, expand=True)'''

    def create_widgets(self):
        style = ttk.Style()
        style.theme_use('clam')

        # Main container
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Input section
        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=5)

        ttk.Label(input_frame, text="Drawing Number:", font=('Arial', 11)).pack(side=tk.LEFT)
        self.entry_number = ttk.Entry(input_frame, width=25, font=('Arial', 11))
        self.entry_number.pack(side=tk.LEFT, padx=5)

        # Action buttons
        self.btn_search = ttk.Button(input_frame, text="Search Files", command=self.start_search)
        self.btn_search.pack(side=tk.LEFT, padx=5)

        self.btn_print_selected = ttk.Button(input_frame, 
                                            text="Print Selected", 
                                            command=self.print_selected,
                                            state=tk.DISABLED)
        self.btn_print_selected.pack(side=tk.LEFT, padx=5)

        # Initialize selection frame (fix for AttributeError)
        self.selection_frame = ttk.LabelFrame(main_frame, text="Found Files", padding=10)
        self.selection_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # Add "Select All" checkbox to the top-right corner of selection frame
        self.select_all_var = tk.BooleanVar(value=False)  # Variable to track "Select All" state
        self.select_all_checkbox = ttk.Checkbutton(
            self.selection_frame,
            text="Select All",
            variable=self.select_all_var,
            command=self.toggle_select_all  # Function to toggle all checkboxes
        )
        self.select_all_checkbox.pack(side=tk.TOP, anchor=tk.NE, padx=5, pady=5)

        # Canvas and scrollbar for scrollable area
        self.canvas = tk.Canvas(self.selection_frame, borderwidth=0)
        self.scrollbar = ttk.Scrollbar(self.selection_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(
            scrollregion=self.canvas.bbox("all")
        ))

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Pack canvas with expand=True and fill=BOTH to allow dynamic resizing
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # Progress and log sections (existing components)
        self.progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, mode='determinate')
        self.progress.pack(fill=tk.X, pady=10)

        log_frame = ttk.Frame(main_frame)
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.log_console = scrolledtext.ScrolledText(
            log_frame, wrap=tk.WORD, font=('Consolas', 10),
            bg='#ffffff', padx=10, pady=10
        )
        self.log_console.pack(fill=tk.BOTH, expand=True)


    

    def create_selection_ui(self, categorized_files):
        # Clear previous selection UI
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        self.check_vars = {}  # Reset checkbox variables

        # Create a CheckboxTreeview for collapsible categories and checkboxes
        self.tree = CheckboxTreeview(self.scrollable_frame, show="tree")  # Use self.tree
        self.tree.pack(fill=tk.BOTH, expand=True)


        # Configure the treeview columns to expand and fill the available space
        self.tree.column("#0", width=5000, stretch=True)  # Adjust width as needed
        # Add categories and files to the Treeview
        for category, files in categorized_files.items():
            if not files:
                continue

            # Add a parent node for the category (collapsed by default)
            parent_node = self.tree.insert("", "end", text=category.upper(), open=False)

            for file_path in files:
                file_name = os.path.basename(file_path)

                # Add child nodes for each file under the category
                child_node = self.tree.insert(parent_node, "end", text=file_name)

                # Attach a BooleanVar to each file for selection tracking
                var = tk.BooleanVar(value=False)  # Default unchecked
                self.check_vars[child_node] = (file_path, var)

                # Set initial checkbox state in the Treeview
                self.tree.change_state(child_node, "unchecked")

        # Handle no results case
        if not any(len(files) > 0 for files in categorized_files.values()):
            ttk.Label(self.scrollable_frame,
                    text="No matching files found in any category",
                    foreground="gray").pack(pady=10)

        # Bind checkbox toggle event to update button state
        def on_check(event):
            item_id = self.tree.identify_row(event.y)
            if item_id in self.check_vars:
                _, var = self.check_vars[item_id]
                new_state = not var.get()
                var.set(new_state)  # Toggle BooleanVar state
                self.tree.change_state(item_id, "checked" if new_state else "unchecked")
            self.update_print_button_state()

        self.tree.bind("<Button-1>", on_check)
        # Bind right-click event to show file menu
        self.tree.bind("<Button-3>", self.show_file_menu)
        self.update_print_button_state()  # Initial check for button state

   

    def print_selected(self):
        if not self.check_vars:
            messagebox.showwarning("Print Error", "No files selected for printing")
            return

        self.btn_print_selected.config(state=tk.DISABLED)
        total_selected = sum(var.get() for _, var in self.check_vars.values())
        printed_count = 0

        for _, (file_path, var) in self.check_vars.items():
            if var.get():  # Only print selected files
                if self.print_file(file_path):
                    printed_count += 1
                    self.log_message(f"Printed: {file_path}", 'success')
                else:
                    self.log_message(f"Failed: {file_path}", 'error')

        # Log summary of printing results
        self.log_message("\n=== Selective Printing Summary ===")
        self.log_message(f"Successfully printed {printed_count}/{total_selected} files")
        self.btn_print_selected.config(state=tk.NORMAL)

    def show_file_menu(self, event):
        item_id = self.tree.identify_row(event.y)
        if item_id in self.check_vars:
            file_path, _ = self.check_vars[item_id]
            self.tree.selection_set(item_id)
            file_menu = tk.Menu(self, tearoff=0)
            file_menu.add_command(label="Open File", command=lambda: os.startfile(file_path))
            file_menu.add_command(label="Open File Location", command=lambda: os.startfile(os.path.dirname(file_path)))
            file_menu.tk_popup(event.x_root, event.y_root)




    def toggle_select_all(self):
        select_all_state = self.select_all_var.get()  # Get current state of "Select All"

        for item_id, (file_path, var) in self.check_vars.items():
            var.set(select_all_state)  # Set each checkbox's BooleanVar
            self.tree.change_state(item_id, "checked" if select_all_state else "unchecked")  # Update visual state

        self.update_print_button_state()  # Update the button state



    def start_search(self):
        number = self.entry_number.get().strip()
        if not number:
            messagebox.showwarning("Input Error", "Please enter a drawing number")
            return

        # Disable the search button during processing
        self.btn_search.config(state=tk.DISABLED)
        self.log_message(f"Searching for files containing '{number}'...")

        # Start the search in a separate thread
        search_thread = threading.Thread(target=self.perform_search, args=(number,))
        search_thread.start()
    
    def perform_search(self, number):
        categorized_files = {category: [] for category in PRINT_CONFIG.keys()}
        missing_files = []  # List to track missing files

        for category, config in PRINT_CONFIG.items():
            extensions = config['extension'] if isinstance(config['extension'], list) else [config['extension']]
            found_in_category = False  # Track if any file is found in this category

            for ext in extensions:
                if not isinstance(ext, str):
                    self.log_message(f"Invalid extension type {ext} for {category}", 'error')
                    continue

                for root, dirs, files in os.walk(config['folder']):
                    for file in files:
                        # Check if the file matches the search number and extension
                        if re.search(rf'(?<!\d){re.escape(number)}(?!\d)', file) and file.endswith(ext):
                            categorized_files[category].append(os.path.join(root, file))
                            found_in_category = True

            # If no file was found for this category, mark it as missing
            if not found_in_category:
                missing_files.append((category, config["folder"], number))

        # Update the UI after completing the search
        self.after(0, self.update_ui_after_search, categorized_files, missing_files)

    def update_ui_after_search(self, categorized_files, missing_files):
        # Re-enable the search button
        self.btn_search.config(state=tk.NORMAL)

        # Update the UI with found files
        self.create_selection_ui(categorized_files)

        # Log found files
        found_categories = sum(len(v) > 0 for v in categorized_files.values())
        self.log_message(f"Found files containing '{self.entry_number.get()}' across {found_categories} categories")

        # Log missing files
        if missing_files:
            self.log_message("\n=== Missing Files ===", 'warning')
            for category, folder, num in missing_files:
                self.log_message(f"- {category.capitalize()}: No file with '{num}' found in {folder}", 'error')



    def update_print_button_state(self):
        """Enable or disable the 'Print Selected' button based on checkbox states."""
        any_selected = any(var.get() for _, var in self.check_vars.values())  # Correctly access values
        self.btn_print_selected.config(state=tk.NORMAL if any_selected else tk.DISABLED)


    def log_message(self, message, tag=None):
            self.log_console.configure(state='normal')
            self.log_console.insert(tk.END, message + '\n', tag)
            self.log_console.configure(state='disabled')
            self.log_console.see(tk.END)

    def find_file(self, root_folder, partial_number, extension):
            matches = []
            for root, dirs, files in os.walk(root_folder):
                for file in files:
                    # Check if number is in filename AND extension matches
                    if partial_number in file and file.endswith(extension):
                        matches.append(os.path.join(root, file))
            return matches


    def print_file(self, file_path):
            if file_path and os.path.exists(file_path):
                default_printer = win32print.GetDefaultPrinter()
                win32api.ShellExecute(
                    0,
                    "print",
                    file_path,
                    f'/d:"{default_printer}"',
                    ".",
                    0
                )
                
                
                print(f"Printing {file_path} on {default_printer}")
                return True
            return False


if __name__ == "__main__":
    app = PrintApp()
    app.mainloop()