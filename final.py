'''
v2.0
Update December 3, 25
packing printer now includes jpeg, png, and docx files
quality alert printer now includes docx files


--------
v1.9 
Update December 1, 25
add Enter key to trigger search
Now printing packing files to include docx files

--------
v1.8 
update sept 17, 25
change path to new folders 

--------
v1.7
update june 10, 25
add new folder for Truc's recipes
reworked UI/UX
--------


update march 8, 25
fixed "FAILED TO REMOVE FOLDER" after closing program.

--------
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
import atexit
import shutil
import sys

if getattr(sys, 'frozen', False):
    current_meipass_dir = getattr(sys, '_MEIPASS', None)

    def cleanup_meipass_dir():
        try:
            # Only attempt to remove our own PyInstaller temp dir
            if (
                current_meipass_dir
                and os.path.isdir(current_meipass_dir)
                and os.path.basename(current_meipass_dir).startswith('MEI')
            ):
                shutil.rmtree(current_meipass_dir, ignore_errors=True)
        except Exception as cleanup_error:
            print(f"Failed to delete PyInstaller temp directory: {cleanup_error}")

    atexit.register(cleanup_meipass_dir)

PRINTCONFIG = {
    "master": {"folder": r'X:\\', "prefix": "", "suffix": "", "extension": [".pdf", ".jpg"]},
    "qc": {"folder": r'S:\QCLAB\Bluebooks Related\QC PDF Drawing', "prefix": "QC", "suffix": "", "extension": ".pdf"},
    "quality alert": {"folder": r'S:\QCLAB\Bluebooks Related\Quality Alert', "prefix": "", "suffix": "", "extension": [".pdf", ".docx"]},
    "sample approval": {"folder": r'S:\QCLAB\Bluebooks Related\Sample Approval', "prefix": "", "suffix": "", "extension": ".pdf"},
    "packing": {"folder": r'S:\QCLAB\Bluebooks Related\Packing', "prefix": "", "suffix": "", "extension": [".pdf", ".jpg", ".jpeg", ".png", ".docx"]},
    "blue book": {"folder": r'S:\QCLAB\Bluebooks Related\Blue Book', "prefix": "", "suffix": "", "extension": ".docx"},
    "Rein-Punches" : {"folder": r'S:\QCLAB\Customers\Customer Punches - Reinforcement', "prefix": "", "suffix": "", "extension": ".pdf"},
    "Fit-Functions" : {"folder": r'S:\QCLAB\Bluebooks Related\Fit-Functions', "prefix": "", "suffix": "", "extension": ".pdf"},
    "Weatherstrip Details" : {"folder": r'S:\QCLAB\Bluebooks Related\Weatherstrip Details per Company', "prefix": "", "suffix": "", "extension": ".pdf"},
    "set-up sheets": {"folder": r'S:\Set-up Sheets', "prefix": "", "suffix": "", "extension": ".pdf"},
    "setup sheets TRUC": {"folder": r'S:\Truc Dau\RECIPES\RECIPES', "prefix": "SETUP-SHEET-", "suffix": "", "extension": ".pdf"}
}

class PrintApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Bluebooks Printer - Vision Profile Extrusions - Made by Duc Mai - v2.0")
        self.geometry("800x600")
        self.configure(bg="#f0f0f0")
        self.check_vars = {}
        self.create_widgets()
        self.center_window()

    def center_window(self):
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    def create_widgets(self):
        style = ttk.Style()
        style.theme_use("clam")

        # Main container
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Input section
        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=5)
        ttk.Label(input_frame, text="Drawing Number", font=("Arial", 11)).pack(side=tk.LEFT)
        self.entry_number = ttk.Entry(input_frame, width=25, font=("Arial", 11))
        self.entry_number.pack(side=tk.LEFT, padx=5)
        self.entry_number.bind("<Return>", lambda event: self.start_search())  # Trigger search on Enter key
        self.btn_search = ttk.Button(input_frame, text="Search Files", command=self.start_search)
        self.btn_search.pack(side=tk.LEFT, padx=5)
        self.btn_print_selected = ttk.Button(input_frame, text="Print Selected", command=self.print_selected, state=tk.DISABLED)
        self.btn_print_selected.pack(side=tk.LEFT, padx=5)

        # Use PanedWindow to split vertically
        paned = ttk.PanedWindow(main_frame, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True)

        self.selection_frame = ttk.LabelFrame(paned, text="Found Files", padding=5)
        paned.add(self.selection_frame, weight=3)  # Give more weight to found files section

        # Select All checkbox
        self.select_all_var = tk.BooleanVar(value=False)
        self.select_all_checkbox = ttk.Checkbutton(
            self.selection_frame, text="Select All", variable=self.select_all_var, command=self.toggle_select_all
        )
        self.select_all_checkbox.pack(side=tk.TOP, anchor=tk.NE, padx=5, pady=5)

        # Canvas and scrollbar for scrollable area
        self.canvas = tk.Canvas(self.selection_frame, borderwidth=0)
        self.scrollbar = ttk.Scrollbar(self.selection_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Bind the canvas <Configure> event to resize the frame to match canvas size
        def resize_frame(event):
            self.canvas.itemconfig(self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw"), width=event.width, height=event.height)

        self.canvas.bind("<Configure>", resize_frame)

        self.canvas.pack(side="left", fill="both", expand=True)

        # Log section with less weight
        log_frame = ttk.Frame(paned)
        self.log_console = scrolledtext.ScrolledText(
            log_frame, wrap=tk.WORD, font=('Consolas', 10), bg='#ffffff', padx=10, pady=10
        )
        self.log_console.pack(fill=tk.BOTH, expand=True)
        paned.add(log_frame, weight=1)

        # Progress bar
        self.progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, mode='determinate')
        self.progress.pack(fill=tk.X, pady=10)

    def create_selection_ui(self, categorized_files):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.check_vars = {}  # Reset checkbox variables

        # Create a CheckboxTreeview with two columns: 'File Name' and 'Source Path'
        self.tree = CheckboxTreeview(self.scrollable_frame, columns=("path",), show="tree headings")
        self.tree.heading("#0", text="File Name")
        self.tree.heading("path", text="Source Path")
        self.tree.column("#0", width=200, stretch=True)
        self.tree.column("path", width=450, stretch=True)
        self.tree.pack(fill=tk.BOTH, expand=True)

        # Add categories and files to the Treeview
        for category, files in categorized_files.items():
            if not files:
                continue
            parent_node = self.tree.insert("", "end", text=category.upper(), open=False)
            for file_path in files:
                file_name = os.path.basename(file_path)
                child_node = self.tree.insert(parent_node, "end", text=file_name, values=(file_path,))
                var = tk.BooleanVar(value=False)
                self.check_vars[child_node] = (file_path, var)
                self.tree.change_state(child_node, "unchecked")

        if not any(len(files) > 0 for files in categorized_files.values()):
            ttk.Label(self.scrollable_frame, text="No matching files found in any category", foreground="gray").pack(pady=10)

        def on_check(event):
            item_id = self.tree.identify_row(event.y)
            if item_id in self.check_vars:
                _, var = self.check_vars[item_id]
                new_state = not var.get()
                var.set(new_state)
                self.tree.change_state(item_id, "checked" if new_state else "unchecked")
                self.update_print_button_state()
        self.tree.bind("<Button-1>", on_check)
        self.tree.bind("<Button-3>", self.show_file_menu)
        self.update_print_button_state()

    def print_selected(self):
        if not self.check_vars:
            messagebox.showwarning("Print Error", "No files selected for printing")
            return
        self.btn_print_selected.config(state=tk.DISABLED)
        total_selected = sum(var.get() for _, var in self.check_vars.values())
        printed_count = 0
        for _, (file_path, var) in self.check_vars.items():
            if var.get():
                if self.print_file(file_path):
                    printed_count += 1
                    self.log_message(f"Printed: {file_path}", "success")
                else:
                    self.log_message(f"Failed: {file_path}", "error")
        self.log_message("\nSelective Printing Summary:")
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
        select_all_state = self.select_all_var.get()
        for item_id, (_, var) in self.check_vars.items():
            var.set(select_all_state)
            self.tree.change_state(item_id, "checked" if select_all_state else "unchecked")
        self.update_print_button_state()

    def start_search(self):
        number = self.entry_number.get().strip()
        if not number:
            messagebox.showwarning("Input Error", "Please enter a drawing number")
            return
        self.btn_search.config(state=tk.DISABLED)
        self.log_message(f"Searching for files containing: {number}...")
        search_thread = threading.Thread(target=self.perform_search, args=(number,))
        search_thread.start()

    def perform_search(self, number):
        categorized_files = {category: [] for category in PRINTCONFIG.keys()}
        missing_files = []
        for category, config in PRINTCONFIG.items():
            extensions = config["extension"] if isinstance(config["extension"], list) else [config["extension"]]
            found_in_category = False
            prefix = config.get("prefix", "")
            for ext in extensions:
                if not isinstance(ext, str):
                    self.log_message(f"Invalid extension type: {ext} for {category}", "error")
                    continue
                for root, dirs, files in os.walk(config["folder"]):
                    for file in files:
                        if (
                            (not prefix or file.startswith(prefix)) and
                            re.search(rf"{re.escape(number)}", file) and
                            file.endswith(ext)
                        ):
                            categorized_files[category].append(os.path.join(root, file))
                            found_in_category = True
            if not found_in_category:
                missing_files.append((category, config["folder"], number))
        self.after(0, self.update_ui_after_search, categorized_files, missing_files)

    def update_ui_after_search(self, categorized_files, missing_files):
        self.btn_search.config(state=tk.NORMAL)
        self.create_selection_ui(categorized_files)
        found_categories = sum(len(v) > 0 for v in categorized_files.values())
        self.log_message(f"Found files containing {self.entry_number.get()} across {found_categories} categories")
        if missing_files:
            self.log_message("\nMissing Files:", "warning")
            for category, folder, num in missing_files:
                self.log_message(f"- {category.capitalize()}: No file with {num} found in {folder}", "error")

    def update_print_button_state(self):
        any_selected = any(var.get() for _, var in self.check_vars.values())
        self.btn_print_selected.config(state=tk.NORMAL if any_selected else tk.DISABLED)

    def log_message(self, message, tag=None):
        self.log_console.configure(state="normal")
        self.log_console.insert(tk.END, message + "\n", tag)
        self.log_console.configure(state="disabled")
        self.log_console.see(tk.END)

    def find_file(self, root_folder, partial_number, extension):
        matches = []
        for root, dirs, files in os.walk(root_folder):
            for file in files:
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

