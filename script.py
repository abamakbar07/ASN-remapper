import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from typing import Dict, List, Tuple, Optional
import warnings
warnings.filterwarnings('ignore')

class ASNTemplateMapper:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("ASN Template Mapper with Lookup")
        self.root.geometry("900x700")
        
        # ASN Template columns with their mapping requirements (1=required, 0=optional/hidden)
        self.asn_template = {
            "Messages": 0,
            "GenericKey": 1,
            "HIDDEN ASN/Receipt:": 0,
            "Item:": 1,
            "Owner": 1,
            "Line #:": 0,
            "Expected Qty:": 1,
            "Received Qty:": 0,
            "UOM:": 0,
            "Pack:": 0,
            "LPN:": 1,
            "Location:": 1,
            "Purchase Order": 0,
            "Hold Code:": 1,
            "HIDDEN Status:": 0,
            "LOTTABLE01": 1,
            "LOTTABLE02": 1,
            "LOTTABLE03": 1,
            "LOTTABLE04": 0,
            "LOTTABLE05": 0,
            "LOTTABLE06": 1,
            "LOTTABLE07": 1,
            "LOTTABLE08": 0,
            "LOTTABLE09": 1,
            "LOTTABLE10": 0,
            "LOTTABLE11": 0,
            "LOTTABLE12": 0,
            "Container Reference:": 0,
            "CUBE": 0,
            "Gross Weight:": 0,
            "Net Weight:": 0,
            "Tare Weight:": 0,
            "Temperature": 0,
            "Received Date:": 0,
            "DISPOSITIONCODE": 0,
            "DISPOSITIONTYPE": 0,
            "Transaction Override Date:": 0,
            "Extended Price:": 0,
            "EXTERNALLOT": 0,
            "External Line #:": 0,
            "External ASN #:": 0,
            "ID": 0,
            "MatchLottable": 0,
            "Notes:": 0,
            "Packing:": 0,
            "POLINENUMBER": 0,
            "QC Auto Adjust:": 0,
            "Inspected:": 0,
            "Rejected:": 0,
            "Reject Reason:": 0,
            "QC Required:": 0,
            "QC Status:": 0,
            "QCUSER": 0,
            "QTYADJUSTED": 0,
            "QTYREJECTED": 0,
            "Reason Code:": 0,
            "RETURNCONDITION": 0,
            "RETURNREASON": 0,
            "RETURNTYPE": 0,
            "RMA Number:": 0,
            "SupplierKey": 0,
            "Ship From Name": 0,
            "UDF1:": 1,
            "UDF2:": 1,
            "UDF3:": 1,
            "UDF4:": 1,
            "UDF5:": 0
        }
        
        self.source_file_path = None
        self.source_columns = []
        self.column_mappings = {}
        self.lookup_files = {}  # Store lookup reference files
        self.multi_select_columns = [
            "LOTTABLE01", "LOTTABLE02", "LOTTABLE03", "LOTTABLE04", "LOTTABLE05",
            "LOTTABLE06", "LOTTABLE07", "LOTTABLE08", "LOTTABLE09", "LOTTABLE10",
            "LOTTABLE11", "LOTTABLE12", "UDF1:", "UDF2:", "UDF3:", "UDF4:", "UDF5:"
        ]
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # File selection frame
        file_frame = ttk.LabelFrame(main_frame, text="Source File", padding="5")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(file_frame, text="Select Source File", 
                  command=self.select_source_file).grid(row=0, column=0, padx=(0, 10))
        
        self.file_label = ttk.Label(file_frame, text="No file selected")
        self.file_label.grid(row=0, column=1, sticky=tk.W)
        
        # Lookup files frame
        lookup_frame = ttk.LabelFrame(main_frame, text="Lookup Reference Files", padding="5")
        lookup_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(lookup_frame, text="Add Lookup File", 
                  command=self.add_lookup_file).grid(row=0, column=0, padx=(0, 10))
        
        ttk.Button(lookup_frame, text="Manage Lookup Files", 
                  command=self.manage_lookup_files).grid(row=0, column=1, padx=(0, 10))
        
        self.lookup_label = ttk.Label(lookup_frame, text="No lookup files loaded")
        self.lookup_label.grid(row=0, column=2, sticky=tk.W, padx=(10, 0))
        
        # Mapping frame
        self.mapping_frame = ttk.LabelFrame(main_frame, text="Column Mapping", padding="5")
        self.mapping_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # Create canvas and scrollbar for mapping
        canvas = tk.Canvas(self.mapping_frame, height=400)
        scrollbar = ttk.Scrollbar(self.mapping_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Control buttons frame
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=3, column=0, columnspan=2, pady=(10, 0))
        
        ttk.Button(control_frame, text="Generate ASN Template", 
                  command=self.generate_template).grid(row=0, column=0, padx=(0, 10))
        
        ttk.Button(control_frame, text="Reset Mappings", 
                  command=self.reset_mappings).grid(row=0, column=1)
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        self.mapping_frame.columnconfigure(0, weight=1)
        self.mapping_frame.rowconfigure(0, weight=1)
    
    def create_manual_input_mapping(self, asn_col):
        """Create manual input field"""
        mapping_info = self.column_mappings[asn_col]
        entry = ttk.Entry(mapping_info['config_frame'], width=30)
        entry.pack(side=tk.LEFT)
        mapping_info['config_widgets']['entry'] = entry

    def add_lookup_file(self):
        """Add a lookup reference file"""
        file_path = filedialog.askopenfilename(
            title="Select Lookup Reference File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls *.xlsb"),
                ("CSV files", "*.csv"),
                ("All supported", "*.xlsx *.xls *.xlsb *.csv"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            # Get a name for this lookup file
            file_name = os.path.basename(file_path)
            lookup_name = tk.simpledialog.askstring(
                "Lookup File Name", 
                f"Enter a name for this lookup file:\n{file_name}",
                initialvalue=os.path.splitext(file_name)[0]
            )
            
            if lookup_name:
                try:
                    # Load the lookup file to get column information
                    lookup_df = self.load_dataframe(file_path)
                    self.lookup_files[lookup_name] = {
                        'path': file_path,
                        'columns': list(lookup_df.columns),
                        'df': lookup_df
                    }
                    self.update_lookup_label()
                    messagebox.showinfo("Success", f"Lookup file '{lookup_name}' added successfully!")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to load lookup file: {str(e)}")
    
    def manage_lookup_files(self):
        """Manage existing lookup files"""
        if not self.lookup_files:
            messagebox.showinfo("Info", "No lookup files loaded.")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Manage Lookup Files")
        dialog.geometry("500x300")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # List of lookup files
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="Loaded Lookup Files:", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 10))
        
        listbox = tk.Listbox(main_frame, height=10)
        for name, info in self.lookup_files.items():
            listbox.insert(tk.END, f"{name} ({len(info['columns'])} columns)")
        listbox.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        def remove_selected():
            selection = listbox.curselection()
            if selection:
                name = list(self.lookup_files.keys())[selection[0]]
                del self.lookup_files[name]
                listbox.delete(selection[0])
                self.update_lookup_label()
        
        ttk.Button(button_frame, text="Remove Selected", command=remove_selected).pack(side=tk.LEFT)
        ttk.Button(button_frame, text="Close", command=dialog.destroy).pack(side=tk.RIGHT)
    
    def update_lookup_label(self):
        """Update the lookup files label"""
        count = len(self.lookup_files)
        if count == 0:
            self.lookup_label.config(text="No lookup files loaded")
        else:
            self.lookup_label.config(text=f"{count} lookup file(s) loaded")
    
    def load_dataframe(self, file_path):
        """Load DataFrame from various file formats"""
        file_ext = os.path.splitext(file_path)[1].lower()
        
        if file_ext == '.csv':
            return pd.read_csv(file_path)
        elif file_ext == '.xls':
            return pd.read_excel(file_path, engine='xlrd')
        elif file_ext == '.xlsb':
            return pd.read_excel(file_path, engine='pyxlsb')
        elif file_ext == '.xlsx':
            return pd.read_excel(file_path, engine='openpyxl')
        else:
            # Try to read as Excel first, then CSV
            try:
                return pd.read_excel(file_path)
            except:
                return pd.read_csv(file_path)
    
    def select_source_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Source File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls *.xlsb"),
                ("CSV files", "*.csv"),
                ("All supported", "*.xlsx *.xls *.xlsb *.csv"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.source_file_path = file_path
            self.file_label.config(text=os.path.basename(file_path))
            self.load_source_columns()
            self.create_mapping_interface()
    
    def load_source_columns(self):
        try:
            df = self.load_dataframe(self.source_file_path)
            self.source_columns = list(df.columns)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read source file: {str(e)}")
            self.source_columns = []
    
    def create_mapping_interface(self):
        # Clear existing mapping widgets
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        
        self.column_mappings = {}
        
        # Header
        ttk.Label(self.scrollable_frame, text="ASN Column", font=("Arial", 10, "bold")).grid(
            row=0, column=0, padx=5, pady=5, sticky=tk.W)
        ttk.Label(self.scrollable_frame, text="Mapping Type", font=("Arial", 10, "bold")).grid(
            row=0, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Label(self.scrollable_frame, text="Configuration", font=("Arial", 10, "bold")).grid(
            row=0, column=2, padx=5, pady=5, sticky=tk.W)
        ttk.Label(self.scrollable_frame, text="Required", font=("Arial", 10, "bold")).grid(
            row=0, column=3, padx=5, pady=5, sticky=tk.W)
        
        # Create separator
        separator = ttk.Separator(self.scrollable_frame, orient="horizontal")
        separator.grid(row=1, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)
        
        # Create mapping rows only for columns with value = 1
        row = 2
        for asn_col, required in self.asn_template.items():
            if required == 1:  # Only show required columns
                # ASN column label
                ttk.Label(self.scrollable_frame, text=asn_col).grid(
                    row=row, column=0, padx=5, pady=2, sticky=tk.W)
                
                # Mapping type dropdown
                mapping_types = ["Direct", "Multi-Select", "Lookup", "Manual Input"]
                type_combo = ttk.Combobox(self.scrollable_frame, values=mapping_types, 
                                         state="readonly", width=15)
                type_combo.set("Direct")
                type_combo.grid(row=row, column=1, padx=5, pady=2, sticky=tk.W)
                
                # Configuration frame
                config_frame = ttk.Frame(self.scrollable_frame)
                config_frame.grid(row=row, column=2, padx=5, pady=2, sticky=tk.W)
                
                # Initialize mapping
                self.column_mappings[asn_col] = {
                    'type_combo': type_combo,
                    'config_frame': config_frame,
                    'current_type': 'Direct',
                    'config_widgets': {}
                }
                
                # Bind type change event
                type_combo.bind('<<ComboboxSelected>>', 
                               lambda e, col=asn_col: self.on_mapping_type_change(col))
                
                # Create initial direct mapping
                self.create_direct_mapping(asn_col)
                
                # Required indicator
                ttk.Label(self.scrollable_frame, text="✓", foreground="red").grid(
                    row=row, column=3, padx=5, pady=2, sticky=tk.W)
                
                row += 1
    
    def on_mapping_type_change(self, asn_col):
        """Handle mapping type change"""
        mapping_info = self.column_mappings[asn_col]
        new_type = mapping_info['type_combo'].get()
        
        if new_type != mapping_info['current_type']:
            # Clear existing configuration
            for widget in mapping_info['config_frame'].winfo_children():
                widget.destroy()
            mapping_info['config_widgets'] = {}
            
            # Create new configuration based on type
            if new_type == "Direct":
                self.create_direct_mapping(asn_col)
            elif new_type == "Multi-Select":
                self.create_multi_select_mapping(asn_col)
            elif new_type == "Lookup":
                self.create_lookup_mapping(asn_col)
            elif new_type == "Manual Input":
                self.create_manual_input_mapping(asn_col)
            
            mapping_info['current_type'] = new_type
    
    def create_direct_mapping(self, asn_col):
        """Create direct column mapping"""
        mapping_info = self.column_mappings[asn_col]
        source_options = ["-- Leave Blank --"] + self.source_columns
        combobox = ttk.Combobox(mapping_info['config_frame'], values=source_options, 
                               state="readonly", width=30)
        combobox.set("-- Leave Blank --")
        combobox.pack(side=tk.LEFT)
        
        mapping_info['config_widgets']['combobox'] = combobox
    
    def create_multi_select_mapping(self, asn_col):
        """Create multi-select mapping with ordered selection"""
        mapping_info = self.column_mappings[asn_col]
        
        # Button to open multi-select dialog
        select_button = ttk.Button(mapping_info['config_frame'], text="Select Multiple...", 
                                 command=lambda: self.open_enhanced_multi_select_dialog(asn_col))
        select_button.pack(side=tk.LEFT, padx=(0, 5))
        
        # Label to show selected columns
        selected_label = ttk.Label(mapping_info['config_frame'], text="None selected", 
                                 foreground="gray", width=40)
        selected_label.pack(side=tk.LEFT)
        
        mapping_info['config_widgets']['button'] = select_button
        mapping_info['config_widgets']['label'] = selected_label
        mapping_info['config_widgets']['selections'] = []
    
    def create_lookup_mapping(self, asn_col):
        """Create lookup mapping configuration"""
        mapping_info = self.column_mappings[asn_col]
        
        # Button to configure lookup
        lookup_button = ttk.Button(mapping_info['config_frame'], text="Configure Lookup...", 
                                 command=lambda: self.open_lookup_config_dialog(asn_col))
        lookup_button.pack(side=tk.LEFT, padx=(0, 5))
        
        # Label to show lookup configuration
        config_label = ttk.Label(mapping_info['config_frame'], text="Not configured", 
                               foreground="gray", width=40)
        config_label.pack(side=tk.LEFT)
        
        mapping_info['config_widgets']['button'] = lookup_button
        mapping_info['config_widgets']['label'] = config_label
        mapping_info['config_widgets']['lookup_config'] = None
    
    def open_enhanced_multi_select_dialog(self, asn_col):
        """Enhanced multi-select dialog with ordered selection"""
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Multi-Select Columns for {asn_col}")
        dialog.geometry("600x500")
        dialog.transient(self.root)
        dialog.grab_set()
        
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Instructions
        ttk.Label(main_frame, text=f"Select columns for {asn_col} (order matters):", 
                 font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 5))
        ttk.Label(main_frame, text="Click columns in order. Selected columns will be combined with '|' separator.", 
                 font=("Arial", 9), foreground="gray").pack(anchor=tk.W, pady=(0, 10))
        
        # Two-column layout
        columns_frame = ttk.Frame(main_frame)
        columns_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Available columns (left side)
        left_frame = ttk.LabelFrame(columns_frame, text="Available Columns", padding="5")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        available_listbox = tk.Listbox(left_frame, height=15)
        available_scroll = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=available_listbox.yview)
        available_listbox.configure(yscrollcommand=available_scroll.set)
        
        for col in self.source_columns:
            available_listbox.insert(tk.END, col)
        
        available_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        available_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Selected columns (right side)
        right_frame = ttk.LabelFrame(columns_frame, text="Selected Columns (in order)", padding="5")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        selected_listbox = tk.Listbox(right_frame, height=15)
        selected_scroll = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=selected_listbox.yview)
        selected_listbox.configure(yscrollcommand=selected_scroll.set)
        
        selected_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        selected_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Middle buttons
        middle_frame = ttk.Frame(columns_frame)
        middle_frame.pack(side=tk.LEFT, padx=5)
        
        # Current selections
        current_selections = self.column_mappings[asn_col]['config_widgets']['selections']
        for col in current_selections:
            selected_listbox.insert(tk.END, col)
        
        def add_selected():
            selection = available_listbox.curselection()
            if selection:
                col_name = self.source_columns[selection[0]]
                if col_name not in [selected_listbox.get(i) for i in range(selected_listbox.size())]:
                    selected_listbox.insert(tk.END, col_name)
        
        def remove_selected():
            selection = selected_listbox.curselection()
            if selection:
                selected_listbox.delete(selection[0])
        
        def move_up():
            selection = selected_listbox.curselection()
            if selection and selection[0] > 0:
                idx = selection[0]
                item = selected_listbox.get(idx)
                selected_listbox.delete(idx)
                selected_listbox.insert(idx - 1, item)
                selected_listbox.selection_set(idx - 1)
        
        def move_down():
            selection = selected_listbox.curselection()
            if selection and selection[0] < selected_listbox.size() - 1:
                idx = selection[0]
                item = selected_listbox.get(idx)
                selected_listbox.delete(idx)
                selected_listbox.insert(idx + 1, item)
                selected_listbox.selection_set(idx + 1)
        
        ttk.Button(middle_frame, text="Add >>", command=add_selected).pack(pady=2)
        ttk.Button(middle_frame, text="<< Remove", command=remove_selected).pack(pady=2)
        ttk.Separator(middle_frame, orient="horizontal").pack(fill=tk.X, pady=10)
        ttk.Button(middle_frame, text="Move Up", command=move_up).pack(pady=2)
        ttk.Button(middle_frame, text="Move Down", command=move_down).pack(pady=2)
        
        # Double-click to add
        available_listbox.bind('<Double-1>', lambda e: add_selected())
        
        # Bottom buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        def save_selections():
            selected_columns = [selected_listbox.get(i) for i in range(selected_listbox.size())]
            self.column_mappings[asn_col]['config_widgets']['selections'] = selected_columns
            
            # Update label
            if selected_columns:
                display_text = " | ".join(selected_columns)
                if len(display_text) > 50:
                    display_text = display_text[:47] + "..."
                self.column_mappings[asn_col]['config_widgets']['label'].config(
                    text=display_text, foreground="black")
            else:
                self.column_mappings[asn_col]['config_widgets']['label'].config(
                    text="None selected", foreground="gray")
            
            dialog.destroy()
        
        def clear_all():
            selected_listbox.delete(0, tk.END)
        
        ttk.Button(button_frame, text="Clear All", command=clear_all).pack(side=tk.LEFT)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="Save", command=save_selections).pack(side=tk.RIGHT, padx=(5, 0))
    
    def open_lookup_config_dialog(self, asn_col):
        """Open lookup configuration dialog"""
        if not self.lookup_files:
            messagebox.showwarning("Warning", "No lookup files loaded. Please add a lookup file first.")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Configure Lookup for {asn_col}")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Instructions
        ttk.Label(main_frame, text=f"Configure lookup for {asn_col}:", 
                 font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 10))
        
        # Lookup file selection
        ttk.Label(main_frame, text="Select Lookup File:").pack(anchor=tk.W, pady=(0, 5))
        lookup_files = list(self.lookup_files.keys())
        file_combo = ttk.Combobox(main_frame, values=lookup_files, state="readonly", width=40)
        if lookup_files:
            file_combo.set(lookup_files[0])
        file_combo.pack(fill=tk.X, pady=(0, 10))
        
        # Source column (for matching)
        ttk.Label(main_frame, text="Source Column (for matching):").pack(anchor=tk.W, pady=(0, 5))
        source_combo = ttk.Combobox(main_frame, values=self.source_columns, state="readonly", width=40)
        source_combo.pack(fill=tk.X, pady=(0, 10))
        
        # Lookup key column
        lookup_key_label = ttk.Label(main_frame, text="Lookup Key Column:")
        lookup_key_label.pack(anchor=tk.W, pady=(0, 5))
        lookup_key_combo = ttk.Combobox(main_frame, state="readonly", width=40)
        lookup_key_combo.pack(fill=tk.X, pady=(0, 10))
        
        # Lookup value column
        lookup_value_label = ttk.Label(main_frame, text="Lookup Value Column:")
        lookup_value_label.pack(anchor=tk.W, pady=(0, 5))
        lookup_value_combo = ttk.Combobox(main_frame, state="readonly", width=40)
        lookup_value_combo.pack(fill=tk.X, pady=(0, 10))
        
        # Update lookup columns when file changes
        def update_lookup_columns(event=None):
            selected_file = file_combo.get()
            if selected_file in self.lookup_files:
                columns = self.lookup_files[selected_file]['columns']
                lookup_key_combo['values'] = columns
                lookup_value_combo['values'] = columns
                if columns:
                    lookup_key_combo.set(columns[0])
                    lookup_value_combo.set(columns[0])
        
        file_combo.bind('<<ComboboxSelected>>', update_lookup_columns)
        update_lookup_columns()  # Initialize
        
        # Load current configuration if exists
        current_config = self.column_mappings[asn_col]['config_widgets'].get('lookup_config')
        if current_config:
            if current_config['lookup_file'] in lookup_files:
                file_combo.set(current_config['lookup_file'])
                update_lookup_columns()
            if current_config['source_column'] in self.source_columns:
                source_combo.set(current_config['source_column'])
            lookup_key_combo.set(current_config.get('lookup_key', ''))
            lookup_value_combo.set(current_config.get('lookup_value', ''))
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(20, 0))
        
        def save_config():
            config = {
                'lookup_file': file_combo.get(),
                'source_column': source_combo.get(),
                'lookup_key': lookup_key_combo.get(),
                'lookup_value': lookup_value_combo.get()
            }
            
            # Validate configuration
            if not all([config['lookup_file'], config['source_column'], 
                       config['lookup_key'], config['lookup_value']]):
                messagebox.showerror("Error", "Please fill in all configuration fields.")
                return
            
            self.column_mappings[asn_col]['config_widgets']['lookup_config'] = config
            
            # Update label
            display_text = f"{config['source_column']} → {config['lookup_file']}.{config['lookup_value']}"
            if len(display_text) > 50:
                display_text = display_text[:47] + "..."
            self.column_mappings[asn_col]['config_widgets']['label'].config(
                text=display_text, foreground="black")
            
            dialog.destroy()
        
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="Save", command=save_config).pack(side=tk.RIGHT)
    
    def perform_lookup(self, source_df, lookup_config):
        """Perform lookup operation"""
        try:
            lookup_file_info = self.lookup_files[lookup_config['lookup_file']]
            lookup_df = lookup_file_info['df']
            
            # Create lookup dictionary
            lookup_dict = dict(zip(
                lookup_df[lookup_config['lookup_key']], 
                lookup_df[lookup_config['lookup_value']]
            ))
            
            # Perform lookup
            source_values = source_df[lookup_config['source_column']]
            result = source_values.map(lookup_dict).fillna('')
            
            return result.values
        except Exception as e:
            messagebox.showerror("Lookup Error", f"Failed to perform lookup: {str(e)}")
            return [""] * len(source_df)
    
    def reset_mappings(self):
        """Reset all mappings to default state"""
        for asn_col, mapping_info in self.column_mappings.items():
            # Reset type to Direct
            mapping_info['type_combo'].set("Direct")
            
            # Clear configuration frame
            for widget in mapping_info['config_frame'].winfo_children():
                widget.destroy()
            mapping_info['config_widgets'] = {}
            
            # Recreate direct mapping
            self.create_direct_mapping(asn_col)
            mapping_info['current_type'] = 'Direct'
    
    def generate_template(self):
        if not self.source_file_path:
            messagebox.showerror("Error", "Please select a source file first.")
            return
        
        # Ask for output file location
        from datetime import datetime
        date_str = datetime.now().strftime("%Y%m%d")
        ref_name = os.path.splitext(os.path.basename(self.source_file_path))[0]
        output_path = os.path.join(os.getcwd(), f"{date_str}_{ref_name}.xlsx")
        
        try:
            # Load source data
            source_df = self.load_dataframe(self.source_file_path)
            
            # Create ASN template DataFrame
            asn_data = {}
            processing_summary = {
                'direct': 0,
                'multi_select': 0,
                'lookup': 0,
                'manual_input': 0,
                'empty': 0
            }
            
            # Process each ASN column
            for asn_col in self.asn_template.keys():
                if asn_col in self.column_mappings:
                    mapping_info = self.column_mappings[asn_col]
                    mapping_type = mapping_info['current_type']
                    
                    if mapping_type == "Direct":
                        # Direct column mapping
                        combobox = mapping_info['config_widgets'].get('combobox')
                        if combobox:
                            selected = combobox.get()
                            if selected != "-- Leave Blank --" and selected in source_df.columns:
                                asn_data[asn_col] = source_df[selected].values
                                processing_summary['direct'] += 1
                            else:
                                asn_data[asn_col] = [""] * len(source_df)
                                processing_summary['empty'] += 1
                        else:
                            asn_data[asn_col] = [""] * len(source_df)
                            processing_summary['empty'] += 1
                    
                    elif mapping_type == "Multi-Select":
                        # Multi-select mapping
                        selected_columns = mapping_info['config_widgets'].get('selections', [])
                        if selected_columns:
                            combined_values = []
                            for index in range(len(source_df)):
                                values = []
                                for source_col in selected_columns:
                                    if source_col in source_df.columns:
                                        value = str(source_df[source_col].iloc[index])
                                        if value and value not in ['nan', 'None', 'NaN']:
                                            values.append(value)
                                combined_values.append("|".join(values) if values else "")
                            asn_data[asn_col] = combined_values
                            processing_summary['multi_select'] += 1
                        else:
                            asn_data[asn_col] = [""] * len(source_df)
                            processing_summary['empty'] += 1
                    
                    elif mapping_type == "Lookup":
                        # Lookup mapping
                        lookup_config = mapping_info['config_widgets'].get('lookup_config')
                        if lookup_config:
                            lookup_values = self.perform_lookup(source_df, lookup_config)
                            asn_data[asn_col] = lookup_values
                            processing_summary['lookup'] += 1
                        else:
                            asn_data[asn_col] = [""] * len(source_df)
                            processing_summary['empty'] += 1
                    
                    elif mapping_type == "Manual Input":
                        entry = mapping_info['config_widgets'].get('entry')
                        if entry:
                            manual_value = entry.get()
                            asn_data[asn_col] = [manual_value] * len(source_df)
                            processing_summary['manual_input'] = processing_summary.get('manual_input', 0) + 1
                        else:
                            asn_data[asn_col] = [""] * len(source_df)
                            processing_summary['empty'] += 1
                    
                    else:
                        asn_data[asn_col] = [""] * len(source_df)
                        processing_summary['empty'] += 1
                else:
                    # Fill with empty values for unmapped columns
                    asn_data[asn_col] = [""] * len(source_df)
                    processing_summary['empty'] += 1
            
            # Create DataFrame and save
            asn_df = pd.DataFrame(asn_data)
            asn_df.to_excel(output_path, index=False)
            
            # Show success message with detailed summary
            total_mapped = processing_summary['direct'] + processing_summary['multi_select'] + processing_summary['lookup']
            messagebox.showinfo("Success", 
                              f"ASN template generated successfully!\n\n"
                              f"Output: {output_path}\n"
                              f"Rows processed: {len(asn_df):,}\n"
                              f"Total ASN columns: {len(self.asn_template)}\n\n"
                              f"Mapping Summary:\n"
                              f"• Direct mappings: {processing_summary['direct']}\n"
                              f"• Multi-select mappings: {processing_summary['multi_select']}\n"
                              f"• Manual input mappings: {processing_summary['manual_input']}\n"
                              f"• Lookup mappings: {processing_summary['lookup']}\n"
                              f"• Empty/unmapped: {processing_summary['empty']}\n"
                              f"• Total mapped: {total_mapped}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate template: {str(e)}")
            import traceback
            print(traceback.format_exc())  # For debugging
    
    def run(self):
        self.root.mainloop()

# Add missing import for simpledialog
import tkinter.simpledialog

# Usage example
if __name__ == "__main__":
    app = ASNTemplateMapper()
    app.run()
