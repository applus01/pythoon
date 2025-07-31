import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Menu
import os
import threading
import subprocess
import platform
from pathlib import Path
import json
import webbrowser
from datetime import datetime
import shutil
import string

class NetworkFileExplorer:
    def __init__(self, root):
        self.root = root
        self.root.title("Network Drive File Explorer")
        self.root.geometry("1200x800")
        self.root.minsize(800, 600)
        
        # File extension categories
        self.file_categories = {
            'Documents': ['.doc', '.docx', '.pdf', '.txt', '.rtf', '.odt'],
            'Presentations': ['.ppt', '.pptx', '.odp'],
            'Spreadsheets': ['.xls', '.xlsx', '.csv', '.ods'],
            'Images': ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.tif', '.svg', '.webp'],
            'CAD Files': ['.dwg', '.dxf', '.dwf', '.dgn'],
            'Archives': ['.zip', '.rar', '.7z', '.tar', '.gz'],
            'Videos': ['.mp4', '.avi', '.mkv', '.mov', '.wmv', '.flv'],
            'Audio': ['.mp3', '.wav', '.flac', '.aac', '.ogg'],
            'Code': ['.py', '.js', '.html', '.css', '.cpp', '.java', '.c']
        }
        
        # Flatten extensions for quick lookup
        self.all_extensions = set()
        for exts in self.file_categories.values():
            self.all_extensions.update(exts)
        
        self.filtered_files = []
        self.scanning = False
        self.diagnosis_results = {}
        
        self.setup_gui()
        self.populate_drives()
        self.load_settings()
    
    def setup_gui(self):
        # Menu bar
        self.create_menu()
        
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # Path selection frame
        path_frame = ttk.LabelFrame(main_frame, text="Drive/Folder Selection", padding="5")
        path_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        path_frame.columnconfigure(2, weight=1)
        
        # Drive selection dropdown
        ttk.Label(path_frame, text="Quick Select:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        
        self.drive_var = tk.StringVar()
        self.drive_combo = ttk.Combobox(path_frame, textvariable=self.drive_var, width=25, state="readonly")
        self.drive_combo.grid(row=0, column=1, sticky=tk.W, padx=(0, 10))
        self.drive_combo.bind('<<ComboboxSelected>>', self.on_drive_selected)
        
        # Path entry with folder icon
        ttk.Label(path_frame, text="Selected Path:").grid(row=1, column=0, sticky=tk.W, padx=(0, 5), pady=(5, 0))
        
        # Create the folder icon button
        folder_button = tk.Button(path_frame, text="📁", font=("Arial", 12), 
                                relief="flat", bd=0, bg="white", fg="blue",
                                command=self.browse_folder, cursor="hand2",
                                width=3, height=1)
        folder_button.grid(row=1, column=1, sticky=tk.W, padx=(0, 5), pady=(5, 0))
        
        # Path entry
        self.path_var = tk.StringVar()
        self.path_var.trace('w', self.on_path_changed)
        self.path_entry = ttk.Entry(path_frame, textvariable=self.path_var, width=50)
        self.path_entry.grid(row=1, column=2, sticky=(tk.W, tk.E), padx=(0, 5), pady=(5, 0))
        
        # Action buttons frame
        buttons_frame = ttk.Frame(path_frame)
        buttons_frame.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=(10, 0))
        
        ttk.Button(buttons_frame, text="📁 Browse Folder", command=self.browse_folder).pack(side=tk.LEFT, padx=(0, 10))
        
        self.scan_folder_button = ttk.Button(buttons_frame, text="🔍 Scan Folder", command=self.diagnose_and_scan)
        self.scan_folder_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.quick_scan_button = ttk.Button(buttons_frame, text="⚡ Quick Scan", command=self.start_scan)
        self.quick_scan_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.stop_button = ttk.Button(buttons_frame, text="⏹ Stop", command=self.stop_scan, state="disabled")
        self.stop_button.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(buttons_frame, text="🔄 Refresh Drives", command=self.populate_drives).pack(side=tk.LEFT)
        
        # Filter frame
        filter_frame = ttk.LabelFrame(main_frame, text="File Filters", padding="5")
        filter_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # Filter checkboxes
        self.filter_vars = {}
        row = 0
        for category, extensions in self.file_categories.items():
            var = tk.BooleanVar(value=True)
            self.filter_vars[category] = var
            
            cb = ttk.Checkbutton(filter_frame, text=f"{category} ({', '.join(extensions)})", 
                               variable=var, command=self.apply_filters)
            cb.grid(row=row, column=0, sticky=tk.W, pady=1)
            row += 1
        
        # Select/Deselect all buttons
        button_frame = ttk.Frame(filter_frame)
        button_frame.grid(row=row, column=0, sticky=tk.W, pady=(10, 0))
        
        ttk.Button(button_frame, text="Select All", command=self.select_all_filters).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="Deselect All", command=self.deselect_all_filters).pack(side=tk.LEFT)
        
        # Results frame
        results_frame = ttk.LabelFrame(main_frame, text="Files Found", padding="5")
        results_frame.grid(row=0, column=1, rowspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(10, 0))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(2, weight=1)
        
        # Tab control for folders and files
        self.notebook = ttk.Notebook(results_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 5))
        
        # Folders tab
        folders_frame = ttk.Frame(self.notebook)
        self.notebook.add(folders_frame, text="📁 Folders")
        
        # Folders treeview
        self.folders_tree = ttk.Treeview(folders_frame, show="tree", height=10)
        folders_scroll_y = ttk.Scrollbar(folders_frame, orient="vertical", command=self.folders_tree.yview)
        folders_scroll_x = ttk.Scrollbar(folders_frame, orient="horizontal", command=self.folders_tree.xview)
        self.folders_tree.configure(yscrollcommand=folders_scroll_y.set, xscrollcommand=folders_scroll_x.set)
        
        self.folders_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        folders_scroll_y.grid(row=0, column=1, sticky=(tk.N, tk.S))
        folders_scroll_x.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        folders_frame.columnconfigure(0, weight=1)
        folders_frame.rowconfigure(0, weight=1)
        
        # Files tab
        files_frame = ttk.Frame(self.notebook)
        self.notebook.add(files_frame, text="📄 Files")
        
        # Search box
        search_frame = ttk.Frame(files_frame)
        search_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        search_frame.columnconfigure(1, weight=1)
        
        ttk.Label(search_frame, text="Search:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.filter_by_search)
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        search_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 5))
        
        ttk.Button(search_frame, text="Clear", command=self.clear_search).grid(row=0, column=2)
        
        # Treeview for file list
        columns = ("Name", "Type", "Size", "Modified", "Path")
        self.tree = ttk.Treeview(files_frame, columns=columns, show="tree headings", height=15)
        
        # Configure columns
        self.tree.heading("#0", text="", anchor=tk.W)
        self.tree.column("#0", width=20, minwidth=20)
        
        for col in columns:
            self.tree.heading(col, text=col, anchor=tk.W)
            if col == "Name":
                self.tree.column(col, width=200, minwidth=100)
            elif col == "Type":
                self.tree.column(col, width=80, minwidth=60)
            elif col == "Size":
                self.tree.column(col, width=80, minwidth=60)
            elif col == "Modified":
                self.tree.column(col, width=120, minwidth=100)
            else:
                self.tree.column(col, width=300, minwidth=200)
        
        # Scrollbars for files treeview
        tree_scroll_y = ttk.Scrollbar(files_frame, orient="vertical", command=self.tree.yview)
        tree_scroll_x = ttk.Scrollbar(files_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
        
        self.tree.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        tree_scroll_y.grid(row=1, column=1, sticky=(tk.N, tk.S))
        tree_scroll_x.grid(row=2, column=0, sticky=(tk.W, tk.E))
        
        files_frame.columnconfigure(0, weight=1)
        files_frame.rowconfigure(1, weight=1)
        
        # Bind events
        self.tree.bind("<Double-1>", self.open_selected_file)
        self.tree.bind("<Button-3>", self.show_context_menu)
        self.folders_tree.bind("<Double-1>", self.on_folder_double_click)
        self.folders_tree.bind("<Button-3>", self.show_folder_context_menu)
        
        # Status frame
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        status_frame.columnconfigure(1, weight=1)
        
        # Progress bar
        self.progress = ttk.Progressbar(status_frame, mode='indeterminate')
        self.progress.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        # Status label
        self.status_var = tk.StringVar(value="Ready")
        status_label = ttk.Label(status_frame, textvariable=self.status_var)
        status_label.grid(row=1, column=0, sticky=tk.W)
        
        # File count label
        self.count_var = tk.StringVar(value="Files: 0")
        count_label = ttk.Label(status_frame, textvariable=self.count_var)
        count_label.grid(row=1, column=1, sticky=tk.E)
    
    def create_menu(self):
        menubar = Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Export Results...", command=self.export_results)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        # View menu
        view_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="View", menu=view_menu)
        view_menu.add_command(label="Refresh", command=self.refresh_results)
        view_menu.add_command(label="Clear Results", command=self.clear_results)
        
        # Tools menu
        tools_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        tools_menu.add_command(label="🔄 Refresh Drives", command=self.populate_drives)
        tools_menu.add_command(label="📁 Open File Location", command=self.open_file_location)
        tools_menu.add_command(label="📋 Copy File Path", command=self.copy_file_path)
        tools_menu.add_separator()
        tools_menu.add_command(label="🌐 Browse Network", command=self.browse_network)
        
        # Help menu
        help_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)
    
    def populate_drives(self):
        drives = []
        
        try:
            if platform.system() == "Windows":
                # Get local drives
                for letter in string.ascii_uppercase:
                    drive = f"{letter}:\\"
                    if os.path.exists(drive):
                        try:
                            total, used, free = shutil.disk_usage(drive)
                            drive_type = self.get_drive_type(drive)
                            size_info = self.format_file_size(total)
                            drives.append(f"{drive} ({drive_type} - {size_info})")
                        except:
                            drives.append(f"{drive} (Drive)")
                
                # Try to get network drives
                try:
                    result = subprocess.run(['net', 'use'], capture_output=True, text=True, timeout=5)
                    if result.returncode == 0:
                        lines = result.stdout.split('\n')
                        for line in lines:
                            if ':' in line and '\\\\' in line:
                                parts = line.strip().split()
                                if len(parts) >= 2 and parts[1].startswith('\\\\'):
                                    drives.append(f"{parts[0]} → {parts[1]} (Network)")
                except:
                    pass
                
                drives.extend([
                    "\\\\localhost (Local Network)",
                    "Browse Network Locations..."
                ])
            
            else:
                drives.extend([
                    "/ (Root)",
                    "/home (Home)",
                    "/mnt (Mount Points)",
                    "/media (Media)",
                    "Browse Folders..."
                ])
        
        except Exception as e:
            drives = ["Browse Folders..."]
        
        self.drive_combo['values'] = drives
        if drives:
            self.drive_combo.set(drives[0])
    
    def get_drive_type(self, drive):
        try:
            if platform.system() == "Windows":
                import ctypes
                drive_type = ctypes.windll.kernel32.GetDriveTypeW(drive)
                types = {1: "Unknown", 2: "Removable", 3: "Local", 4: "Network", 5: "CD-ROM", 6: "RAM"}
                return types.get(drive_type, "Unknown")
        except:
            pass
        return "Local"
    
    def on_drive_selected(self, event=None):
        selected = self.drive_var.get()
        if not selected:
            return
        
        if "Browse" in selected:
            self.browse_folder()
        elif "Network Locations" in selected:
            self.browse_network()
        else:
            if "→" in selected:
                drive_path = selected.split("→")[0].strip()
            else:
                drive_path = selected.split("(")[0].strip()
            
            if os.path.exists(drive_path):
                self.path_var.set(drive_path)
    
    def on_path_changed(self, *args):
        path = self.path_var.get().strip()
        
        if path and os.path.exists(path):
            self.scan_folder_button.config(state="normal")
            self.quick_scan_button.config(state="normal")
            
            if self.is_network_path(path):
                self.scan_folder_button.config(text="🌐 Scan Network Folder")
                self.status_var.set(f"Network path selected: {path}")
            else:
                self.scan_folder_button.config(text="🔍 Scan Folder")
                self.status_var.set(f"Local path selected: {path}")
            
            # Auto-load folders when path changes
            self.load_folders(path)
        else:
            self.scan_folder_button.config(state="normal")
            self.quick_scan_button.config(state="normal")
            if path:
                self.status_var.set("Path may be invalid - click scan to verify")
            else:
                self.status_var.set("Ready - select a folder or enter path manually")
    
    def load_folders(self, path):
        """Load and display folders in the selected path"""
        # Clear existing folders
        for item in self.folders_tree.get_children():
            self.folders_tree.delete(item)
        
        if not path or not os.path.exists(path):
            return
        
        # Start folder loading in background
        self.status_var.set("Loading folders...")
        self.progress.start()
        
        folder_thread = threading.Thread(target=self.scan_folders_background, args=(path,))
        folder_thread.daemon = True
        folder_thread.start()
    
    def scan_folders_background(self, path):
        """Background thread to scan folders - optimized for speed"""
        folders = []
        folder_count = 0
        
        try:
            # Get immediate subdirectories only (no deep scanning)
            items = os.listdir(path)
            
            for item in items:
                item_path = os.path.join(path, item)
                try:
                    # Quick check if it's a directory
                    if os.path.isdir(item_path):
                        # Quick folder info without deep scanning
                        try:
                            stat_info = os.stat(item_path)
                            modified = datetime.fromtimestamp(stat_info.st_mtime)
                        except (OSError, PermissionError):
                            modified = datetime.now()
                        
                        # Quick count of immediate contents only
                        subfolders = 0
                        files = 0
                        try:
                            # Use faster os.scandir instead of os.listdir
                            with os.scandir(item_path) as entries:
                                for entry in entries:
                                    if entry.is_dir(follow_symlinks=False):
                                        subfolders += 1
                                    else:
                                        files += 1
                                    
                                    # Limit counting to avoid slowdown
                                    if subfolders + files > 100:
                                        subfolders = f"{subfolders}+"
                                        files = f"{files}+"
                                        break
                        except (PermissionError, OSError):
                            subfolders = "?"
                            files = "?"
                        
                        folders.append({
                            'name': item,
                            'path': item_path,
                            'modified': modified,
                            'subfolders': subfolders,
                            'files': files
                        })
                        folder_count += 1
                        
                        # Update UI more frequently for better responsiveness
                        if folder_count % 5 == 0:
                            self.root.after(0, lambda c=folder_count: self.status_var.set(f"Loading folders... Found {c}"))
                
                except (PermissionError, OSError, FileNotFoundError):
                    # Skip items we can't access
                    continue
            
            # Sort folders by name (case-insensitive)
            folders.sort(key=lambda x: x['name'].lower())
            
        except Exception as e:
            self.root.after(0, lambda: self.status_var.set(f"Error loading folders: {str(e)}"))
            return
        
        # Update UI on main thread
        self.root.after(0, lambda: self.folders_loaded(folders, folder_count))
    
    def folders_loaded(self, folders, folder_count):
        """Handle completion of folder loading"""
        self.progress.stop()
        
        # Add folders to tree with better performance
        for folder in folders:
            # Create display text with folder info
            folder_text = f"📁 {folder['name']}"
            
            # Format subfolder/file counts
            if isinstance(folder['subfolders'], str) and isinstance(folder['files'], str):
                if folder['subfolders'] != "?" and folder['files'] != "?":
                    folder_text += f" ({folder['subfolders']} folders, {folder['files']} files)"
                elif folder['subfolders'] != "?" or folder['files'] != "?":
                    folder_text += f" (Access limited)"
                else:
                    folder_text += f" (Access denied)"
            else:
                folder_text += f" ({folder['subfolders']} folders, {folder['files']} files)"
            
            # Insert folder into tree
            item_id = self.folders_tree.insert("", "end", text=folder_text, 
                                             values=(folder['path'],), open=False)
            
            # Only add expandable marker if folder has subfolders and we can access them
            if (isinstance(folder['subfolders'], int) and folder['subfolders'] > 0) or \
               (isinstance(folder['subfolders'], str) and folder['subfolders'].endswith('+')):
                # Add lazy loading capability - don't add dummy child for now
                pass
        
        # Update status
        path_type = "Network" if self.is_network_path(self.path_var.get()) else "Local"
        self.status_var.set(f"{path_type} folders loaded: {folder_count} folders found")
        
        # Update notebook tab title
        self.notebook.tab(0, text=f"📁 Folders ({folder_count})")
    
    def on_folder_double_click(self, event):
        """Handle double-click on folder to navigate or scan"""
        selection = self.folders_tree.selection()
        if not selection:
            return
        
        item = selection[0]
        folder_text = self.folders_tree.item(item)['text']
        
        # Get folder path
        values = self.folders_tree.item(item)['values']
        if values:
            folder_path = values[0]
            
            # Show quick action dialog instead of context menu
            self.show_folder_quick_actions(folder_path)
    
    def show_folder_quick_actions(self, folder_path):
        """Show quick action dialog for selected folder"""
        dialog = tk.Toplevel(self.root)
        dialog.title("📁 Folder Actions")
        dialog.geometry("400x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (200)
        y = (dialog.winfo_screenheight() // 2) - (100)
        dialog.geometry(f"400x200+{x}+{y}")
        
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill="both", expand=True)
        
        # Folder name label
        folder_name = os.path.basename(folder_path)
        ttk.Label(main_frame, text=f"📁 {folder_name}", 
                 font=("Arial", 12, "bold")).pack(pady=(0, 10))
        
        ttk.Label(main_frame, text=folder_path, 
                 font=("Arial", 8), foreground="gray").pack(pady=(0, 15))
        
        # Action buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=10)
        
        ttk.Button(button_frame, text="📂 Open in Explorer", 
                  command=lambda: [dialog.destroy(), self.open_folder_in_explorer(folder_path)]).pack(fill="x", pady=2)
        
        ttk.Button(button_frame, text="🔍 Scan This Folder", 
                  command=lambda: [dialog.destroy(), self.scan_specific_folder(folder_path)]).pack(fill="x", pady=2)
        
        ttk.Button(button_frame, text="📁 Navigate Here", 
                  command=lambda: [dialog.destroy(), self.navigate_to_folder(folder_path)]).pack(fill="x", pady=2)
        
        # Close button
        ttk.Button(main_frame, text="❌ Close", 
                  command=dialog.destroy).pack(pady=(15, 0))
    
    def show_folder_action_menu(self, event, folder_path):
        """Show action menu for selected folder"""
        action_menu = Menu(self.root, tearoff=0)
        action_menu.add_command(label="📂 Open Folder", 
                              command=lambda: self.open_folder_in_explorer(folder_path))
        action_menu.add_command(label="🔍 Scan This Folder", 
                              command=lambda: self.scan_specific_folder(folder_path))
        action_menu.add_command(label="📁 Navigate to This Folder", 
                              command=lambda: self.navigate_to_folder(folder_path))
        action_menu.add_separator()
        action_menu.add_command(label="📋 Copy Path", 
                              command=lambda: self.copy_folder_path(folder_path))
        
        action_menu.post(event.x_root, event.y_root)
    
    def show_folder_context_menu(self, event):
        """Show context menu for folder tree"""
        selection = self.folders_tree.selection()
        if not selection:
            return
        
        item = selection[0]
        if self.folders_tree.item(item)['text'] == "Loading...":
            return
        
        values = self.folders_tree.item(item)['values']
        if values:
            folder_path = values[0]
            self.show_folder_action_menu(event, folder_path)
    
    def open_folder_in_explorer(self, folder_path):
        """Open folder in system file explorer"""
        try:
            if platform.system() == "Windows":
                subprocess.run(f'explorer "{folder_path}"')
            elif platform.system() == "Darwin":
                subprocess.run(["open", folder_path])
            else:
                subprocess.run(["xdg-open", folder_path])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open folder: {str(e)}")
    
    def scan_specific_folder(self, folder_path):
        """Scan a specific folder"""
        self.path_var.set(folder_path)
        self.start_scan()
    
    def navigate_to_folder(self, folder_path):
        """Navigate to a specific folder and load its subfolders"""
        self.path_var.set(folder_path)
        # The on_path_changed will automatically trigger folder loading
    
    def copy_folder_path(self, folder_path):
        """Copy folder path to clipboard"""
        self.root.clipboard_clear()
        self.root.clipboard_append(folder_path)
        self.status_var.set("Folder path copied to clipboard")
    
    def is_network_path(self, path):
        return (path.startswith('\\\\') or 
                (len(path) >= 2 and path[1] == ':' and self.get_drive_type(path[:3]) == "Network"))
    
    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select Network Drive or Folder")
        if folder:
            self.path_var.set(folder)
    
    def browse_network(self):
        if platform.system() == "Windows":
            try:
                subprocess.run(['explorer', 'shell:NetworkPlacesFolder'], check=True)
            except:
                messagebox.showinfo("Info", "Please use the Browse button to select a network location manually.")
        else:
            self.browse_folder()
    
    def diagnose_and_scan(self):
        path = self.path_var.get().strip()
        if not path:
            self.browse_folder()
            path = self.path_var.get().strip()
            if not path:
                return
        
        if not os.path.exists(path):
            messagebox.showerror("Error", f"The path '{path}' does not exist or is not accessible.")
            return
        
        self.scanning = True
        self.stop_button.config(state="normal")
        self.scan_folder_button.config(state="disabled")
        self.quick_scan_button.config(state="disabled")
        self.progress.start()
        self.status_var.set("Diagnosing folder structure...")
        
        self.diagnosis_thread = threading.Thread(target=self.perform_diagnosis, args=(path,))
        self.diagnosis_thread.daemon = True
        self.diagnosis_thread.start()
    
    def perform_diagnosis(self, path):
        diagnosis = {
            'path': path,
            'is_network': self.is_network_path(path),
            'accessible': True,
            'total_folders': 0,
            'total_files': 0,
            'file_types': {},
            'large_folders': [],
            'errors': [],
            'estimated_scan_time': 0
        }
        
        try:
            self.root.after(0, lambda: self.status_var.set("Analyzing folder structure..."))
            
            folder_count = 0
            file_count = 0
            file_types = {}
            start_time = datetime.now()
            
            sample_folders = 0
            max_sample = 10
            
            for root, dirs, files in os.walk(path):
                if not self.scanning:
                    break
                
                folder_count += 1
                file_count += len(files)
                
                for file in files:
                    ext = os.path.splitext(file)[1].lower()
                    if ext in self.all_extensions:
                        file_types[ext] = file_types.get(ext, 0) + 1
                
                if len(files) > 100:
                    diagnosis['large_folders'].append({
                        'path': root,
                        'file_count': len(files)
                    })
                
                sample_folders += 1
                if sample_folders >= max_sample:
                    elapsed = (datetime.now() - start_time).total_seconds()
                    if elapsed > 0:
                        estimated_total = (folder_count / sample_folders) * elapsed
                        diagnosis['estimated_scan_time'] = estimated_total
                    break
                
                if folder_count % 50 == 0:
                    self.root.after(0, lambda: self.status_var.set(f"Analyzed {folder_count} folders, {file_count} files..."))
            
            diagnosis['total_folders'] = folder_count
            diagnosis['total_files'] = file_count
            diagnosis['file_types'] = file_types
            
        except Exception as e:
            diagnosis['accessible'] = False
            diagnosis['errors'].append(str(e))
        
        self.root.after(0, lambda: self.diagnosis_complete(diagnosis))
    
    def diagnosis_complete(self, diagnosis):
        self.scanning = False
        self.stop_button.config(state="disabled")
        self.scan_folder_button.config(state="normal")
        self.quick_scan_button.config(state="normal")
        self.progress.stop()
        
        self.diagnosis_results = diagnosis
        self.show_diagnosis_results(diagnosis)
    
    def show_diagnosis_results(self, diagnosis):
        dialog = tk.Toplevel(self.root)
        dialog.title("📊 Folder Diagnosis Results")
        dialog.geometry("600x500")
        dialog.transient(self.root)
        dialog.grab_set()
        
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (600 // 2)
        y = (dialog.winfo_screenheight() // 2) - (500 // 2)
        dialog.geometry(f"600x500+{x}+{y}")
        
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        results_text = tk.Text(text_frame, wrap="word", height=20, width=70)
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=results_text.yview)
        results_text.configure(yscrollcommand=scrollbar.set)
        
        results_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        report = self.generate_diagnosis_report(diagnosis)
        results_text.insert("1.0", report)
        results_text.config(state="disabled")
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x")
        
        if diagnosis['accessible']:
            ttk.Button(button_frame, text="🔍 Start Full Scan", 
                      command=lambda: [dialog.destroy(), self.start_scan()]).pack(side="left", padx=(0, 5))
            
            ttk.Button(button_frame, text="⚡ Quick Scan (Filter Applied)", 
                      command=lambda: [dialog.destroy(), self.start_scan()]).pack(side="left", padx=(0, 5))
        
        ttk.Button(button_frame, text="📋 Copy Report", 
                  command=lambda: self.copy_diagnosis_report(report)).pack(side="left", padx=(0, 5))
        
        ttk.Button(button_frame, text="❌ Close", command=dialog.destroy).pack(side="right")
    
    def generate_diagnosis_report(self, diagnosis):
        report = f"""🔍 FOLDER DIAGNOSIS REPORT
{'='*50}

📁 Path: {diagnosis['path']}
🌐 Network Location: {'Yes' if diagnosis['is_network'] else 'No'}
✅ Accessible: {'Yes' if diagnosis['accessible'] else 'No'}

📊 SUMMARY STATISTICS
{'='*25}
📂 Total Folders: {diagnosis['total_folders']:,}
📄 Total Files: {diagnosis['total_files']:,}
⏱️ Estimated Scan Time: {diagnosis['estimated_scan_time']:.1f} seconds

🎯 FILE TYPES FOUND
{'='*20}"""
        
        if diagnosis['file_types']:
            for ext, count in sorted(diagnosis['file_types'].items(), key=lambda x: x[1], reverse=True):
                category = self.get_file_category(ext)
                report += f"\n{ext.upper():>6} files: {count:>6,} ({category})"
        else:
            report += "\nNo matching file types found in sample."
        
        if diagnosis['large_folders']:
            report += f"\n\n📁 LARGE FOLDERS (>100 files)\n{'='*30}"
            for folder in diagnosis['large_folders'][:10]:
                report += f"\n{folder['file_count']:>4} files: {folder['path']}"
            if len(diagnosis['large_folders']) > 10:
                report += f"\n... and {len(diagnosis['large_folders']) - 10} more folders"
        
        if diagnosis['errors']:
            report += f"\n\n⚠️ ERRORS ENCOUNTERED\n{'='*20}"
            for error in diagnosis['errors']:
                report += f"\n• {error}"
        
        report += f"\n\n💡 RECOMMENDATIONS\n{'='*18}"
        
        if diagnosis['estimated_scan_time'] > 60:
            report += f"\n• ⏰ Large folder detected - scan may take {diagnosis['estimated_scan_time']/60:.1f} minutes"
            report += "\n• 💡 Consider using Quick Scan with specific file type filters"
        
        if diagnosis['is_network']:
            report += "\n• 🌐 Network location - ensure stable connection during scan"
            report += "\n• 🔄 Network scan may be slower than local scan"
        
        if len(diagnosis['large_folders']) > 5:
            report += "\n• 📁 Many large folders detected - consider scanning subfolders individually"
        
        if not diagnosis['file_types']:
            report += "\n• ❓ No matching file types found in sample - check your file filters"
        
        report += f"\n\n📅 Diagnosis completed: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        
        return report
    
    def copy_diagnosis_report(self, report):
        self.root.clipboard_clear()
        self.root.clipboard_append(report)
        messagebox.showinfo("Copied", "Diagnosis report copied to clipboard!")
    
    def start_scan(self):
        path = self.path_var.get().strip()
        if not path:
            self.browse_folder()
            path = self.path_var.get().strip()
            if not path:
                return
        
        if not os.path.exists(path):
            messagebox.showerror("Error", f"The path '{path}' does not exist or is not accessible.")
            return
        
        self.scanning = True
        self.stop_button.config(state="normal")
        self.scan_folder_button.config(state="disabled")
        self.quick_scan_button.config(state="disabled")
        self.progress.start()
        self.status_var.set("Scanning files...")
        self.clear_results()
        
        self.scan_thread = threading.Thread(target=self.scan_files, args=(path,))
        self.scan_thread.daemon = True
        self.scan_thread.start()
    
    def stop_scan(self):
        self.scanning = False
        self.stop_button.config(state="disabled")
        self.scan_folder_button.config(state="normal")
        self.quick_scan_button.config(state="normal")
        self.progress.stop()
        self.status_var.set("Scan stopped by user")
    
    def scan_files(self, root_path):
        found_files = []
        
        try:
            for root, dirs, files in os.walk(root_path):
                if not self.scanning:
                    break
                
                self.root.after(0, lambda: self.status_var.set(f"Scanning: {root}"))
                
                for file in files:
                    if not self.scanning:
                        break
                    
                    file_path = os.path.join(root, file)
                    file_ext = os.path.splitext(file)[1].lower()
                    
                    if file_ext in self.all_extensions:
                        try:
                            stat = os.stat(file_path)
                            file_info = {
                                'name': file,
                                'path': file_path,
                                'ext': file_ext,
                                'size': stat.st_size,
                                'modified': datetime.fromtimestamp(stat.st_mtime),
                                'category': self.get_file_category(file_ext)
                            }
                            found_files.append(file_info)
                        except (OSError, IOError):
                            continue
        
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", f"Error during scan: {str(e)}"))
        
        self.root.after(0, lambda: self.scan_complete(found_files))
    
    def scan_complete(self, files):
        self.scanning = False
        self.stop_button.config(state="disabled")
        self.scan_folder_button.config(state="normal")
        self.quick_scan_button.config(state="normal")
        self.progress.stop()
        
        self.filtered_files = files
        self.apply_filters()
        
        scan_type = "Network" if self.is_network_path(self.path_var.get()) else "Local"
        self.status_var.set(f"{scan_type} scan complete. Found {len(files)} matching files.")
        
        # Update files tab title with count
        self.notebook.tab(1, text=f"📄 Files ({len(files)})")
        
        # Switch to files tab to show results
        self.notebook.select(1)
    
    def get_file_category(self, extension):
        for category, extensions in self.file_categories.items():
            if extension in extensions:
                return category
        return "Other"
    
    def apply_filters(self):
        selected_categories = [cat for cat, var in self.filter_vars.items() if var.get()]
        
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        displayed_files = []
        search_term = self.search_var.get().lower()
        
        for file_info in self.filtered_files:
            if file_info['category'] not in selected_categories:
                continue
            
            if search_term and search_term not in file_info['name'].lower():
                continue
            
            displayed_files.append(file_info)
            self.add_file_to_tree(file_info)
        
        self.count_var.set(f"Files: {len(displayed_files)}")
    
    def add_file_to_tree(self, file_info):
        size_str = self.format_file_size(file_info['size'])
        mod_str = file_info['modified'].strftime("%Y-%m-%d %H:%M")
        
        self.tree.insert("", "end", values=(
            file_info['name'],
            file_info['category'],
            size_str,
            mod_str,
            file_info['path']
        ))
    
    def format_file_size(self, size_bytes):
        if size_bytes == 0:
            return "0 B"
        
        size_names = ["B", "KB", "MB", "GB", "TB"]
        i = 0
        while size_bytes >= 1024 and i < len(size_names) - 1:
            size_bytes /= 1024.0
            i += 1
        
        return f"{size_bytes:.1f} {size_names[i]}"
    
    def filter_by_search(self, *args):
        self.apply_filters()
    
    def clear_search(self):
        self.search_var.set("")
    
    def select_all_filters(self):
        for var in self.filter_vars.values():
            var.set(True)
        self.apply_filters()
    
    def deselect_all_filters(self):
        for var in self.filter_vars.values():
            var.set(False)
        self.apply_filters()
    
    def create_context_menu(self):
        self.context_menu = Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="📂 Open File", command=self.open_selected_file)
        self.context_menu.add_command(label="📁 Open File Location", command=self.open_file_location)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="📋 Copy File Path", command=self.copy_file_path)
        self.context_menu.add_command(label="📝 Copy File Name", command=self.copy_file_name)
    
    def show_context_menu(self, event):
        if not hasattr(self, 'context_menu'):
            self.create_context_menu()
        
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)
    
    def open_selected_file(self, event=None):
        selection = self.tree.selection()
        if not selection:
            return
        
        item = selection[0]
        file_path = self.tree.item(item)['values'][4]
        
        try:
            if platform.system() == "Windows":
                os.startfile(file_path)
            elif platform.system() == "Darwin":
                subprocess.run(["open", file_path])
            else:
                subprocess.run(["xdg-open", file_path])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file: {str(e)}")
    
    def open_file_location(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a file first.")
            return
        
        item = selection[0]
        file_path = self.tree.item(item)['values'][4]
        folder_path = os.path.dirname(file_path)
        
        try:
            if platform.system() == "Windows":
                subprocess.run(f'explorer /select,"{file_path}"')
            elif platform.system() == "Darwin":
                subprocess.run(["open", "-R", file_path])
            else:
                subprocess.run(["xdg-open", folder_path])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file location: {str(e)}")
    
    def copy_file_path(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a file first.")
            return
        
        item = selection[0]
        file_path = self.tree.item(item)['values'][4]
        self.root.clipboard_clear()
        self.root.clipboard_append(file_path)
        self.status_var.set("File path copied to clipboard")
    
    def copy_file_name(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a file first.")
            return
        
        item = selection[0]
        file_name = self.tree.item(item)['values'][0]
        self.root.clipboard_clear()
        self.root.clipboard_append(file_name)
        self.status_var.set("File name copied to clipboard")
    
    def export_results(self):
        if not self.tree.get_children():
            messagebox.showwarning("Warning", "No results to export.")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="Export Results",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, 'w', newline='', encoding='utf-8') as f:
                    if file_path.endswith('.csv'):
                        import csv
                        writer = csv.writer(f)
                        writer.writerow(["Name", "Type", "Size", "Modified", "Path"])
                        
                        for item in self.tree.get_children():
                            values = self.tree.item(item)['values']
                            writer.writerow(values)
                    else:
                        f.write("Name\tType\tSize\tModified\tPath\n")
                        for item in self.tree.get_children():
                            values = self.tree.item(item)['values']
                            f.write("\t".join(str(v) for v in values) + "\n")
                
                messagebox.showinfo("Success", f"Results exported to {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Could not export results: {str(e)}")
    
    def refresh_results(self):
        if self.path_var.get():
            self.start_scan()
    
    def clear_results(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for item in self.folders_tree.get_children():
            self.folders_tree.delete(item)
        self.filtered_files = []
        self.count_var.set("Files: 0")
        self.notebook.tab(0, text="📁 Folders")
        self.notebook.tab(1, text="📄 Files")
    
    def load_settings(self):
        settings_file = "file_explorer_settings.json"
        try:
            if os.path.exists(settings_file):
                with open(settings_file, 'r') as f:
                    settings = json.load(f)
                    if 'last_path' in settings:
                        self.path_var.set(settings['last_path'])
        except:
            pass
    
    def save_settings(self):
        settings_file = "file_explorer_settings.json"
        try:
            settings = {'last_path': self.path_var.get()}
            with open(settings_file, 'w') as f:
                json.dump(settings, f)
        except:
            pass
    
    def show_about(self):
        about_text = """Network Drive File Explorer
Version 1.0

A powerful tool for exploring and filtering files across network drives.

Features:
• Scan network drives and local folders
• Filter files by category and extension
• Search functionality
• File size and modification date display
• Export results to CSV/text
• Double-click to open files
• Right-click context menu

Supported file types:
• Documents (DOC, PDF, TXT, etc.)
• Presentations (PPT, etc.)
• Spreadsheets (XLS, CSV, etc.)
• Images (JPG, PNG, TIF, etc.)
• CAD files (DWG, DXF, etc.)
• Archives, Videos, Audio, Code files

Created with Python and Tkinter"""
        
        messagebox.showinfo("About", about_text)
    
    def on_closing(self):
        self.scanning = False
        self.save_settings()
        self.root.destroy()

def main():
    root = tk.Tk()
    app = NetworkFileExplorer(root)
    
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    try:
        root.iconbitmap("icon.ico")
    except:
        pass
    
    root.mainloop()

if __name__ == "__main__":
    main()