#!/usr/bin/env python3
"""
Comic Metadata Bulk Editor - Python GUI Version
A tool to edit ComicInfo.xml metadata in CBZ/CBR files

Version 2.31: Autovoluming Re-added
- FIX: Re-integrated the "Autovoluming" and "Manual Sorting" features 
  that were accidentally removed in the v2.30 layout update.
- Maintained: "View Metadata" button and clear 3-column button layout.
- Maintained: Robust merge logic and all ComicRack standard options.
"""

import os
import sys 
import zipfile
import shutil
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Optional
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import re 
import time
import datetime 
import json 

try:
    import rarfile
    RARFILE_AVAILABLE = True
except ImportError:
    RARFILE_AVAILABLE = False
    print("Warning: rarfile not available. CBR support disabled.")
    print("Install with: pip install rarfile")
    
try:
    import winsound
    WINSOUND_AVAILABLE = True
except ImportError:
    WINSOUND_AVAILABLE = False


class ToolTip:
    """Creates a tooltip for a given widget."""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        self.widget.bind("<Enter>", self.show_tip)
        self.widget.bind("<Leave>", self.hide_tip)
    
    def show_tip(self, event=None):
        if self.tip_window or not self.text:
            return
        
        # Position the tooltip relative to the widget
        x = self.widget.winfo_rootx() + self.widget.winfo_width()
        y = self.widget.winfo_rooty() + self.widget.winfo_height()
        
        self.tip_window = tk.Toplevel(self.widget)
        self.tip_window.wm_overrideredirect(True) 
        self.tip_window.wm_geometry(f"+{x}+{y}")
        
        label = tk.Label(self.tip_window, text=self.text, justify=tk.LEFT,
                         background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                         font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)
        
    def hide_tip(self, event=None):
        if self.tip_window:
            self.tip_window.destroy()
        self.tip_window = None


class ComicMetadataEditor:
    """Handles reading and writing ComicInfo.xml in comic archives"""
    
    # Extensive mapping of internal keys to XML tags for ComicInfo v2.1
    FIELD_MAPPING = {
        'title': 'Title', 'series': 'Series', 'volume': 'Volume', 'number': 'Number',
        'issuecount': 'IssueCount', 'year': 'Year', 'month': 'Month', 'day': 'Day',
        'alternateSeries': 'AlternateSeries', 'alternatenumber': 'AlternateNumber',
        'alternateissuecount': 'AlternateIssueCount', 'storyArc': 'StoryArc', 
        'seriesgroup': 'SeriesGroup', 'seriescomplete': 'SeriesComplete', 
        'volume_count': 'VolumeCount', 'format': 'Format', 'agerating': 'AgeRating', 
        'manga': 'Manga', 'publisher': 'Publisher', 'imprint': 'Imprint', 
        'blackandwhite': 'BlackAndWhite', 'language': 'LanguageISO', 'genre': 'Genre', 
        'tags': 'Tags', 'writer': 'Writer', 'penciller': 'Penciller', 'inker': 'Inker',
        'colorist': 'Colorist', 'letterer': 'Letterer', 'coverartist': 'CoverArtist', 
        'editor': 'Editor', 'authorsort': 'AuthorSort', 'summary': 'Summary', 
        'maincharacter': 'MainCharacterOrTeam', 'characters': 'Characters', 'teams': 'Teams', 
        'locations': 'Locations', 'notes': 'Notes', 'review': 'Review',
        'scaninformation': 'ScanInformation', 'web': 'Web',
        'communityrating': 'CommunityRating', 'gtin': 'GTIN', 'read': 'Read',
        'country': 'Country', 'pages': 'PageCount', 'isfolder': 'IsFolder'
    }
    
    # Reverse mapping for reading XML
    REVERSE_MAPPING = {v: k for k, v in FIELD_MAPPING.items()}
    
    def __init__(self, file_path: str):
        self.file_path = Path(file_path)
        self.is_cbz = self.file_path.suffix.lower() == '.cbz'
        self.is_cbr = self.file_path.suffix.lower() == '.cbr'
        
        if not (self.is_cbz or self.is_cbr):
            raise ValueError("File must be .cbz or .cbr")

    def _read_xml_from_archive(self) -> Optional[str]:
        """Reads ComicInfo.xml content from the archive."""
        if self.is_cbz:
            try:
                with zipfile.ZipFile(self.file_path, 'r') as zf:
                    xml_name = next((f for f in zf.namelist() if f.lower().endswith('comicinfo.xml')), None)
                    if xml_name:
                        return zf.read(xml_name).decode('utf-8')
            except Exception as e:
                pass
        elif self.is_cbr and RARFILE_AVAILABLE:
            try:
                with rarfile.RarFile(self.file_path, 'r') as rf:
                    xml_name = next((f for f in rf.namelist() if f.lower().endswith('comicinfo.xml')), None)
                    if xml_name:
                        return rf.read(xml_name).decode('utf-8')
            except Exception as e:
                pass
        return None

    def read_metadata(self) -> Dict[str, str]:
        """Parses ComicInfo.xml content and returns a dictionary of metadata."""
        xml_content = self._read_xml_from_archive()
        metadata = {}
        if xml_content:
            try:
                # ET.fromstring handles the root attributes like 'xmlns:xsi' 
                root = ET.fromstring(xml_content)
                for element in root:
                    # Convert XML tag name (e.g., 'Title') to internal key (e.g., 'title')
                    # Strip namespace prefix if present (e.g., '{http://namespace}Tag')
                    tag = element.tag.split('}')[-1]
                    internal_key = self.REVERSE_MAPPING.get(tag)
                    
                    if internal_key and element.text is not None:
                        # Sanitize and store only non-empty strings
                        text_val = element.text.strip()
                        if text_val:
                            metadata[internal_key] = text_val
            except ET.ParseError as e:
                print(f"XML Parse Error in {self.file_path}: {e}")
        return metadata

    def _create_xml(self, metadata: Dict) -> str:
        """Creates the ComicInfo XML string from a dictionary of metadata."""
        # Filter metadata for non-empty/non-None values
        filtered_metadata = {
            k: v for k, v in metadata.items() 
            if v is not None and v is not False and (str(v).strip() or str(v) == '0')
        }
        
        # Add required XML schema attributes to the root element
        root = ET.Element(
            'ComicInfo',
            {
                'xmlns:xsi': "http://www.w3.org/2001/XMLSchema-instance",
                'xsi:noNamespaceSchemaLocation': "ComicInfo.xsd"
            }
        )
        root.text = '\n  '
        root.tail = '\n'
        
        last_elem = None
        
        # Sort keys based on XML tag name for consistent XML file structure
        sorted_keys = sorted(filtered_metadata.keys(), key=lambda k: self.FIELD_MAPPING.get(k, k))
        
        for key in sorted_keys:
            value = filtered_metadata[key]
            if key in self.FIELD_MAPPING:
                xml_tag = self.FIELD_MAPPING[key]
                elem = ET.SubElement(root, xml_tag)
                elem.text = str(value)
                elem.tail = '\n  '
                last_elem = elem
        
        if last_elem is not None:
            last_elem.tail = '\n'
        
        # Convert the ElementTree to a string
        xml_str = ET.tostring(root, encoding='utf-8', short_empty_elements=True, xml_declaration=False).decode('utf-8')
        
        # Add XML declaration and return
        return '<?xml version="1.0" encoding="utf-8"?>\n' + xml_str

    def write_metadata(self, metadata: Dict) -> Optional[str]:
        """Writes new ComicInfo.xml into the archive, returning the new file path if successful (for CBR->CBZ conversion)."""
        xml_content = self._create_xml(metadata)
        
        temp_dir = Path(tempfile.mkdtemp())
        temp_xml_path = temp_dir / 'ComicInfo.xml'
        
        try:
            # 1. Write XML to a temporary file
            with open(temp_xml_path, 'w', encoding='utf-8') as f:
                f.write(xml_content)
            
            # 2. Update the archive
            new_file_path = str(self.file_path) # Default to original path
            
            if self.is_cbz:
                # CBZ (ZIP) - Safely update by copying and overwriting
                temp_archive = temp_dir / self.file_path.name
                shutil.copy2(self.file_path, temp_archive)
                
                with zipfile.ZipFile(temp_archive, 'a') as zf_temp:
                    # Write the new/updated XML, overwriting the old one if it exists
                    zf_temp.write(temp_xml_path, 'ComicInfo.xml')
                
                # Overwrite the original file with the updated temporary file
                shutil.move(temp_archive, self.file_path)

            elif self.is_cbr and RARFILE_AVAILABLE:
                # CBR (RAR) - Must convert to CBZ (ZIP)
                
                # Create a list of files in the RAR archive to copy (excluding existing XML)
                file_list = []
                with rarfile.RarFile(self.file_path, 'r') as rf:
                    for f in rf.namelist():
                        if not f.lower().endswith('comicinfo.xml'):
                            file_list.append(f)
                    
                    # Create new CBZ path
                    new_file_path = str(self.file_path).replace('.cbr', '.cbz').replace('.CBR', '.CBZ')
                    if os.path.exists(new_file_path):
                        os.remove(new_file_path) # Remove previous version if exists
                        
                    # Create the new CBZ archive
                    with zipfile.ZipFile(new_file_path, 'w', zipfile.ZIP_DEFLATED) as zf_new:
                        # Add the new ComicInfo.xml
                        zf_new.write(temp_xml_path, 'ComicInfo.xml')
                        
                        # Add all original files from RAR
                        for f in file_list:
                            # Extract to temp, then add to zip
                            temp_file = rf.extract(f, temp_dir)
                            zf_new.write(temp_file, f)
                            os.remove(temp_file) # Clean up temp file
                            
                # Delete the old CBR file
                os.remove(self.file_path)
                
            else:
                raise ValueError("CBR is not supported (rarfile missing)")
                
            return new_file_path

        except Exception as e:
            print(f"Error writing XML to archive {self.file_path}: {e}")
            return None
            
        finally:
            # Clean up temporary directory
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)


class MetadataViewer:
    """A secondary window to display the existing metadata of a single comic file."""
    
    FIELD_MAP = ComicMetadataEditor.FIELD_MAPPING
    
    def __init__(self, master, metadata: Dict[str, str], file_path: str):
        self.master = master
        self.metadata = metadata
        self.file_path = file_path
        
        self.string_vars = {} 
        
        self.top = tk.Toplevel(master)
        self.top.title(f"Metadata Viewer: {os.path.basename(file_path)}")
        self.top.resizable(True, True)
        self.top.grab_set() 

        style = ttk.Style()
        style.configure("Viewer.TFrame", background="#f0f0f0")
        
        main_frame = ttk.Frame(self.top, padding="10", style="Viewer.TFrame")
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="Current Metadata (Read-Only)", font=("Arial", 12, "bold")).pack(pady=(0, 10))
        ttk.Label(main_frame, text=f"Source: {os.path.basename(file_path)}", font=("Arial", 9, "italic")).pack(pady=(0, 5))
        
        # --- Canvas and Scrollbar Setup (The Scrollable Area) ---
        canvas_frame = ttk.Frame(main_frame)
        canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        # We must use tk.Canvas for proper cross-platform scrolling within a Toplevel window
        canvas = tk.Canvas(canvas_frame, borderwidth=0)
        v_scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        
        v_scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        canvas.configure(yscrollcommand=v_scrollbar.set)
        
        scrollable_frame = ttk.Frame(canvas, padding="5 0 5 0")

        def _on_frame_configure(event):
            # This is critical for the scrollbar to know the size of the inner frame
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        scrollable_frame.bind("<Configure>", _on_frame_configure)
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        # --- End Canvas Setup ---

        self.create_widgets(scrollable_frame)
        self.show_metadata()

        # Footer
        footer = ttk.Frame(main_frame)
        footer.pack(fill='x', pady=10)
        ToolTip(footer, "Clicking a value copies it to your clipboard. Hovering highlights the text.")
        ttk.Label(footer, text="Click value to copy to clipboard", font=("Arial", 8, "italic"), foreground="gray").pack(side=tk.LEFT)
        ttk.Button(footer, text="Close", command=self.top.destroy).pack(side=tk.RIGHT)

        # Center the window
        self.top.update_idletasks()
        master_x = master.winfo_x()
        master_y = master.winfo_y()
        master_w = master.winfo_width()
        top_w = self.top.winfo_width()

        x = master_x + (master_w // 2) - (top_w // 2)
        y = master_y + 50 # Slightly offset from master top

        self.top.geometry(f"+{x}+{y}")


    def create_widgets(self, parent):
        """Creates label/value pairs for all possible fields."""
        
        ordered_keys = sorted(self.FIELD_MAP.keys(), key=lambda k: self.FIELD_MAP[k])
        
        row_idx = 0
        
        parent.grid_columnconfigure(0, weight=0) 
        parent.grid_columnconfigure(1, weight=1) 
        
        # Define hover colors
        NORMAL_BG = "#ffffff"
        HOVER_BG = "#e0e0ff" # Light blue background
        NORMAL_FG = "#000000" # Black text
        HOVER_FG = "#0000FF" # Blue text
        
        for internal_key in ordered_keys:
            # Skip keys that are usually calculated
            if internal_key in ['pages', 'isfolder']:
                continue 

            xml_tag = self.FIELD_MAP.get(internal_key, internal_key)

            ttk.Label(parent, text=f"{xml_tag}:", font=("Arial", 9, "bold")).grid(row=row_idx, column=0, sticky=tk.W, padx=5, pady=2)
            
            var = tk.StringVar(parent)
            self.string_vars[internal_key] = var 

            value_label = ttk.Label(parent, textvariable=var, wraplength=400, justify=tk.LEFT, 
                                    background=NORMAL_BG, foreground=NORMAL_FG, 
                                    relief=tk.FLAT, anchor=tk.NW)
            value_label.grid(row=row_idx, column=1, sticky=tk.W + tk.E + tk.N + tk.S, padx=5, pady=2, ipadx=5, ipady=2)
            
            # --- BINDING: Pass the StringVar and the Label widget ---
            def _copy_bind(var_ref, label_ref):
                # The lambda function captures the current references
                return lambda e: self._copy_to_clipboard(var_ref, label_ref)

            value_label.bind("<Button-1>", _copy_bind(var, value_label))
            # ----------------------------------------------------------------

            # Hover feedback (background and foreground change)
            value_label.bind("<Enter>", lambda e, l=value_label: l.config(background=HOVER_BG, foreground=HOVER_FG))
            value_label.bind("<Leave>", lambda e, l=value_label: l.config(background=NORMAL_BG, foreground=NORMAL_FG))
            
            parent.grid_rowconfigure(row_idx, weight=1) 
            
            row_idx += 1

    def show_metadata(self):
        """Populates the StringVars with the actual metadata values."""
        
        for key in self.string_vars:
            self.string_vars[key].set("(Not Set)")
        
        for key, value in self.metadata.items():
            if key in self.string_vars:
                display_value = str(value).strip() if value else "(Not Set)"
                self.string_vars[key].set(display_value)

    def _copy_to_clipboard(self, var: tk.StringVar, label_widget: ttk.Label):
        """Copies the given text to the clipboard and shows visual confirmation."""
        text = var.get()
        if text and text != "(Not Set)":
            self.top.clipboard_clear()
            self.top.clipboard_append(text)
            self.top.update() 

            # --- Visual Confirmation ---
            original_display_text = text # Store original text for revert
            
            # Use the existing StringVar for the text, so the label updates automatically
            CONFIRM_TEXT = "🚀 Copied to Clipboard! (Ctrl+C on that thang!)"
            CONFIRM_FG = "#006400" # Dark Green
            CONFIRM_BG = "#ccffcc" # Light Green
            
            # Temporarily change label properties
            label_widget.config(text=CONFIRM_TEXT, foreground=CONFIRM_FG, background=CONFIRM_BG, font=("Arial", 9, "bold"))

            def revert():
                # Revert to original text and colors (using the standard colors from create_widgets)
                # We check the current text to avoid reverting if the user immediately hovers away and the Enter binding resets the color
                if label_widget.cget("text") == CONFIRM_TEXT:
                    label_widget.config(text=original_display_text, 
                                        foreground="#000000", 
                                        background="#ffffff", 
                                        font=("Arial", 9, "normal"))
                
            # Schedule the revert function to run after 750 milliseconds
            self.top.after(750, revert)
            # --- End Visual Confirmation ---


class ComicMetadataGUI:
    """The main application window and logic controller."""
    
    METADATA_FIELDS = ComicMetadataEditor.FIELD_MAPPING

    # --- Standard Lists for Comboboxes ---
    ISO_LANGUAGES = [
        'en (English)', 'es (Spanish)', 'fr (French)', 'de (German)', 
        'ja (Japanese)', 'ko (Korean)', 'zh (Chinese)', 'it (Italian)', 
        'ru (Russian)', 'pt (Portuguese)', 'nl (Dutch)', 'pl (Polish)',
        'sv (Swedish)', 'fi (Finnish)', 'nb (Norwegian)', 'da (Danish)',
        'la (Latin)', 'ar (Arabic)', 'tr (Turkish)', 'he (Hebrew)'
    ]
    
    COMMON_COUNTRIES = [
        'USA', 'UK', 'Canada', 'Australia', 'New Zealand', 'France', 'Germany', 
        'Spain', 'Italy', 'Japan', 'South Korea', 'China', 'Brazil', 'Mexico',
        'Argentina', 'India', 'Russia'
    ]
    
    AGE_RATINGS = [
        '', 'Adults Only', 'Early Childhood', 'Everyone', 'Everyone 10+', 'G', 
        'Kids', 'M', 'MA15+', 'Mature', 'PG', 'PG-13', 'Teen', 'T+', 'X', 'Young Adult'
    ]
    
    # UPDATED: Added YesAndRightToLeft 
    MANGA_TYPES = ['', 'Yes', 'No', 'Seinen', 'Shoujo', 'Shonen', 'Josei', 'Kodomo', 'YesAndRightToLeft']
    
    # NEW: Added full list of ComicInfo standard formats
    COMMON_FORMATS = [
        '', 'Digital', 'Print', 'Hardcover', 'Softcover', 'Trade Paperback', 'Graphic Novel',
        'Magazine', 'Fanzine', 'Web Comic', 'Anthology', 'Collected Edition', 'TPB'
    ]
    # -----------------------------------

    def __init__(self, root):
        self.root = root
        self.root.title("Comic Metadata Bulk Editor - v2.31 (Autovoluming Re-added)")
        
        self.files: List[str] = []
        self.control_vars: Dict[str, tk.Variable] = {}
        self.input_widgets: Dict[str, tk.Widget] = {} 
        self.autonumber_start_var = tk.StringVar(value='1')
        
        # New member variables for the new buttons
        self.copy_all_btn: Optional[ttk.Button] = None
        self.view_metadata_btn: Optional[ttk.Button] = None
        
        self._setup_styles()
        self._setup_main_layout()
        self._create_metadata_tab() # Single, consolidated tab
        
        self.show_welcome_message()
        
    def _setup_styles(self):
        """Configure Ttk styles."""
        style = ttk.Style()
        style.theme_use('clam')
        
        style.configure("List.TFrame", padding="10") 
        style.configure("Sash.TPanedwindow", sashwidth=8, sashrelief='groove', background="#CCCCCC")
        
        style.configure('TButton', font=('Arial', 10), padding=5)
        style.map('TButton', background=[('active', '#e0e0e0')])

    def _setup_main_layout(self):
        """Create the paned window and the overall structure."""
        
        main_frame = ttk.Frame(self.root, padding="5")
        main_frame.pack(fill=tk.BOTH, expand=True)

        self.paned_window = ttk.Panedwindow(main_frame, orient=tk.HORIZONTAL, style="Sash.TPanedwindow")
        self.paned_window.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # --- Left Pane: File List and Controls ---
        left_pane = ttk.Frame(self.paned_window, style="List.TFrame")
        self.paned_window.add(left_pane, weight=1) # File List Pane (smaller)

        # --- Right Pane: Notebook (Tabs) ---
        self.notebook = ttk.Notebook(self.paned_window)
        self.paned_window.add(self.notebook, weight=3) # Metadata Pane (3x larger)

        self.scrollable_frames = {}
        self._create_left_pane(left_pane)
        
    def _create_left_pane(self, parent):
        """Build the file list, controls, and status bar."""
        
        # --- 1. File List Frame ---
        list_frame = ttk.Frame(parent)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        
        ttk.Label(list_frame, text="Loaded Comic Files:", font=("Arial", 10, "bold")).pack(anchor=tk.W)

        listbox_vscroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL)
        listbox_vscroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.file_listbox = tk.Listbox(
            list_frame, 
            selectmode=tk.EXTENDED, 
            height=15, 
            yscrollcommand=listbox_vscroll.set,
            exportselection=False, 
            font=("Arial", 9)
        )
        self.file_listbox.pack(fill=tk.BOTH, expand=True)
        listbox_vscroll.config(command=self.file_listbox.yview)
        
        self.file_listbox.bind('<<ListboxSelect>>', self.on_file_select)
        
        # --- 2. Action Buttons Frame (Add/Remove/View) ---
        action_frame = ttk.Frame(parent)
        action_frame.pack(fill=tk.X, pady=(5, 0))
        # Use 3 columns for a clean layout
        action_frame.columnconfigure(0, weight=1)
        action_frame.columnconfigure(1, weight=1)
        action_frame.columnconfigure(2, weight=1) 

        ttk.Button(action_frame, text="Add Files...", command=self.add_files).grid(row=0, column=0, sticky=tk.W+tk.E, padx=2, pady=2)
        ttk.Button(action_frame, text="Remove Selected", command=self.remove_selected).grid(row=0, column=1, sticky=tk.W+tk.E, padx=2, pady=2)
        
        # NEW BUTTON: View Metadata
        self.view_metadata_btn = ttk.Button(action_frame, text="👁️ View Metadata", command=self.load_metadata, state=tk.DISABLED)
        self.view_metadata_btn.grid(row=0, column=2, sticky=tk.W+tk.E, padx=2, pady=2)
        ToolTip(self.view_metadata_btn, "Loads and displays the ComicInfo.xml metadata from the **single** selected file in a new window.")

        # "Copy All" button renamed and moved to row 1, spanning all 3 columns
        self.copy_all_btn = ttk.Button(action_frame, text="Import Metadata from Selected", command=self.copy_all_to_main_fields, state=tk.DISABLED)
        self.copy_all_btn.grid(row=1, column=0, columnspan=3, sticky=tk.W+tk.E, padx=2, pady=2)
        ToolTip(self.copy_all_btn, "Loads metadata from the **single** selected file and populates the main form fields, ticking their checkboxes.")


        # --- 3. Sorting & Autonumbering Frame ---
        sort_num_frame = ttk.Frame(parent, relief=tk.RIDGE, padding=5)
        sort_num_frame.pack(fill=tk.X, pady=(5, 5))
        sort_num_frame.columnconfigure(0, weight=1)
        sort_num_frame.columnconfigure(1, weight=1)
        
        # Sorting Buttons
        ttk.Label(sort_num_frame, text="Manual Sorting:", font=("Arial", 9, "bold")).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 2))
        ttk.Button(sort_num_frame, text="▲ Move Up", command=self.move_selected_up).grid(row=1, column=0, sticky=tk.W+tk.E, padx=2, pady=2)
        ttk.Button(sort_num_frame, text="▼ Move Down", command=self.move_selected_down).grid(row=1, column=1, sticky=tk.W+tk.E, padx=2, pady=2)
        
        ttk.Separator(sort_num_frame, orient=tk.HORIZONTAL).grid(row=2, column=0, columnspan=2, sticky=tk.EW, pady=5)
        
        # Autonumbering Controls
        ttk.Label(sort_num_frame, text="Auto-Voluming:", font=("Arial", 9, "bold")).grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=(0, 2))
        
        num_entry_frame = ttk.Frame(sort_num_frame)
        num_entry_frame.grid(row=4, column=0, sticky=tk.W+tk.E, padx=2, pady=2)
        num_entry_frame.columnconfigure(1, weight=1)
        ttk.Label(num_entry_frame, text="Start Num:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(num_entry_frame, textvariable=self.autonumber_start_var, width=5).grid(row=0, column=1, sticky=tk.W, padx=(5, 0))
        ToolTip(num_entry_frame, "The volume number to start counting from (e.g., 1).")
        
        ttk.Button(sort_num_frame, text="🔢 Autovolume Selected", command=self.autonumber_selected).grid(row=4, column=1, sticky=tk.W+tk.E, padx=2, pady=2)
        ToolTip(sort_num_frame.winfo_children()[-1], "Sequentially number the selected files based on their order, starting from the Start Num. Updates the 'Volume #' and 'Total Volumes' fields, and ticks their checkboxes.")


        # --- 4. Apply and Status ---
        self.btn_apply = ttk.Button(parent, text="APPLY METADATA to Selected", command=self.apply_metadata, style='Accent.TButton')
        self.btn_apply.pack(fill=tk.X, pady=5)
        
        self.status_bar = ttk.Label(parent, text="Ready. Load files to begin.", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(fill=tk.X, ipady=2, pady=(5, 0))
        
        self.progress_bar = ttk.Progressbar(parent, orient="horizontal", length=100, mode="determinate")
        self.progress_bar.pack(fill=tk.X, pady=(0, 5))
        self.progress_bar.pack_forget() 

        self.update_status()

    def _on_mousewheel(self, event, canvas):
        """Universal scroll wheel binding for a canvas."""
        # Windows and Linux use <MouseWheel> with a delta. macOS uses <Button-4>/<Button-5> (which appear as <MouseWheel> events with different deltas).
        
        # Determine scroll units based on OS/Event details
        if sys.platform.startswith('win'):
            # Windows: event.delta is typically +/- 120 per click
            scroll_delta = int(-1 * (event.delta / 120))
        elif sys.platform.startswith('linux') or sys.platform == 'darwin':
            # Linux/macOS: event.delta is usually not available, but Button-4/5 are sometimes mapped to <MouseWheel>
            # For simplicity and cross-platform compatibility, we'll use a fixed scroll amount for non-Windows:
            if event.num == 4: # Button-4 (Scroll Up)
                scroll_delta = -1
            elif event.num == 5: # Button-5 (Scroll Down)
                scroll_delta = 1
            elif event.delta < 0: # Modern Linux/macOS
                scroll_delta = 1
            else:
                scroll_delta = -1
        else:
            # Fallback for others
            scroll_delta = -1 if event.delta > 0 else 1
            
        canvas.yview_scroll(scroll_delta, "units")

    def _create_metadata_tab_frame(self, tab_name):
        """Helper to create a scrollable frame within a new notebook tab."""
        
        tab_frame = ttk.Frame(self.notebook)
        self.notebook.add(tab_frame, text=tab_name)
        
        canvas = tk.Canvas(tab_frame, borderwidth=0)
        v_scrollbar = ttk.Scrollbar(tab_frame, orient="vertical", command=canvas.yview)
        v_scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        canvas.configure(yscrollcommand=v_scrollbar.set)
        
        scrollable_frame = ttk.Frame(canvas, padding="10")
        
        def _on_frame_configure(event):
            # Ensure the canvas view is properly set and includes the minimum width 
            # for the content to fully stretch.
            canvas.configure(scrollregion=canvas.bbox("all"), width=scrollable_frame.winfo_width())
        
        scrollable_frame.bind("<Configure>", _on_frame_configure)
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        
        # --- SCROLL WHEEL BINDING (Cross-Platform) ---
        canvas.bind("<MouseWheel>", lambda e: self._on_mousewheel(e, canvas))
        # This is for X11/macOS compatibility for scroll up/down
        canvas.bind("<Button-4>", lambda e: self._on_mousewheel(e, canvas)) 
        canvas.bind("<Button-5>", lambda e: self._on_mousewheel(e, canvas))
        # Bind events to the inner frame as well, so it works when the cursor is over a widget inside
        scrollable_frame.bind("<MouseWheel>", lambda e: self._on_mousewheel(e, canvas))
        scrollable_frame.bind("<Button-4>", lambda e: self._on_mousewheel(e, canvas)) 
        scrollable_frame.bind("<Button-5>", lambda e: self._on_mousewheel(e, canvas))

        self.scrollable_frames[tab_name] = scrollable_frame
        return scrollable_frame

    def _create_metadata_widget(self, parent, internal_key, label_text, row, col, widget_type='entry', widget_opts=None, tooltip_text=""):
        """
        Creates a Checkbutton (Apply), Label, and Input Widget (Value).
        col should be 0 or 4 for the left/right groups.
        """
        
        # --- 1. Apply Checkbox (controls whether this field is written) ---
        var_check = tk.IntVar(value=0)
        self.control_vars[f'check_{internal_key}'] = var_check
        
        check_apply = ttk.Checkbutton(parent, variable=var_check)
        check_apply.grid(row=row, column=col, sticky=tk.W, padx=(5, 0), pady=2)
        ToolTip(check_apply, f"Check this box to **APPLY** the field value below to selected files. If unchecked, the existing value will be preserved.")

        # --- 2. Label ---
        label = ttk.Label(parent, text=label_text)
        label.grid(row=row, column=col + 1, sticky=tk.W, padx=(0, 5), pady=2)
        
        # --- 3. Input Widget (Value) ---
        widget = None
        
        # Use a larger default width for better horizontal stretching
        WIDER_WIDTH = 60 

        if widget_type == 'checkbutton':
            # For boolean metadata: the value is a tk.IntVar (0 or 1)
            var_value = tk.IntVar()
            self.control_vars[internal_key] = var_value
            
            # Place the checkbutton in the value column (col + 2)
            widget = ttk.Checkbutton(parent, text="", variable=var_value) 
            widget.grid(row=row, column=col + 2, sticky=tk.W, padx=(0, 5), pady=2)
            
        elif widget_type == 'entry':
            var_value = tk.StringVar()
            self.control_vars[internal_key] = var_value
            widget = ttk.Entry(parent, textvariable=var_value, width=WIDER_WIDTH)
            widget.grid(row=row, column=col + 2, sticky=tk.W + tk.E, padx=(0, 5), pady=2)

        elif widget_type == 'combobox':
            var_value = tk.StringVar()
            self.control_vars[internal_key] = var_value
            opts = {'state': 'readonly', 'width': WIDER_WIDTH - 2, **(widget_opts or {})}
            widget = ttk.Combobox(parent, textvariable=var_value, **opts)
            widget.grid(row=row, column=col + 2, sticky=tk.W + tk.E, padx=(0, 5), pady=2)
        
        self.input_widgets[internal_key] = widget

        if tooltip_text and widget:
            ToolTip(widget, tooltip_text)

        # Ensure the input column (col + 2) expands horizontally to fill space
        parent.grid_columnconfigure(col + 2, weight=1)

    def _create_long_text_widget(self, parent, internal_key, label_text, row, tooltip_text=""):
        """Creates the special layout for multi-line text fields (Summary, Notes, Review)."""
        
        # --- 1. Apply Checkbox ---
        var_check = tk.IntVar(value=0)
        self.control_vars[f'check_{internal_key}'] = var_check
        
        check_apply = ttk.Checkbutton(parent, variable=var_check)
        check_apply.grid(row=row, column=0, sticky=tk.W, padx=(5, 0), pady=2)
        ToolTip(check_apply, f"Check this box to **APPLY** the field value below to selected files. If unchecked, the existing value will be preserved.")

        # --- 2. Label (Spans two columns for better look) ---
        label = ttk.Label(parent, text=label_text)
        label.grid(row=row, column=1, columnspan=2, sticky=tk.W, padx=(0, 5), pady=2)

        # --- 3. ScrolledText Widget (Value) ---
        var_value = tk.StringVar() 
        self.control_vars[internal_key] = var_value
        
        # Increase default width for better horizontal space usage
        widget = scrolledtext.ScrolledText(parent, wrap=tk.WORD, width=90, height=5, font=("Arial", 9))
        var_value.widget_ref = widget 
        
        # Widget spans columns 0 through 6 for maximum width
        widget.grid(row=row + 1, column=0, columnspan=7, sticky=tk.W + tk.E, padx=5, pady=2) 
        
        # Update StringVar from widget content
        widget.bind("<FocusOut>", lambda e, w=widget, v=var_value: v.set(w.get("1.0", tk.END).strip()))
        widget.bind("<KeyRelease>", lambda e, w=widget, v=var_value: v.set(w.get("1.0", tk.END).strip()))
        
        self.input_widgets[internal_key] = widget

        if tooltip_text:
            ToolTip(widget, tooltip_text)

        # Ensure the frame allows text widgets to expand horizontally
        parent.grid_columnconfigure(6, weight=1)

    def _create_metadata_tab(self):
        """Consolidated tab for all metadata, using a dense two-column grid."""
        scrollable_frame = self._create_metadata_tab_frame("All Metadata")
        
        # --- Date Fields Group (Top Left) ---
        date_fields = [
            ('year', 'Year:', 'entry', None, "Publication year (YYYY)."),
            ('month', 'Month:', 'entry', None, "Publication month (1-12)."),
            ('day', 'Day:', 'entry', None, "Publication day (1-31)."),
        ]

        # --- Core/Volume/Series Fields (Left Column, below Dates) ---
        left_fields = [
            ('title', 'Title:', 'entry', None, "The main title of the comic."),
            ('series', 'Series:', 'entry', None, "The main series name."),
            ('volume', 'Volume #:', 'entry', None, "The volume number (e.g., 1 for Volume 1)."),
            ('volume_count', 'Total Volumes:', 'entry', None, "Total number of volumes in the series."),
            ('number', 'Issue #:', 'entry', None, "The sequential issue number."),
            ('issuecount', 'Total Issues:', 'entry', None, "Total number of issues in the story arc/series."),
            ('seriescomplete', 'Series Complete:', 'checkbutton', None, "Is this series finished? (True/False)"),
            ('storyArc', 'Story Arc:', 'entry', None, "The name of the overall story arc (e.g., 'The Dark Phoenix Saga')."),
        ]
        
        # --- Auxiliary/Publisher/Rating Fields (Right Column) ---
        right_fields = [
            ('publisher', 'Publisher:', 'entry', None, "The comic book publisher."),
            ('imprint', 'Imprint:', 'entry', None, "The imprint or sub-brand of the publisher."),
            ('seriesgroup', 'Series Group:', 'entry', None, "A grouping for related series (e.g., 'Marvel Legacy')."),
            # UPDATED: Format now uses a combobox with a list of COMMON_FORMATS
            ('format', 'Format:', 'combobox', {'values': self.COMMON_FORMATS, 'state': 'normal'}, "Format type (e.g., Digital, Print, Hardcover)."),
            ('agerating', 'Maturity Rating:', 'combobox', {'values': self.AGE_RATINGS, 'state': 'readonly'}, "The maturity rating for the content."),
            # UPDATED: Manga uses the expanded MANGA_TYPES list
            ('manga', 'Manga:', 'combobox', {'values': self.MANGA_TYPES, 'state': 'readonly'}, "Is this a Manga? Specifies reading direction."),
            ('blackandwhite', 'Black & White:', 'checkbutton', None, "Is the content in black and white? (True/False)"),
            ('language', 'Language (ISO):', 'combobox', {'values': self.ISO_LANGUAGES, 'state': 'normal'}, "Language code (e.g., en, fr)."),
            ('country', 'Country:', 'combobox', {'values': self.COMMON_COUNTRIES, 'state': 'normal'}, "Country of publication (e.g., USA)."),
            ('genre', 'Genre:', 'entry', None, "The primary genre (e.g., Action, Sci-Fi)."),
            ('communityrating', 'Community Rating (0-5):', 'entry', None, "A community or personal rating (e.g., 3.5)."),
            ('gtin', 'GTIN/EAN/ISBN:', 'entry', None, "Product identifier code."),
            ('read', 'Read Status:', 'checkbutton', None, "Has the comic been read? (For tracking purposes)."),
            ('alternatenumber', 'Alt. Number:', 'entry', None, "Alternate issue number for variants or reprints."),
            ('alternateissuecount', 'Alt. Total Issues:', 'entry', None, "Total issues in the alternate series/volume."),
            ('alternateSeries', 'Alt. Series:', 'entry', None, "Name of an alternate series/volume."),
        ]
        
        # 1. Date fields (Left column, starting at row 0)
        current_row = 0
        for i, (internal_key, label, widget_type, widget_opts, tooltip) in enumerate(date_fields):
            self._create_metadata_widget(scrollable_frame, internal_key, label, current_row + i, 0, widget_type, widget_opts, tooltip)
        current_row += len(date_fields) # Now at the row right after the date fields

        # 2. Left fields (Left column, continuing from date fields)
        for i, (internal_key, label, widget_type, widget_opts, tooltip) in enumerate(left_fields):
            self._create_metadata_widget(scrollable_frame, internal_key, label, current_row + i, 0, widget_type, widget_opts, tooltip)
        max_left_row = current_row + len(left_fields)

        # 3. Right fields (Right column, starting at row 0)
        scrollable_frame.grid_columnconfigure(3, minsize=50) # Spacer column
        for i, (internal_key, label, widget_type, widget_opts, tooltip) in enumerate(right_fields):
            self._create_metadata_widget(scrollable_frame, internal_key, label, i, 4, widget_type, widget_opts, tooltip)
        max_right_row = len(right_fields)


        # CREW & SIMPLE TEXT (Middle section, single column)
        simple_text_start_row = max(max_left_row, max_right_row) + 1 
        
        ttk.Separator(scrollable_frame, orient=tk.HORIZONTAL).grid(row=simple_text_start_row, column=0, columnspan=7, sticky=tk.EW, pady=(10, 5))
        simple_text_start_row += 1 

        simple_text_fields = [
            ('writer', 'Writer:', "The writer(s) of the story."),
            ('penciller', 'Penciller:', "The artist(s) responsible for the pencils."),
            ('inker', 'Inker:', "The artist(s) responsible for the inking."),
            ('colorist', 'Colorist:', "The artist(s) responsible for coloring."),
            ('letterer', 'Letterer:', "The person(s) responsible for lettering."),
            ('coverartist', 'Cover Artist:', "The artist(s) who drew the cover."),
            ('editor', 'Editor:', "The editor(s) responsible for the book."),
            ('authorsort', 'Author Sort:', "Field used for sorting by author/creator."),
            ('maincharacter', 'Main Character:', "The central character or team."),
            ('characters', 'Characters (Comma separated):', "All major characters in the comic."),
            ('teams', 'Teams (Comma separated):', "All teams featured in the comic."),
            ('locations', 'Locations (Comma separated):', "Key locations where the story takes place."),
            ('tags', 'Tags (Comma separated):', "Additional keywords or tags."),
            ('scaninformation', 'Scan Info:', "Details about the digital scan/source."),
            ('web', 'Web/URL:', "URL of the comic or source.")
        ]
        
        # Use columns 0-2 (left group) for crew/text fields
        for i, (internal_key, label, tooltip) in enumerate(simple_text_fields):
            self._create_metadata_widget(scrollable_frame, internal_key, label, simple_text_start_row + i, 0, 'entry', None, tooltip)


        # MULTILINE FIELDS (Bottom section, full width)
        # Start where simple text left off, plus a row for visual break
        text_start_row = simple_text_start_row + len(simple_text_fields) + 1
        
        ttk.Separator(scrollable_frame, orient=tk.HORIZONTAL).grid(row=text_start_row, column=0, columnspan=7, sticky=tk.EW, pady=(10, 5))
        text_start_row += 1 

        long_text_fields = [
            ('summary', 'Summary:', "A brief description of the comic's plot."),
            ('notes', 'Notes:', "General notes about the file or comic."),
            ('review', 'Review:', "A personal review or rating for the comic."),
        ]
        
        current_row = text_start_row
        for internal_key, label, tooltip in long_text_fields:
            # This creates the label at current_row and the ScrolledText widget at current_row + 1
            self._create_long_text_widget(scrollable_frame, internal_key, label, current_row, tooltip)
            current_row += 2 # Long text takes up the label row + text widget row
            
    def show_welcome_message(self):
        """Displays a welcome and explanation window."""
        top = tk.Toplevel(self.root)
        top.title("Welcome to the Comic Metadata Bulk Editor")
        top.resizable(False, False)
        top.transient(self.root)
        top.grab_set()

        frame = ttk.Frame(top, padding="15")
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Comic Metadata Bulk Editor (v2.31)", font=("Arial", 14, "bold")).pack(pady=10)
        
        st = scrolledtext.ScrolledText(frame, width=60, height=15, font=("Arial", 10), relief=tk.FLAT)
        st.pack(pady=10, fill=tk.BOTH, expand=True)
        
        explanation = """
This version re-integrates the **Autovoluming** and **Manual Sorting** tools alongside the **View Metadata** button.

### 1. Auto-Voluming & Manual Sorting 📚
* **Manual Sort:** Use the **'▲ Move Up'** and **'▼ Move Down'** buttons to precisely order the files in the listbox.
* **Auto-Voluming:** Select the files, set the **'Start Num'** (e.g., 1), and click **'Autovolume Selected'**.
    * This will set the `Volume #` (Volume) field to the sequence (1, 2, 3...) based on the file order.
    * It will also calculate and set the `Total Volumes` (VolumeCount) field.
    * The **Apply** checkboxes for both fields will be automatically **ticked**.

### 2. View/Import Metadata 🔍
* **👁️ View Metadata Button:** Click this button when **one file** is selected to see its existing ComicInfo.xml metadata in a new, scrollable, read-only window. Values are **clickable to copy**.
* **Import Button:** Click "Import Metadata from Selected" (when one file is selected) to populate the main form with that file's data.

### 3. Core Logic: Tickbox Merge
* Only fields with a **TICKED APPLY CHECKBOX** will be included in the update.
* Unticked fields will be **PRESERVED** from the original file.
        """
        
        st.insert("1.0", explanation)
        st.config(state=tk.DISABLED)
        
        ttk.Button(frame, text="Got It! Start Editing", command=top.destroy).pack(pady=10)
        
        self.root.update_idletasks() 
        top.update_idletasks() 
        
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_w = self.root.winfo_width()
        top_w = top.winfo_width()
        
        x = root_x + (root_w // 2) - (top_w // 2)
        y = root_y + 50 

        top.wm_geometry(f"+{x}+{y}")

    def update_status(self, message: Optional[str] = None):
        """Updates the status bar message and checks file list integrity."""
        
        if not self.files:
            msg = "Ready. Load files to begin."
            self.btn_apply.config(state=tk.DISABLED)
        else:
            selected_count = len(self.file_listbox.curselection())
            msg = f"Files Loaded: {len(self.files)} | Selected: {selected_count}"
            self.btn_apply.config(state=tk.NORMAL if selected_count > 0 else tk.DISABLED)
            
            # Update 'Copy All' and 'View Metadata' button state
            is_single_selection = selected_count == 1
            if self.copy_all_btn:
                self.copy_all_btn.config(state=tk.NORMAL if is_single_selection else tk.DISABLED)
            if self.view_metadata_btn:
                 self.view_metadata_btn.config(state=tk.NORMAL if is_single_selection else tk.DISABLED)

        self.status_bar.config(text=message or msg)

    def _update_listbox_display(self, selected_indices=None):
        """Refreshes the listbox content to reflect changes in self.files."""
        current_selection = self.file_listbox.curselection()
        
        self.file_listbox.delete(0, tk.END)
        for f in self.files:
            self.file_listbox.insert(tk.END, os.path.basename(f))
            
        # Restore previous selection or apply new selection
        indices_to_select = selected_indices if selected_indices is not None else current_selection
        for idx in indices_to_select:
            try:
                self.file_listbox.select_set(idx)
            except tk.TclError:
                pass # Index out of range
                
        self.update_status()

    def add_files(self):
        """Opens file dialog to select CBZ/CBR files."""
        file_types = [
            ("Comic Archives", "*.cbz *.cbr"),
            ("CBZ files", "*.cbz"),
            ("CBR files", "*.cbr"),
            ("All files", "*.*")
        ]
        
        new_files = filedialog.askopenfilenames(
            title="Select Comic Files (.cbz / .cbr)",
            filetypes=file_types
        )
        
        if new_files:
            for f in new_files:
                if f not in self.files:
                    self.files.append(f)
            self._update_listbox_display()

    def remove_selected(self):
        """Removes selected files from the list."""
        selected_indices = self.file_listbox.curselection()
        if not selected_indices:
            return

        # Use reversed order to correctly delete from the list
        for index in selected_indices[::-1]:
            del self.files[index]
            
        self._update_listbox_display()
        self.clear_fields()

    def move_selected_up(self):
        """Moves selected items up one position in the list."""
        selected_indices = list(self.file_listbox.curselection())
        
        if not selected_indices or 0 in selected_indices:
            return

        for i in selected_indices:
            # Only move if the item directly above it is NOT selected
            if i > 0 and i - 1 not in selected_indices:
                # Swap the elements in the underlying list
                self.files[i], self.files[i-1] = self.files[i-1], self.files[i]
                
        # Calculate the new selection indices
        new_selection = [i - 1 for i in selected_indices if i > 0]
        # Include any selected indices that were at index 0 (they didn't move)
        new_selection += [0 for i in selected_indices if i == 0]
        
        # Filter and sort the final selection list
        new_selection = sorted(list(set(new_selection)))
        
        self._update_listbox_display(new_selection)

    def move_selected_down(self):
        """Moves selected items down one position in the list."""
        selected_indices = list(self.file_listbox.curselection())
        
        if not selected_indices or (len(self.files) - 1) in selected_indices:
            return

        # Iterate in reverse order to ensure indices remain correct during swaps
        for i in selected_indices[::-1]:
            # Only move if the item directly below it is NOT selected
            if i < len(self.files) - 1 and i + 1 not in selected_indices:
                # Swap the elements in the underlying list
                self.files[i], self.files[i+1] = self.files[i+1], self.files[i]

        # Calculate the new selection indices
        new_selection = [i + 1 for i in selected_indices if i < len(self.files) - 1]
        
        # Filter and sort the final selection list
        new_selection = sorted(list(set(new_selection)))
        
        self._update_listbox_display(new_selection)
        
    def autonumber_selected(self):
        """
        Sequentially numbers the selected files based on their order in the list.
        Updates the 'Volume' and 'VolumeCount' fields and ticks their 'apply' checkboxes.
        """
        selected_indices = list(self.file_listbox.curselection())
        if not selected_indices:
            messagebox.showwarning("No Selection", "Please select one or more files to set volume numbers.")
            return

        try:
            start_num = int(self.autonumber_start_var.get())
            if start_num <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Invalid Input", "The 'Start Num' must be a positive integer.")
            return
            
        # Get the number of selected files (which will be the VolumeCount)
        volume_count = len(selected_indices)
        
        # The volume field to update
        volume_var = self.control_vars.get('volume')
        volume_check = self.control_vars.get('check_volume')
        
        # The volume_count field to update
        volume_count_var = self.control_vars.get('volume_count')
        volume_count_check = self.control_vars.get('check_volume_count')
        
        # Set the total count
        if volume_count_var and volume_count_check:
            volume_count_var.set(str(volume_count))
            volume_count_check.set(1) # Tick the checkbox for Total Volumes

        # Set the starting number in the 'Volume' field
        if volume_var and volume_check:
            volume_var.set(str(start_num))
            volume_check.set(1) # Tick the checkbox for Volume #
            
        self.update_status(f"Auto-Voluming set: Start={start_num}, Count={volume_count}. 'Volume #' and 'Total Volumes' fields are ready.")


    def clear_fields(self):
        """Clears all input fields and unticks all checkboxes."""
        for key in self.METADATA_FIELDS:
            # Clear value variable (StringVar or IntVar)
            if key in self.control_vars:
                var = self.control_vars[key]
                if isinstance(var, tk.StringVar):
                    # Handle Text widget specially
                    if hasattr(var, 'widget_ref'):
                        var.widget_ref.delete("1.0", tk.END)
                    var.set("")
                elif isinstance(var, tk.IntVar):
                    var.set(0) # Clear boolean checkbox

            # Clear 'apply' checkbox
            check_key = f'check_{key}'
            if check_key in self.control_vars:
                self.control_vars[check_key].set(0)

        self.update_status("Fields cleared. Ready for new input.")

    def get_metadata_values(self) -> Dict[str, str]:
        """Retrieves current values from all input fields, casting to appropriate string format for ComicInfo.xml.
           ONLY returns fields that are ticked/checked."""
        metadata = {}
        for key in self.METADATA_FIELDS:
            # Only process if the 'apply' checkbox is ticked
            check = self.control_vars.get(f'check_{key}')
            if not check or check.get() != 1:
                continue

            var = self.control_vars.get(key)
            if not var:
                continue

            value = None
            
            if isinstance(var, tk.StringVar):
                # Handle StringVar (Entry, Combobox, Text)
                if hasattr(var, 'widget_ref'):
                    # Retrieve the content directly from the ScrolledText widget
                    value = var.widget_ref.get("1.0", tk.END).strip()
                else:
                    value = var.get().strip()
                
                # Special casing for LanguageISO: extract the code (e.g., 'en')
                if key == 'language':
                    # Extract ISO part from 'en (English)' format
                    match = re.match(r'([a-z]{2,3})\s+\(.+\)', value, re.IGNORECASE)
                    if match:
                        value = match.group(1).lower()

            elif isinstance(var, tk.IntVar):
                # Handle IntVar (Checkbutton)
                raw_value = var.get()
                
                # Convert the internal 0/1 to the standard ComicInfo.xml string ('Yes'/'No' or 'True'/'False')
                if key in ['seriescomplete', 'blackandwhite', 'read']: 
                    value = 'Yes' if raw_value == 1 else 'No'
                elif key == 'isfolder':
                    # This field is usually handled internally, but included for completeness
                    value = 'true' if raw_value == 1 else 'false'
                
            if value is not None and value != '':
                metadata[key] = value
        return metadata
        
    def load_metadata(self):
        """Loads metadata from the currently selected file and displays it in a viewer window."""
        selected_indices = self.file_listbox.curselection()
        
        if len(selected_indices) != 1:
            messagebox.showwarning("Selection Error", "Please select exactly ONE file to view its metadata.")
            return

        selected_index = selected_indices[0]
        file_path = self.files[selected_index]
        
        try:
            editor = ComicMetadataEditor(file_path)
            metadata = editor.read_metadata()
            
            if not metadata:
                messagebox.showwarning(
                    "No Metadata Found",
                    f"The file '{os.path.basename(file_path)}' does not contain readable ComicInfo.xml metadata (or it's empty)."
                )
                return 
            
            MetadataViewer(self.root, metadata, file_path)
            
        except Exception as e:
            messagebox.showerror("Error Reading File", f"Could not read metadata from file: {e}")

    def copy_all_to_main_fields(self):
        """Copies metadata from the selected file into the main input fields."""
        selected_indices = self.file_listbox.curselection()
        
        if len(selected_indices) != 1:
            # This check is also done in the button state, but good to keep here
            messagebox.showwarning("Selection Error", "Please select exactly ONE file to import metadata from.")
            return

        selected_index = selected_indices[0]
        file_path = self.files[selected_index]
        self.update_status(f"Importing metadata from: {os.path.basename(file_path)}...")

        try:
            editor = ComicMetadataEditor(file_path)
            metadata = editor.read_metadata()
            
            if not metadata:
                 messagebox.showwarning(
                    "No Metadata Found",
                    f"The file '{os.path.basename(file_path)}' does not contain readable ComicInfo.xml metadata (or it's empty)."
                )
                 return

            self.clear_fields()
            
            updated_count = 0
            for key, value in metadata.items():
                if key in self.control_vars:
                    var = self.control_vars[key]
                    
                    if isinstance(var, tk.StringVar):
                        # Handle Text widget
                        if hasattr(var, 'widget_ref'):
                            var.widget_ref.delete("1.0", tk.END)
                            var.widget_ref.insert("1.0", value)
                            var.set(value)
                        else:
                            # For Language, try to match ISO code ('en') to the full Combobox string ('en (English)')
                            if key == 'language':
                                iso_code = value.lower()
                                matched_val = next((s for s in self.ISO_LANGUAGES if s.startswith(iso_code + ' ')), value)
                                var.set(matched_val)
                            else:
                                var.set(value)
                                
                    elif isinstance(var, tk.IntVar):
                        # Handle Boolean values: 'Yes', 'No', 'True', 'False', 0, 1
                        val_str = str(value).lower()
                        if val_str in ['yes', 'true', '1']:
                            var.set(1)
                        else:
                            var.set(0)
                            
                    # Set the corresponding 'apply' checkbox to Ticked (1)
                    check_key = f'check_{key}'
                    if check_key in self.control_vars:
                        self.control_vars[check_key].set(1)
                        updated_count += 1
                        
            self.update_status(f"Successfully imported {updated_count} fields and ticked checkboxes.")

        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to import metadata: {e}")
            self.update_status("Error during metadata import.")

    def apply_metadata(self):
        """Applies the current metadata values to all selected files."""
        selected_indices = self.file_listbox.curselection()
        
        if not selected_indices:
            messagebox.showwarning("No Selection", "Please select one or more files to apply metadata.")
            return

        # Get the fields that the user WANTS to update (i.e., the checked fields)
        gui_metadata_updates = self.get_metadata_values()
        
        # Check for autovoluming case: if 'volume' and 'volume_count' are checked
        volume_check = self.control_vars.get('check_volume', tk.IntVar()).get()
        volume_count_check = self.control_vars.get('check_volume_count', tk.IntVar()).get()

        # Check if the autovoluming condition is met
        if volume_check == 1 and volume_count_check == 1 and 'volume' in self.control_vars:
            # Autovoluming is checked and active. We need to iterate and apply the volume sequentially.
            try:
                # The bulk form 'volume' field contains the starting number
                start_num = int(self.control_vars['volume'].get())
            except ValueError:
                messagebox.showerror("Invalid Start Num", "Please set a valid positive integer in the 'Volume #' field for autovoluming.")
                return
            
            # Remove volume number from the bulk update dictionary so we can set it per-file
            if 'volume' in gui_metadata_updates:
                del gui_metadata_updates['volume'] 
            
            gui_metadata_updates['volume_count'] = str(len(selected_indices))
            
            autovolume_mode = True
        else:
            autovolume_mode = False

        
        # Final check if any field is ticked (including the volume/volume_count pair)
        if not gui_metadata_updates and not autovolume_mode:
            messagebox.showwarning("No Fields Selected", "No metadata fields are ticked. Nothing will be updated.")
            return

        result = messagebox.askyesno(
            "Confirm",
            f"Apply metadata to {len(selected_indices)} file(s)?\n\n"
            "Only **TICKED** fields will be updated/merged. Unticked fields will be **PRESERVED**."
        )
        
        if not result:
            return
        
        self.file_listbox.config(state=tk.DISABLED)
        self.btn_apply.config(state=tk.DISABLED)
        self.progress_bar.config(mode='determinate', maximum=len(selected_indices))
        self.progress_bar.pack(fill=tk.X, pady=(0, 5))
        
        success_count = 0
        error_count = 0
        errors = []
        
        current_num = start_num if autovolume_mode else 0

        # Iterate over the selected indices in the display order
        for i, list_index in enumerate(selected_indices):
            file_path = self.files[list_index] # Get the file path from the underlying list
            self.update_status(f"Processing file {i+1}/{len(selected_indices)}: {os.path.basename(file_path)}")
            self.root.update_idletasks()
            
            try:
                editor = ComicMetadataEditor(file_path)
                
                # 1. Read existing metadata (for preservation)
                existing_metadata = editor.read_metadata()
                
                # 2. Merge: Start with existing data, then overwrite/add only the new, checked values
                merged_metadata = existing_metadata.copy()
                merged_metadata.update(gui_metadata_updates)

                # 3. Apply sequential volume number if in autovolume mode
                if autovolume_mode:
                    merged_metadata['volume'] = str(current_num)
                    current_num += 1

                # 4. Write the merged result
                new_path_str = editor.write_metadata(merged_metadata)
                
                if new_path_str:
                    success_count += 1
                    if new_path_str != file_path:
                        # File type changed (CBR -> CBZ), update the list
                        self.files[list_index] = new_path_str
                else:
                    error_count += 1
                    errors.append(f"{os.path.basename(file_path)}: Write failed")
            except Exception as e:
                error_count += 1
                errors.append(f"{os.path.basename(file_path)}: {str(e)}")
            
            self.progress_bar['value'] = i + 1
            self.root.update_idletasks()
        
        self.file_listbox.config(state=tk.NORMAL)
        self._update_listbox_display(selected_indices) # Refresh display and re-select
        
        self.btn_apply.config(state=tk.NORMAL)
        self.progress_bar.pack_forget() 
        self.progress_bar['value'] = 0
        
        msg = f"Successfully updated: {success_count}\nFailed: {error_count}"
        if errors:
            msg += "\n\nErrors:\n" + "\n".join(errors[:5])
            if len(errors) > 5:
                msg += f"\n... and {len(errors) - 5} more"
        
        if error_count == 0:
            if WINSOUND_AVAILABLE: winsound.MessageBeep(winsound.MB_ICONASTERISK)
            messagebox.showinfo("Success", msg)
        else:
            if WINSOUND_AVAILABLE: winsound.MessageBeep(winsound.MB_ICONWARNING)
            messagebox.showwarning("Completed with Errors", msg)
        
        self.update_status()

    def on_file_select(self, event):
        """Handle selection changes in the file listbox."""
        selected_indices = self.file_listbox.curselection()
        
        is_single_selection = len(selected_indices) == 1
        
        # Update 'View Metadata' button state based on selection count
        if self.view_metadata_btn:
            self.view_metadata_btn.config(state=tk.NORMAL if is_single_selection else tk.DISABLED)
        
        # Update 'Import Metadata' button state based on selection count
        if self.copy_all_btn:
            self.copy_all_btn.config(state=tk.NORMAL if is_single_selection else tk.DISABLED)
            
        self.update_status()


def main():
    """Main function to start the GUI application."""
    root = tk.Tk()
    app = ComicMetadataGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()