#!/usr/bin/env python3
"""
Comic Metadata Bulk Editor - Python GUI Version
A tool to edit ComicInfo.xml metadata in CBZ/CBR files

Version 2.10: Strong Column Separation
- Increased the separator column minsize from 50 to 100 for a much clearer gap.
- Added extra left padding to the right-side elements to enhance separation.
"""

import os
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
        
        self.tip_window = tk.Toplevel(self.widget)
        self.tip_window.wm_overrideredirect(True) 
        
        label = tk.Label(self.tip_window, text=self.text, justify=tk.LEFT,
                         background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                         font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)
        
        self.tip_window.update_idletasks()
        
        tooltip_height = self.tip_window.winfo_height()
        widget_x = self.widget.winfo_rootx()
        widget_y = self.widget.winfo_rooty()
        
        # Position the tooltip above the widget
        x = widget_x + 10 
        y = widget_y - tooltip_height - 5 
        
        self.tip_window.wm_geometry(f"+{x}+{y}")

    def hide_tip(self, event=None):
        if self.tip_window:
            self.tip_window.destroy()
        self.tip_window = None


class ComicMetadataEditor:
    """Handles reading and writing ComicInfo.xml in comic archives"""
    
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
        'maincharacter': 'MainCharacterOrTeam', 'characters': 'Characters', 
        'teams': 'Teams', 'locations': 'Locations', 'notes': 'Notes', 'review': 'Review',
        'scaninformation': 'ScanInformation', 'web': 'Web',
        'communityrating': 'CommunityRating', 'gtin': 'GTIN', 'read': 'Read',
        'country': 'Country'
    }
    
    FIELD_MAPPING_REVERSE = {v: k for k, v in FIELD_MAPPING.items()}
    
    def __init__(self, file_path: str):
        self.file_path = Path(file_path)
        self.is_cbz = self.file_path.suffix.lower() == '.cbz'
        self.is_cbr = self.file_path.suffix.lower() == '.cbr'
        
        if not (self.is_cbz or self.is_cbr):
            raise ValueError("File must be .cbz or .cbr")

    def _parse_xml(self, xml_content: str) -> Dict[str, str]:
        data = {}
        try:
            root = ET.fromstring(xml_content)
            for elem in root:
                if elem.tag in self.FIELD_MAPPING_REVERSE:
                    key = self.FIELD_MAPPING_REVERSE[elem.tag]
                    data[key] = elem.text or ""
        except ET.ParseError as e:
            print(f"Error parsing XML for {self.file_path}: {e}")
        return data

    def read_metadata_cbz(self) -> Dict[str, str]:
        """Reads ComicInfo.xml from a CBZ file."""
        try:
            with zipfile.ZipFile(self.file_path, 'r') as zf:
                xml_name = next((f for f in zf.namelist() if f.lower().endswith('comicinfo.xml')), None)
                if xml_name:
                    with zf.open(xml_name) as f:
                        xml_content = f.read().decode('utf-8')
                        return self._parse_xml(xml_content)
        except zipfile.BadZipFile:
            print(f"Bad zip file: {self.file_path}")
        except Exception as e:
            print(f"Error reading CBZ {self.file_path}: {e}")
        return {}

    def read_metadata_cbr(self) -> Dict[str, str]:
        """Reads ComicInfo.xml from a CBR file."""
        if not RARFILE_AVAILABLE:
            return {}
        try:
            with rarfile.RarFile(self.file_path, 'r') as rf:
                xml_name = next((f for f in rf.namelist() if f.lower().endswith('comicinfo.xml')), None)
                if xml_name:
                    with rf.open(xml_name) as f:
                        xml_content = f.read().decode('utf-8')
                        return self._parse_xml(xml_content)
        except rarfile.BadRarFile:
            print(f"Bad rar file: {self.file_path}")
        except Exception as e:
            print(f"Error reading CBR {self.file_path}: {e}")
        return {}
        
    def read_metadata(self) -> Dict[str, str]:
        if self.is_cbz:
            return self.read_metadata_cbz()
        elif self.is_cbr:
            return self.read_metadata_cbr()
        return {}

    def _create_xml(self, metadata: Dict) -> str:
        filtered_metadata = {
            k: v for k, v in metadata.items() 
            if v is not None and v is not False and (str(v).strip() or str(v) == '0')
        }
        
        root = ET.Element('ComicInfo')
        root.text = '\n  '
        root.tail = '\n'
        
        last_elem = None
        for key in sorted(filtered_metadata.keys()):
            value = filtered_metadata[key]
            if key in self.FIELD_MAPPING:
                xml_tag = self.FIELD_MAPPING[key]
                elem = ET.SubElement(root, xml_tag)
                elem.text = str(value)
                elem.tail = '\n  '
                last_elem = elem
        
        if last_elem is not None:
            last_elem.tail = '\n'
        
        xml_str = '<?xml version="1.0" encoding="utf-8"?>\n'
        xml_str += ET.tostring(root, encoding='unicode')
        return xml_str
    
    def _write_file_to_archive(self, file_path: Path, new_content: Optional[str] = None) -> bool:
        temp_path = ""
        try:
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.cbz')
            temp_path = temp_file.name
            temp_file.close()
            
            with zipfile.ZipFile(file_path, 'r') as zf_in:
                with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as zf_out:
                    xml_name = next((f.filename for f in zf_in.infolist() if f.filename.lower().endswith('comicinfo.xml')), None)
                    
                    for item in zf_in.infolist():
                        if item.filename != xml_name:
                            data = zf_in.read(item.filename)
                            zf_out.writestr(item, data)
                    
                    if new_content:
                        zf_out.writestr('ComicInfo.xml', new_content)
            
            shutil.move(temp_path, file_path)
            return True
        except zipfile.BadZipFile:
            print(f"Bad zip file for modification: {file_path}")
            return False
        except Exception as e:
            print(f"Error modifying archive {file_path}: {e}")
            if os.path.exists(temp_path):
                os.unlink(temp_path)
            return False

    def write_metadata_cbz(self, metadata: Dict) -> bool:
        xml_content = self._create_xml(metadata)
        return self._write_file_to_archive(self.file_path, xml_content)

    def write_metadata_cbr(self, metadata: Dict) -> Optional[Path]:
        if not RARFILE_AVAILABLE: return None
        temp_dir = ""
        try:
            xml_content = self._create_xml(metadata)
            temp_dir = tempfile.mkdtemp()
            temp_cbz = os.path.join(temp_dir, 'temp.cbz')
            
            with rarfile.RarFile(self.file_path, 'r') as rf:
                with zipfile.ZipFile(temp_cbz, 'w', zipfile.ZIP_DEFLATED) as zf:
                    xml_name = next((f.filename for f in rf.infolist() if f.filename.lower().endswith('comicinfo.xml')), None)
                    for item in rf.infolist():
                        if not item.filename.endswith('/') and item.filename != xml_name:
                            data = rf.read(item.filename)
                            zf.writestr(item.filename, data)
                    zf.writestr('ComicInfo.xml', xml_content)
            
            new_path = self.file_path.with_suffix('.cbz')
            shutil.move(temp_cbz, new_path)
            if self.file_path.exists(): os.unlink(self.file_path)
            shutil.rmtree(temp_dir, ignore_errors=True)
            return new_path
        except Exception as e:
            print(f"Error writing metadata to {self.file_path}: {e}")
            if temp_dir and os.path.exists(temp_dir): shutil.rmtree(temp_dir, ignore_errors=True)
            return None
    
    def write_metadata(self, metadata_from_gui: Dict) -> Optional[str]:
        existing_data = self.read_metadata()
        final_data = existing_data
        final_data.update(metadata_from_gui)
        
        if self.is_cbz:
            if self.write_metadata_cbz(final_data):
                return str(self.file_path)
        elif self.is_cbr:
            new_path = self.write_metadata_cbr(final_data)
            if new_path:
                return str(new_path)
                
        return None
    
    def delete_metadata(self) -> Optional[str]:
        if self.is_cbz:
            if self._write_file_to_archive(self.file_path, new_content=None):
                return str(self.file_path)
        elif self.is_cbr:
            if not RARFILE_AVAILABLE: return None
            temp_dir = ""
            try:
                temp_dir = tempfile.mkdtemp()
                temp_cbz = os.path.join(temp_dir, 'temp.cbz')
                
                with rarfile.RarFile(self.file_path, 'r') as rf:
                    with zipfile.ZipFile(temp_cbz, 'w', zipfile.ZIP_DEFLATED) as zf:
                        xml_name = next((f.filename for f in rf.infolist() if f.filename.lower().endswith('comicinfo.xml')), None)
                        for item in rf.infolist():
                            if not item.filename.endswith('/') and item.filename != xml_name:
                                data = rf.read(item.filename)
                                zf.writestr(item.filename, data)
                
                new_path = self.file_path.with_suffix('.cbz')
                shutil.move(temp_cbz, new_path)
                if self.file_path.exists(): os.unlink(self.file_path)
                shutil.rmtree(temp_dir, ignore_errors=True)
                return str(new_path)
            except Exception as e:
                print(f"Error deleting metadata from {self.file_path}: {e}")
                if temp_dir and os.path.exists(temp_dir): shutil.rmtree(temp_dir, ignore_errors=True)
                return None

        return None


TOOLTIPS = {
    'series': "The title of the series (e.g., 'Land of the Lustrous').",
    'title': "The specific title of the issue (e.g., 'Volume 1').",
    'volume': "The volume number of the series this issue belongs to.",
    'number': "The issue number within the volume.",
    'issuecount': "The total number of issues in the volume (optional).",
    'year': "The year the issue was published.",
    'month': "The month the issue was published (1-12).",
    'day': "The day the issue was published (1-31).",
    'alternateSeries': "Title of an alternate or parallel series.",
    'alternatenumber': "Issue number in the alternate series.",
    'alternateissuecount': "Total number of issues in the alternate series.",
    'storyArc': "The name of the current story arc.",
    'seriesgroup': "The name of a larger group/universe the series belongs to.",
    'seriescomplete': "Is this series considered complete? (Yes/No).",
    'volume_count': "Total number of volumes in the series (optional).",
    'format': "The format of the comic (e.g., 'Oneshot', 'Trade Paperback').",
    'agerating': "The intended maturity rating (e.g., 'Teen', 'Mature').",
    'manga': "Is this Manga? (Yes/No/YesAndRightToLeft).",
    'publisher': "The publisher (e.g., 'Marvel Comics', 'Kodansha').",
    'imprint': "The specific imprint of the publisher.",
    'blackandwhite': "Is this a black and white comic?",
    'language': "The language of the comic (e.g., 'en', 'de').",
    'genre': "The primary genre (e.g., 'Sci-Fi', 'Fantasy').",
    'tags': "Comma-separated list of keywords/tags.",
    'writer': "The writer(s) of the story.",
    'penciller': "The penciller(s) or line artist(s).",
    'inker': "The inker(s).",
    'colorist': "The colorist(s).",
    'letterer': "The letterer(s).",
    'coverartist': "The cover artist(s).",
    'editor': "The editor(s).",
    'authorsort': "A key used for sorting by author (e.g., LastName, FirstName).",
    'summary': "A brief summary or description of the plot.",
    'maincharacter': "The main character or team.",
    'characters': "A list of other characters appearing.",
    'teams': "A list of teams/groups appearing.",
    'locations': "Key locations where the story takes place.",
    'notes': "Any private notes or internal information.",
    'review': "A public review or rating text.",
    'scaninformation': "Information about the scan or digital source.",
    'web': "Official website or link related to the comic.",
    'communityrating': 'A community rating (0-5 stars).',
    'gtin': 'Global Trade Item Number (e.g., ISBN or UPC).',
    'read': "Reading status (Yes/No).",
    'country': 'Country of origin/publication.'
}

class ComicMetadataGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Comic Metadata Bulk Editor (v2.10 - Starke Spalten-Trennung)")
        self.root.geometry("1200x900") 
        
        # --- Core State & Data ---
        # metadata stores: {'key': {'var': tk.BooleanVar, 'widget': tk.Widget}}
        self.metadata: Dict[str, Dict] = {} 
        # files stores a list of full paths of loaded files
        self.files: List[str] = [] 
        # file_script_status maps full path -> bool (True if modified in this session)
        self.file_script_status: Dict[str, bool] = {} 
        
        # --- UI Widgets (Initialized later in setup_ui) ---
        self.file_listbox: Optional[ttk.Treeview] = None 
        self.status_label: Optional[ttk.Label] = None
        self.btn_apply: Optional[ttk.Button] = None
        self.progress_bar: Optional[ttk.Progressbar] = None

        # --- Status & Flags ---
        self._explanation_shown: bool = False
        
        self.setup_ui()
        self.show_startup_explanation()
    
    # --- UI Setup and Layout ---
    
    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        title = ttk.Label(main_frame, text="Comic Metadata Bulk Editor", 
                         font=('Arial', 16, 'bold'))
        title.grid(row=0, column=0, columnspan=2, pady=10)
        
        # --- File Selection and List Frame (Row 1) ---
        file_frame = ttk.LabelFrame(main_frame, text="Files", padding="10")
        file_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # Make the file_frame resizable (Vertical weight 1)
        main_frame.rowconfigure(1, weight=1) 
        
        btn_frame = ttk.Frame(file_frame)
        btn_frame.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        ttk.Button(btn_frame, text="Select Files", 
                  command=self.select_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Select Folder", 
                  command=self.select_folder).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Select All", 
                  command=self.select_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Deselect All", 
                  command=self.deselect_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Remove Selected", 
                  command=self.remove_selected).pack(side=tk.LEFT, padx=5)
        
        list_frame = ttk.Frame(file_frame)
        list_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        file_frame.columnconfigure(0, weight=1)
        file_frame.rowconfigure(1, weight=1) # Make the list_frame resizable
        
        # --- Treeview Setup ---
        self.file_listbox = self.setup_treeview(list_frame)
        
        self.status_label = ttk.Label(file_frame, text="0 files loaded, 0 selected")
        self.status_label.grid(row=2, column=0, sticky=tk.W, pady=5)
        
        # --- Metadata Notebook (Tabs) (Row 2) ---
        notebook = ttk.Notebook(main_frame)
        # Row 2 is NOT configured with weight=1, allowing Row 1 (File List) to take the majority of the extra space.
        notebook.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        main_frame.columnconfigure(0, weight=1)
        
        tab1 = ttk.Frame(notebook)
        notebook.add(tab1, text="Main Information")
        self.create_main_info_tab(tab1) 
        
        tab2 = ttk.Frame(notebook)
        notebook.add(tab2, text="Artists & People")
        self.create_people_tab(tab2) 
        
        tab3 = ttk.Frame(notebook)
        notebook.add(tab3, text="Plot & Notes")
        self.create_plot_tab(tab3) 
        
        tab4 = ttk.Frame(notebook)
        notebook.add(tab4, text="Format & Details")
        self.create_format_tab(tab4) 
        
        # --- Utility Toolbar (Row 3) ---
        self.setup_toolbar(main_frame) 
        
        # --- Apply/Progress Bar Frame (Row 4) ---
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        bottom_frame.columnconfigure(0, weight=1)

        self.btn_apply = ttk.Button(bottom_frame, text="Apply Metadata to Selected Files",
                              command=self.apply_metadata)
        self.btn_apply.grid(row=0, column=0, sticky=tk.E, padx=5)

        self.progress_bar = ttk.Progressbar(bottom_frame, orient='horizontal', mode='determinate')
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 220))
        self.progress_bar['value'] = 0

    def setup_treeview(self, parent_frame):
        """Sets up the Treeview widget with custom columns."""
        
        scrollbar = ttk.Scrollbar(parent_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        tree = ttk.Treeview(parent_frame, columns=('last_changed', 'script_changed'), 
                            show='tree headings', selectmode='extended', 
                            yscrollcommand=scrollbar.set)
        
        tree.heading('#0', text='File Name', anchor=tk.W)
        tree.column('#0', width=600, anchor=tk.W, stretch=tk.YES)
        
        tree.heading('last_changed', text='Last Changed', anchor=tk.W)
        tree.column('last_changed', width=180, anchor=tk.W, stretch=tk.NO)
        
        tree.heading('script_changed', text='Script Changed', anchor=tk.CENTER)
        tree.column('script_changed', width=120, anchor=tk.CENTER, stretch=tk.NO)
        
        tree.bind('<<TreeviewSelect>>', lambda e: self.update_status())
        tree.tag_configure('modified', background='#e0ffe0') 
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=tree.yview)
        
        return tree
        
    def setup_toolbar(self, parent_frame):
        """Sets up the toolbar at the new location (below the notebook, row 3)."""
        toolbar_frame = ttk.Frame(parent_frame, padding="5 5 5 5", relief=tk.RIDGE)
        toolbar_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(toolbar_frame, text="Utility Toolbar:").pack(side=tk.LEFT, padx=5)

        btn_clear = ttk.Button(toolbar_frame, text="Clear ALL Fields", 
                               command=self.clear_all_fields)
        btn_clear.pack(side=tk.LEFT, padx=15, pady=0)
        ToolTip(btn_clear, "Clears all text/value fields in the tabs (keeps checkboxes unchecked).")

        ttk.Separator(toolbar_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        style = ttk.Style()
        style.configure('Danger.TButton', foreground='red', font=('Arial', 10, 'bold'))
        btn_delete = ttk.Button(toolbar_frame, text="DELETE Metadata from File", 
                                command=self.delete_metadata_from_selected_files,
                                style='Danger.TButton')
        btn_delete.pack(side=tk.LEFT, padx=15, pady=0)
        ToolTip(btn_delete, "REMOVES the ComicInfo.xml file from all selected archives. DANGEROUS!")

    # --- Reintegrated Metadata Widget Creation (Essential for startup) ---

    def _create_metadata_widget(self, parent_frame, key, label_text, row, col, widget_type='entry', widget_opts=None, columnspan=3, row_span=1):
        """
        Creates the label, control checkbox, and input widget for a metadata field.
        The layout uses: Checkbox (col), Label (col+1), Widget (col+2 to col+columnspan-1).
        """
        # --- Checkbox and Label ---
        check_var = tk.BooleanVar(value=False)
        self.metadata[key] = {'var': check_var} 
        
        # Custom padding for separation
        check_padx = (5, 0)
        label_padx = (5, 5)
        
        # If it's the right column (col=3), add extra left padding to the checkbox
        if col == 3:
            check_padx = (20, 0) 

        check_box = ttk.Checkbutton(parent_frame, variable=check_var)
        check_box.grid(row=row, column=col, sticky=tk.W, padx=check_padx, pady=4, rowspan=row_span)
        ToolTip(check_box, f"Tick this box to INCLUDE the value of the '{label_text}' field when writing metadata. If unticked, the original file's value is PRESERVED.")

        label = ttk.Label(parent_frame, text=label_text)
        label.grid(row=row, column=col + 1, sticky=tk.W, padx=label_padx, pady=4, rowspan=row_span)
        
        # --- Input Widget ---
        input_widget = None
        widget_opts = widget_opts or {}
        
        if widget_type == 'entry':
            input_widget = ttk.Entry(parent_frame, width=30) 
        elif widget_type == 'spinbox':
            input_widget = ttk.Spinbox(parent_frame, width=28, **widget_opts) 
        elif widget_type == 'combobox':
            input_widget = ttk.Combobox(parent_frame, width=28, **widget_opts) 
        elif widget_type == 'scrolledtext':
            # ScrolledText needs a frame wrapper for proper grid/pack interaction
            text_frame = ttk.Frame(parent_frame)
            # ScrolledText spans all 7 columns (0 to 6)
            text_frame.grid(row=row + 1, column=0, columnspan=7, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=4)
            input_widget = scrolledtext.ScrolledText(text_frame, height=5, wrap=tk.WORD)
            input_widget.pack(fill=tk.BOTH, expand=True)
            parent_frame.rowconfigure(row + 1, weight=1) 
            
        if input_widget:
            if widget_type != 'scrolledtext':
                # Input widget placement adjusted for the new 7-column layout (0-6):
                # Left Column starts at col=0, widget at col=2, spans 1 (column 2)
                # Right Column starts at col=3, widget at col=5, spans 2 (columns 5 and 6)
                if col == 0:
                     # Left column widget, spans 1 column
                     input_widget.grid(row=row, column=col + 2, sticky=(tk.W, tk.E), padx=(0, 10), pady=4, columnspan=1, rowspan=row_span)
                elif col == 3:
                     # Right column widget, spans 2 columns (col 5 & 6)
                     input_widget.grid(row=row, column=col + 2, sticky=(tk.W, tk.E), padx=(0, 10), pady=4, columnspan=2, rowspan=row_span)
                
            self.metadata[key]['widget'] = input_widget
            if key in TOOLTIPS:
                ToolTip(label, TOOLTIPS[key])
                if widget_type != 'scrolledtext':
                    ToolTip(input_widget, TOOLTIPS[key])
        
        return input_widget

    def _create_scrollable_tab(self, tab):
        """Standard boilerplate for creating a scrollable frame within a tab."""
        canvas = tk.Canvas(tab, borderwidth=0)
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, padding="10")

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Configure the grid in the scrollable frame for 7 columns:
        # Col 0: Checkbox (Left)
        # Col 1: Label (Left)
        # Col 2: Widget (Left) - Takes weight 1
        # Col 3: Separator (Spacer) - Increased minsize for strong separation
        # Col 4: Checkbox (Right)
        # Col 5: Label (Right)
        # Col 6: Widget (Right) - Takes weight 1
        
        scrollable_frame.columnconfigure(2, weight=1) # Left widget column
        scrollable_frame.columnconfigure(3, minsize=100) # **Physical Separator Column (Increased)**
        scrollable_frame.columnconfigure(6, weight=1) # Right widget column
        
        return scrollable_frame

    def create_main_info_tab(self, tab):
        scrollable_frame = self._create_scrollable_tab(tab)

        # Left Column (Core Info) - Uses columns 0, 1, 2
        self._create_metadata_widget(scrollable_frame, 'series', 'Series:', 0, 0, columnspan=3)
        self._create_metadata_widget(scrollable_frame, 'title', 'Title:', 1, 0, columnspan=3)
        
        # Right Column (Publisher/Imprint) - Uses columns 3, 4, 5, 6
        self._create_metadata_widget(scrollable_frame, 'publisher', 'Publisher:', 0, 3, columnspan=4)
        self._create_metadata_widget(scrollable_frame, 'imprint', 'Imprint:', 1, 3, columnspan=4)
        
        # Spacer Row 2
        ttk.Separator(scrollable_frame, orient='horizontal').grid(row=2, column=0, columnspan=7, sticky=(tk.W, tk.E), pady=10)
        
        # Left Column (Issue Details)
        self._create_metadata_widget(scrollable_frame, 'volume', 'Volume:', 3, 0, columnspan=3)
        self._create_metadata_widget(scrollable_frame, 'volume_count', 'Total Volumes:', 4, 0, columnspan=3)
        self._create_metadata_widget(scrollable_frame, 'number', 'Issue Number:', 5, 0, columnspan=3)
        self._create_metadata_widget(scrollable_frame, 'issuecount', 'Issue Count:', 6, 0, columnspan=3)
        
        # Right Column (Date)
        self._create_metadata_widget(scrollable_frame, 'year', 'Year:', 3, 3, columnspan=4)
        self._create_metadata_widget(scrollable_frame, 'month', 'Month:', 4, 3, columnspan=4)
        self._create_metadata_widget(scrollable_frame, 'day', 'Day:', 5, 3, columnspan=4)
        
        # Spacer Row 7
        ttk.Separator(scrollable_frame, orient='horizontal').grid(row=7, column=0, columnspan=7, sticky=(tk.W, tk.E), pady=10)

        # Bottom Row (Completion/Format Flags)
        self._create_metadata_widget(scrollable_frame, 'seriescomplete', 'Series Complete:', 8, 0, 'combobox', 
                                     {'values': ['', 'Yes', 'No'], 'state': 'readonly'}, columnspan=3)
        self._create_metadata_widget(scrollable_frame, 'manga', 'Manga Format:', 8, 3, 'combobox', 
                                     {'values': ['', 'Yes', 'No', 'YesAndRightToLeft'], 'state': 'readonly'}, columnspan=4)


    def create_people_tab(self, tab):
        scrollable_frame = self._create_scrollable_tab(tab)
        
        # Left Column (Core Roles) - Columns 0, 1, 2
        self._create_metadata_widget(scrollable_frame, 'writer', 'Writer(s):', 0, 0, columnspan=3)
        self._create_metadata_widget(scrollable_frame, 'penciller', 'Penciller(s):', 1, 0, columnspan=3)
        self._create_metadata_widget(scrollable_frame, 'inker', 'Inker(s):', 2, 0, columnspan=3)
        self._create_metadata_widget(scrollable_frame, 'colorist', 'Colorist(s):', 3, 0, columnspan=3)
        
        # Right Column (Supporting Roles) - Columns 3, 4, 5, 6
        self._create_metadata_widget(scrollable_frame, 'letterer', 'Letterer(s):', 0, 3, columnspan=4)
        self._create_metadata_widget(scrollable_frame, 'coverartist', 'Cover Artist(s):', 1, 3, columnspan=4)
        self._create_metadata_widget(scrollable_frame, 'editor', 'Editor(s):', 2, 3, columnspan=4)
        self._create_metadata_widget(scrollable_frame, 'authorsort', 'Author Sort Key:', 3, 3, columnspan=4)
        
        # Spacer Row 4
        ttk.Separator(scrollable_frame, orient='horizontal').grid(row=4, column=0, columnspan=7, sticky=(tk.W, tk.E), pady=10)

        # Left Column (Character/Team Info)
        self._create_metadata_widget(scrollable_frame, 'maincharacter', 'Main Character/Team:', 5, 0, columnspan=3)
        self._create_metadata_widget(scrollable_frame, 'characters', 'Other Characters:', 6, 0, columnspan=3)
        self._create_metadata_widget(scrollable_frame, 'teams', 'Teams/Groups:', 7, 0, columnspan=3)
        self._create_metadata_widget(scrollable_frame, 'locations', 'Locations:', 8, 0, columnspan=3)
    
    def create_plot_tab(self, tab):
        # This tab uses the standard scrollable frame but the widgets span all columns
        scrollable_frame = self._create_scrollable_tab(tab)
        
        def _create_full_span_scrolled_text(parent_frame, key, label_text, row):
            """Helper for ScrolledText to span all 7 columns."""
            check_var = tk.BooleanVar(value=False)
            self.metadata[key] = {'var': check_var} 
            
            check_box = ttk.Checkbutton(parent_frame, variable=check_var)
            check_box.grid(row=row, column=0, sticky=tk.W, padx=(5, 0), pady=4)
            ToolTip(check_box, f"Tick this box to INCLUDE the value of the '{label_text}' field when writing metadata. If unticked, the original file's value is PRESERVED.")

            label = ttk.Label(parent_frame, text=label_text)
            label.grid(row=row, column=1, sticky=tk.W, padx=5, pady=4)
            
            # The Text widget frame spans all 7 columns (0 to 6)
            text_frame = ttk.Frame(parent_frame)
            text_frame.grid(row=row + 1, column=0, columnspan=7, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=4)
            input_widget = scrolledtext.ScrolledText(text_frame, height=5, wrap=tk.WORD)
            input_widget.pack(fill=tk.BOTH, expand=True)
            parent_frame.rowconfigure(row + 1, weight=1) 
            
            self.metadata[key]['widget'] = input_widget
            if key in TOOLTIPS:
                ToolTip(label, TOOLTIPS[key])
            
            return input_widget
            
        _create_full_span_scrolled_text(scrollable_frame, 'summary', 'Summary:', 0)
        _create_full_span_scrolled_text(scrollable_frame, 'notes', 'Notes:', 3)
        _create_full_span_scrolled_text(scrollable_frame, 'review', 'Review:', 6)

        # Standard Entries (below text boxes, spanning all 7 columns)
        # We use column 0 and a high columnspan to force it across all 7 columns
        self._create_metadata_widget(scrollable_frame, 'tags', 'Tags (Comma-Separated):', 9, 0, columnspan=7)
        self._create_metadata_widget(scrollable_frame, 'scaninformation', 'Scan Information:', 10, 0, columnspan=7)


    def create_format_tab(self, tab):
        scrollable_frame = self._create_scrollable_tab(tab)
        
        # Left Column (Format/Identifiers) - Columns 0, 1, 2
        self._create_metadata_widget(scrollable_frame, 'format', 'Format:', 0, 0, columnspan=3)
        self._create_metadata_widget(scrollable_frame, 'gtin', 'GTIN (ISBN/UPC):', 1, 0, columnspan=3)
        self._create_metadata_widget(scrollable_frame, 'web', 'Web Link:', 2, 0, columnspan=3)
        self._create_metadata_widget(scrollable_frame, 'genre', 'Genre:', 3, 0, columnspan=3) 
        self._create_metadata_widget(scrollable_frame, 'country', 'Country:', 4, 0, columnspan=3)
        
        # Right Column (Ratings/Flags) - Columns 3, 4, 5, 6
        self._create_metadata_widget(scrollable_frame, 'communityrating', 'Community Rating (0-5):', 0, 3, 'spinbox', 
                                     {'from_': 0, 'to': 5, 'increment': 0.1}, columnspan=4)
        self._create_metadata_widget(scrollable_frame, 'agerating', 'Maturity Rating:', 1, 3, 'combobox', 
                                     {'values': ['', 'Adults Only', 'Early Childhood', 'Everyone', 'Everyone 10+', 'G', 'Kids', 'M', 'MA15+', 'Mature', 'PG', 'PG-13', 'Teen', 'T+', 'X', 'Young Adult'], 'state': 'readonly'}, columnspan=4)
        self._create_metadata_widget(scrollable_frame, 'language', 'Language ISO:', 2, 3, 'combobox', 
                                     {'values': ['', 'English', 'German', 'Spanish', 'French', 'Japanese', 'Korean', 'Chinese (Simplified)', 'Vietnamese'], 'state': 'editable'}, columnspan=4)
                                     
        # Alternate Series / Group
        self._create_metadata_widget(scrollable_frame, 'storyArc', 'Story Arc Name:', 5, 0, columnspan=3)
        self._create_metadata_widget(scrollable_frame, 'seriesgroup', 'Series Group Name:', 6, 0, columnspan=3)
        self._create_metadata_widget(scrollable_frame, 'alternateSeries', 'Alternate Series:', 5, 3, columnspan=4)
        self._create_metadata_widget(scrollable_frame, 'alternatenumber', 'Alternate Issue No.:', 6, 3, columnspan=4)
        self._create_metadata_widget(scrollable_frame, 'alternateissuecount', 'Alternate Issue Count:', 7, 3, columnspan=4)

        # Spacer Row 8
        ttk.Separator(scrollable_frame, orient='horizontal').grid(row=8, column=0, columnspan=7, sticky=(tk.W, tk.E), pady=10)

        # Checkboxes (Boolean Flags)
        self._create_metadata_widget(scrollable_frame, 'blackandwhite', 'Black & White:', 9, 0, widget_type='entry', columnspan=3) 
        self._create_metadata_widget(scrollable_frame, 'read', 'Read Status:', 9, 3, widget_type='entry', columnspan=4) 

    # --- New Helper Methods for File Handling ---
    
    def _get_file_attributes(self, file_path: str) -> Dict:
        """Calculates and formats file attributes for Treeview."""
        try:
            mod_time_timestamp = os.path.getmtime(file_path)
            mod_time_str = datetime.datetime.fromtimestamp(mod_time_timestamp).strftime('%Y-%m-%d %H:%M')
        except:
            mod_time_str = 'N/A'
        
        # Check if the script modified it in this session
        script_changed_status = "Yes" if self.file_script_status.get(file_path, False) else "-"
        
        return {
            'filename': os.path.basename(file_path),
            'last_changed': mod_time_str,
            'script_changed': script_changed_status
        }
    
    def _add_file_to_treeview(self, file_path: str):
        """Adds a new file to the internal list and Treeview."""
        if file_path not in self.files:
            self.files.append(file_path)
            
            attrs = self._get_file_attributes(file_path)
            
            tags = ('modified',) if self.file_script_status.get(file_path, False) else ()
            
            # Ensure file_listbox is initialized
            if self.file_listbox:
                self.file_listbox.insert(
                    '', 
                    tk.END, 
                    iid=file_path, 
                    text=attrs['filename'], 
                    values=(attrs['last_changed'], attrs['script_changed']),
                    tags=tags
                )

    def _update_treeview_item(self, file_path: str, new_file_path: Optional[str] = None):
        """Updates the Treeview item (e.g., after modification or CBR->CBZ conversion)."""
        
        current_iid = file_path
        new_iid = new_file_path if new_file_path else current_iid
        
        if not self.file_listbox: return

        # 1. Update internal state if path changed (CBR -> CBZ)
        if new_iid != current_iid:
            if current_iid in self.files:
                self.files.remove(current_iid)
            if new_iid not in self.files:
                self.files.append(new_iid)
            
            if current_iid in self.file_script_status:
                self.file_script_status[new_iid] = self.file_script_status.pop(current_iid)
                
            # Remove old item and insert new one to update IID
            if self.file_listbox.exists(current_iid):
                self.file_listbox.delete(current_iid)
            current_iid = new_iid
        
        # 2. Update status and display attributes
        if current_iid not in self.file_listbox.get_children():
            self.file_listbox.insert('', tk.END, iid=current_iid)
            
        self.file_script_status[current_iid] = True 
        attrs = self._get_file_attributes(current_iid)
        
        self.file_listbox.item(
            current_iid, 
            text=attrs['filename'], 
            values=(attrs['last_changed'], attrs.get('script_changed')),
            tags=('modified',)
        )
    
    def get_selected_files_paths(self) -> List[str]:
        """Returns the full paths of all selected items in the Treeview."""
        if self.file_listbox:
            return list(self.file_listbox.selection())
        return []
    
    # --- UI Action Methods ---

    def clear_all_fields(self):
        result = messagebox.askyesno(
            "Confirm Clear",
            "Are you sure you want to clear the content of ALL metadata fields in the GUI?"
        )
        if not result:
            if WINSOUND_AVAILABLE:
                winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS) 
            return

        for key, items in self.metadata.items():
            widget = items.get('widget')
            
            items['var'].set(False)
            
            if isinstance(widget, scrolledtext.ScrolledText):
                widget.config(state=tk.NORMAL)
                widget.delete("1.0", tk.END)
            elif isinstance(widget, ttk.Combobox):
                widget.set('')
            elif widget and hasattr(widget, 'delete'):
                widget.delete(0, tk.END)
                if hasattr(widget, 'insert'):
                     widget.insert(0, '') 
        
        if WINSOUND_AVAILABLE:
            winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS) 

        messagebox.showinfo("Fields Cleared", "All metadata fields have been cleared.")


    def delete_metadata_from_selected_files(self):
        selected_paths = self.get_selected_files_paths()
        if not selected_paths:
            messagebox.showwarning("No Files Selected", "Please select one or more files to delete metadata from.")
            return

        num_files = len(selected_paths)
        
        result1 = messagebox.askyesno(
            "!!! First Warning !!!",
            f"Are you ABSOLUTELY sure you want to DELETE the ComicInfo.xml metadata from {num_files} selected file(s)?\n\n"
            "This action PERMANENTLY removes the metadata and CANNOT be undone.",
            icon='warning'
        )
        if not result1:
            if WINSOUND_AVAILABLE: winsound.PlaySound("SystemExit", winsound.SND_ALIAS)
            return

        if WINSOUND_AVAILABLE:
            winsound.PlaySound("SystemCritical", winsound.SND_ALIAS) 
        result2 = messagebox.askyesno(
            "!!! CRITICAL FINAL WARNING !!!",
            f"CONFIRM DELETION: The ComicInfo.xml file will be PERMANENTLY REMOVED from {num_files} file(s).\n\n"
            "Click 'Yes' to proceed with irreversible deletion.",
            icon='error'
        )

        if not result2: return
        
        # UI Freeze and Progress Setup
        if self.btn_apply and self.file_listbox and self.progress_bar:
            self.btn_apply.config(state=tk.DISABLED)
            self.file_listbox.config(selectmode='none') 
            self.progress_bar['maximum'] = num_files
            self.progress_bar['value'] = 0
            self.root.update_idletasks()
        
        success_count = 0
        error_count = 0
        errors = []
        
        for i, file_path in enumerate(selected_paths):
            try:
                editor = ComicMetadataEditor(file_path)
                new_path_str = editor.delete_metadata()
                
                if new_path_str:
                    success_count += 1
                    self._update_treeview_item(file_path, new_path_str)
                else:
                    error_count += 1
                    errors.append(f"{os.path.basename(file_path)}: Deletion failed")
            except Exception as e:
                error_count += 1
                errors.append(f"{os.path.basename(file_path)}: {str(e)}")
            
            if self.progress_bar:
                self.progress_bar['value'] = i + 1
                self.root.update_idletasks()
            
        # UI Re-enable
        if self.file_listbox and self.btn_apply and self.progress_bar:
            self.file_listbox.config(selectmode='extended')
            self.btn_apply.config(state=tk.NORMAL)
            self.progress_bar['value'] = 0
        
        # Show results
        msg = f"Successfully deleted metadata from: {success_count}\nFailed: {error_count}"
        if errors:
            msg += "\n\nErrors:\n" + "\n".join(errors[:5])
            if len(errors) > 5:
                msg += f"\n... and {len(errors) - 5} more"
        
        if error_count == 0:
            messagebox.showinfo("Deletion Complete", msg)
        else:
            messagebox.showwarning("Deletion Completed with Errors", msg)
        
        self.update_status()

    # --- File Selection Methods ---

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Select Comic Files",
            filetypes=[("Comic Files", "*.cbz *.cbr"), ("All Files", "*.*")]
        )
        for file in files:
            self._add_file_to_treeview(file)
        self.update_status()
    
    def select_folder(self):
        folder = filedialog.askdirectory(title="Select Folder")
        if folder:
            file_glob = "*.cb[zr]" if RARFILE_AVAILABLE else "*.cbz"
            if not RARFILE_AVAILABLE:
                messagebox.showinfo("CBR Support Disabled",
                                    "rarfile library not found. Only scanning for .cbz files.")
                                    
            for file in Path(folder).glob(file_glob):
                self._add_file_to_treeview(str(file))
        self.update_status()
    
    def select_all(self):
        if self.file_listbox:
            self.file_listbox.selection_set(self.file_listbox.get_children())
        self.update_status()
    
    def deselect_all(self):
        if self.file_listbox:
            self.file_listbox.selection_remove(self.file_listbox.selection())
        self.update_status()
    
    def remove_selected(self):
        if self.file_listbox:
            selected = self.file_listbox.selection()
            for iid in selected:
                self.file_listbox.delete(iid)
                if iid in self.files:
                    self.files.remove(iid)
                if iid in self.file_script_status:
                    del self.file_script_status[iid]
        self.update_status()
    
    def update_status(self, event=None):
        total = len(self.files)
        selected = len(self.file_listbox.selection()) if self.file_listbox else 0
        if self.status_label:
            self.status_label.config(text=f"{total} files loaded, {selected} selected")
    
    # --- Apply Metadata Method ---

    def apply_metadata(self):
        selected_paths = self.get_selected_files_paths()
        
        if not selected_paths:
            messagebox.showwarning("No Files Selected", 
                                 "Please select one or more files to apply metadata to.")
            return
        
        metadata_values_to_write = self.get_metadata_values()
        
        if not metadata_values_to_write:
            messagebox.showwarning("Nothing to Do",
                                 "You didn't tick any checkboxes.\n"
                                 "Tick the box next to a field to include it in the write.")
            return

        result = messagebox.askyesno(
            "Confirm Bulk Write",
            f"Apply metadata to {len(selected_paths)} file(s)?\n\n"
            f"The following {len(metadata_values_to_write)} fields will be merged/overwritten:\n"
            f"- {', '.join(metadata_values_to_write.keys())}\n\n"
            "Fields with **unticked** boxes will be preserved from the original file."
        )
        
        if not result:
            return
        
        # UI Freeze and Progress Setup
        if self.btn_apply and self.file_listbox and self.progress_bar:
            self.btn_apply.config(state=tk.DISABLED)
            self.file_listbox.config(selectmode='none')
            self.progress_bar['maximum'] = len(selected_paths)
            self.progress_bar['value'] = 0
            self.root.update_idletasks()
        
        success_count = 0
        error_count = 0
        errors = []
        
        for i, file_path in enumerate(selected_paths):
            try:
                editor = ComicMetadataEditor(file_path)
                new_path_str = editor.write_metadata(metadata_values_to_write) 
                
                if new_path_str:
                    success_count += 1
                    self._update_treeview_item(file_path, new_path_str)
                else:
                    error_count += 1
                    errors.append(f"{os.path.basename(file_path)}: Write failed")
            except Exception as e:
                error_count += 1
                errors.append(f"{os.path.basename(file_path)}: {str(e)}")
            
            if self.progress_bar:
                self.progress_bar['value'] = i + 1
                self.root.update_idletasks()
            
        # UI Re-enable
        if self.file_listbox and self.btn_apply and self.progress_bar:
            self.file_listbox.config(selectmode='extended')
            self.file_listbox.selection_set(self.file_listbox.get_children()) 
            self.btn_apply.config(state=tk.NORMAL)
            self.progress_bar['value'] = 0
        
        # Show results
        msg = f"Successfully updated: {success_count}\nFailed: {error_count}"
        if errors:
            msg += "\n\nErrors:\n" + "\n".join(errors[:5])
            if len(errors) > 5:
                msg += f"\n... and {len(errors) - 5} more"
        
        if error_count == 0:
            messagebox.showinfo("Success", msg)
        else:
            messagebox.showwarning("Completed with Errors", msg)
        
        self.update_status()

    # --- Other Helper Methods ---
    
    def get_metadata_values(self) -> Dict:
        """
        Extract metadata values from UI, ONLY for fields
        where the control checkbox is ticked.
        """
        values = {}
        for key, items in self.metadata.items():
            widget = items.get('widget')
            check_var = items['var']
            
            if check_var.get():
                if key == 'language' and widget:
                    values[key] = self._map_language_to_iso(widget.get().strip())
                elif key in ['blackandwhite', 'read']:
                    values[key] = 'Yes' 
                elif isinstance(widget, scrolledtext.ScrolledText):
                    values[key] = widget.get("1.0", tk.END).strip()
                elif widget:
                    values[key] = widget.get().strip()
                    
        return values
    
    def _map_language_to_iso(self, language_name: str) -> str:
        mapping = {
            'English': 'en', 'Spanish': 'es', 'French': 'fr', 'German': 'de',
            'Italian': 'it', 'Japanese': 'ja', 'Korean': 'ko', 
            'Chinese (Simplified)': 'zh-Hans', 'Vietnamese': 'vi'
        }
        return mapping.get(language_name, language_name)
    
    def show_startup_explanation(self):
        """Displays a modal window explaining the new features."""
        if self._explanation_shown: return

        top = tk.Toplevel(self.root)
        top.title("Welcome to v2.10: Stronger Layout Separation")
        top.geometry("600x450")
        top.resizable(False, False)
        top.grab_set() 

        frame = ttk.Frame(top, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="New in v2.10: Enhanced Field Separation", font=('Arial', 14, 'bold')).pack(pady=10)

        st = scrolledtext.ScrolledText(frame, height=15, wrap=tk.WORD, font=('Arial', 10), relief=tk.FLAT)
        st.pack(pady=10, fill=tk.BOTH, expand=True)
        
        explanation = """
This version significantly increases the **visual separation** between the left and right field groups in the metadata tabs.

### Enhanced Separation Details
1.  **Separator Size:** The empty spacer column between the two field groups is now **100 pixels wide** (previously 50).
2.  **Right-Side Padding:** The elements in the right field group now have extra horizontal padding to ensure they start further away from the separator.

This guarantees a clearly visible, fixed gap between the two columns, even when the main window is maximized.
        """
        
        st.insert("1.0", explanation)
        st.config(state=tk.DISABLED)
        
        ttk.Button(frame, text="Got It! Start Editing", command=top.destroy).pack(pady=10)
        
        self.root.update_idletasks() 
        top.update_idletasks() 
        
        # Center the window
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_w = self.root.winfo_width()
        top_w = top.winfo_width()
        
        x = root_x + (root_w // 2) - (top_w // 2)
        y = root_y + 50 

        top.wm_geometry(f"+{x}+{y}")
        self._explanation_shown = True 


def main():
    root = tk.Tk()
    style = ttk.Style()
    style.configure('TFrame', padding=5)
    
    app = ComicMetadataGUI(root)
    root.mainloop()


if __name__ == '__main__':
    main()