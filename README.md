````markdown
# Comic Metadata Bulk Editor

A Python GUI tool for bulk-editing comic metadata inside `.cbz` and `.cbr` archives.  
The program writes and updates `ComicInfo.xml` files using the ComicRack metadata standard.

This tool is designed for users who want to mass-edit metadata for large comic libraries without manually opening each archive.

---

## Features

- **Batch Editing** – Apply metadata to dozens or hundreds of comics at once.
- **GUI Interface** – A multi-tab Tkinter interface.
- **CBZ Support** – Reads and writes metadata directly to `.cbz` files.
- **CBR Support** – Reads `.cbr` through the `rarfile` library and automatically converts them to `.cbz` when saving.
- **35+ Metadata Fields** – Includes series information, creators, plot details, publishing info, ratings, and more.

---

## Requirements

- Python 3.x  
- Tkinter (included with most Python installations)  
- `rarfile` (only required for reading `.cbr` files)

Install the dependency:

```bash
pip install rarfile
````

---

## Installation

1. Install Python 3.
    
2. Install the `rarfile` dependency (if you need `.cbr` support).
    
3. Download or clone this repository.
    
4. Run the tool:
    

```bash
python3 comic-editor-python-gui.py
```

---

## How to Use

1. Launch the program.
    
2. Click **Select Files** or **Select Folder** to load `.cbz` and `.cbr` files.
    
3. Select one or more comics from the file list.
    
4. Fill in the metadata fields you want to apply.
    

**Note:**  
Fields left empty will not overwrite existing data. Only filled fields are written or updated.

5. Click **Apply Metadata to Selected Files**.
    
6. A confirmation message will display how many files were successfully updated.
    

---

## ⚠️ Important: How Readers Detect Your Metadata

This editor writes metadata inside each archive as a `ComicInfo.xml` file.  
Different readers handle embedded metadata differently.

### KOReader (USB Transfer)

KOReader does **not** read the internal `ComicInfo.xml` by default.  
To enable metadata support:

- Install the **comicmeta.koplugin** on your KOReader device.
    

This plugin allows KOReader to read ComicRack metadata stored inside `.cbz` and `.cbr` files.  
Without it, KOReader ignores metadata embedded in the archive.

---

### Calibre

Calibre uses its own metadata database and does not automatically detect file changes.

After editing comics:

- Remove the book from Calibre.
    
- Re-add the updated file.
    

Calibre will read the new metadata and generate a `metadata.calibre` sidecar when sending files to devices.

---

### Other Readers (CDisplayEx, YACReader, etc.)

Most desktop comic readers natively read `ComicInfo.xml`.  
Your updated metadata should appear immediately.

---

## Supported Metadata Fields

### **Main Information**

- Title
    
- Series
    
- Volume
    
- Number
    
- Year
    
- Month
    
- Day
    
- Alternate Series
    
- Alternate Number
    
- Story Arc
    
- Series Group
    
- Series Complete
    
- Read Status
    

### **Artists & People**

- Writer
    
- Penciller
    
- Inker
    
- Colorist
    
- Letterer
    
- Cover Artist
    
- Editor
    
- Author Sort
    

### **Plot & Notes**

- Summary
    
- Main Character / Team
    
- Characters
    
- Teams
    
- Locations
    
- Scan Information
    
- Web
    
- Notes
    
- Review
    

### **Format & Details**

- Format
    
- Age Rating
    
- Publisher
    
- Imprint
    
- Language (ISO)
    
- Genre
    
- Tags
    
- ISBN
    
- Manga (Yes / No / Right-to-Left)
    
- Black and White
    
- Community Rating
    

---

## License

This project is licensed under the **MIT License**.

```

If you need badges, a logo, a screenshot section, or auto-generated TOC, I can append them.
```
