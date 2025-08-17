# Cryo Label Studio

Cryo Label Studio is a simple, streamlined label-making application for cryogenic vial labels, microcentrifuge labels, and more.
It generates .docx files of fully formatted label sheets using built-in templates, designed to work best with Microsoft Word for easy printing. 
You can pull label data directly from CSV or Excel files, create a series of incremented labels with or without a prefix, or produce a full sheet of identical labels in just a few clicks.
---

## ✨ Features

- **Preset-based workflow** – Save and load label format settings instantly.
- **Supports CSV, Excel, or direct text input**.
- **Serial or identical labels** – Generate numbered sets or repeat the same label.
- **Multiple label templates** – Includes Cryo Babies, Cryo Tags, and Tough Spots.  Developers can customize to add additional templates.
- **Partial sheet printing** – Start anywhere on a sheet to avoid waste.
- **Export to DOCX** – Print from any standard printer. Works best with Microsoft Word.
- **User-friendly UI** – Minimal clutter, quick to learn.
---


## 🖥 System Requirements

- **Windows 10 or 11** (64-bit)
- Microsoft **Visual C++ 2015–2022 Redistributable (x64)**
-  **Microsoft Word** (for best results when opening and printing .docx label sheets)
-  *(Other word processors, such as LibreOffice, may work but can sometimes cause extra pages or formatting issues.)*
- Standard printer capable of printing label sheets
---

## 🚀 Quick Start

1. Install and launch **Cryo Label Studio**.
2. Select a preset from the dropdown, or create a new one.
3. Enter your label text or load a CSV/XLSX file.
5. Save as DOCX and print.
---

## 📂 Presets & Templates

The app comes with built-in presets and label sheet templates, including formats for **Cryo Babies**, **Cryo Tags**, and **Tough Spots** labels.

You can create your own presets in the app to generate the exact labels you need.  
If you want to use the app with other types of labels, you can also add new templates by editing the source code and rebuilding the app yourself. Developers are welcome to fork this repository, add additional templates, and customize the app for their own workflows.

## 🧩 Adding a New Template

Cryo Label Studio comes with built-in templates for Cryo Babies, Cryo Tags, and Tough Spots labels.  
You can add your own templates to use with other label types.

**Steps to add a new template:**

1. **Create a blank template in Word**  
   - Set up the page layout, table grid, and margins so they match your label sheet.
   - Save the file as a `.docx` in the `templates/` folder of the source code.  
     Example: `templates/MYNEW-1000.docx`

2. **Edit `label_templates.py`**  
   - Open `label_templates.py` in a text editor.
   - Add a new entry to the `label_templates` dictionary with the following fields:
     ```python
     "MYNEW-1000": {
         "display_name": "My New Label 1.00 x 0.50",
         "template_path": "templates/MYNEW-1000.docx",
         "label_width": 1.00,
         "labels_across": 5,
         "labels_down": 20,
         "chars_per_line": 30,
         "lines_per_label": 5,
         "default_font_size": 8,
         "table_format": "checkerboard",  # or "LSL stripes"
         "needs_page_break": False,
     },
     ```
     Adjust the values for your label’s dimensions and layout.

3. **Rebuild the app**  
   - If running from source:  
     ```
     python gui.py
     ```
   - If creating a standalone EXE: rebuild with PyInstaller:
     ```
     pyinstaller --clean --noconfirm gui.spec

     ```

Once rebuilt, your new template will appear in the template dropdown inside the app.

---

## 🛠 Development

This app is built with:

- **Python** (3.12+ recommended)
- **Tkinter** (GUI)
- **python-docx** and **openpyxl**
- **PyInstaller** (packaging)

### Running from source
```bash
pip install -r requirements.txt
python gui.py
```

### 🚧 Known Limitations

Only Cryo Babies and Tough Spots templates are included for now.

Installer isn’t signed → expect a SmartScreen warning.

### 🤝 Contributing

Bug reports and feature requests welcome under Issues.

Pull requests welcome if you want to extend functionality.

### 📝 License

Released as free software under the MIT License.

Donations/support welcome at [my ko-fi](https://ko-fi.com/shannonbarrera).
