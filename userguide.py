import tkinter as tk
from tkinter import ttk, Canvas, Scrollbar
from tkhtmlview import HTMLLabel
import markdown

main_page_overview = """**Main Window Overview**

The Main Window is where you interact with your currently selected **Preset** to generate labels.<br>  
A Preset is a saved set of instructions that tells the app how to format and create your labels.  <br>
When you select a preset from the dropdown, the page will automatically update based on the preset type.

---

**Label Template**  <br>
Displays the selected label layout. This determines size and spacing of the labels.

**Quantity Radio Buttons**  <br>
If multiple quantities were configured in the preset (e.g. 1, 3, 5 labels per sample), they appear as radio buttons.  <br>
Choose one to select the quantity of each label. 

**Text Input Box** *(For Text Input Presets only)*  <br>
- **Identical**: Input the text which will appear on each label. The box is prefilled with the text entered in the preset editor.  <br>
- **Incremental**: Input the starting value for an incrementing sequence.<br>
Supported Serial Number Formats (12 digits maximum):<br>
  • Numbers only (e.g., 1234)<br>
  • Prefix (1–5 letters/numbers) + dash + digits (e.g., ab-123)<br>
  • Prefix + underscore + digits (e.g., xy_0999)<br>
  • Prefix (letters only) + digits (e.g., ab0001)<br>

**First Label Preview Box** *(For File Input Presets only)*  
When a file is loaded, this box displays a preview of the text for the first label.

**Load File Button** *(For File Input only)*  
Upload a `.csv` or `.xlsx` file with label data. A preview of the first label appears after loading.

**Partial Sheet Selector**  
Choose where on the label sheet to begin and end printing. Useful for using up partially used sheets.  
If the number of labels to generate exceeds the number of labels on the partial sheet, the remaining labels will roll over onto a full sheet.

**Pages of Labels Dropdown** *(For Serial logic only)*  
Choose how many full pages of serial labels to generate.

**Save Labels Button**  
Generates and saves the labels. You’ll choose a location and filename.  
The default filename chosen when creating the preset will be suggested.

**Footer Bar**  
The Footer Bar displays helpful information about the currently loaded preset.  
Depending on the preset type (Text Input or File Input), the footer may include:

- **Logic**: (Text Input only) Displays whether the logic is **Identical** or **Serial**.  
- **Sample File Name**: (File Input only) Displays the name of the uploaded CSV or Excel file, if available.  
- **Copies Per Label**: Displays the quantity of copies of each label to be generated.  

This bar is meant to help you quickly verify the parameters of the currently loaded preset.
"""

file_input_presets = """
**Creating and Editing File Input Presets**

To create a new File Input Preset:
Select **Presets > New File Input Preset** from the menu bar at the top of the main window.
This will open the File Input Preset Editor, where you can design a preset to create labels from a CSV or XLSX file.

---

**Preset Name**  
Enter the name of your preset. This name will appear in the **Preset Selector** dropdown on the main page.

**Label Sheet Template**  
Choose which label sheet template to use for formatting the labels.

**Copies Per Label**  
- A single number (e.g. `1`)  
- A range (e.g. `1-3`)  
- A comma-separated list (e.g. `1, 2, 3`)  
This will determine how many labels will be generated for each row of data.

**Date Format**  
Choose a date format from the dropdown to automatically format any dates found in the spreadsheet.  
Select **"Leave as is"** to use the date exactly as shown in the Excel/CSV file (without formatting changes).

**Font Name**  
Select the font to use for the label text.

**Font Size**  
Choose the font size for your labels.

**Label Text Alignment**  
Choose **Left**, **Center**, or **Right** alignment for your label text.

**Default Output Filename**  
Enter a default filename to be suggested when saving your labels.  
This helps save time and keep files organized.

**Add Datetime Stamp to Filename**  
Enable this to automatically append the current date and time to the suggested filename when saving.  
This ensures each file has a unique name and avoids overwriting.

**Partial Sheet Selection**  
Check this box to allow the user to select a starting and ending position on a label sheet — useful for using partial sheets.

**Color Scheme**  
Choose a color scheme to customize the appearance of the main app window when this preset is loaded.

---

**Upload Sample File**  
Upload a sample `.csv` or `.xlsx` file that contains the expected column headers for your data.  
The name of this file will be displayed in the footer on the main window when your preset is loaded.\n
The column headers will appear in the preset editor window as **buttons** that you can click to insert placeholders into the **Label Format** box.

---

**Label Format Box**  
This defines how each label will be constructed.
- Use curly braces `{}` to reference column headers.
- Click on a column header button to insert it into the format string.

**Example:**  
If your column is named `Sample ID`, inserting it will create: `{Sample ID}`

You can also slice part of a value using Python-style slicing:
- To get characters from position 6 onward: `{Sample ID}[6:]`
- To get the first 4 characters: `{Sample ID}[:4]`

These features allow for flexible formatting of your label text.

**Save Preset Button**
Saves the preset and closes the window.
"""

text_input_presets = """
**Creating and Editing Text Input Presets**

To create a new Text Input Preset:  
Select **Presets > New Text Input Preset** from the menu bar at the top of the main window.  
This will open the Text Input Preset Editor, where you can design a preset to create labels from manually entered text.

---

**Preset Name**  
Enter the name of your preset. This name will appear in the **Preset Selector** dropdown on the main page.

**Label Sheet Template**  
Choose which label sheet template to use for formatting the labels.

**Logic**  
- **Incremental**: Creates a series of labels incremented by one.  
- **Identical**: Creates multiple identical labels using the same text.

**Copies Per Label**  
Enter:  
- A single number (e.g. `1`)  
- A range (e.g. `1-3`)  
- A comma-separated list (e.g. `1, 2, 3`)  

For **Incremental** labels: This controls how many copies to print for each serial number.  
For **Identical** labels: This controls how many copies of the same text to print.  
Leave blank to auto-fill the first page.

**Font Name**  
Select the font to use for the label text.

**Font Size**  
Choose the font size for your labels.

**Label Text Alignment**  
Choose **Left**, **Center**, or **Right** alignment for your label text.

**Default Output Filename**  
Enter a default filename to be suggested when saving your labels.  
This helps save time and keeps files organized.

**Remove Duplicate Labels**
Check this box to make sure that if there are multiple lines in your spreadsheet that result in identical labels, only one set of each label will be generated.

**Add Datetime Stamp to Filename**  
Enable this to automatically append the current date and time to the suggested filename when saving.  
This ensures each file has a unique name and avoids overwriting.

**Partial Sheet Selection**  
Check this box to allow the user to select a starting and ending position on a label sheet.  
This is useful when working with sheets that have already been partially used.

**Color Scheme**  
Choose a color scheme to customize the appearance of the main app window when this preset is loaded.

---

**Label Text Format Box**  

For **Incremental** logic:  
Click the **{LABEL_TEXT}** button to insert the placeholder into the format box.  
Add blank lines and additional formatting around it as needed.

For **Identical** logic:  
You can enter default text here that will be pre-filled when loading the preset.  
Users will be able to customize or overwrite this text before generating the labels.

---

**Save Preset Button**  
Click this to save the preset and close the editor window.
"""



HELP_CONTENT = {
    "Main Window Overview": main_page_overview,
    "File Input Preset Editor": file_input_presets,
    "Text Input Preset Editor": text_input_presets,
    "About": "CryoPop Label Studio v1.0\nCreated by Shannon Barrera",
}



def show_help_window(parent):
    help_win = tk.Toplevel(parent)
    help_win.title("User Guide")
    help_win.geometry("800x700")
    help_win.resizable(True, True)

    # Sidebar
    sidebar = tk.Frame(help_win, width=80, bg="#f0f0f0")
    sidebar.pack(side="left", fill="y")

    # Right side with canvas + scrollbar
    right_container = tk.Frame(help_win)
    right_container.pack(side="right", fill="both", expand=True)

    canvas = tk.Canvas(right_container)
    canvas.pack(side="left", fill="both", expand=True)

    scrollbar = tk.Scrollbar(right_container, orient="vertical", command=canvas.yview)
    scrollbar.pack(side="right", fill="y")
    canvas.configure(yscrollcommand=scrollbar.set)

    # Create scrollable frame inside canvas
    scroll_frame = tk.Frame(canvas)
    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")

    def on_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))

    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _bind_to_mousewheel(event):
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def _unbind_from_mousewheel(event):
        canvas.unbind_all("<MouseWheel>")

    scroll_frame.bind("<Configure>", on_configure)
    scroll_frame.bind("<Enter>", _bind_to_mousewheel)
    scroll_frame.bind("<Leave>", _unbind_from_mousewheel)


    # HTMLLabel inside scroll_frame
    html_label = HTMLLabel(scroll_frame, html="", background="white")
    html_label.config(width=70, height=40, wrap="word")
    html_label.pack(fill="both", expand=True, padx=10, pady=10)


    # Load markdown content
    def load_topic(topic):
        raw_md = HELP_CONTENT.get(topic, "Select a topic from the menu.")
        html = markdown.markdown(raw_md)
        html_label.set_html(html)
        html_label.fit_height()  # Ensure it resizes height to match content

    # Sidebar buttons
    for topic in HELP_CONTENT:
        btn = tk.Button(sidebar, text=topic, anchor="w", width=25, relief="flat", bg="#f0f0f0",
                        command=lambda t=topic: load_topic(t))
        btn.pack(fill="x", pady=1)

    load_topic("Main Page Overview")



