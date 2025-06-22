import tkinter as tk
from label_templates import label_templates
from tkinter import filedialog, messagebox, ttk
import json
import os
import csv
import uuid 
import openpyxl as xlsx


from .file_helpers import get_csv_headers, get_xlsx_headers
from .format_helpers import get_textbox_dimensions

DATE_FORMAT_DISPLAY_MAP = {
    # Four-digit year formats
    "MM-DD-YYYY": "%m-%d-%Y",
    "DD-MM-YYYY": "%d-%m-%Y",
    "YYYY-MM-DD": "%Y-%m-%d",
    "MM/DD/YYYY": "%m/%d/%Y",
    "DD/MM/YYYY": "%d/%m/%Y",
    "YYYY/MM/DD": "%Y/%m/%d",
    "Month DD, YYYY": "%B %d, %Y",   # January 01, 2025
    "DD Mon YYYY": "%d %b %Y",       # 01 Jan 2025

    # Two-digit year formats
    "MM-DD-YY": "%m-%d-%y",
    "DD-MM-YY": "%d-%m-%y",
    "YY-MM-DD": "%y-%m-%d",
    "MM/DD/YY": "%m/%d/%y",
    "DD/MM/YY": "%d/%m/%y",
    "YY/MM/DD": "%y/%m/%d"
}


class PresetEditor(tk.Toplevel):
    def __init__(self, master, preset_type="File", preset_data=None, preset_path=None, on_save=None):
        super().__init__(master)
        self.preset_type = self._resolve_type(preset_type, preset_data)
        self.preset_data = preset_data or {}
        self.sample_filename = self.preset_data.get("sample_filename")
        self.preset_path = preset_path
        self.on_save = on_save
        self.entries = {}
        self.header_buttons_frame = None
        self.textbox_format = None

        self._setup_window()
        self._init_template_maps()

        self._define_fields()
        self._create_fields_ui()
        self._handle_format_input_section()
        self._create_save_button()
        self.update_textbox_size()



    def _resolve_type(self, preset_type, preset_data):
        return preset_data.get("presettype") if preset_data and "presettype" in preset_data else preset_type

    def _setup_window(self):
        title = ("Text Input Preset Editor" if self.preset_type == "Text" else "File Input Preset Editor")
        self.title(title)
        if self.preset_type == "Text":
            self.geometry("500x655+70+1") 
        else:
            self.geometry("600x820+70+1")

    def _init_template_maps(self):
        self.template_display_map = {v["display_name"]: k for k, v in label_templates.items()}
        self.template_internal_map = {k: v["display_name"] for k, v in label_templates.items()}

    def _define_fields(self):
        self.fields = [
            ("name", "Preset Name"),
            ("labeltemplate", "Label Sheet Template"),
            ("copiesperlabel", "Copies Per Label"),
            ("fontname", "Font Name"),
            ("fontsize", "Font Size"),
            ("text_alignment", "Label Text Alignment"),
            ("outputfilenameprefix", "Default Output Filename"),
            ("output_add_date", "Add Datetime Stamp to Filename"),
            ("partialsheet", "Partial Sheet Selection"),
            ("color_theme", "Color Scheme")
        ]
        if self.preset_type == "Text":
            self.fields.insert(2, ("identical_or_incremental", "Logic"))
        
        if self.preset_type == "File":
            self.fields.insert(3, ("date_format", "Date Format"))
            self.fields.insert(8, ("remove_duplicates", "Remove Duplicate Labels"))


    def _create_fields_ui(self):
        field_row = 0
        for key, label in self.fields:
            tk.Label(self, text=label).grid(row=field_row, column=0, sticky="w", padx=10, pady=2)

            if key in ["partialsheet", "output_add_date", "remove_duplicates"]:
                var = tk.BooleanVar()
                var.set(self.preset_data.get(key, False))
                cb = tk.Checkbutton(self, variable=var)
                cb.grid(row=field_row, column=1, padx=10, pady=2, sticky="w")
                self.entries[key] = var

            elif key == "color_theme":
                cb = ttk.Combobox(self, values=["Grey", "Pink", "Green", "Blue", "Yellow", "Purple"], state="readonly")
                cb.set(self.preset_data.get(key, "Pink"))
                cb.grid(row=field_row, column=1, padx=10, pady=2)
                self.entries[key] = cb

            elif key == "text_alignment":
                cb = ttk.Combobox(self, values=["Left", "Center", "Right"], state="readonly")
                cb.set(self.preset_data.get(key, "Center"))
                cb.grid(row=field_row, column=1, padx=10, pady=2)
                self.entries[key] = cb


            elif key == "date_format":
                cb = ttk.Combobox(
                    self,
                    values=["Leave as is"] + list(DATE_FORMAT_DISPLAY_MAP.keys()),
                    state="readonly"
                )

                # Find display label that matches saved backend format
                saved_format = self.preset_data.get(key)
                if not saved_format:
                    display_label = "Leave as is"
                else:
                    display_label = next(
                        (label for label, fmt in DATE_FORMAT_DISPLAY_MAP.items() if fmt == saved_format),
                        "MM-DD-YYYY"
                    )


                cb.set(display_label)
                cb.grid(row=field_row, column=1, padx=10, pady=2)
                self.entries[key] = cb

            elif key == "labeltemplate":
                display_names = list(self.template_display_map.keys())
                cb = ttk.Combobox(self, values=display_names, state="readonly", width=27)
                cb.bind("<<ComboboxSelected>>", self.update_textbox_size)
                internal_value = self.preset_data.get(key)
                display_name = self.template_internal_map.get(internal_value, display_names[0])
                cb.set(display_name)
                cb.grid(row=field_row, column=1, padx=10, pady=2)
                self.entries[key] = cb

            elif key == "fontname":
                cb = ttk.Combobox(self, values=["Arial", "Courier", "Helvetica", "Times", "Verdana"], state="readonly")
                cb.set(self.preset_data.get(key, "Arial"))
                cb.grid(row=field_row, column=1, padx=10, pady=2)
                self.entries[key] = cb

            elif key == "fontsize":
                cb = ttk.Combobox(
                    self, values=["5.5", "6", "6.5", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20"],
                    state="readonly"
                )
                cb.set(str(self.preset_data.get(key, "6")))
                cb.bind("<<ComboboxSelected>>", self.update_textbox_size)
                cb.grid(row=field_row, column=1, padx=10, pady=2)
                self.entries[key] = cb

            elif key == "identical_or_incremental":
                cb = ttk.Combobox(self, values=["Identical", "Incremental"], state="readonly")
                cb.set(self.preset_data.get(key, "Identical"))
                cb.grid(row=field_row, column=1, padx=10, pady=2)
                self.entries[key] = cb


            elif key == "copiesperlabel":
                entry = tk.Entry(self, width=40)
                entry.insert(0, str(self.preset_data.get(key, "")))
                entry.grid(row=field_row, column=1, padx=10, pady=2)
                self.entries[key] = entry

            else:
                entry = tk.Entry(self, width=40)
                entry.insert(0, str(self.preset_data.get(key, "")))
                entry.grid(row=field_row, column=1, padx=10, pady=2)
                self.entries[key] = entry

            field_row += 1

        self.final_field_row = field_row


    def _handle_format_input_section(self):
        # Start below the last fields row
        format_row = self.final_field_row

        if self.preset_type == "File":
            tk.Button(self, text="Upload Sample File", command=self.load_sample_file).grid(row=format_row, column=0, columnspan=2, pady=10)
            format_row += 1

            self.header_buttons_frame = tk.Frame(self)
            self.header_buttons_frame.grid(row=format_row, column=0, columnspan=2, padx=10, pady=5, sticky="n")
            format_row += 1

            tk.Label(self, text="Label Format:").grid(row=format_row, column=0, columnspan=2, sticky="n", pady=(10, 0))
            format_row += 1

            self.textbox_format = tk.Text(self, width=75, height=8)
            self.textbox_format.grid(row=format_row, column=0, columnspan=2, padx=10, pady=(0, 10))
            if self.preset_data.get("textboxformatinput"):
                self.textbox_format.insert("1.0", self.preset_data["textboxformatinput"])
            self.entries["textboxformatinput"] = self.textbox_format

            # ✅ If preset has saved headers, repopulate header buttons
            if self.preset_data.get("saved_headers"):
                self.current_file_headers = self.preset_data["saved_headers"]

                # Clear previous buttons (if any)
                for widget in self.header_buttons_frame.winfo_children():
                    widget.destroy()

                grid_frame = tk.Frame(self.header_buttons_frame, width=400)
                grid_frame.pack(pady=10)
                grid_frame.pack_propagate(False)

                for i, header in enumerate(self.current_file_headers):
                    btn = tk.Button(
                        grid_frame,
                        text=header,
                        command=lambda h=header: self.insert_field(h)
                    )
                    btn.grid(row=i // 4, column=i % 4, padx=5, pady=5)


        elif self.preset_type == "Text":
            tk.Label(self, text="Label Text Format:").grid(row=format_row, column=0, columnspan=2, sticky="n", pady=(10, 0))
            format_row += 1

            self.format_entry = tk.Text(self, height=4, width=50)
            self.format_entry.grid(row=format_row, column=0, columnspan=2, padx=10, pady=(0, 10))
            self.entries["textboxformatinput"] = self.format_entry
            self.textbox_format = self.format_entry

            if self.preset_data.get("textboxformatinput"):
                self.textbox_format.insert("1.0", self.preset_data["textboxformatinput"])

            format_row += 1

            def insert_label_text():
                self.format_entry.insert(tk.INSERT, "{LABEL_TEXT}")

            self.insert_button = tk.Button(self, text="{LABEL_TEXT}", command=insert_label_text)
            self.insert_button.grid(row=format_row, column=1, sticky="w", padx=10, pady=(0, 10))

            def update_insert_button_state(event=None):
                logic_value = self.entries.get("identical_or_incremental").get()
                if logic_value == "Incremental":
                    self.insert_button.config(state="normal")
                else:
                    self.insert_button.config(state="disabled")

            # Set initial state
            update_insert_button_state()

            # Bind logic selector to update the state dynamically
            self.entries["identical_or_incremental"].bind("<<ComboboxSelected>>", update_insert_button_state)


    def _create_save_button(self):
        self.save_button = tk.Button(self, text="Save Preset", command=self.save_preset)
        self.save_button.grid(row=self.final_field_row + 5, column=0, columnspan=2, pady=20)

    def save_preset(self):
        preset = {"presettype": self.preset_type}

        # Gather widget data
        for key, widget in self.entries.items():
            if key == "labeltemplate":
                display_value = widget.get()
                internal_value = self.template_display_map.get(display_value, display_value)
                preset[key] = internal_value

            elif isinstance(widget, ttk.Combobox):
                val = widget.get()
                if key == "date_format":
                    if val == "Leave as is":
                        preset[key] = "Leave as is"  # ⬅️ Explicitly stores no format
                    else:
                        preset[key] = DATE_FORMAT_DISPLAY_MAP.get(val, "%m-%d-%Y")
                else:
                    preset[key] = int(val) if val.isdigit() else val


            elif isinstance(widget, tk.BooleanVar):
                preset[key] = widget.get()


            elif isinstance(widget, tk.Text):
                val = widget.get("0.0", "end-1c")  # Keep exact text, no extra strip/rstrip unless desired
                preset[key] = val

            else:
                try:
                    val = widget.get()
                    preset[key] = int(val) if val.isdigit() else val
                except Exception:
                    print(f"Skipping unknown widget for key: {key}")


        # Save file headers if available
        if self.preset_type == "File" and hasattr(self, "current_file_headers"):
            preset["saved_headers"] = self.current_file_headers
        if self.preset_type == "File" and hasattr(self, "sample_filename"):
            preset["sample_filename"] = self.sample_filename


        # Preserve or assign a unique preset_id
        if self.preset_path and os.path.exists(self.preset_path):
            with open(self.preset_path, "r") as f:
                existing_data = json.load(f)
                preset["preset_id"] = existing_data.get("preset_id", str(uuid.uuid4()))
        else:
            preset["preset_id"] = str(uuid.uuid4())

        # Save UI layout (optional)
        if self.preset_type == "Text":
            preset["ui_layout"] = {
                "elements": [
                    {"type": "textbox", "id": "user_input"},
                    {"type": "button", "id": "generate", "label": "Save Labels"},
                ]
            }
        elif self.preset_type == "File":
            preset["ui_layout"] = {
                "elements": [
                    {"type": "textpreview", "id": "preview_area"},
                    {"type": "button", "id": "upload_file", "label": "Load File"},
                    {"type": "button", "id": "generate", "label": "Save Labels"},
                ]
            }

            preset["sample_filename"] = self.sample_filename or self.preset_data.get("sample_filename", "")


        # Generate filename if not provided
        if not self.preset_path:
            safe_name = "".join(
                c if c.isalnum() or c in (" ", "-", "_") else "_" for c in preset.get("name", "preset")
            ).strip().replace(" ", "_")

            counter = 1
            filename = f"{safe_name}.json"
            full_path = os.path.join("presets", filename)

            while os.path.exists(full_path):
                filename = f"{safe_name}_{counter}.json"
                full_path = os.path.join("presets", filename)
                counter += 1

            self.preset_path = full_path

        # Write to file
        if self.preset_path:
            os.makedirs(os.path.dirname(self.preset_path), exist_ok=True)
            with open(self.preset_path, "w") as f:
                json.dump(preset, f, indent=4)

            messagebox.showinfo("Success", "Preset saved successfully.")
            if self.on_save:
                self.on_save(preset)  # pass preset to the callback
            self.destroy()


    def _get_ui_layout(self):
        if self.preset_type == "Text":
            return {
                "elements": [
                    {"type": "textbox", "id": "user_input"},
                    {"type": "button", "id": "generate", "label": "Save Labels"},
                ]
            }
        elif self.preset_type == "File":
            return {
                "elements": [
                    {"type": "file_upload", "id": "upload_sample"},
                    {"type": "textbox", "id": "textboxformatinput"},
                    {"type": "button", "id": "generate", "label": "Save Labels"},
                ]
            }
        return {}


    def update_textbox_size(self, event=None):
        template_name = self.template_display_map.get(
            self.entries.get("labeltemplate").get(), "default"
        )
        font_size_str = self.entries.get("fontsize").get()
        width, height = get_textbox_dimensions(template_name, font_size_str)
        if self.textbox_format:
            self.textbox_format.config(width=width, height=height)


    def load_sample_file(self):
        path = filedialog.askopenfilename(filetypes=[("CSV or Excel files", "*.csv *.xlsx")])
        if path:
            self.sample_filename = os.path.basename(path)
            self.lift()
            self.focus_force()
            self.preset_data["sample_filename"] = self.sample_filename


            # Get headers based on file extension
            if path.lower().endswith(".csv"):
                headers = get_csv_headers(path)
            elif path.lower().endswith(".xlsx"):
                headers = get_xlsx_headers(path)
            else:
                print("Unsupported file type")


            # Clear previous buttons
            for widget in self.header_buttons_frame.winfo_children():
                widget.destroy()

            filtered_headers = [h for h in headers if h and str(h).strip()]
            self.current_file_headers = filtered_headers  # store them on the instance


            # Set up an inner frame with a fixed width that will be centered by pack
            grid_frame = tk.Frame(self.header_buttons_frame, width=400)
            grid_frame.pack(pady=3)
            grid_frame.pack_propagate(False)  # Prevent it from shrinking to fit

            # Grid buttons inside that fixed-size frame
            for i, header in enumerate(filtered_headers):
                btn = tk.Button(
                    grid_frame,
                    text=header,
                    command=lambda h=header: self.insert_field(h)
                )
                btn.grid(row=i // 4, column=i % 4, padx=5, pady=5)




    
    def insert_field(self, column_name):
        if self.textbox_format:
            self.textbox_format.insert(tk.INSERT, f"{{{column_name}}}")


