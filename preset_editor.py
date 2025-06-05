import tkinter as tk
from label_templates import label_templates
from tkinter import filedialog, messagebox, ttk
import json
import os
import csv
import uuid 
import openpyxl as xlsx

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
        title = ("Edit " if self.preset_data else "New ") + ("Text Input Preset" if self.preset_type == "Text" else "File Input Preset")
        self.title(title)
        # self.geometry("550x600")

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


    def _create_fields_ui(self):
        field_row = 0
        for key, label in self.fields:
            tk.Label(self, text=label).grid(row=field_row, column=0, sticky="w", padx=10, pady=4)

            if key in ["partialsheet", "output_add_date"]:
                var = tk.BooleanVar()
                var.set(self.preset_data.get(key, False))
                cb = tk.Checkbutton(self, variable=var)
                cb.grid(row=field_row, column=1, padx=10, pady=4, sticky="w")
                self.entries[key] = var

            elif key == "color_theme":
                cb = ttk.Combobox(self, values=["Grey", "Pink", "Green", "Blue", "Yellow", "Purple"], state="readonly")
                cb.set(self.preset_data.get(key, "Pink"))
                cb.grid(row=field_row, column=1, padx=10, pady=4)
                self.entries[key] = cb

            elif key == "text_alignment":
                cb = ttk.Combobox(self, values=["Left", "Center", "Right"], state="readonly")
                cb.set(self.preset_data.get(key, "Center"))
                cb.grid(row=field_row, column=1, padx=10, pady=4)
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
                cb.grid(row=field_row, column=1, padx=10, pady=4)
                self.entries[key] = cb

            elif key == "labeltemplate":
                display_names = list(self.template_display_map.keys())
                cb = ttk.Combobox(self, values=display_names, state="readonly", width=27)
                cb.bind("<<ComboboxSelected>>", self.update_textbox_size)
                internal_value = self.preset_data.get(key)
                display_name = self.template_internal_map.get(internal_value, display_names[0])
                cb.set(display_name)
                cb.grid(row=field_row, column=1, padx=10, pady=4)
                self.entries[key] = cb

            elif key == "fontname":
                cb = ttk.Combobox(self, values=["Arial", "Courier", "Helvetica", "Times", "Verdana"], state="readonly")
                cb.set(self.preset_data.get(key, "Arial"))
                cb.grid(row=field_row, column=1, padx=10, pady=4)
                self.entries[key] = cb

            elif key == "fontsize":
                cb = ttk.Combobox(
                    self, values=["5.5", "6", "6.5", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20"],
                    state="readonly"
                )
                cb.set(str(self.preset_data.get(key, "6")))
                cb.bind("<<ComboboxSelected>>", self.update_textbox_size)
                cb.grid(row=field_row, column=1, padx=10, pady=4)
                self.entries[key] = cb

            elif key == "identical_or_incremental":
                cb = ttk.Combobox(self, values=["Identical", "Incremental"], state="readonly")
                cb.set(self.preset_data.get(key, "Identical"))
                cb.grid(row=field_row, column=1, padx=10, pady=4)
                self.entries[key] = cb


            elif key == "copiesperlabel":
                entry = tk.Entry(self, width=40)
                entry.insert(0, str(self.preset_data.get(key, "")))
                entry.grid(row=field_row, column=1, padx=10, pady=4)
                self.entries[key] = entry

            else:
                entry = tk.Entry(self, width=40)
                entry.insert(0, str(self.preset_data.get(key, "")))
                entry.grid(row=field_row, column=1, padx=10, pady=4)
                self.entries[key] = entry

            field_row += 1

        self.final_field_row = field_row



    def toggle_labels_per_serial(self, event=None):

        if self.entries["identical_or_incremental"].get() == "Incremental":
            self.labels_per_serial_label.grid(row=self.labels_per_serial_row, column=0, sticky="w", padx=10, pady=4)
            self.labels_per_serial_dropdown.grid(row=self.labels_per_serial_row, column=1, padx=10, pady=4)
        else:
            self.labels_per_serial_label.grid_remove()
            self.labels_per_serial_dropdown.grid_remove()


    def _handle_format_input_section(self):
        # Start below the last fields row
        format_row = self.final_field_row

        if self.preset_type == "File":
            tk.Button(self, text="Upload Sample File", command=self.load_sample_file).grid(row=format_row, column=0, columnspan=2, pady=10)
            format_row += 1

            self.header_buttons_frame = tk.Frame(self)
            self.header_buttons_frame.grid(row=format_row, column=0, columnspan=2, padx=10, pady=5, sticky="w")
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
                    btn.grid(row=i // 3, column=i % 3, padx=5, pady=5)


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

            insert_button = tk.Button(self, text="{LABEL_TEXT}", command=insert_label_text)
            insert_button.grid(row=format_row, column=1, sticky="w", padx=10, pady=(0, 10))



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
        if self.preset_type == "File" and hasattr(self, "sample_file_name"):
            preset["sample_file_name"] = self.sample_file_name


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

    def toggle_multi_mode(self, event=None):
        # Handle File preset Multi mode
        if hasattr(self, "copiesperlabel_dropdown"):
            if self.copiesperlabel_dropdown.get() == "Multi":
                self._show_multi_values_entry(self.file_multi_values_row)
            else:
                self._hide_multi_values_entry()

        # Handle Text preset Multi mode
        if hasattr(self, "labels_per_serial_dropdown"):
            if self.labels_per_serial_dropdown.get() == "Multi":
                self._show_multi_values_entry(self.text_multi_values_row)
            else:
                self._hide_multi_values_entry()

    def _show_multi_values_entry(self, target_row):
        if not hasattr(self, "multi_values_label"):
            self.multi_values_label = tk.Label(self, text="Multi Values (comma-separated):")
            self.multi_values_label.grid(row=target_row, column=0, sticky="w", padx=10, pady=4)

        if not hasattr(self, "multi_values_entry"):
            self.multi_values_entry = tk.Entry(self, width=30)
            self.multi_values_entry.grid(row=target_row, column=1, padx=10, pady=4)

    def _hide_multi_values_entry(self):
        if hasattr(self, "multi_values_label"):
            self.multi_values_label.destroy()
            del self.multi_values_label

        if hasattr(self, "multi_values_entry"):
            self.multi_values_entry.destroy()
            del self.multi_values_entry


    def get_multi_value_list(self, raw_value):
        """Return cleaned list of numeric strings from comma-separated input."""
        raw_items = raw_value.split(",")
        cleaned_list = [item.strip() for item in raw_items if item.strip().isdigit()]
        return cleaned_list

    def validate_multi_values(self, raw_value):
        """Check if all non-empty entries are numbers; show warning if not."""
        raw_items = raw_value.split(",")
        invalid_items = [item.strip() for item in raw_items if item.strip() and not item.strip().isdigit()]
        if invalid_items:
            messagebox.showwarning("Invalid Input", f"These values are invalid: {', '.join(invalid_items)}")
            return False
        return True


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
        labeltemplate_widget = self.entries.get("labeltemplate")
        fontsize_widget = self.entries.get("fontsize")

        if hasattr(labeltemplate_widget, "get") and hasattr(fontsize_widget, "get"):
            display_name = labeltemplate_widget.get()
            internal_key = self.template_display_map.get(display_name, display_name)
            template = label_templates.get(internal_key)

            if template:
                chars_per_line = template.get("chars_per_line", 20)
                lines_per_label = template.get("lines_per_label", 3)

                try:
                    font_size = float(fontsize_widget.get())
                except ValueError:
                    font_size = 6

                # Apply size to the box
                if self.textbox_format:
                    self.textbox_format.config(width=chars_per_line, height=lines_per_label)


    def load_sample_file(self):
        path = filedialog.askopenfilename(filetypes=[("CSV or Excel files", "*.csv *.xlsx")])
        if path:
            self.sample_file_name = os.path.basename(path)
            self.lift()
            self.focus_force()

            # Get headers based on file extension
            if path.lower().endswith(".csv"):
                with open(path, newline="", encoding="utf-8") as f:
                    reader = csv.reader(f)
                    headers = next(reader)
            elif path.lower().endswith(".xlsx"):
                wb = xlsx.load_workbook(path, read_only=True)
                sheet = wb.active
                headers = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1))]
            else:
                print("Unsupported file type")
                return

            # Clear previous buttons
            for widget in self.header_buttons_frame.winfo_children():
                widget.destroy()

            filtered_headers = [h for h in headers if h and str(h).strip()]
            self.current_file_headers = filtered_headers  # store them on the instance


            # Set up an inner frame with a fixed width that will be centered by pack
            grid_frame = tk.Frame(self.header_buttons_frame, width=400)
            grid_frame.pack(pady=10)
            grid_frame.pack_propagate(False)  # Prevent it from shrinking to fit

            # Grid buttons inside that fixed-size frame
            for i, header in enumerate(filtered_headers):
                btn = tk.Button(
                    grid_frame,
                    text=header,
                    command=lambda h=header: self.insert_field(h)
                )
                btn.grid(row=i // 3, column=i % 3, padx=5, pady=5)




    
    def insert_field(self, column_name):
        if self.textbox_format:
            self.textbox_format.insert(tk.INSERT, f"{{{column_name}}}")


