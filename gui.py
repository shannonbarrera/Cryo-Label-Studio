import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from label_spec import LabelSpec
from main import main
import json
import os
from datetime import datetime
from label_templates import label_templates
from userguide import show_help_window
from preset_editor import PresetEditor
from data_extract import get_data_list_csv, get_data_list_xlsx
from label_format import apply_format_to_row
from data_process import is_valid_serial_format, parse_copiesperlabel_input
from preset_editor.file_helpers import get_csv_headers, get_xlsx_headers

class CryoPopLabelStudioLite:
    def __init__(self, root):
        """
        Initialize the CryoPop Label Studio main application window.

        Args:
            root (tk.Tk): The root Tkinter window object.
        """

        self.root = root
        self.root.title("CryoPop Label Studio")
        self.root.geometry("600x650+20+20")
        self.root.resizable(False, False)  # lock resizing

        self.top_frame = tk.Frame(self.root)
        self.top_frame.pack(fill=tk.X)

        self.body_frame = tk.Frame(self.root)
        self.body_frame.pack(fill=tk.BOTH, expand=True)


        self.current_spec = None
        self.presets_dir = "presets"
        os.makedirs(self.presets_dir, exist_ok=True)

        self.setup_menu()
        self.setup_main_ui()
        self.welcome_frame = tk.Frame(self.top_frame)
        self.welcome_frame.pack(expand=True)

        tk.Label(self.welcome_frame, text="Welcome to CryoPop Label Studio", font=("Arial", 16, "bold")).pack(pady=(60, 10))
        tk.Label(self.welcome_frame, text="Select or Create a Preset to Begin", font=("Arial", 12)).pack(pady=5)

        self.load_all_presets()
        self.widgets = {}
        self.selected_label_count = tk.StringVar()

        self.multi_radio_frame = tk.Frame(self.body_frame)
        self.multi_radio_frame.pack(pady=10)

        self.status_var = tk.StringVar()
        self.status_var.set("No Preset Loaded")

        self.footer_bar = tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor='w')
        self.footer_bar.pack(side=tk.BOTTOM, fill=tk.X)


    def update_footer(self, input_type=None, logic_type=None, sample_file=None, copies=None):
        """
        Update the footer bar with information about the current preset, including logic type,
        sample file name, and number of copies.

        Args:
            logic_type (str, optional): The logic used for label generation ("Identical" or "Incremental").
            sample_file (str, optional): Name of the sample file for File Input presets.
            copies (str or int, optional): Number of label copies to display.
        """
        footer_parts = []
        if copies is not None:
            strcopies = str(copies)
            if "," in strcopies or "-" in strcopies:
                footer_copies = None
            else:
                if copies == "":
                    footer_copies = "Fill Sheet"
                else:
                    footer_copies = copies
        else:
            footer_copies = None

        if input_type == "Text":
            if logic_type:
                footer_parts.append(f"Logic: {logic_type}")
            if footer_copies is not None:
                footer_parts.append(f"Copies: {footer_copies}")
            
        elif input_type == "File":
            footer_parts.append(f"Sample File: {sample_file or '(none)'}")
            if footer_copies is not None:
                footer_parts.append(f"Copies: {footer_copies}")
            
        else:
            footer_parts.append("Please Select a Preset")


        self.status_var.set("  •  ".join(footer_parts))

    def update_footer_copies_only(self):
        """
        Refresh only the "Copies" part of the footer based on the selected radio button.
        """
        current = self.selected_label_count.get()
        existing_text = self.status_var.get()
        parts = existing_text.split("  •  ")
        # Remove any old Copies entry
        parts = [p for p in parts if not p.strip().startswith("Copies:")]
        parts.append(f"Copies: {current}")
        self.status_var.set("  •  ".join(parts))

    def setup_menu(self):
        """
        Initialize the main application menu bar, including Preset and Help menus.
        """
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)

        # Presets menu
        preset_menu = tk.Menu(self.menu_bar, tearoff=0)
        preset_menu.add_command(label="New File Input Preset", command=lambda: self.new_preset_window("File"))
        preset_menu.add_command(label="New Text Input Preset", command=lambda: self.new_preset_window("Text"))
        preset_menu.add_separator()
        preset_menu.add_command(label="Edit Presets", command=self.edit_presets_window)
        self.menu_bar.add_cascade(label="Presets", menu=preset_menu)

        # Help menu
        help_menu = tk.Menu(self.menu_bar, tearoff=0)
        help_menu.add_command(label="User Guide", command=lambda: show_help_window(self.root))
        self.menu_bar.add_cascade(label="Help", menu=help_menu)


    def setup_main_ui(self):
        """
        Build the initial main user interface, including the preset dropdown and status label.
        """
        # Preset Selector Dropdown
        tk.Label(self.top_frame, text="Select a Preset:", font=("Arial", 12)).pack(pady=(20, 5))
        self.preset_var = tk.StringVar()
        self.preset_dropdown = ttk.Combobox(self.top_frame, textvariable=self.preset_var, state="readonly", width=35)
        self.preset_dropdown.pack(pady=5)
        self.preset_dropdown.bind("<<ComboboxSelected>>", self.load_selected_preset)

        self.status_label = tk.Label(self.top_frame, text="", font=("Arial", 10))
        self.status_label.pack(pady=10)


    def apply_color_theme(self, theme_name):
        """
        Apply a background color theme to the application UI.

        Args:
            theme_name (str): Name of the color theme to apply (e.g., "Pink", "Grey").
        """

        themes = {
            "Pink": "#FBDAE9",
            "Green": "#D3F8E2",
            "Blue": "#C6E9FB",
            "Yellow": "#F4F0CD",
            "Purple": "#EFDAFB",
            "Grey": "#F0F0F0"
        }
        bg = themes.get(theme_name, "Grey")

        # Setup custom styles for ttk widgets
        style = ttk.Style()
        style.configure('Custom.TButton', background='white')
        style.configure('Custom.TLabel', background=bg, foreground='black')
        style.configure('Custom.TEntry', fieldbackground='white')
        style.configure('Custom.TCombobox', fieldbackground='white', background='white')

        def apply_bg_recursively(widget):
            try:
                widget_class = widget.winfo_class()
            except Exception:
                return  # Skip destroyed widgets

            # Skip ttk.Combobox entirely — leave system-native look
            if isinstance(widget, ttk.Combobox) or widget_class == 'TCombobox':
                return

            # Make buttons white.
            try:
                if isinstance(widget, (tk.Button, tk.Text)):
                    widget.configure(bg='white')
                else:
                    widget.configure(bg=bg)
            except:
                pass

            # ttk widgets (skip Combobox!)
            if isinstance(widget, ttk.Button) or widget_class == 'TButton':
                widget.configure(style='Custom.TButton')
            elif isinstance(widget, ttk.Label) or widget_class == 'TLabel':
                widget.configure(style='Custom.TLabel')
            elif isinstance(widget, ttk.Entry) or widget_class == 'TEntry':
                widget.configure(style='Custom.TEntry')

            # Recurse into children
            for child in widget.winfo_children():
                apply_bg_recursively(child)




        # Apply to your main body and top frames only (avoid menu bar!)
        apply_bg_recursively(self.body_frame)
        apply_bg_recursively(self.top_frame)

        if hasattr(self, 'multi_radio_frame'):
            apply_bg_recursively(self.multi_radio_frame)

        if hasattr(self, 'pages_of_labels_frame'):
            apply_bg_recursively(self.pages_of_labels_frame)

    def load_all_presets(self):
        """
        Load all saved presets from disk into the preset dropdown.
        """
        self.presets = {}
        if os.path.exists(self.presets_dir):
            for file in os.listdir(self.presets_dir):
                if file.endswith(".json"):
                    path = os.path.join(self.presets_dir, file)
                    with open(path, "r") as f:
                        data = json.load(f)
                        name = data.get("name", file)
                        self.presets[name] = (path, data)
        self.preset_dropdown['values'] = list(self.presets.keys())

    def load_selected_preset(self, event=None, preset_name=None):
        """
        Load and apply the selected preset's settings and dynamically build the UI.

        Args:
            event (Event, optional): Triggering event (usually from the dropdown).
            preset_name (str, optional): Name of the preset to load manually.
        """

        if hasattr(self, "welcome_frame") and self.welcome_frame.winfo_exists():
            self.welcome_frame.pack_forget()

        name = preset_name or self.preset_var.get()

        if name in self.presets:
            path, data = self.presets[name]
            self.current_spec = LabelSpec(**data)
            self.apply_preset_to_ui(self.current_spec)
            self.clear_ui()
            self.build_ui_from_spec(self.current_spec.ui_layout)

        # ✅ Only clear the existing multi_radio_frame, never recreate it
        if hasattr(self, "multi_radio_frame") and self.multi_radio_frame.winfo_exists():
            for widget in self.multi_radio_frame.winfo_children():
                widget.destroy()

        # ✅ Only clear or create the pages_of_labels dropdown if needed
        if self.current_spec.presettype == "Text" and getattr(self.current_spec, "identical_or_incremental", "") == "Incremental":
            if not hasattr(self, "pages_of_labels_var"):
                self.pages_of_labels_var = tk.StringVar(value="1")

            if not hasattr(self, "pages_of_labels_frame") or not self.pages_of_labels_frame.winfo_exists():
                self.pages_of_labels_frame = tk.Frame(self.body_frame)
                self.pages_of_labels_frame.pack(pady=5)

                tk.Label(self.pages_of_labels_frame, text="Number of Pages:").pack(side="left")
                self.pages_of_labels_dropdown = ttk.Combobox(
                    self.pages_of_labels_frame,
                    textvariable=self.pages_of_labels_var,
                    values=[str(i) for i in range(1, 11)],
                    state="readonly"
                )
                self.pages_of_labels_dropdown.pack(side="left", padx=5)
            else:
                self.pages_of_labels_var.set("1")
                self.pages_of_labels_frame.pack(pady=5)
        else:
            if hasattr(self, "pages_of_labels_frame") and self.pages_of_labels_frame.winfo_exists():
                self.pages_of_labels_frame.pack_forget()

        raw_input = str(getattr(self.current_spec, "copiesperlabel", "1"))
        copies_list = parse_copiesperlabel_input(raw_input)

        if copies_list:
            self.selected_label_count.set(copies_list[0])

            if len(copies_list) > 1:
                tk.Label(self.multi_radio_frame, text="Copies Per Label:").pack(anchor="w")

                for val in copies_list:
                    rb = tk.Radiobutton(
                        self.multi_radio_frame,
                        text=val,
                        variable=self.selected_label_count,
                        value=val,
                        command=self.update_footer_copies_only  # 🟢 Trigger on click
                    )
                    rb.pack(side="left", padx=5)

            # 🟢 Call after radio buttons are set
        self.update_footer(
            input_type=self.current_spec.presettype,
            logic_type=getattr(self.current_spec, "identical_or_incremental", None),
            sample_file=getattr(self.current_spec, "sample_filename", None),
            copies=self.selected_label_count.get()
        )

        

        # ✅ Apply color theme LAST, after everything is built
        if self.current_spec and getattr(self.current_spec, "color_theme", "Grey"):
            self.apply_color_theme(self.current_spec.color_theme)




    def apply_preset_to_ui(self, spec):
        """
        Update UI elements to reflect the values in a given LabelSpec.

        Args:
            spec (LabelSpec): The loaded preset specification.
        """

        # Update status label
        
        template_id = getattr(spec, "labeltemplate", "Unknown")
        template_display = label_templates.get(template_id, {}).get("display_name", template_id)
        self.status_label.config(text=f"Label Template: {template_display}")

        # Update text preview (if applicable)
        if hasattr(spec, "textboxformatinput") and hasattr(self, "preview_box"):
            self.preview_box.delete("1.0", "end")
            self.preview_box.insert("1.0", spec.textboxformatinput)

        # Update font info
        if hasattr(spec, "fontname") and hasattr(self, "font_label"):
            self.font_label.config(text=f"Font: {spec.fontname}, {spec.fontsize}pt")
    
    def build_ui_from_spec(self, layout_data):
        """
        Dynamically construct the main body UI based on the preset layout specification.

        Args:
            layout_data (dict): UI layout configuration dict with element definitions.
        """

        # ✅ Always reserve the multi_radio_frame early, so it stays above the buttons
        if not hasattr(self, "multi_radio_frame") or not self.multi_radio_frame.winfo_exists():
            self.multi_radio_frame = tk.Frame(self.body_frame)
            self.multi_radio_frame.pack(pady=10)

        for element in layout_data.get("elements", []):
            etype = element["type"]
            eid = element["id"]

            if etype == "textbox":
                template_id = self.current_spec.labeltemplate
                template = label_templates.get(template_id, {})
                if self.current_spec.presettype == "Text" and getattr(self.current_spec, "identical_or_incremental", "") == "Incremental":
                    box_width = 27
                    box_height = 1

                else:
                    box_width = template.get("chars_per_line", 45)
                    box_height = template.get("lines_per_label", 6)

                #Add description label based on preset type and logic
                label_text = "Label Text:"  # default
                if self.current_spec.presettype == "Text":
                    logic = getattr(self.current_spec, "identical_or_incremental", "").lower()
                    if logic == "incremental":
                        label_text = "First Number in Series:"
                    elif logic == "identical":
                        label_text = "Label Text:"


                desc_label = tk.Label(self.body_frame, text=label_text)
                desc_label.pack(padx=10, pady=(10, 2))

                text_box = tk.Text(self.body_frame, width=box_width, height=box_height)
                text_box.pack()
                self.widgets[eid] = text_box

                # ✅ Optional: prefill if the preset has plain text
                fmt = getattr(self.current_spec, "textboxformatinput", "")
                if self.current_spec.presettype == "Text" and fmt and "{" not in fmt and "}" not in fmt:
                    text_box.insert("1.0", fmt)

            elif etype == "textpreview":
                desc_label = tk.Label(self.body_frame, text="First Label Preview:")
                desc_label.pack(padx=10, pady=(10, 2))
                txt = tk.Text(self.body_frame, height=10, width=50)
                txt.pack()

                self.widgets[eid] = txt

            elif etype == "label":
                lbl = tk.Label(self.body_frame, text=element.get("text", ""))
                lbl.pack(padx=10, pady=15)
                self.widgets[eid] = lbl

        if self.current_spec and getattr(self.current_spec, "partialsheet", False):
            template_id = self.current_spec.labeltemplate
            template = label_templates.get(template_id, {})
            labels_down = template.get("labels_down", 99)
            labels_across = template.get("labels_across", 99)

            row_range = [str(i) for i in range(1, labels_down + 1)]
            col_range = [str(j) for j in range(1, labels_across + 1)]

            # First Label Row & Column (top line)
            first_row_frame = tk.Frame(self.body_frame)
            first_row_frame.pack(pady=(25, 2))

            tk.Label(first_row_frame, text="First Label Row:").pack(side=tk.LEFT)
            self.row_start_var = tk.StringVar(value="1")
            ttk.Combobox(first_row_frame, textvariable=self.row_start_var, values=row_range, width=5).pack(side=tk.LEFT, padx=5)

            tk.Label(first_row_frame, text="Column:").pack(side=tk.LEFT)
            self.col_start_var = tk.StringVar(value="1")
            ttk.Combobox(first_row_frame, textvariable=self.col_start_var, values=col_range, width=5).pack(side=tk.LEFT, padx=5)

            # Last Label Row & Column (bottom line)
            last_row_frame = tk.Frame(self.body_frame)
            last_row_frame.pack(pady=(5, 20))

            tk.Label(last_row_frame, text="Last Label Row:").pack(side=tk.LEFT)
            self.row_end_var = tk.StringVar(value=str(labels_down))
            ttk.Combobox(last_row_frame, textvariable=self.row_end_var, values=row_range, width=5).pack(side=tk.LEFT, padx=5)

            tk.Label(last_row_frame, text="Column:").pack(side=tk.LEFT)
            self.col_end_var = tk.StringVar(value=str(labels_across))
            ttk.Combobox(last_row_frame, textvariable=self.col_end_var, values=col_range, width=5).pack(side=tk.LEFT, padx=5)

        # Group buttons horizontally in a row BELOW the radio frame
        btn_row = tk.Frame(self.body_frame)
        btn_row.pack(pady=15)

        for element in layout_data.get("elements", []):
            if element["type"] == "button":
                eid = element["id"]
                label = element["label"]

                if eid == "generate":
                    btn = tk.Button(btn_row, text=label, command=self.generate_labels)
                elif eid == "upload_file":
                    btn = tk.Button(btn_row, text=label, command=self.upload_sample_file)
                else:
                    btn = tk.Button(btn_row, text=label)

                btn.pack(side=tk.LEFT, padx=10)
                self.widgets[eid] = btn


    def generate_labels(self):
        """
        Generate labels based on the current preset and user input, saving the output to file.
        Handles both text and file input presets.
        """

        spec = self.current_spec
        if not spec:
            messagebox.showerror("Error", "No preset loaded.")
            return

        # Get pages of labels (for Text Incremental presets)
        if hasattr(self, "pages_of_labels_var"):
            pages_of_labels = int(self.pages_of_labels_var.get())
        else:
            pages_of_labels = 1
        spec.pages_of_labels = pages_of_labels

        if hasattr(self, "selected_label_count"):
            val = self.selected_label_count.get()
            spec.copiesperlabel = int(val) if val.strip().isdigit() else ""

        filename_base = self.current_spec.outputfilenameprefix

        if self.current_spec.output_add_date:
            datetime_str = datetime.now().strftime("%m%d%y%H%M")  # e.g. 0527241435
            initial_filename = f"{filename_base}_{datetime_str}"
        else:
            initial_filename = filename_base

        output_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx")],
            initialfile=initial_filename
        )

        if not output_path:
            return

        if hasattr(self, "row_start_var"):
            spec.row_start = int(self.row_start_var.get())
        if hasattr(self, "row_end_var"):
            spec.row_end = int(self.row_end_var.get())
        if hasattr(self, "col_start_var"):
            spec.col_start = int(self.col_start_var.get())
        if hasattr(self, "col_end_var"):
            spec.col_end = int(self.col_end_var.get())

        try:
            if spec.presettype == "Text":
                if spec.identical_or_incremental.lower() == "incremental":
                    text = self.widgets["user_input"].get("1.0", "end").strip()
                    if not is_valid_serial_format(text):
                        messagebox.showerror(
                            "Error",
                            "Serial format must:\n"
                            "- Be 12 characters or fewer\n"
                            "- Match one of these formats:\n"
                            "  • Numbers only (e.g., 1234)\n"
                            "  • Prefix (1–5 letters/numbers) + dash + digits (e.g., ab-123)\n"
                            "  • Prefix + underscore + digits (e.g., xy_0999)\n"
                            "  • Prefix + digits (e.g., ab0001)"
                        )
                        return 
                    main(spec, text_box_input=text, output_file_path=output_path)
                elif spec.identical_or_incremental.lower() == "identical":
                    text = self.widgets["user_input"].get("1.0", "end").rstrip()
                    main(spec, text_box_input=text, output_file_path=output_path)

            elif spec.presettype == "File":
                if not hasattr(self, "input_file_path") or not self.input_file_path:
                    messagebox.showerror("Error", "Please upload a CSV or file.")
                    return 
                main(spec, input_file_path=self.input_file_path, output_file_path=output_path)

        except Exception as e:
            messagebox.showerror("Error", f"Label generation failed:\n{e}")

    def upload_sample_file(self):
        """
        Let the user upload a CSV or Excel file and preview the first formatted label.
        """
        path = filedialog.askopenfilename(filetypes=[("CSV or Excel files", "*.csv *.xlsx")])

        if path:
            self.input_file_path = path
            try:

                if path.endswith(".csv"):
                    data_list = get_data_list_csv(path, self.current_spec.textboxformatinput, self.current_spec.date_format)
                else:
                    data_list = get_data_list_xlsx(path, self.current_spec.textboxformatinput, self.current_spec.date_format)

                if data_list and len(data_list) > 0:
                    preview = apply_format_to_row(self.current_spec.textboxformatinput, data_list[0], self.current_spec.date_format)
                else:
                    preview = "No data found or invalid format."

                self.widgets["preview_area"].delete("1.0", "end")
                self.widgets["preview_area"].insert("1.0", preview)

            except Exception as e:
                messagebox.showwarning("Warning", f"Preview failed:\n{e}")


    def clear_ui(self):
        """
        Remove all dynamically generated UI widgets from the main window.
        """

        for widget in self.body_frame.winfo_children():
            widget.destroy()

        self.widgets.clear()

    def new_preset_window(self, preset_type):
        """
        Launch the Preset Editor to create a new preset.

        Args:
            preset_type (str): The type of preset to create ("Text" or "File").
        """
        PresetEditor(self.root, preset_type=preset_type, on_save=self.on_preset_saved)

    def edit_presets_window(self):
        """
        Open a window to manage existing presets, allowing users to edit or delete them.
        """
        win = tk.Toplevel(self.root)
        win.title("Edit Presets")
        win.geometry("400x300")

        lb = tk.Listbox(win, selectmode=tk.MULTIPLE)
        lb.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        for name in self.presets:
            lb.insert(tk.END, name)

        status_label = tk.Label(win, text="", fg="green")
        status_label.pack(pady=5)

        def edit_selected():
            selected = lb.curselection()
            if not selected:
                return
            name = lb.get(selected[0])
            path, data = self.presets[name]

            # Call PresetEditor to open the edit window
            editor = PresetEditor(self.root, preset_data=data, preset_path=path, on_save=self.on_preset_saved)

            # NEW: If it's a File preset and has an input file path, refresh the buttons
            if data.get("presettype") == "File" and data.get("input_file_path"):
                self.refresh_column_buttons_from_file(
                    data["input_file_path"],
                    data.get("textboxformatinput", "")
                )

            win.destroy()


        def delete_selected():
            selected = lb.curselection()
            if not selected:
                return
            
            if len(selected) >= 1:
                confirm = messagebox.askyesno("Confirm Delete", f"Delete {len(selected)} presets?")
                if not confirm:
                    return

            deleted_names = []
            for i in selected:
                name = lb.get(i)
                path, _ = self.presets[name]
                os.remove(path)
                deleted_names.append(name)

            # Refresh listbox without closing the window
            self.load_all_presets()
            lb.delete(0, tk.END)
            for name in self.presets:
                lb.insert(tk.END, name)

            status_label.config(text=f"Deleted: {', '.join(deleted_names)}")

            win.lift()
            win.focus_force()


        btn_frame = tk.Frame(win)
        btn_frame.pack(pady=5)
        tk.Button(btn_frame, text="Edit", command=edit_selected).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="Delete", command=delete_selected).pack(side=tk.LEFT, padx=10)

    def on_preset_saved(self, saved_preset):
        """
        Reload the preset list and update the UI after a preset is saved.

        Args:
            saved_preset (dict): Dictionary containing saved preset metadata.
        """

        saved_id = saved_preset.get("preset_id")
        self.load_all_presets()

        # Try to find the saved preset and reload it
        for name, (path, data) in self.presets.items():
            if data.get("preset_id") == saved_id:
                self.preset_dropdown.set(name)
                self.current_spec = LabelSpec(**data)
                self.load_selected_preset(preset_name=name)
                break

    def refresh_column_buttons_from_file(self, file_path, format_string):
        """
        Update column selection buttons in the Preset Editor UI based on the uploaded file.

        Args:
            file_path (str): Path to the sample CSV or XLSX file.
            format_string (str): Formatting string used for preview generation.
        """

        if file_path.endswith(".csv"):
            headers = get_csv_headers(file_path)
        else:
            headers = get_xlsx_headers(file_path)

        # Clear previous buttons
        for widget in self.header_buttons_frame.winfo_children():
            widget.destroy()

        # Build new buttons
        for i, header in enumerate(headers):
            btn = tk.Button(
                self.header_buttons_frame,
                text=header,
                command=lambda h=header: self.insert_field_into_format(h)
            )
            btn.grid(row=i // 3, column=i % 3, padx=5, pady=5)



if __name__ == "__main__":
    root = tk.Tk()
    app = CryoPopLabelStudioLite(root)
    root.mainloop()
