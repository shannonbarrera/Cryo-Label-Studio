import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from label_spec import LabelSpec
from main import main
import json
import os
from datetime import datetime
from label_templates import label_templates
from preset_editor import PresetEditor
from data_extract import get_data_list_csv, get_data_list_xlsx
from label_format import apply_format_to_row
from data_process import is_valid_serial_format
import re

class CryoPopLabelStudioLite:
    def __init__(self, root):
        self.root = root
        self.root.title("CryoPop Label Studio Lite")
        self.root.geometry("600x500")
        self.top_frame = tk.Frame(self.root)
        self.top_frame.pack(fill=tk.X)

        self.body_frame = tk.Frame(self.root)
        self.body_frame.pack(fill=tk.BOTH, expand=True)


        self.current_spec = None
        self.presets_dir = "presets"
        os.makedirs(self.presets_dir, exist_ok=True)

        self.setup_menu()
        self.setup_main_ui()
        self.load_all_presets()
        self.widgets = {}
        self.selected_label_count = tk.StringVar()

        self.multi_radio_frame = tk.Frame(self.body_frame)
        self.multi_radio_frame.pack(pady=10)

        self.status_var = tk.StringVar()
        self.status_var.set("No Preset Loaded")

        self.footer_bar = tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor='w')
        self.footer_bar.pack(side=tk.BOTTOM, fill=tk.X)


    def update_footer(self, input_type=None, logic_type=None, sample_file=None, current_file=None, status="Ready"):
        footer_parts = []

        if input_type == "Text":
            footer_parts.append("Input: Text")
            if logic_type:
                footer_parts.append(f"Logic: {logic_type}")
            footer_parts.append(f"Status: {status}")

        elif input_type == "File":
            footer_parts.append(f"Sample File: {sample_file or '(none)'}")
            footer_parts.append(f"Current File: {current_file or '(none)'}")
            footer_parts.append(f"Status: {status}")

        else:
            footer_parts.append("Input: Unknown")
            footer_parts.append(f"Status: {status}")

        self.status_var.set("  •  ".join(footer_parts))




    def setup_menu(self):
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)

        preset_menu = tk.Menu(self.menu_bar, tearoff=0)
        preset_menu.add_command(label="New File Input Preset", command=lambda: self.new_preset_window("File"))
        preset_menu.add_command(label="New Text Input Preset", command=lambda: self.new_preset_window("Text"))
        preset_menu.add_separator()
        preset_menu.add_command(label="Edit Presets", command=self.edit_presets_window)
        self.menu_bar.add_cascade(label="Presets", menu=preset_menu)

    def setup_main_ui(self):
        tk.Label(self.top_frame, text="Select a Preset:", font=("Arial", 12)).pack(pady=(20, 5))
        self.preset_var = tk.StringVar()
        self.preset_dropdown = ttk.Combobox(self.top_frame, textvariable=self.preset_var, state="readonly", width=35)
        self.preset_dropdown.pack(pady=5)
        self.preset_dropdown.bind("<<ComboboxSelected>>", self.load_selected_preset)

        self.status_label = tk.Label(self.top_frame, text="", font=("Arial", 10))
        self.status_label.pack(pady=10)


    def apply_color_theme(self, theme_name):
        themes = {
            "Pink": "#FBDAE9",
            "Green": "#D3F8E2",
            "Blue": "#C6E9FB",
            "Yellow": "#F4F0CD",
            "Purple": "#EFDAFB",
            "Grey": "D3D3D3"
        }
        bg = themes.get(theme_name, "white")

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

            # Standard Tk widgets
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

    def load_selected_preset(self, event=None):
        name = self.preset_var.get()
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

                tk.Label(self.pages_of_labels_frame, text="Pages of Labels:").pack(side="left")
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

                # default to first

                for val in copies_list:
                    rb = tk.Radiobutton(
                        self.multi_radio_frame,
                        text=val,
                        variable=self.selected_label_count,
                        value=val
                    )
                    rb.pack(side="left", padx=5)
        

            # ✅ Apply color theme LAST, after everything is built
        if self.current_spec and getattr(self.current_spec, "color_theme", None):
            self.apply_color_theme(self.current_spec.color_theme)

        self.update_footer(
            input_type=self.current_spec.presettype,
            logic_type=getattr(self.current_spec, "identical_or_incremental", None),
            sample_file=getattr(self.current_spec, "sample_file_name", None),
            current_file=getattr(self, "input_file_path", None),
            status="Preset loaded"
        )




    def apply_preset_to_ui(self, spec):
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
                box_width = template.get("chars_per_line", 45)
                box_height = template.get("lines_per_label", 6)

                text_box = tk.Text(self.body_frame, width=box_width, height=box_height)
                text_box.pack()
                self.widgets[eid] = text_box

            elif etype == "textpreview":
                txt = tk.Text(self.body_frame, height=10, width=50)
                txt.pack(padx=10, pady=15)
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
            first_row_frame.pack(pady=5)

            tk.Label(first_row_frame, text="First Label Row:").pack(side=tk.LEFT)
            self.row_start_var = tk.StringVar(value="1")
            ttk.Combobox(first_row_frame, textvariable=self.row_start_var, values=row_range, width=5).pack(side=tk.LEFT, padx=5)

            tk.Label(first_row_frame, text="Column:").pack(side=tk.LEFT)
            self.col_start_var = tk.StringVar(value="1")
            ttk.Combobox(first_row_frame, textvariable=self.col_start_var, values=col_range, width=5).pack(side=tk.LEFT, padx=5)

            # Last Label Row & Column (bottom line)
            last_row_frame = tk.Frame(self.body_frame)
            last_row_frame.pack(pady=5)

            tk.Label(last_row_frame, text="Last Label Row:").pack(side=tk.LEFT)
            self.row_end_var = tk.StringVar(value=str(labels_down))
            ttk.Combobox(last_row_frame, textvariable=self.row_end_var, values=row_range, width=5).pack(side=tk.LEFT, padx=5)

            tk.Label(last_row_frame, text="Column:").pack(side=tk.LEFT)
            self.col_end_var = tk.StringVar(value=str(labels_across))
            ttk.Combobox(last_row_frame, textvariable=self.col_end_var, values=col_range, width=5).pack(side=tk.LEFT, padx=5)

        # ✅ Group buttons horizontally in a row BELOW the radio frame
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
        spec = self.current_spec
        if not spec:
            messagebox.showerror("Error", "No preset loaded.")
            return

        # ✅ Get pages of labels (for Text Incremental presets)
        if hasattr(self, "pages_of_labels_var"):
            pages_of_labels = int(self.pages_of_labels_var.get())
        else:
            pages_of_labels = 1
        spec.pages_of_labels = pages_of_labels

        if hasattr(self, "selected_label_count"):
            spec.copiesperlabel = int(self.selected_label_count.get())
        


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
                text = self.widgets["user_input"].get("1.0", "end").strip()
                print(text)
                if spec.identical_or_incremental.lower() == "incremental":
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
                    text = [text]
                    main(spec, text_box_input=text, output_file_path=output_path)

            elif spec.presettype == "File":
                if not hasattr(self, "input_file_path") or not self.input_file_path:
                    messagebox.showerror("Error", "Please upload a CSV or file.")
                    return 
                main(spec, input_file_path=self.input_file_path, output_file_path=output_path)

        except Exception as e:
            messagebox.showerror("Error", f"Label generation failed:\n{e}")

    def upload_sample_file(self):
        path = filedialog.askopenfilename(filetypes=[("CSV or Excel files", "*.csv *.xlsx")])

        if path:
            self.input_file_path = path
            try:

                if path.endswith(".csv"):
                    data_list = get_data_list_csv(path, self.current_spec.textboxformatinput)
                else:
                    data_list = get_data_list_xlsx(path, self.current_spec.textboxformatinput)
                print(data_list)
                if data_list and len(data_list) > 0:
                    preview = apply_format_to_row(self.current_spec.textboxformatinput, data_list[0], self.current_spec.date_format)
                else:
                    preview = "No data found or invalid format."
                print(preview)
                self.widgets["preview_area"].delete("1.0", "end")
                self.widgets["preview_area"].insert("1.0", preview)

            except Exception as e:
                messagebox.showwarning("Warning", f"Preview failed:\n{e}")


    def clear_ui(self):
        for widget in self.root.winfo_children():
            for widget in self.body_frame.winfo_children():
                widget.destroy()
        for widget in self.widgets.values():
            try:
                widget.destroy()
            except:
                pass
        self.widgets.clear()

    def new_preset_window(self, preset_type):
        PresetEditor(self.root, preset_type=preset_type, on_save=self.on_preset_saved)

    def edit_presets_window(self):
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
        saved_id = saved_preset.get("preset_id")
        self.load_all_presets()

        # Always clear selection after saving
        self.preset_dropdown.set("")
        self.current_spec = None  # or whatever variable holds the loaded preset

    def refresh_column_buttons_from_file(self, file_path, format_string):
        if file_path.endswith(".csv"):
            headers = get_data_list_csv(file_path, format_string, headers_only=True)
        else:
            headers = get_data_list_xlsx(file_path, format_string, headers_only=True)

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



    def run_label_maker(self):
        if self.current_spec:
            try:
                main.main(self.current_spec)
                messagebox.showinfo("Success", "Labels generated successfully.")
            except Exception as e:
                messagebox.showerror("Error", str(e))


def parse_copiesperlabel_input(raw_input):
    result = []
    parts = [part.strip() for part in raw_input.split(",")]
    for part in parts:
        if "-" in part:
            match = re.match(r"(\d+)-(\d+)", part)
            if match:
                start, end = map(int, match.groups())
                result.extend([str(i) for i in range(start, end + 1)])
        elif part.isdigit():
            result.append(part)
    return result


if __name__ == "__main__":
    root = tk.Tk()
    app = CryoPopLabelStudioLite(root)
    root.mainloop()
