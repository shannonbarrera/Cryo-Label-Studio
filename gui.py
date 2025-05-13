
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from label_spec import LabelSpec
from main import main
import json
import os
from label_templates import label_templates
from preset_editor import PresetEditor
from data_extract import get_data_list_csv, get_data_list_xlsx

class CryoPopLabelStudioLite:
    def __init__(self, root):
        self.root = root
        self.root.title("CryoPop Label Studio Lite")
        self.root.geometry("600x500")

        self.current_spec = None
        self.presets_dir = "presets"
        os.makedirs(self.presets_dir, exist_ok=True)

        self.setup_menu()
        self.setup_main_ui()
        self.load_all_presets()
        self.widgets = {}


    def setup_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        preset_menu = tk.Menu(menubar, tearoff=0)
        preset_menu.add_command(label="New File Input Preset", command=lambda: self.new_preset_window("File"))
        preset_menu.add_command(label="New Text Input Preset", command=lambda: self.new_preset_window("Text"))
        preset_menu.add_separator()
        preset_menu.add_command(label="Edit Presets", command=self.edit_presets_window)
        menubar.add_cascade(label="Presets", menu=preset_menu)

    def setup_main_ui(self):
        tk.Label(self.root, text="Select a Preset:", font=("Arial", 12)).pack(pady=(20, 5))
        self.preset_var = tk.StringVar()
        self.preset_dropdown = ttk.Combobox(self.root, textvariable=self.preset_var, state="readonly")
        self.preset_dropdown.pack(pady=5)
        self.preset_dropdown.bind("<<ComboboxSelected>>", self.load_selected_preset)

        self.status_label = tk.Label(self.root, text="", font=("Arial", 10))
        self.status_label.pack(pady=10)


    def apply_color_theme(self, theme_name):
        themes = {
            "Pink": "#FBDAE9",
            "Green": "#D3F8E2",
            "Blue": "#C6E9FB",
            "Yellow": "#F4F0CD",
            "Purple": "#EFDAFB"
        }
        bg = themes.get(theme_name, "white")
        self.root.configure(bg=bg)
        for child in self.root.winfo_children():
            try:
                child.configure(bg=bg)
            except:
                pass


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

            print("Preset loaded")

    def apply_preset_to_ui(self, spec):
        # Apply color theme
        if hasattr(spec, "color_theme"):
            self.apply_color_theme(spec.color_theme)

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
        for element in layout_data.get("elements", []):
            etype = element["type"]
            eid = element["id"]

            if etype == "textbox":

                entry = tk.Entry(self.root)
                entry.pack()
                self.widgets[eid] = entry

            elif etype == "button":
                if eid == "generate":
                    btn = tk.Button(self.root, text=element["label"], command=self.generate_labels)
                elif eid == "upload_file":
                    btn = tk.Button(self.root, text=element["label"], command=self.upload_sample_file)
                else:
                    btn = tk.Button(self.root, text=element["label"])
                btn.pack(padx=10, pady=15)
                self.widgets[eid] = btn

            elif etype == "textpreview":
                txt = tk.Text(self.root, height=10, width=50)
                txt.pack(padx=10, pady=15)
                self.widgets[eid] = txt

            elif etype == "label":
                lbl = tk.Label(self.root, text=element.get("text", ""))
                lbl.pack(padx=10, pady=15)
                self.widgets[eid] = lbl

    def build_ui_from_preset(self, spec):
        self.clear_ui()

        # Label at top: "Cryo Dots – 1.28 x 0.50 (Serial)"
        label_name = label_templates.get(spec.labeltemplate, {}).get("display_name", spec.labeltemplate)
        logic_label = getattr(spec, "identical_or_incremental", "Unknown").capitalize()
        self.template_label = tk.Label(self.root, text=f"{label_name} ({logic_label})", font=("Arial", 12, "bold"))
        self.template_label.pack(pady=10)

        if spec.presettype == "Text":
            # Build dynamic text input box
            template = label_templates.get(spec.labeltemplate, {})
            width_in = template.get("label_width_in", 1.0)
            height_in = template.get("label_height_in", 0.5)
            font_size = int(getattr(spec, "fontsize", 10))

            chars = int(width_in / (font_size * 0.07))
            lines = int(height_in / (font_size * 0.17))

            self.text_entry = tk.Text(self.root, width=max(40, chars), height=max(4, lines), font=(spec.fontname, font_size))
            self.text_entry.pack(padx=10, pady=5)

            # Optional row selectors if partial sheet is enabled
            if getattr(spec, "partialsheet", False):
                self.row_start_var = tk.StringVar(value="0")
                self.row_end_var = tk.StringVar(value="10")
                tk.Label(self.root, text="Start Row").pack()
                tk.Entry(self.root, textvariable=self.row_start_var).pack()
                tk.Label(self.root, text="End Row").pack()
                tk.Entry(self.root, textvariable=self.row_end_var).pack()

            # Save file button
            self.save_button = tk.Button(self.root, text="Save Labels", command=self.save_labels)
            self.save_button.pack(pady=10)

        elif spec.presettype == "File":
            # TODO: build File Input UI next
            pass

    def generate_labels(self):
        spec = self.current_spec

        if not spec:
            messagebox.showerror("Error", "No preset loaded.")
            return

        output_path = filedialog.asksaveasfilename(
            defaultextension=".docx" if spec.outputformat.lower() == "docx" else ".pdf",
            filetypes=[("Word Document", "*.docx"), ("PDF Document", "*.pdf")]
        )
        if not output_path:
            return

        try:
            if spec.presettype == "Text":
                text = self.widgets["user_input"].get("1.0", "end").strip()
                if spec.identical_or_incremental.lower() == "serial" and not text.isnumeric():
                    messagebox.showerror("Error", "Serial format must be a number.")
                    return
                main(spec, text_box_input=text, output_file_path=output_path)

            elif spec.presettype == "File":
                if not hasattr(self, "input_file_path") or not self.input_file_path:
                    messagebox.showerror("Error", "Please upload a CSV or XLSX file.")
                    return
                main(spec, input_file_path=self.input_file_path, output_file_path=output_path)

            messagebox.showinfo("Success", f"Labels saved to:\n{output_path}")
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

                if data_list and len(data_list) > 0:
                    preview = data_list[0]
                else:
                    preview = "No data found or invalid format."

                self.widgets["preview_area"].delete("1.0", "end")
                self.widgets["preview_area"].insert("1.0", preview)

            except Exception as e:
                messagebox.showwarning("Warning", f"Preview failed:\n{e}")


    def clear_ui(self):
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

        lb = tk.Listbox(win)
        lb.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        for name in self.presets:
            lb.insert(tk.END, name)

        def edit_selected():
            selected = lb.curselection()
            if not selected:
                return
            name = lb.get(selected[0])
            path, data = self.presets[name]
            PresetEditor(self.root, preset_data=data, preset_path=path, on_save=self.on_preset_saved)
            win.destroy()

        def delete_selected():
            selected = lb.curselection()
            if not selected:
                return
            name = lb.get(selected[0])
            path, _ = self.presets[name]
            os.remove(path)
            messagebox.showinfo("Deleted", f"Deleted preset: {name}")
            win.destroy()
            self.load_all_presets()

        btn_frame = tk.Frame(win)
        btn_frame.pack(pady=5)
        tk.Button(btn_frame, text="Edit", command=edit_selected).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="Delete", command=delete_selected).pack(side=tk.LEFT, padx=10)

    def on_preset_saved(self):
        self.load_all_presets()

    def run_label_maker(self):
        if self.current_spec:
            try:
                main.main(self.current_spec)
                messagebox.showinfo("Success", "Labels generated successfully.")
            except Exception as e:
                messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = CryoPopLabelStudioLite(root)
    root.mainloop()
