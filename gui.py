
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from label_spec import LabelSpec
import main
import json
import os

from preset_editor import PresetEditor

class CryoPopLabelStudioLite:
    def __init__(self, root):
        self.root = root
        self.root.title("CryoPop Label Studio Lite")
        self.root.geometry("600x400")

        self.current_spec = None
        self.presets_dir = "presets"
        os.makedirs(self.presets_dir, exist_ok=True)

        self.setup_menu()
        self.setup_main_ui()
        self.load_all_presets()

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

        self.run_button = tk.Button(self.root, text="Generate Labels", command=self.run_label_maker, state=tk.DISABLED)
        self.run_button.pack(pady=10)

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
            self.status_label.config(text=f"Loaded preset: {name}")
            self.run_button.config(state=tk.NORMAL)

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
