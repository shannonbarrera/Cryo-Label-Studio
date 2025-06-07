import tkinter as tk

HELP_TEXT = (
    "📘 Welcome to CryoPop Label Studio Lite!\n\n"
    "🔹 PRESETS\n"
    "- Use 'New File Input Preset' to pull label data from a CSV or Excel file.\n"
    "- Use 'New Text Input Preset' to enter label text manually (Identical or Serial).\n"
    "- 'Edit Presets' lets you rename, update, or delete saved presets.\n\n"
    "🔹 GENERATING LABELS\n"
    "- Load a preset from the dropdown.\n"
    "- Upload your file or enter your text.\n"
    "- Click 'Generate Labels' to export a Word or PDF file.\n\n"
    "🔹 FORMATTING\n"
    "- Use {FIELD} to pull data from column headers.\n"
    "- Add newlines with \\n or customize font, size, and layout in the preset.\n\n"
    "🔹 DATE FORMATTING\n"
    "- Select a format (e.g., MM/DD/YY) in the preset editor.\n"
    "- Or choose 'Leave as is' to preserve original Excel formatting.\n\n"
    "🔹 TROUBLESHOOTING\n"
    "- If labels aren't generating, check for missing fields or formatting typos.\n"
    "- Contact support at cryopopsoftware.com for help.\n"
)

def show_help_window(parent):
    help_win = tk.Toplevel(parent)
    help_win.title("Help")
    help_win.geometry("600x500")
    help_win.resizable(True, True)

    text_widget = tk.Text(help_win, wrap="word", padx=10, pady=10)
    text_widget.insert("1.0", HELP_TEXT)
    text_widget.config(state="disabled")
    text_widget.pack(expand=True, fill="both", padx=10, pady=10)

def show_about_window(parent):
    about_win = tk.Toplevel(parent)
    about_win.title("About")
    about_win.geometry("400x300")
    about_win.resizable(False, False)

    text = (
        "CryoPop Label Studio\n"
        "Version 1.0\n\n"
        "Created by Shannon Barrera\n"
        "Copyright 2025\n\n"
        "Website:\ncryopopsoftware.com\n\n"
        "This app was designed to make\n"
        "lab labeling faster, simpler, and\n"
        "a little more joyful.\n\n"
        "🧪✨"
    )

    label = tk.Label(about_win, text=text, justify="left", padx=20, pady=20, font=("Segoe UI", 10))
    label.pack(expand=True, fill="both")
