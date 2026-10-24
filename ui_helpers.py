"""
Pure UI-logic helpers for the main window: no Tkinter widgets are created or
touched here, which makes this code testable without a display.
"""

THEME_COLORS = {
    "Pink": "#FBDAE9",
    "Green": "#D3F8E2",
    "Blue": "#C6E9FB",
    "Yellow": "#F4F0CD",
    "Purple": "#EFDAFB",
    "Grey": "#F0F0F0",
}


def resolve_theme_color(theme_name):
    """
    Returns the background color for a theme name. Falls back to the literal
    string "Grey" (not the "#F0F0F0" hex value) for an unrecognized name, to
    match the original behavior; Tk resolves "Grey" as a named X11 color.
    """
    return THEME_COLORS.get(theme_name, "Grey")


def build_footer_text(input_type=None, logic_type=None, sample_file=None, copies=None):
    """
    Builds the footer bar text summarizing the current preset: its logic type
    or sample file, and the number of copies per label.
    """
    footer_parts = []
    if copies is not None:
        strcopies = str(copies)
        if "," in strcopies or "-" in strcopies:
            footer_copies = None
        elif copies == "":
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

    return "  •  ".join(footer_parts)


def replace_copies_in_footer(existing_text, new_copies):
    """
    Returns existing_text with its "Copies: ..." segment replaced (or added)
    to reflect new_copies, leaving the other segments untouched.
    """
    parts = existing_text.split("  •  ")
    parts = [p for p in parts if not p.strip().startswith("Copies:")]
    parts.append(f"Copies: {new_copies}")
    return "  •  ".join(parts)
