from label_templates import label_templates

def get_textbox_dimensions(label_template_key, font_size_str):
    template = label_templates.get(label_template_key, {})
    chars_per_line = template.get("chars_per_line", 20)
    lines_per_label = template.get("lines_per_label", 3)

    try:
        font_size = float(font_size_str)
    except ValueError:
        font_size = 6

    return chars_per_line, lines_per_label
