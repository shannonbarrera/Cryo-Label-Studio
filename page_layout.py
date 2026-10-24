"""
PageLayout: the page geometry for a label template - which table rows/columns
hold labels, and (for a partial sheet) which of those are available on the
first page - plus the logic to paginate a list of label data across pages and
build the resulting Word document from it.

This bundles up the row/column-index bookkeeping that main.py previously had
to thread through every call by hand.
"""

from label_format import (
    get_row_and_column_indices,
    get_first_page_row_indices,
    get_first_page_col_indices,
    get_max_labels_first_page,
    paginate_labels,
    format_labels_page,
    combine_docs,
)
from file_io import resource_path


class PageLayout:
    def __init__(self, spec, template_meta):
        """
        Args:
            spec (LabelSpec): only partialsheet/row_start/row_end/col_start/
                col_end are read here.
            template_meta (dict): the label_templates[...] entry matching
                spec.labeltemplate.
        """
        self.template_meta = template_meta
        self.templatepath = resource_path(template_meta["template_path"])
        self.needs_page_break = template_meta["needs_page_break"]

        self.row_indices, self.column_indices = get_row_and_column_indices(
            self.templatepath, template_meta["table_format"]
        )
        self.max_labels_per_page = len(self.row_indices) * len(self.column_indices)

        if spec.partialsheet:
            start_row = getattr(spec, "row_start", 1)
            end_row = getattr(spec, "row_end", template_meta.get("labels_down", 99))
            start_col = getattr(spec, "col_start", 1)
            end_col = getattr(spec, "col_end", template_meta.get("labels_across", 99))

            self.first_page_row_indices = get_first_page_row_indices(
                start_row, end_row, self.row_indices
            )
            (
                self.first_page_first_row_col_indices,
                self.first_page_last_row_col_indices,
            ) = get_first_page_col_indices(
                start_col, end_col, start_row, end_row, self.column_indices
            )
        else:
            self.first_page_row_indices = self.row_indices
            self.first_page_first_row_col_indices = self.column_indices
            self.first_page_last_row_col_indices = self.column_indices

        self.first_page_max_labels = get_max_labels_first_page(
            self.first_page_row_indices,
            self.column_indices,
            self.first_page_first_row_col_indices,
            self.first_page_last_row_col_indices,
        )

    def build_document(self, data_list, copiesperlabel, spec):
        """
        Paginates data_list across this layout's pages, formats each page
        from the template, and returns the combined Document.
        """
        first_page, other_pages = paginate_labels(
            self.first_page_max_labels, self.max_labels_per_page, data_list, copiesperlabel
        )
        pages = [first_page] + other_pages

        final_doc = None
        for i, page in enumerate(pages):
            is_last = i == len(pages) - 1
            formatted_page = format_labels_page(
                page,
                self.templatepath,
                self.first_page_row_indices if i == 0 else self.row_indices,
                self.column_indices,
                self.first_page_first_row_col_indices if i == 0 else self.column_indices,
                self.first_page_last_row_col_indices if i == 0 else self.column_indices,
                spec,
                self.needs_page_break,
                is_last_page=is_last,
            )

            if final_doc is None:
                final_doc = formatted_page
            else:
                final_doc = combine_docs(final_doc, formatted_page)

        return final_doc
