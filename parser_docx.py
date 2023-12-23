from docx import Document


class Parser():
    def __init__(self):
        self.document = None
    def get_table(self, index):
        return self.document.tables[index]
    def get_value(self,property, index_table, index_row = 0, index_cell = -1, idex_property = 0):
        _index_row = 0
        for row in self.get_table(index_table).rows:
            if row.cells[idex_property].text == property:
                if _index_row == index_row:
                    return row.cells[index_cell].text
                _index_row += 1
    def set_path(self, path):
        if path == "get_name":
            return
        self.document = Document(path)