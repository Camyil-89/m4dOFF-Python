from docx import Document


class Parser():
    def __init__(self):
        self.document = None
    def get_table(self, index):
        return self.document.tables[index]
    def get_value(self,property, index_table):
        for row in self.get_table(index_table).rows:
            if row.cells[0].text == property:
                return row.cells[-1].text
    def set_path(self, path):
        if path == "get_name":
            return
        self.document = Document(path)