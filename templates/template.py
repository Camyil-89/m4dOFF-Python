import xml.etree.ElementTree as ET
from uuid import uuid4
from xml.dom import minidom
import datetime
from parser_docx import Parser
import re


class Template():
    def __init__(self, path):
        '''
        Создает основную часть документа
        '''
        self.path_to_source = path
        self.name_template = "None"

        self.parser = Parser()  # парсер .docx файлов
        self.parser.set_path(path)  # выбираем файл

        self.name = f"ON_EMCHD_{datetime.datetime.now().strftime('%Y%m%d')}"
        self.uid = f"{uuid4()}"
        attribute = {'xmlns:xsi': "http://www.w3.org/2001/XMLSchema-instance",
                     'xmlns:xsd': "http://www.w3.org/2001/XMLSchema",
                     'ВерсФорм': "EMCHD_1",
                     'ПрЭлФорм': "00000000",
                     'ИдФайл': f"{self.name}_{self.uid}",
                     'xmlns': "urn://x-artefacts/EMCHD_1",
                     }
        self.root = ET.Element('Доверенность', attrib=attribute)

    def save(self, path_folder):
        rough_string = ET.tostring(self.root, 'utf-8')
        reparsed = minidom.parseString(rough_string)
        xml_text = reparsed.toprettyxml(indent="    ")
        with open(f'{path_folder}{os.sep}{self.name}_{self.uid}.xml', "w", encoding="utf-8") as file:
            file.write(xml_text)
        #with open(f'test.xml', "w", encoding="utf-8") as file:

        return f'{path_folder}{os.sep}{self.name}_{self.uid}.xml'

    def __from_string(self, text: str):
        xml_string = re.sub('>[\s
]+<', '><', text)
        return ET.fromstring(xml_string)
    def __add_attribute(self, path, property, value):
        element = self.__find_element(path)
        element.attrib.update({f"{property}": f"{value}"})
    def __find_element(self, path):
        return self.root.find(path)


    def custom(self):
        '''
        данная часть модифицируется
        :return:
        '''
        pass
