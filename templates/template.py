import xml.etree.ElementTree as ET
from uuid import uuid4
from xml.dom import minidom
import datetime
from parser_docx import Parser
import re
import os
import logging


class Template():
    def __init__(self, path):
        self.name_template = "Пример работы шаблона"
        '''
        Создает основную часть документа
        '''
        self.path_to_source = path

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
        return f'{path_folder}{os.sep}{self.name}_{self.uid}.xml'

    def __from_string(self, text: str):
        try:
            xml_string = re.sub('>[\s\n\r]+<', '><', text)
            return ET.fromstring(xml_string)
        except Exception as ex:
            print(f"Ошибка в XML:\n{text}")
            logging.exception(ex)
            raise ex

    def __add_attribute(self, path, property, value):
        element = self.__find_element(path)
        element.attrib.update({f"{property}": f"{value}"})

    def __find_element(self, path):
        return self.root.find(path)

    def __get_patronymic(self, fio_parts):
        if isinstance(fio_parts, list) == False:
            fio_parts = f"{fio_parts}".split(" ")
        patronymic = f' Отчество="{fio_parts[-1].upper()}"' if len(fio_parts) == 3 else ''
        return patronymic

    def __get_fio(self, fio):
        fio_parts = fio.split(" ")
        return f'Фамилия="{fio_parts[0].upper()}" Имя="{fio_parts[1].upper()}" {self.__get_patronymic(fio_parts)}'

    def __get_passport_number(self, number):
        numbers_only = number.replace(" ", "")
        formatted = f"{numbers_only[:2]} {numbers_only[2:4]} {numbers_only[4:]}"
        return formatted

    def custom(self):
        '''
        данная часть модифицируется
        :return:
        '''
