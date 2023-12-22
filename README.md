# m4dOFF
> Приложение для быстрого создания большого количества МЧД.

## Свой шаблон
Для создания шаблона необходимо:
- Создать файл [имя файла].py
- Скопировать содержимое template.py
- Написать свою логику обработки таблиц из .docx файла.

Обработка таблиц пишется просто.
> __from_string - позволяет вставлять "куски" МЧД и форматировать их не создавая сложную логику.
```python
 def __from_string(self, text: str):
    xml_string = re.sub('>[\s\n\r]+<', '><', text)
    return ET.fromstring(xml_string)
```
> self.uid - уникальный идентификатор МЧД
```python
self.uid
```
> в функции custom пишется логика обработки .docx файла.
```python
def custom(self):
    '''
    данная часть модифицируется
    :return:
    '''
    document = ET.SubElement(self.root, 'Документ') # создаем xml элемент
    trust = ET.SubElement(document, 'Довер') # создаем xml элемент
    trust.append(self.__from_string(f'''<СвДов ВидДовер="1" ПрПередов="1" ВнНомДовер="{self.parser.get_value("Внутренний номер", 0)}" НомДовер="{self.uid}" ДатаВыдДовер="{self.parser.get_value("Дата выдачи", 0)}" СрокДейст="{self.parser.get_value("Срок действия", 0)}">
        <СведСист>https://m4d.nalog.gov.ru/EMCHD/check-status?guid={self.uid}
        </СведСист>
    </СвДов>'''))
```
> self.parser.get_value("[имя первого столбца в таблице]", [номер таблицы в .docx файле (отсчет начинается с 0)])
```python
self.parser.get_value("Дата выдачи", 0)
```
