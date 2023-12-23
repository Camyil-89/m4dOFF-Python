import os
from executing import Executing
import gui
import argparse
import sys
import logging

def create_folders():
    try:
        os.mkdir("source")
    except:
        pass

    try:
        os.mkdir("output")
    except:
        pass

def main():
    create_folders()
    logging.basicConfig(level=logging.DEBUG,
                        format='%(asctime)s %(pathname)s:%(lineno)d %(levelname)s [Function: %(funcName)s] %(message)s',
                        datefmt='%d-%b-%y %H:%M:%S')


def start_cmd(path_to_source, path_to_output, name_template):
    executing = Executing()
    files = [f for f in os.listdir(f"{os.getcwd()}{os.sep}templates") if os.path.isfile(os.path.join(f"{os.getcwd()}{os.sep}templates", f))]
    for temp in files:
        if temp == name_template:
            for i in executing.load_template(f"{os.getcwd()}{os.sep}templates{os.sep}{temp}", [f"{path_to_source}{os.sep}{f}" for f in os.listdir(path_to_source) if os.path.isfile(os.path.join(path_to_source, f))], path_to_output):
                print(f"МЧД создана по пути: {i}")

parser = argparse.ArgumentParser(description="m4dOFF облегчает создание большого количества МЧД, путем создания шаблонов.")

main()
if len(sys.argv) > 1:
    parser.add_argument('--source', '-s', help='Путь к папке с .docx файлами из которых необходимо создать МЧД.', required=True)
    parser.add_argument('--output', '-o', help='Директория в которую будут создаваться МЧД.', required=True)
    parser.add_argument('--template', '-t', help='Имя файла (только имя без полного пути) шаблона.', required=True)
    args = parser.parse_args()

    start_cmd(args.source, args.output, args.template)
    print("\nСоздание МЧД завершено!")
else:
    gui.start_gui()

