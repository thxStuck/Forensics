import os
from oletools.olevba import VBA_Parser


file_name = 'файл.doc'

if os.path.exists(file_name):
    vba_parser = VBA_Parser(file_name)

    if vba_parser.detect_macros():
        print("Макросы обнаружены.")


        macros = vba_parser.extract_macros()


        for macro in macros:

            print(f'Имя файла: {macro[0]}')
            print(f'Путь к макросу: {macro[1]}')
            print(f'Имя модуля: {macro[2]}')
            print('Код VBA:')
            print(macro[3])  # Код VBA
            print('-' * 40)
    else:
        print("Макросы не обнаружены.")


    vba_parser.close()
else:
    print(f'Файл {file_name} не найден.')