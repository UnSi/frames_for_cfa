import os
import sys

import xlrd


def lot_sep_print(something='', sep='=', number=25):
    print(f'{sep * number}\n{something}')


# TODO: make ini-file with path
def ask_excel_list():  # TODO: rename to request
    frames = ''
    print('Введите список кадров. Для окончания ввода введите end')
    while True:
        # frames.append(str(input().split(' ')))
        frame = input()
        if frame == 'end':
            break
        elif frame.lower() == 'return':
            print('Выхожу')
            return ['return'] # TODO: найти, как сделать return_to_main_menu
        try:
            for item in frame.split():
                int(item)
        except Exception:  # TODO: clarify the exception
            print("\nЗначения должны быть числовые. Попробуйте снова")
            print(f"'{frame}' не учтено в списке. Если часть элемента должна быть в нём - доавьте его заново")
            frame = ''
        # else:
        #     print("Значения должны быть числовые. Попробуйте снова")
        #     frame = ''
        frames += frame + ' '
    result = frames.split()
    print('Список прочитан')
    # print(result)
    return result


def check_different_frames(result=[]):  # TODO: remove input param
    if not result:
        result = get_excel_list()
    if result == ['return']:
        return # TODO: найти, как сделать return_to_main_menu
    print("Проверяю кадры с несколькими объектами. Будут выведены номера кадров, которые использованы несколько раз")
    # print(result)
    lot_sep_print()
    print(f"введено кадров: {len(result)}")
    print(f"уникальных кадров: {len(set(result))}")
    for item in set(result):
        result.remove(item)
    print(f"Совпадающие кадры: {result}")
    lot_sep_print()


def find_img_dir():
    try:
        # TODO: save way before changing OR:
        # TODO: Return path instead changing dir and leave try except
        new_path = input("Введите путь к папке с кадрами: \n")
        if new_path.split(os.path.sep)[-1] != 'img':
            print("Путь не к папке img. Ищу img ")
            new_path = os.path.join(new_path, 'img')
            if not os.path.isdir(new_path):
                raise WindowsError('img не найдено')
        os.chdir(new_path)
        lot_sep_print("Нашел")
    except WindowsError as exc:
        print(exc)
        print("Ищу папку img из пути запуска скрипта")
        print("Где же она?")  # Здесь была Лиза
        if os.getcwd().split(os.path.sep)[-1] == 'img':
            lot_sep_print(f"Найдена. Адрес: {os.getcwd()}")
        elif os.path.isdir('img'):
            os.chdir('img')
            lot_sep_print(f"Найдена. Адрес: {os.getcwd()}")
        else:
            print(f'Не найдена папка img, закрываю. Адрес: {os.getcwd()}')
            sys.exit()


def find_excel_file():  # TODO: leave try except, add  "0.Размеченные кадры.xlsx" to ini
    try:
        path = input("Введите путь к xlsx-файлу: \n")
        if path.split('.')[-1] != 'xlsx':
            temp_path = os.path.join(path, input("Введите название файла: \n"))
            if not os.path.isfile(temp_path):
                path = os.path.join(path, "0.Размеченные кадры.xlsx")
                if not os.path.isfile(path):
                    raise FileNotFoundError("Файл не найден")
            else:
                path = temp_path
        xlrd.open_workbook(path)
        return path
    except FileNotFoundError as exc:
        print(exc)
        print("Ищу файл '0.Размеченные кадры.xlsx' из пути запуска скрипта")
        path = os.path.join(os.getcwd(), "0.Размеченные кадры.xlsx")
        if os.path.isfile(path):
            print(f"Найден. Адрес: {path}")
            xlrd.open_workbook(path)
            return path
        else:
            print(f'Не найден файл "0.Размеченные кадры.xlsx", необходимо ввести список вручную:')
            # print(os.getcwd())
            # print(new_path)
            return ''


def get_dir_frames_list():
    find_img_dir()
    result = []
    for item in os.listdir():
        item_list = str(item).split('.')
        if item_list[-1] == 'jpg':
            result.append(item_list[0].split('_')[-1])
    os.chdir('..')
    # print(os.getcwd())
    return result


def compare_results(excel_set=[]):
    excel_set = set(excel_set if excel_set else get_excel_list())
    if excel_set == {'return'}:
        return  # TODO: найти, как сделать return_to_main_menu
    dir_set = set(get_dir_frames_list())
    print("Кадры, которые есть в екселе, но нет в папке img:", (excel_set - dir_set) or "Всё сошлось")
    print("Кадры, которые есть в папке img, но нет в excel:", (dir_set - excel_set) or "Всё сошлось")
    lot_sep_print()


def get_excel_list():
    path = find_excel_file()
    if path:
        excel_file = xlrd.open_workbook(path)
        sheet = excel_file.sheet_by_index(0)
        frames_list = []
        if sheet.nrows > 0:
            for col in range(sheet.ncols):
                for row in range(1, sheet.nrows):
                    cell = sheet.row(row)[col]
                    if cell.value == '':
                        break
                    frames_list.append(str(int(cell.value)))  # TODO: try except
            # print(batch_names_wanted)
            return frames_list
        else:
            print('пустой файл')
            frames_list = ask_excel_list()
            return frames_list
    else:
        frames_list = ask_excel_list()
        return frames_list


if __name__ == "__main__":
    # check_different_frames()
    # compare_results(some_list)
    command = ''
    while True:
        try:
            if not command:
                command = sys.argv[1]
        except IndexError:
            print('Не введено аргументов.')
        if command.lower().find('compare') != -1:
            compare_results()
            command = ''
        elif command.lower().find('df') != -1:
            check_different_frames()
            command = ''
        elif command.lower().find('exit') != -1:
            break
        else:
            print('Запуск с параметром compare - сравнит файлы из папки img с введенными в программу номерами кадров')
            print('Запуск с параметром df - проверит наличие дублей в excel файле ')
            print('exit - выход ')
            print('Введите параметр:')
            command = input()


#  TODO: на кнопки