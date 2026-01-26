#TODO (in progress) Функция проверки проекта на комплектность документов и исправления привязок
#TODO Функция комплектовки документов определенной сборки
#TODO Функция печати pdf
#TODO Функция открытия файлов по клику
#TODO Функция копирования проекта с новым именем проекта и заменой всех внутренних ссылок
#TODO Функция выявления лишних файлов, чертежей, деталей, сборок

import API


if __name__ == '__main__':
    #попросить компас вернуть текущий документ
    kompasAPI = API.API()
    kompasAPI.scan(r"C:\Users\Leafan\PycharmProjects\Checker\kompas_files\ГКЮШ.ТЕСТ.00.000.a3d")
    #kompasAPI.open(r"D:\Projects\Checker\Kompas_files\ГКЮШ.ТЕСТ.00.000.a3d")
    #print(kompasAPI.main_tree[0].path)
    '''for i in kompasAPI.main_tree:
        print('Drawing: ', i.drawing, i.path)
        try:
            print('BOM: ', i.bill_of_material, i.path)
        except:
            pass
        print('-----------------------')'''

    kompasAPI.files.print_all_data()
    #TODO полное закрытие приложения, если работало в скрытом режиме