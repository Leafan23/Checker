import API


if __name__ == '__main__':
    #попросить компас вернуть текущий документ
    kompasAPI = API.API()
    #kompasAPI.scan(r'C:\Users\Leafan\PycharmProjects\Checker\kompas_files\Сборка2.a3d')
    kompasAPI.open(r'D:\Projects\Checker\Kompas_files\ГКЮШ.ТЕСТ.00.001.m3d')
    print(kompasAPI.main_tree[-1].drawing)