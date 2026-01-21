import API


if __name__ == '__main__':
    #попросить компас вернуть текущий документ
    kompasAPI = API.API()
    #kompasAPI.scan(r'C:\Users\Leafan\PycharmProjects\Checker\kompas_files\Сборка2.a3d')
    kompasAPI.open(r'C:\Users\Leafan\PycharmProjects\Checker\kompas_files\Деталь2.m3d')