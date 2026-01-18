import API


if __name__ == '__main__':
    #попросить компас вернуть текущий документ
    kompasAPI = API.API()
    kompasAPI.scan(r"D:\Projects\Checker\Kompas_files\Сборка.a3d")