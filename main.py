import API


if __name__ == '__main__':
    #попросить компас вернуть текущий документ
    kompasAPI = API.API()
    kompasAPI.scan(r"C:\Users\Leafan\Desktop\Сборка.a3d")
    for i in kompasAPI.main_tree:
        print(i.type)
        print(i.path)
        print(i.id)
        print(i.drawing_number)
    print(kompasAPI.main_tree)