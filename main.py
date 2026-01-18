import GUI


if __name__ == '__main__':
    #попросить компас вернуть текущий документ
    window = GUI.MyCtk()
    window.mainloop()
    for i in window.main_tree:
        print(i.type)
        print(i.path)
        print(i.id)
        print(i.drawing_number)
    print(window.main_tree)