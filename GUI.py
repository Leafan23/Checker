import customtkinter as ctk
import API
from Classes import File, Part, Pdf, Assemble
import os


class MyCtk(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title('Checker')
        self.main_tree: list[File] = []

        self.kompas_start()
        self.add_to_main_tree(r"C:\Users\Leafan\Desktop\Сборка.a3d")

    def kompas_start(self):
        self.kompas = API.API()

    def add_to_main_tree(self, path):
        self.kompas.open(path)
        if os.path.splitext(path)[1] == '.m3d':
            self.main_tree.append(Part(self.kompas, id_number=len(self.main_tree)))
        elif os.path.splitext(path)[1] == '.a3d':
            self.main_tree.append(Assemble(self.kompas, id_number=len(self.main_tree)))

    def scan(self):
        # открыть сборку
        # добавить в список
        # если есть детали, открыть их поочередно и добавить в список с добавлением родителя изначальной сборки, но только если они небыли раньше обработаны
        # закрыть детали после добавления
        # посчитать количество подсборок (например 3)
        # добавить подсборки в таблицу и занести их адрес в список на проверку
        # закрыть основную сборку
        # открыть сборку из списка и удалить адрес из этого же списка
        # повторять пока список не будет пуст
        pass