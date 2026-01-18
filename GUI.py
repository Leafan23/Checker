import customtkinter as ctk
import API
from Classes import File


class MyCtk(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title('Checker')
        self.main_tree: list[File] = []

        self.kompas_start()
        self.add_to_main_tree()

    def kompas_start(self):
        self.kompas = API.API()

    def add_to_main_tree(self):
        self.kompas.open(r"C:\Users\Leafan\Desktop\Деталь.m3d")
        self.main_tree.append(File())
        pass