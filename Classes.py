import os
from typing import Any, Self


class File:
    def __init__(self, kompas: str | Any, id=0):
        self.kompas = kompas
        self.id = None
        if isinstance(kompas, str):
            self.path = kompas
        else:
            self.path = kompas.path
        self.type = 0
        self.child = []
        self.parent = None

    def add_child(self, child: list[int] | int) -> None:
        if isinstance(child, list):
            self.child.extend(child)
        else:
            self.child.append(child)
        #TODO дописать удаление повторений


class Pdf(File):
    def __init__(self, kompas, id_number):
        super().__init__(kompas, id_number)
        self.type = 1


class Part(File):
    def __init__(self, kompas, id_number):
        super().__init__(kompas, id_number)
        self.type = 2
        self.drawing = False # чертеж присутствует - True; чертеж отсутствует - False
        self.drawing_number = kompas.drawing_number
        self.drawing_name = kompas.drawing_name


class Assemble(Part):
    def __init__(self, kompas, id_number):
        super().__init__(kompas, id_number)
        self.type = 3
        self.bill_of_material = False # спецификация присутствует - True; спецификация отсутствует - False


class Files:
    def __init__(self):
        self.all_objects = [] # все объекты
        self.a3d_list = [] # документы сборки
        self.m3d_list = []  # документы детали
        self.cdw_list = [] # документы чертежей
        self.spw_list = []  # документы спецификаций
        self.pdf_list = [] # pdf документы

        self._id_count = 0

    def add_file(self, file: str):
        """Принимает строку со ссылкой на файл"""
        self._id_count += 1
        file_object = File(file)
        file_object.id = self._id_count

        self.all_objects.append(file_object)
        if file_object.type == 2:
            self.m3d_list.append(file_object)
        if file_object.type == 3:
            self.a3d_list.append(file_object)

    def find_missing_drawing(self):

        for i in self.a3d_list:

            if i.path in self.cdw_list:
                i.drawing = True

    def id_return(self, id: int):
        for i in self.all_objects:
            if i.id == id: return i
        return None

    def last_added(self):
        return self.all_objects[-1]

    def print_all_data(self):
        for i in self.all_objects:
            print('id: ', i.id, '   path: ',i.path, '   parent: ',i.parent, '   child: ',i.child)