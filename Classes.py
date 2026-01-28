import os
from typing import Any, Self


class File:
    def __init__(self, kompas: str | Any):
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
    def __init__(self, kompas):
        super().__init__(kompas)
        self.type = 1


class Part(File):
    def __init__(self, kompas):
        super().__init__(kompas)
        self.type = 2
        self.drawing = False # чертеж присутствует - True; чертеж отсутствует - False
        self.drawing_number = ''
        self.drawing_name = ''


class Assemble(Part):
    def __init__(self, kompas):
        super().__init__(kompas)
        self.type = 3
        self.bill_of_material = False # спецификация присутствует - True; спецификация отсутствует - False


class Drawing(File):
    def __init__(self, kompas):
        super().__init__(kompas)
        self.type = 4


class Files:
    def __init__(self):
        self.all_objects = [] # все объекты
        self.a3d_list = [] # документы сборки
        self.m3d_list = []  # документы детали
        self.cdw_list = [] # документы чертежей
        self.spw_list = []  # документы спецификаций
        self.pdf_list = [] # pdf документы

        self._id_count = 0

    def add_file(self, file: str): #TODO дописать для заполнения всех элементов
        """Принимает строку со ссылкой на файл"""
        self._id_count += 1
        if os.path.splitext(file)[1] == '.a3d': file_object = Assemble(file)
        elif os.path.splitext(file)[1] == '.m3d': file_object = Part(file)
        elif os.path.splitext(file)[1] == '.cdw': file_object = Drawing(file)
        else: file_object = File(file)
        file_object.id = self._id_count

        self.all_objects.append(file_object)
        if file_object.type == 2: self.m3d_list.append(file_object)
        elif file_object.type == 3: self.a3d_list.append(file_object)
        elif file_object.type == 4: self.cdw_list.append(file_object)

    def find_missing_drawing(self): #TODO Переписать под изменения
        list_to_check = []

        for i in self.cdw_list:
            list_to_check.append(os.path.splitext(i.path)[0])
        print(self.cdw_list)
        print(list_to_check)
        for i in self.a3d_list:
            print(os.path.splitext(i.path)[0])
            if os.path.splitext(i.path)[0] in self.cdw_list:
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
            if type(i) is Part: print('   drawing: ', i.drawing)
            if type(i) is Assemble: print('   drawing: ', i.drawing,'   BOM: ',i.bill_of_material)