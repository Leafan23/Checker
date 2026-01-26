

class Files:
    def __init__(self):
        self.all_objects = [] # все объекты
        self.a3d_list = [] # документы сборки
        self.m3d_list = []  # документы детали
        self.cdw_list =[] # документы чертежей
        self.spw_list = []  # документы спецификаций
        self.pdf_list = [] # pdf документы

    def add_file(self, file):
        self.all_objects.extend(file)

class File:
    def __init__(self, kompas, id_number):
        self.kompas = kompas
        self.id = id_number
        self.path = kompas.path
        self.type = 0


class Pdf(File):
    def __init__(self, kompas, id_number):
        super().__init__(kompas, id_number)
        self.type = 1


class Part(File):
    def __init__(self, kompas, id_number):
        super().__init__(kompas, id_number)
        self.type = 2
        self.parent = None
        self.drawing = False # чертеж присутствует - True; чертеж отсутствует - False
        self.drawing_number = kompas.drawing_number
        self.drawing_name = kompas.drawing_name



class Assemble(Part):
    def __init__(self, kompas, id_number):
        super().__init__(kompas, id_number)
        self.type = 3
        self.bill_of_material = False # спецификация присутствует - True; спецификация отсутствует - False
