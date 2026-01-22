

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
